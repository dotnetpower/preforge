"""
HTML(.html) 문서를 PowerPoint(.pptx)로 변환하는 컨버터

HTML 구조를 분석하여 섹션별로 슬라이드를 생성합니다.
테이블, 텍스트, 리스트, 이미지 등을 포함한 다양한 요소를 지원합니다.
"""
from pathlib import Path
from typing import List, Optional, Dict, Any, Tuple
from io import BytesIO
import logging
import re
import base64
from urllib.parse import urlparse

from bs4 import BeautifulSoup, Tag, NavigableString
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR, MSO_AUTO_SIZE
from pptx.enum.shapes import MSO_SHAPE
from pptx.dml.color import RGBColor
from pptx.oxml.xmlchemy import OxmlElement

try:
    from PIL import Image
    HAS_PIL = True
except ImportError:
    HAS_PIL = False

logger = logging.getLogger(__name__)


class HtmlToPptxConverter:
    """HTML을 PowerPoint로 변환하는 컨버터"""
    
    def __init__(self):
        """컨버터 초기화"""
        self.prs = None
        self.current_slide = None
        self.html_path = None  # 스크린샷용 HTML 경로
        
        # 테이블 분할 설정
        self.max_rows_per_slide = 8  # 슬라이드당 최대 행 수 (헤더 제외) - 가독성 향상
        
        # 색상 정의 (HTML CSS에서 추출)
        self.colors = {
            'primary_red': RGBColor(220, 38, 38),  # #dc2626
            'primary_red_light': RGBColor(254, 242, 242),  # #fef2f2
            'primary_red_dark': RGBColor(153, 27, 27),  # #991b1b
            'gray_50': RGBColor(249, 250, 251),
            'gray_100': RGBColor(243, 244, 246),
            'gray_200': RGBColor(229, 231, 235),
            'gray_600': RGBColor(75, 85, 99),
            'gray_800': RGBColor(31, 41, 55),
            'gray_900': RGBColor(17, 24, 39),
            'white': RGBColor(255, 255, 255),
            'black': RGBColor(0, 0, 0),
        }
        
        # 슬라이드 크기 설정 (16:9 비율)
        self.slide_width = Inches(10)
        self.slide_height = Inches(7.5)
        
        # 여백 설정 (더 좁게 조정)
        self.margin_left = Inches(0.3)
        self.margin_right = Inches(0.3)
        self.margin_top = Inches(0.5)
        self.margin_bottom = Inches(0.3)
        
        self.content_width = self.slide_width - self.margin_left - self.margin_right
        self.content_height = self.slide_height - self.margin_top - self.margin_bottom
    
    def _extract_cell_styles(self, cell_elem: Tag) -> Dict[str, Any]:
        """HTML 셀에서 스타일(Bold, Color 등) 추출"""
        styles = {
            'bold': False,
            'color': None,
            'background': None,
            'link': None,
        }
        
        # 셀 자체의 스타일 확인
        style_attr = cell_elem.get('style', '')
        
        # color 추출
        color_match = re.search(r'color:\s*(#[a-fA-F0-9]{6}|#[a-fA-F0-9]{3}|rgb\([^)]+\))', style_attr)
        if color_match:
            color_str = color_match.group(1)
            styles['color'] = self._parse_color(color_str)
        
        # background-color 추출
        bg_match = re.search(r'background(?:-color)?:\s*(#[a-fA-F0-9]{6}|#[a-fA-F0-9]{3}|rgb\([^)]+\))', style_attr)
        if bg_match:
            styles['background'] = self._parse_color(bg_match.group(1))
        
        # font-weight: bold 또는 700 확인
        if 'font-weight' in style_attr:
            weight_match = re.search(r'font-weight:\s*(\w+)', style_attr)
            if weight_match:
                weight = weight_match.group(1)
                if weight in ('bold', '700', '800', '900'):
                    styles['bold'] = True
        
        # 내부의 bold 태그 확인 (b, strong)
        if cell_elem.find(['b', 'strong']):
            styles['bold'] = True
        
        # 내부의 색상 스타일 확인 (span 등)
        colored_elem = cell_elem.find(style=True)
        if colored_elem and not styles['color']:
            inner_style = colored_elem.get('style', '')
            inner_color = re.search(r'color:\s*(#[a-fA-F0-9]{6}|#[a-fA-F0-9]{3}|rgb\([^)]+\))', inner_style)
            if inner_color:
                styles['color'] = self._parse_color(inner_color.group(1))
        
        # 링크 확인
        link = cell_elem.find('a')
        if link:
            styles['link'] = link.get('href', '')
        
        return styles
    
    def _parse_color(self, color_str: str) -> Optional[RGBColor]:
        """색상 문자열을 RGBColor로 변환"""
        if not color_str:
            return None
        
        try:
            # hex color (#rrggbb or #rgb)
            if color_str.startswith('#'):
                hex_color = color_str[1:]
                if len(hex_color) == 3:
                    hex_color = ''.join([c*2 for c in hex_color])
                r = int(hex_color[0:2], 16)
                g = int(hex_color[2:4], 16)
                b = int(hex_color[4:6], 16)
                return RGBColor(r, g, b)
            
            # rgb(r, g, b)
            if color_str.startswith('rgb'):
                match = re.search(r'rgb\s*\(\s*(\d+)\s*,\s*(\d+)\s*,\s*(\d+)\s*\)', color_str)
                if match:
                    r, g, b = int(match.group(1)), int(match.group(2)), int(match.group(3))
                    return RGBColor(r, g, b)
        except:
            pass
        
        return None
    
    def convert(self, html_path: Path, output_path: Path) -> None:
        """
        HTML 파일을 PPTX로 변환
        
        Args:
            html_path: 입력 HTML 파일 경로
            output_path: 출력 PPTX 파일 경로
        """
        logger.info(f"HTML -> PPTX 변환 시작: {html_path} -> {output_path}")
        
        # HTML 파일 경로 저장 (스크린샷용)
        self.html_path = Path(html_path).absolute()
        
        # HTML 파일 읽기
        with open(html_path, 'r', encoding='utf-8') as f:
            html_content = f.read()
        
        soup = BeautifulSoup(html_content, 'lxml')
        
        # 프레젠테이션 초기화
        self.prs = Presentation()
        self.prs.slide_width = self.slide_width
        self.prs.slide_height = self.slide_height
        
        # 타이틀 슬라이드 생성
        self._create_title_slide(soup)
        
        # Analysis Summary 섹션 생성
        self._create_analysis_summary_slides(soup)
        
        # 메인 컨텐츠 처리
        self._process_main_content(soup)
        
        # PPTX 저장
        self.prs.save(str(output_path))
        logger.info(f"변환 완료: {output_path} (총 {len(self.prs.slides)}개 슬라이드)")
    
    def _create_title_slide(self, soup: BeautifulSoup) -> None:
        """타이틀 슬라이드 생성"""
        slide = self.prs.slides.add_slide(self.prs.slide_layouts[6])  # 빈 레이아웃
        
        # 제목 추출
        title_elem = soup.find('div', class_='header-title')
        subtitle_elem = soup.find('div', class_='header-subtitle')
        
        title_text = title_elem.get_text(strip=True) if title_elem else "GeneSeq Vista AI Agent"
        subtitle_text = subtitle_elem.get_text(strip=True) if subtitle_elem else ""
        
        # 배경 사각형
        background = slide.shapes.add_shape(
            MSO_SHAPE.RECTANGLE,
            0, 0,
            self.slide_width, self.slide_height
        )
        background.fill.solid()
        background.fill.fore_color.rgb = self.colors['gray_50']
        background.line.fill.background()
        
        # 상단 빨간 박스
        header_box = slide.shapes.add_shape(
            MSO_SHAPE.RECTANGLE,
            0, Inches(2),
            self.slide_width, Inches(1.5)
        )
        header_box.fill.solid()
        header_box.fill.fore_color.rgb = self.colors['primary_red']
        header_box.line.fill.background()
        
        # 타이틀 텍스트
        title_frame = header_box.text_frame
        title_frame.text = title_text
        title_frame.paragraphs[0].alignment = PP_ALIGN.CENTER
        title_frame.paragraphs[0].font.size = Pt(40)
        title_frame.paragraphs[0].font.bold = True
        title_frame.paragraphs[0].font.color.rgb = self.colors['white']
        
        # 부제목 텍스트
        if subtitle_text:
            subtitle_box = slide.shapes.add_textbox(
                Inches(1), Inches(4),
                self.slide_width - Inches(2), Inches(1.5)
            )
            subtitle_frame = subtitle_box.text_frame
            subtitle_frame.text = subtitle_text
            subtitle_frame.paragraphs[0].alignment = PP_ALIGN.CENTER
            subtitle_frame.paragraphs[0].font.size = Pt(16)
            subtitle_frame.paragraphs[0].font.color.rgb = self.colors['gray_800']
            subtitle_frame.word_wrap = True
    
    def _create_analysis_summary_slides(self, soup: BeautifulSoup) -> None:
        """Analysis Summary 섹션 슬라이드 생성"""
        analysis_div = soup.find('div', class_='analysis-summary')
        if not analysis_div:
            return
        
        # 전체 요약 슬라이드
        summary_section = analysis_div.find('div', class_='summary-section')
        if summary_section:
            self._create_summary_slide(summary_section)
        
        # Target Gene Ranking 슬라이드
        summary_sections = analysis_div.find_all('div', class_='summary-section')
        if len(summary_sections) > 1:
            self._create_ranking_slide(summary_sections[1])
    
    def _create_summary_slide(self, section: Tag) -> None:
        """전체 요약 슬라이드 생성"""
        slide = self.prs.slides.add_slide(self.prs.slide_layouts[6])
        
        # 헤더 추출
        header = section.find('div', class_='section-header')
        header_text = header.get_text(strip=True) if header else "Analysis Summary"
        
        # 제목 추가
        title_box = slide.shapes.add_textbox(
            self.margin_left, self.margin_top - Inches(0.2),
            self.content_width, Inches(0.5)
        )
        title_frame = title_box.text_frame
        title_frame.text = header_text
        title_para = title_frame.paragraphs[0]
        title_para.font.size = Pt(28)
        title_para.font.bold = True
        title_para.font.color.rgb = self.colors['primary_red']
        
        # 테이블 추출 - 키-값 형태는 카드 스타일로 표시
        table_elem = section.find('table')
        if table_elem:
            # 키-값 형태인지 확인 (thead 없고, 2열)
            thead = table_elem.find('thead')
            tbody = table_elem.find('tbody')
            
            if not thead and tbody:
                rows = tbody.find_all('tr')
                first_row = rows[0] if rows else None
                if first_row:
                    cells = first_row.find_all(['th', 'td'])
                    if len(cells) == 2 and len(rows) <= 5:
                        # 카드 스타일로 표시
                        self._add_key_value_cards(
                            slide, table_elem,
                            self.margin_left, self.margin_top + Inches(0.6),
                            self.content_width, Inches(5)
                        )
                        return
            
            # 일반 테이블
            self._add_table_to_slide(
                slide, table_elem,
                self.margin_left, self.margin_top + Inches(0.6),
                self.content_width, Inches(5)
            )
    
    def _create_ranking_slide(self, section: Tag) -> None:
        """Target Gene Ranking 슬라이드 생성"""
        slide = self.prs.slides.add_slide(self.prs.slide_layouts[6])
        
        # 헤더 추출
        header = section.find('div', class_='section-header')
        header_text = header.get_text(strip=True) if header else "Target Gene Ranking"
        
        # 제목 추가
        title_box = slide.shapes.add_textbox(
            self.margin_left, self.margin_top - Inches(0.2),
            self.content_width, Inches(0.5)
        )
        title_frame = title_box.text_frame
        title_frame.text = header_text
        title_para = title_frame.paragraphs[0]
        title_para.font.size = Pt(28)
        title_para.font.bold = True
        title_para.font.color.rgb = self.colors['primary_red']
        
        # 테이블 추출 및 추가
        table_elem = section.find('table')
        if table_elem:
            self._add_table_to_slide(
                slide, table_elem,
                self.margin_left, self.margin_top + Inches(0.6),
                self.content_width, Inches(5.5)
            )
    
    def _process_main_content(self, soup: BeautifulSoup) -> None:
        """메인 컨텐츠 전체 처리"""
        # content-container 내의 모든 섹션 찾기
        content_container = soup.find('div', class_='content-container')
        if not content_container:
            return
        
        # Gene 타이틀 (대제목)
        gene_title_elem = content_container.find('h1', class_='gene-title')
        main_gene_title = gene_title_elem.get_text(strip=True) if gene_title_elem else "Gene Analysis"
        
        # 모든 gene-section 처리
        gene_sections = content_container.find_all('div', class_='gene-section')
        seq_viewer_index = 0  # SeqViewerApp 전용 인덱스
        
        for idx, gene_section in enumerate(gene_sections, 1):
            # 섹션 제목 추출
            subsection_title = gene_section.find('h2', class_='subsection-title')
            section_title = subsection_title.get_text(strip=True) if subsection_title else f"Section {idx}"
            
            # SeqViewerApp 캡처 (Playwright 사용) - find_all로 모든 요소 처리
            seq_viewers = gene_section.find_all('div', class_='SeqViewerApp', recursive=False)
            for sv_idx, seq_viewer in enumerate(seq_viewers):
                if self.html_path:
                    viewer_title = section_title + f" - Sequence Viewer"
                    if len(seq_viewers) > 1:
                        viewer_title += f" ({sv_idx + 1})"
                    self._capture_element_screenshot(
                        '.SeqViewerApp', 
                        viewer_title,
                        seq_viewer_index
                    )
                    seq_viewer_index += 1
            
            # 이미지가 있는 경우 처리
            image_placeholder = gene_section.find('div', class_='image-placeholder')
            if image_placeholder:
                img = image_placeholder.find('img')
                if img and img.get('src', '').startswith('data:image'):
                    self._create_image_slide(img, section_title, main_gene_title)
            
            # 테이블이 있는 경우 - 작은 테이블들은 하나의 슬라이드에 합침
            tables = gene_section.find_all('table', class_='data-table')
            if tables:
                # 테이블 정보 수집
                table_infos = []
                for table in tables:
                    row_count = len(table.find_all('tr'))
                    table_infos.append({'table': table, 'rows': row_count})
                
                # 작은 테이블들 그룹화 (합쳐서 8행 이하인 경우)
                i = 0
                while i < len(table_infos):
                    current_group = [table_infos[i]]
                    total_rows = table_infos[i]['rows']
                    
                    # 다음 테이블과 합칠 수 있는지 확인 (동적 여백 계산)
                    # 슬라이드에서 사용 가능한 높이 계산
                    title_space = Inches(0.85)  # 제목 공간 (main_title + section_title)
                    table_gap = Inches(0.3)  # 테이블 간 간격
                    available_height = self.slide_height - self.margin_top - self.margin_bottom - title_space
                    row_height = Inches(0.28)  # 행당 높이 (폰트 크기 + 여백 고려)
                    
                    # 현재 테이블 높이 계산
                    current_height = row_height * total_rows
                    
                    while i + 1 < len(table_infos):
                        next_rows = table_infos[i + 1]['rows']
                        next_table_height = row_height * next_rows + table_gap
                        
                        # 남은 여백에 다음 테이블이 들어갈 수 있는지 확인
                        remaining_space = available_height - current_height
                        
                        if next_table_height <= remaining_space:
                            # 다음 테이블 추가 가능
                            i += 1
                            current_group.append(table_infos[i])
                            total_rows += next_rows
                            current_height += next_table_height
                        else:
                            break
                    
                    # 그룹 처리
                    if len(current_group) == 1:
                        # 단일 테이블
                        table_title = f"{section_title}"
                        self._create_data_table_slide(current_group[0]['table'], table_title, main_gene_title)
                    else:
                        # 여러 테이블을 하나의 슬라이드에 합침
                        self._create_combined_table_slide(
                            [info['table'] for info in current_group], 
                            section_title, 
                            main_gene_title
                        )
                    
                    i += 1
            
            # 하위 섹션 (h3) 처리
            h3_sections = gene_section.find_all('h3')
            for h3 in h3_sections:
                h3_title = h3.get_text(strip=True)
                # h3 다음의 테이블 찾기
                next_table = h3.find_next('table')
                if next_table and next_table.parent == gene_section:
                    self._create_data_table_slide(next_table, h3_title, main_gene_title)
        
        # subsection-title만 있는 섹션들도 처리
        all_subsections = content_container.find_all('h2', class_='subsection-title')
        for subsection in all_subsections:
            section_title = subsection.get_text(strip=True)
            
            # Evidence 테이블 처리
            next_elem = subsection.find_next_sibling()
            while next_elem:
                if next_elem.name == 'h2' or next_elem.get('class') and 'gene-section' in next_elem.get('class', []):
                    break
                
                if next_elem.name == 'div' and 'evidence-table' in next_elem.get('class', []):
                    self._create_evidence_table_slide(next_elem, section_title)
                
                next_elem = next_elem.find_next_sibling()
    
    def _create_data_table_slide(self, table_elem: Tag, section_title: str, main_title: str = "") -> None:
        """데이터 테이블 슬라이드 생성 (가독성 향상, 자동 분할 지원)"""
        slide = self.prs.slides.add_slide(self.prs.slide_layouts[6])
        
        # 메인 타이틀 (작게)
        if main_title:
            main_box = slide.shapes.add_textbox(
                self.margin_left, Inches(0.1),
                self.content_width, Inches(0.25)
            )
            main_frame = main_box.text_frame
            main_frame.text = main_title
            main_para = main_frame.paragraphs[0]
            main_para.font.size = Pt(12)
            main_para.font.color.rgb = self.colors['gray_600']
        
        # 섹션 타이틀
        title_top = Inches(0.35) if main_title else self.margin_top - Inches(0.1)
        title_box = slide.shapes.add_textbox(
            self.margin_left, title_top,
            self.content_width, Inches(0.4)
        )
        title_frame = title_box.text_frame
        title_frame.text = section_title
        title_para = title_frame.paragraphs[0]
        title_para.font.size = Pt(18)
        title_para.font.bold = True
        title_para.font.color.rgb = self.colors['primary_red']
        
        # 테이블 추가 (더 큰 영역)
        table_top = title_top + Inches(0.5)
        table_height = self.slide_height - table_top - self.margin_bottom
        
        # 테이블 추가 (여러 슬라이드 반환 가능)
        created_slides = self._add_improved_table(
            slide, table_elem,
            self.margin_left, table_top,
            self.content_width, table_height
        )
        
        # 추가 슬라이드에 제목 추가
        if created_slides and len(created_slides) > 1:
            for idx, extra_slide in enumerate(created_slides[1:], 2):
                # 추가 슬라이드에도 제목 추가
                if extra_slide:
                    extra_title_box = extra_slide.shapes.add_textbox(
                        self.margin_left, Inches(0.1),
                        self.content_width, Inches(0.4)
                    )
                    extra_title_frame = extra_title_box.text_frame
                    extra_title_frame.text = f"{section_title} (계속 {idx})"
                    extra_title_para = extra_title_frame.paragraphs[0]
                    extra_title_para.font.size = Pt(18)
                    extra_title_para.font.bold = True
                    extra_title_para.font.color.rgb = self.colors['primary_red']
    
    def _create_combined_table_slide(self, tables: List[Tag], section_title: str, main_title: str = "") -> None:
        """여러 작은 테이블을 하나의 슬라이드에 합쳐서 표시"""
        slide = self.prs.slides.add_slide(self.prs.slide_layouts[6])
        
        # 메인 타이틀 (작게)
        if main_title:
            main_box = slide.shapes.add_textbox(
                self.margin_left, Inches(0.1),
                self.content_width, Inches(0.25)
            )
            main_frame = main_box.text_frame
            main_frame.text = main_title
            main_para = main_frame.paragraphs[0]
            main_para.font.size = Pt(12)
            main_para.font.color.rgb = self.colors['gray_600']
        
        # 섹션 타이틀
        title_top = Inches(0.35) if main_title else self.margin_top - Inches(0.1)
        title_box = slide.shapes.add_textbox(
            self.margin_left, title_top,
            self.content_width, Inches(0.4)
        )
        title_frame = title_box.text_frame
        title_frame.text = section_title
        title_para = title_frame.paragraphs[0]
        title_para.font.size = Pt(18)
        title_para.font.bold = True
        title_para.font.color.rgb = self.colors['primary_red']
        
        # 테이블들을 순차적으로 추가
        current_top = title_top + Inches(0.5)
        available_height = self.slide_height - current_top - self.margin_bottom
        
        for table_idx, table_elem in enumerate(tables):
            # 각 테이블의 행 수 계산
            rows = table_elem.find_all('tr')
            row_count = len(rows)
            
            # 테이블 높이 추정 (행당 0.25인치)
            table_height = min(Inches(0.25) * row_count, available_height * 0.4)
            
            if table_idx > 0:
                # 테이블 간 간격
                current_top += Inches(0.2)
            
            # 테이블 추가
            self._add_improved_table(
                slide, table_elem,
                self.margin_left, current_top,
                self.content_width, table_height
            )
            
            # 다음 테이블 위치 계산
            current_top += table_height + Inches(0.1)

    def _create_gene_overview_slide(self, gene_section: Tag, gene_title: str) -> None:
        """Gene 개요 슬라이드 생성"""
        slide = self.prs.slides.add_slide(self.prs.slide_layouts[6])
        
        # 제목
        title_box = slide.shapes.add_textbox(
            self.margin_left, self.margin_top - Inches(0.2),
            self.content_width, Inches(0.5)
        )
        title_frame = title_box.text_frame
        title_frame.text = gene_title
        title_para = title_frame.paragraphs[0]
        title_para.font.size = Pt(32)
        title_para.font.bold = True
        title_para.font.color.rgb = self.colors['primary_red']
        
        # Background 섹션 찾기
        background_div = gene_section.find('div', class_='background-text')
        if background_div:
            y_position = self.margin_top + Inches(0.7)
            
            # "Background" 부제목
            subtitle_box = slide.shapes.add_textbox(
                self.margin_left, y_position,
                self.content_width, Inches(0.3)
            )
            subtitle_frame = subtitle_box.text_frame
            subtitle_frame.text = "Background"
            subtitle_para = subtitle_frame.paragraphs[0]
            subtitle_para.font.size = Pt(20)
            subtitle_para.font.bold = True
            subtitle_para.font.color.rgb = self.colors['gray_800']
            
            # Background 텍스트
            background_text = background_div.get_text(strip=True)
            text_box = slide.shapes.add_textbox(
                self.margin_left, y_position + Inches(0.4),
                self.content_width, Inches(4)
            )
            text_frame = text_box.text_frame
            text_frame.text = background_text
            text_frame.word_wrap = True
            
            for paragraph in text_frame.paragraphs:
                paragraph.font.size = Pt(12)
                paragraph.font.color.rgb = self.colors['gray_800']
                paragraph.line_spacing = 1.5
    
    def _create_image_slide(self, img_tag: Tag, section_title: str, main_title: str = "") -> None:
        """이미지를 포함한 슬라이드 생성"""
        import base64
        from io import BytesIO
        
        src = img_tag.get('src', '')
        if not src.startswith('data:image'):
            return
        
        try:
            # base64 디코딩
            header, data = src.split(',', 1)
            img_bytes = base64.b64decode(data)
            
            # PIL로 이미지 열기
            pil_img = Image.open(BytesIO(img_bytes))
            img_width, img_height = pil_img.size
            
            # 새 슬라이드 생성
            slide = self.prs.slides.add_slide(self.prs.slide_layouts[6])
            
            # 제목 추가
            title_box = slide.shapes.add_textbox(
                self.margin_left, Inches(0.2),
                self.content_width, Inches(0.4)
            )
            title_frame = title_box.text_frame
            title_frame.text = section_title if section_title else "Analysis Chart"
            title_para = title_frame.paragraphs[0]
            title_para.font.size = Pt(20)
            title_para.font.bold = True
            title_para.font.color.rgb = self.colors['primary_red']
            
            # 이미지 크기 계산 (슬라이드에 맞게 조정)
            available_width = self.content_width
            available_height = self.slide_height - Inches(1.2)  # 제목 공간 제외
            
            # 비율 유지하며 크기 조정
            img_ratio = img_width / img_height
            available_ratio = available_width / available_height
            
            if img_ratio > available_ratio:
                # 너비 기준
                final_width = available_width
                final_height = available_width / img_ratio
            else:
                # 높이 기준
                final_height = available_height
                final_width = available_height * img_ratio
            
            # 중앙 정렬
            img_left = self.margin_left + (available_width - final_width) / 2
            img_top = Inches(0.7) + (available_height - final_height) / 2
            
            # 이미지를 BytesIO로 저장
            img_stream = BytesIO()
            pil_img.save(img_stream, format='PNG')
            img_stream.seek(0)
            
            # 슬라이드에 이미지 추가
            slide.shapes.add_picture(
                img_stream,
                img_left, img_top,
                final_width, final_height
            )
            
            logger.info(f"이미지 슬라이드 생성: {section_title} ({img_width}x{img_height})")
            
        except Exception as e:
            logger.error(f"이미지 슬라이드 생성 실패: {e}")
    
    def _create_gene_content_slides(self, gene_section: Tag, gene_title: str) -> None:
        """Gene 섹션의 상세 내용 슬라이드 생성"""
        
        # 주요 기관 권고 현황 테이블
        major_table = gene_section.find('table', class_='major-institution-table')
        if major_table:
            self._create_table_slide(major_table, f"{gene_title} - 주요 기관 권고 현황")
        
        # 제조사별 상용화 키트 테이블
        company_table = gene_section.find('table', class_='company-product-table')
        if company_table:
            self._create_table_slide(company_table, f"{gene_title} - 제조사별 상용화 키트")
        
        # Reference 카드들
        reference_cards = gene_section.find_all('div', class_='reference-card')
        for i, card in enumerate(reference_cards, 1):
            self._create_reference_slide(card, f"{gene_title} - Reference {i}")
    
    def _create_table_slide(self, table_elem: Tag, slide_title: str) -> None:
        """테이블 전용 슬라이드 생성"""
        slide = self.prs.slides.add_slide(self.prs.slide_layouts[6])
        
        # 제목
        title_box = slide.shapes.add_textbox(
            self.margin_left, self.margin_top - Inches(0.2),
            self.content_width, Inches(0.5)
        )
        title_frame = title_box.text_frame
        title_frame.text = slide_title
        title_para = title_frame.paragraphs[0]
        title_para.font.size = Pt(20)
        title_para.font.bold = True
        title_para.font.color.rgb = self.colors['primary_red']
        
        # 테이블 추가
        self._add_table_to_slide(
            slide, table_elem,
            self.margin_left, self.margin_top + Inches(0.6),
            self.content_width, Inches(5.5)
        )
    
    def _create_reference_slide(self, reference_card: Tag, slide_title: str) -> None:
        """Reference 슬라이드 생성"""
        slide = self.prs.slides.add_slide(self.prs.slide_layouts[6])
        
        y_position = self.margin_top - Inches(0.2)
        
        # Reference 번호
        ref_number = reference_card.find('div', class_='reference-number')
        if ref_number:
            ref_box = slide.shapes.add_textbox(
                self.margin_left, y_position,
                self.content_width, Inches(0.4)
            )
            ref_frame = ref_box.text_frame
            ref_frame.text = ref_number.get_text(strip=True)
            ref_para = ref_frame.paragraphs[0]
            ref_para.font.size = Pt(24)
            ref_para.font.bold = True
            ref_para.font.color.rgb = self.colors['primary_red']
            y_position += Inches(0.5)
        
        # Reference 타이틀
        ref_title = reference_card.find('div', class_='reference-title')
        if ref_title:
            title_box = slide.shapes.add_textbox(
                self.margin_left, y_position,
                self.content_width, Inches(0.6)
            )
            title_frame = title_box.text_frame
            title_frame.text = ref_title.get_text(strip=True)
            title_frame.word_wrap = True
            title_para = title_frame.paragraphs[0]
            title_para.font.size = Pt(16)
            title_para.font.bold = True
            title_para.font.color.rgb = self.colors['gray_800']
            y_position += Inches(0.7)
        
        # Reference 메타정보
        ref_meta = reference_card.find('div', class_='reference-meta')
        if ref_meta:
            meta_items = ref_meta.find_all('div', class_='reference-meta-item')
            meta_text = " | ".join([item.get_text(strip=True) for item in meta_items])
            
            meta_box = slide.shapes.add_textbox(
                self.margin_left, y_position,
                self.content_width, Inches(0.4)
            )
            meta_frame = meta_box.text_frame
            meta_frame.text = meta_text
            meta_frame.word_wrap = True
            meta_para = meta_frame.paragraphs[0]
            meta_para.font.size = Pt(11)
            meta_para.font.color.rgb = self.colors['gray_600']
            y_position += Inches(0.5)
        
        # Reference 요약
        ref_summary = reference_card.find('div', class_='reference-summary')
        if ref_summary:
            summary_box = slide.shapes.add_textbox(
                self.margin_left, y_position,
                self.content_width, Inches(3)
            )
            summary_frame = summary_box.text_frame
            summary_frame.text = ref_summary.get_text(strip=True)
            summary_frame.word_wrap = True
            
            for paragraph in summary_frame.paragraphs:
                paragraph.font.size = Pt(12)
                paragraph.font.color.rgb = self.colors['gray_800']
                paragraph.line_spacing = 1.4
            y_position += Inches(3.2)
        
        # Evidence 테이블
        evidence_table = reference_card.find('div', class_='evidence-table')
        if evidence_table:
            # 테이블을 HTML table 형태로 변환하여 추가
            evidence_rows = evidence_table.find_all('div', class_='evidence-row')
            if evidence_rows:
                self._add_evidence_to_slide(slide, evidence_rows, y_position)
    
    def _add_evidence_to_slide(self, slide, evidence_rows: List[Tag], y_position: float) -> None:
        """Evidence 정보를 슬라이드에 추가"""
        # Evidence 제목
        evidence_title_box = slide.shapes.add_textbox(
            self.margin_left, y_position,
            self.content_width, Inches(0.3)
        )
        evidence_title_frame = evidence_title_box.text_frame
        evidence_title_frame.text = "Evidence Details"
        evidence_title_para = evidence_title_frame.paragraphs[0]
        evidence_title_para.font.size = Pt(14)
        evidence_title_para.font.bold = True
        evidence_title_para.font.color.rgb = self.colors['gray_800']
        
        y_position += Inches(0.4)
        
        # 각 evidence row를 텍스트로 변환
        for i, row in enumerate(evidence_rows[:2]):  # 최대 2개만 표시
            evidence_header = row.find('div', class_='evidence-header')
            evidence_cell = row.find('div', class_='evidence-cell')
            
            if evidence_cell:
                # 모든 텍스트를 하나로 합치기
                texts = [elem.strip() for elem in evidence_cell.stripped_strings]
                combined_text = " | ".join(texts)
                
                text_box = slide.shapes.add_textbox(
                    self.margin_left, y_position,
                    self.content_width, Inches(0.8)
                )
                text_frame = text_box.text_frame
                text_frame.text = combined_text
                text_frame.word_wrap = True
                
                for paragraph in text_frame.paragraphs:
                    paragraph.font.size = Pt(9)
                    paragraph.font.color.rgb = self.colors['gray_800']
                
                y_position += Inches(0.9)
    
    def _add_key_value_cards(
        self,
        slide,
        table_elem: Tag,
        left: float,
        top: float,
        width: float,
        height: float
    ) -> List[Any]:
        """키-값 형태의 테이블을 심플한 카드 스타일로 표시"""
        
        tbody = table_elem.find('tbody')
        if not tbody:
            return []
        
        rows = tbody.find_all('tr')
        if not rows:
            return []
        
        # 심플한 단색 (진한 회색)
        label_bg_color = RGBColor(55, 65, 81)      # Gray-700
        value_bg_color = RGBColor(249, 250, 251)   # Gray-50
        border_color = RGBColor(209, 213, 219)     # Gray-300
        
        y_position = top
        card_height = Inches(1.3)
        card_spacing = Inches(0.12)
        label_width = Inches(1.6)
        
        for i, tr in enumerate(rows):
            cells = tr.find_all(['th', 'td'])
            if len(cells) < 2:
                continue
            
            label = self._clean_text(cells[0].get_text(strip=True))
            value = self._clean_text(cells[1].get_text(strip=True))
            
            # 라벨 영역 (각진 사각형)
            label_shape = slide.shapes.add_shape(
                MSO_SHAPE.RECTANGLE,  # 각진 사각형
                left, y_position,
                label_width, card_height
            )
            label_shape.fill.solid()
            label_shape.fill.fore_color.rgb = label_bg_color
            label_shape.line.fill.background()  # 테두리 없음
            
            # 라벨 텍스트
            label_tf = label_shape.text_frame
            label_tf.word_wrap = True
            label_tf.paragraphs[0].text = label
            label_tf.paragraphs[0].font.size = Pt(13)
            label_tf.paragraphs[0].font.bold = True
            label_tf.paragraphs[0].font.color.rgb = RGBColor(255, 255, 255)
            label_tf.paragraphs[0].alignment = PP_ALIGN.CENTER
            label_shape.text_frame.margin_left = Pt(10)
            label_shape.text_frame.margin_right = Pt(10)
            label_shape.text_frame.margin_top = Pt(10)
            label_shape.text_frame.margin_bottom = Pt(10)
            
            # 수직 중앙 정렬
            from pptx.enum.text import MSO_ANCHOR
            label_tf.auto_size = None
            
            # 값 영역 (각진 사각형)
            value_left = left + label_width
            value_width = width - label_width
            
            value_shape = slide.shapes.add_shape(
                MSO_SHAPE.RECTANGLE,  # 각진 사각형
                value_left, y_position,
                value_width, card_height
            )
            value_shape.fill.solid()
            value_shape.fill.fore_color.rgb = value_bg_color
            value_shape.line.color.rgb = border_color
            value_shape.line.width = Pt(1)
            
            # 값 텍스트
            value_tf = value_shape.text_frame
            value_tf.word_wrap = True
            value_tf.paragraphs[0].text = value
            value_tf.paragraphs[0].font.size = Pt(11)
            value_tf.paragraphs[0].font.color.rgb = self.colors['gray_800']
            value_tf.paragraphs[0].alignment = PP_ALIGN.LEFT
            value_shape.text_frame.margin_left = Pt(12)
            value_shape.text_frame.margin_right = Pt(12)
            value_shape.text_frame.margin_top = Pt(10)
            value_shape.text_frame.margin_bottom = Pt(10)
            
            y_position += card_height + card_spacing
        
        return [slide]
    
    def _add_improved_table(
        self, 
        slide, 
        table_elem: Tag, 
        left: float, 
        top: float, 
        width: float, 
        height: float
    ) -> List[Any]:
        """가독성이 향상된 테이블 추가 (자동 분할 지원, 논문 스타일)"""
        
        # 키-값 형태의 테이블인지 확인 (thead 없고, 2열, 3행 이하)
        thead = table_elem.find('thead')
        tbody = table_elem.find('tbody')
        
        if not thead and tbody:
            rows = tbody.find_all('tr')
            if len(rows) <= 5:  # 5행 이하
                first_row = rows[0] if rows else None
                if first_row:
                    cells = first_row.find_all(['th', 'td'])
                    if len(cells) == 2:  # 2열 (키-값 형태)
                        # 키-값 카드 스타일로 표시
                        return self._add_key_value_cards(slide, table_elem, left, top, width, height)
        
        # 테이블 데이터와 width 정보 추출 (colspan 포함)
        rows_data = []
        header_rows = []
        body_rows = []
        col_widths_html = []
        has_header = False
        merge_info = []  # [(row_idx, col_idx, colspan, rowspan), ...]
        cell_styles = {}  # {(row_idx, col_idx): {'bold': bool, 'color': RGBColor, 'link': str}, ...}
        
        def extract_row_data(tr, row_idx, is_header=False):
            """행 데이터 추출 (colspan 처리 포함)"""
            cells = tr.find_all(['th', 'td'])
            row_data = []
            col_idx = 0
            
            for cell in cells:
                text = self._clean_text(cell.get_text(strip=True))
                colspan = int(cell.get('colspan', 1))
                rowspan = int(cell.get('rowspan', 1))
                
                # 스타일 추출
                styles = self._extract_cell_styles(cell)
                if styles['bold'] or styles['color'] or styles['link']:
                    cell_styles[(row_idx, col_idx)] = styles
                
                row_data.append(text)
                # colspan이 있으면 빈 셀 추가
                for _ in range(colspan - 1):
                    row_data.append('')
                
                # 머지 정보 저장
                if colspan > 1 or rowspan > 1:
                    merge_info.append((row_idx, col_idx, colspan, rowspan))
                
                col_idx += colspan
            
            return row_data
        
        # thead 처리
        if thead:
            has_header = True
            header_trs = thead.find_all('tr')
            for idx, tr in enumerate(header_trs):
                row_data = extract_row_data(tr, len(rows_data))
                header_rows.append(row_data)
                rows_data.append(row_data)
                
                # 첫 번째 헤더 행에서 width 정보 추출
                if not col_widths_html:
                    cells = tr.find_all(['th', 'td'])
                    col_widths_html = self._extract_column_widths(cells)
        
        # tbody 처리
        tbody = table_elem.find('tbody')
        if tbody:
            body_trs = tbody.find_all('tr')
            for idx, tr in enumerate(body_trs):
                row_data = extract_row_data(tr, len(rows_data))
                body_rows.append(row_data)
                rows_data.append(row_data)
                
                # thead가 없는 경우 첫 행에서 width 추출
                if not has_header and idx == 0 and not col_widths_html:
                    cells = tr.find_all(['th', 'td'])
                    col_widths_html = self._extract_column_widths(cells)
        
        # thead도 없고 tbody도 없는 경우 (직접 tr 사용)
        if not has_header and not tbody:
            all_rows = table_elem.find_all('tr')
            for idx, tr in enumerate(all_rows):
                row_data = extract_row_data(tr, len(rows_data))
                body_rows.append(row_data)
                rows_data.append(row_data)
                
                # 첫 번째 행에서 width 정보 추출
                if idx == 0 and not col_widths_html:
                    cells = tr.find_all(['th', 'td'])
                    col_widths_html = self._extract_column_widths(cells)
        
        if not rows_data:
            return []
        
        # 열 개수 결정
        max_cols = max(len(row) for row in rows_data)
        
        # 모든 행의 열 개수를 동일하게 맞춤
        for row in rows_data:
            while len(row) < max_cols:
                row.append("")
        
        # 테이블이 너무 크면 여러 슬라이드로 분할
        header_count = len(header_rows)
        body_count = len(body_rows)
        
        created_slides = []
        
        # 분할이 필요한 경우
        if body_count > self.max_rows_per_slide:
            # 여러 슬라이드로 분할
            num_chunks = (body_count + self.max_rows_per_slide - 1) // self.max_rows_per_slide
            
            logger.info(f"테이블 분할: {body_count}행을 {num_chunks}개 슬라이드로 나눔")
            
            for chunk_idx in range(num_chunks):
                start_idx = chunk_idx * self.max_rows_per_slide
                end_idx = min(start_idx + self.max_rows_per_slide, body_count)
                
                # 청크 데이터 준비
                chunk_data = header_rows + body_rows[start_idx:end_idx]
                
                # 해당 청크에 해당하는 merge_info 필터링
                chunk_merge_info = []
                for row_idx, col_idx, colspan, rowspan in merge_info:
                    # header_count 이후의 행들에 대해 조정
                    if row_idx < header_count:
                        chunk_merge_info.append((row_idx, col_idx, colspan, rowspan))
                    elif row_idx - header_count >= start_idx and row_idx - header_count < end_idx:
                        new_row_idx = header_count + (row_idx - header_count - start_idx)
                        chunk_merge_info.append((new_row_idx, col_idx, colspan, rowspan))
                
                # 해당 청크에 해당하는 cell_styles 필터링
                chunk_cell_styles = {}
                for (r, c), styles in cell_styles.items():
                    if r < header_count:
                        chunk_cell_styles[(r, c)] = styles
                    elif r - header_count >= start_idx and r - header_count < end_idx:
                        new_r = header_count + (r - header_count - start_idx)
                        chunk_cell_styles[(new_r, c)] = styles
                
                # 새 슬라이드 생성 (첫 번째 청크는 기존 슬라이드 사용)
                chunk_slide = slide if chunk_idx == 0 else None
                
                # 테이블 생성
                result = self._create_ppt_table(
                    chunk_slide,
                    chunk_data,
                    header_count,
                    col_widths_html,
                    left, top, width, height,
                    chunk_idx + 1, num_chunks,
                    chunk_merge_info,
                    max_cols,
                    chunk_cell_styles
                )
                
                if result:
                    created_slides.append(result)
            
            return created_slides
        else:
            # 단일 슬라이드에 표시
            result = self._create_ppt_table(
                slide,
                rows_data,
                header_count,
                col_widths_html,
                left, top, width, height,
                1, 1,
                merge_info,
                max_cols,
                cell_styles
            )
            return [result] if result else []
    
    def _extract_column_widths(self, cells: List[Tag]) -> List[Optional[int]]:
        """HTML 테이블 셀에서 width 속성 추출"""
        widths = []
        for cell in cells:
            width = None
            
            # style 속성에서 width 추출
            style = cell.get('style', '')
            if 'width:' in style:
                import re
                match = re.search(r'width:\s*(\d+)(?:px|%)?', style)
                if match:
                    width = int(match.group(1))
            
            # width 속성 직접 확인
            elif cell.get('width'):
                try:
                    width = int(cell.get('width').replace('px', '').replace('%', ''))
                except:
                    pass
            
            widths.append(width)
        
        return widths
    
    def _create_ppt_table(
        self,
        slide,
        rows_data: List[List[str]],
        header_count: int,
        col_widths_html: List[Optional[int]],
        left: float,
        top: float,
        width: float,
        height: float,
        chunk_num: int = 1,
        total_chunks: int = 1,
        merge_info: List[tuple] = None,
        max_cols_override: int = None,
        cell_styles: Dict[Tuple[int, int], Dict[str, Any]] = None
    ) -> Any:
        """실제 PowerPoint 테이블 생성 (논문 스타일)"""
        
        if merge_info is None:
            merge_info = []
        if cell_styles is None:
            cell_styles = {}
        
        # 새 슬라이드가 필요한 경우
        if slide is None:
            slide = self.prs.slides.add_slide(self.prs.slide_layouts[6])
            # 추가 슬라이드의 경우 테이블 시작 위치 조정
            top = Inches(0.6)  # 제목을 위한 공간
            height = self.slide_height - top - self.margin_bottom
        
        if not rows_data:
            return None
        
        max_cols = max_cols_override if max_cols_override else len(rows_data[0])
        
        # 행/열에 따른 폰트 크기 조정
        base_font_size = 8 if len(rows_data) > 15 or max_cols > 6 else 9
        header_font_size = base_font_size + 1
        
        # 테이블 높이 계산 - 콘텐츠에 맞게 최소화
        row_count = len(rows_data)
        # 행당 최소 높이 계산 (더 작게 설정)
        min_row_height = Inches(0.22)  # 최소 행 높이 줄임
        required_height = min_row_height * row_count
        
        # 테이블 높이는 필요한 만큼만 사용 (슬라이드 높이 이하로)
        height = min(required_height, height)
        
        if row_count > 20:
            base_font_size = 7
            header_font_size = 8
        
        try:
            # PowerPoint 테이블 생성
            ppt_table = slide.shapes.add_table(
                row_count, max_cols,
                left, top, width, height
            ).table
            
            # 논문 스타일: 모든 셀의 테두리를 먼저 설정
            from pptx.oxml.ns import qn
            from pptx.oxml import parse_xml
            
            # 테이블 데이터 채우기 및 논문 스타일 적용
            for i, row_data in enumerate(rows_data):
                for j, cell_data in enumerate(row_data):
                    if j >= max_cols:
                        continue
                    cell = ppt_table.cell(i, j)
                    
                    # 텍스트 설정
                    cell.text = str(cell_data) if j < len(row_data) else ""
                    
                    # ✨ 수직 중앙 정렬 (MIDDLE)
                    cell.vertical_anchor = MSO_ANCHOR.MIDDLE
                    
                    # 셀 여백 최소화
                    cell.margin_left = Pt(4)
                    cell.margin_right = Pt(4)
                    cell.margin_top = Pt(2)
                    cell.margin_bottom = Pt(2)
                    
                    # 논문 스타일: 배경색 없음 (투명)
                    cell.fill.background()
                    
                    # HTML에서 추출한 스타일 가져오기
                    html_style = cell_styles.get((i, j), {})
                    has_custom_bold = html_style.get('bold', False)
                    custom_color = html_style.get('color')
                    has_link = html_style.get('link')
                    
                    # 단락 포맷 설정
                    for paragraph in cell.text_frame.paragraphs:
                        # 헤더 행
                        if i < header_count:
                            paragraph.font.size = Pt(header_font_size)
                            paragraph.font.bold = True
                            paragraph.font.color.rgb = self.colors['black']
                            paragraph.alignment = PP_ALIGN.CENTER
                            # 헤더는 줄바꿈 안함 (Gene 등이 줄바꿈되지 않도록)
                            cell.text_frame.word_wrap = False
                        else:
                            paragraph.font.size = Pt(base_font_size)
                            
                            # HTML 스타일 적용 (Bold)
                            if has_custom_bold:
                                paragraph.font.bold = True
                            
                            # HTML 스타일 적용 (Color)
                            if custom_color:
                                paragraph.font.color.rgb = custom_color
                            else:
                                paragraph.font.color.rgb = self.colors['gray_800']
                            
                            # 링크가 있으면 파란색 + 밑줄
                            if has_link:
                                paragraph.font.color.rgb = RGBColor(0, 102, 204)
                                paragraph.font.underline = True
                            
                            # 좌측 정렬이 필요한 열 (긴 텍스트)
                            if len(cell_data) > 30 or '\n' in cell_data:
                                paragraph.alignment = PP_ALIGN.LEFT
                            else:
                                paragraph.alignment = PP_ALIGN.CENTER
                            
                            # 데이터 행은 줄바꿈 허용
                            cell.text_frame.word_wrap = True
                        
                        # 줄 간격
                        paragraph.line_spacing = 1.1
            
            # 논문 스타일 테두리 적용 (상단, 헤더 하단, 하단에만 선)
            self._apply_academic_table_borders(ppt_table, header_count, row_count, max_cols)
            
            # 셀 머지 적용
            for row_idx, col_idx, colspan, rowspan in merge_info:
                try:
                    if row_idx < row_count and col_idx < max_cols:
                        start_cell = ppt_table.cell(row_idx, col_idx)
                        end_row = min(row_idx + rowspan - 1, row_count - 1)
                        end_col = min(col_idx + colspan - 1, max_cols - 1)
                        
                        end_cell = ppt_table.cell(end_row, end_col)
                        start_cell.merge(end_cell)
                        
                        # 머지된 셀 중앙 정렬
                        for paragraph in start_cell.text_frame.paragraphs:
                            paragraph.alignment = PP_ALIGN.CENTER
                except Exception as merge_err:
                    logger.debug(f"셀 머지 실패: {merge_err}")
            
            # HTML width 속성을 기반으로 열 너비 조정
            if col_widths_html and any(w is not None for w in col_widths_html):
                self._apply_html_column_widths(ppt_table, col_widths_html, width)
            else:
                # 자동 조정
                self._adjust_column_widths(ppt_table, rows_data)
            
            return slide
            
        except Exception as e:
            logger.error(f"테이블 추가 실패: {e}")
    
    def _apply_academic_table_borders(self, ppt_table, header_count: int, row_count: int, col_count: int) -> None:
        """테이블 테두리 적용 (헤더에 굵은 선, 데이터 행에 얇은 가로선)"""
        from pptx.util import Pt
        from pptx.dml.color import RGBColor
        from pptx.oxml.ns import nsmap
        
        # 선 두께 정의
        thick_line = Pt(1.5)  # 굵은 선 (헤더 상하단)
        thin_line = Pt(0.5)   # 얇은 선 (데이터 행)
        no_line = Pt(0)       # 선 없음
        
        black = RGBColor(0, 0, 0)
        gray_line = RGBColor(200, 200, 200)  # 연한 회색 선
        
        for i in range(row_count):
            for j in range(col_count):
                try:
                    cell = ppt_table.cell(i, j)
                    
                    # 각 셀의 테두리 설정
                    tc = cell._tc
                    tcPr = tc.get_or_add_tcPr()
                    
                    # 상단 선
                    if i == 0:
                        # 첫 행: 상단에 굵은 선
                        self._set_cell_border(cell, 'top', thick_line, black)
                    elif i == header_count and header_count > 0:
                        # 데이터 첫 행: 헤더 하단 굵은 선이 이미 있음
                        self._set_cell_border(cell, 'top', no_line, black)
                    else:
                        # 데이터 행 상단: 얇은 회색 선
                        self._set_cell_border(cell, 'top', thin_line, gray_line)
                    
                    # 하단 선
                    if i == row_count - 1:
                        # 마지막 행: 하단에 굵은 선
                        self._set_cell_border(cell, 'bottom', thick_line, black)
                    elif i == header_count - 1 and header_count > 0:
                        # 헤더 마지막 행: 하단에 굵은 선
                        self._set_cell_border(cell, 'bottom', thick_line, black)
                    else:
                        # 데이터 행: 하단에 얇은 회색 선
                        self._set_cell_border(cell, 'bottom', thin_line, gray_line)
                    
                    # 좌우 테두리 없음
                    self._set_cell_border(cell, 'left', no_line, black)
                    self._set_cell_border(cell, 'right', no_line, black)
                    
                except Exception as e:
                    pass  # 테두리 설정 실패 시 무시
    
    def _set_cell_border(self, cell, side: str, width, color: RGBColor) -> None:
        """셀의 특정 테두리 설정 (개선된 버전)"""
        from pptx.oxml.ns import qn
        from lxml import etree
        
        tc = cell._tc
        tcPr = tc.get_or_add_tcPr()
        
        # 테두리 요소 이름
        border_map = {
            'top': 'a:lnT',
            'bottom': 'a:lnB', 
            'left': 'a:lnL',
            'right': 'a:lnR'
        }
        
        border_elem_name = border_map.get(side)
        if not border_elem_name:
            return
        
        # 기존 테두리 요소 제거
        for existing in list(tcPr):
            if existing.tag == qn(border_elem_name):
                tcPr.remove(existing)
        
        # EMU 단위로 변환 (1 pt = 12700 EMU)
        width_emu = int(width) if width > 0 else 0
        
        # 새 테두리 요소 생성
        ln = etree.Element(qn(border_elem_name))
        
        if width_emu > 0:
            ln.set('w', str(width_emu))
            ln.set('cap', 'flat')
            ln.set('cmpd', 'sng')
            ln.set('algn', 'ctr')
            
            # 색상 설정
            solidFill = etree.SubElement(ln, qn('a:solidFill'))
            srgbClr = etree.SubElement(solidFill, qn('a:srgbClr'))
            srgbClr.set('val', '%02X%02X%02X' % (color[0], color[1], color[2]))
            
            # 프리셋 대시
            prstDash = etree.SubElement(ln, qn('a:prstDash'))
            prstDash.set('val', 'solid')
        else:
            ln.set('w', '0')
            noFill = etree.SubElement(ln, qn('a:noFill'))
        
        # tcPr의 첫 번째 자식으로 추가 (순서 중요)
        tcPr.insert(0, ln)
    
    def _apply_html_column_widths(
        self, 
        ppt_table, 
        col_widths_html: List[Optional[int]], 
        total_width: float
    ) -> None:
        """HTML에서 추출한 width 속성을 PowerPoint 테이블에 적용
        
        HTML의 width는 보통 픽셀 단위이며, 지정된 열은 그 너비를,
        지정되지 않은 열은 나머지 공간을 차지합니다.
        """
        try:
            col_count = len(col_widths_html)
            if col_count == 0:
                return
            
            # width가 지정된 열과 지정되지 않은 열 분리
            specified_widths = [w for w in col_widths_html if w is not None]
            unspecified_count = col_widths_html.count(None)
            
            if not specified_widths:
                return
            
            # HTML에서 지정된 width를 PowerPoint 단위로 변환
            # 일반적으로 HTML 테이블은 800-1000px 기준
            # PowerPoint 슬라이드는 약 9.4인치 = 약 900px (96dpi 기준)
            html_to_ppt_ratio = total_width / 900  # 1px ≈ 이 비율의 EMU
            
            # 지정된 열들의 너비 먼저 계산
            specified_total_ppt = 0
            for html_width in col_widths_html:
                if html_width is not None:
                    ppt_width = int(html_width * html_to_ppt_ratio)
                    specified_total_ppt += ppt_width
            
            # 남은 너비 계산
            remaining_width = total_width - specified_total_ppt
            
            # 미지정 열에 충분한 공간이 없으면 비율 조정
            if remaining_width < 0 or (unspecified_count > 0 and remaining_width < total_width * 0.3):
                # 지정된 열을 전체의 30%로 제한, 나머지는 미지정 열에
                specified_portion = 0.3 if unspecified_count > 0 else 1.0
                total_specified_html = sum(specified_widths)
                
                for j, html_width in enumerate(col_widths_html):
                    if html_width is not None:
                        proportion = html_width / total_specified_html
                        ppt_table.columns[j].width = int(total_width * specified_portion * proportion)
                
                if unspecified_count > 0:
                    remaining = total_width * (1 - specified_portion)
                    equal_width = int(remaining / unspecified_count)
                    for j, html_width in enumerate(col_widths_html):
                        if html_width is None:
                            ppt_table.columns[j].width = equal_width
            else:
                # 충분한 공간이 있으면 그대로 적용
                for j, html_width in enumerate(col_widths_html):
                    if html_width is not None:
                        ppt_table.columns[j].width = int(html_width * html_to_ppt_ratio)
                
                # 남은 너비를 미지정 열에 균등 분배
                if unspecified_count > 0:
                    equal_width = int(remaining_width / unspecified_count)
                    for j, html_width in enumerate(col_widths_html):
                        if html_width is None:
                            ppt_table.columns[j].width = equal_width
        
        except Exception as e:
            logger.debug(f"HTML width 적용 실패, 자동 조정으로 전환: {e}")
    
    def _adjust_column_widths(self, ppt_table, rows_data: List[List[str]]) -> None:
        """열 너비 자동 조정 (HTML width가 없는 경우) - 텍스트 길이 기반"""
        try:
            col_count = len(rows_data[0]) if rows_data else 0
            if col_count == 0:
                return
            
            # 현재 전체 테이블 너비 계산
            total_table_width = sum(col.width for col in ppt_table.columns)
            
            # 각 열의 최대 텍스트 길이 계산 (가중치 적용)
            max_lengths = [0] * col_count
            for row in rows_data:
                for j, cell in enumerate(row):
                    cell_text = str(cell)
                    # 한글은 더 넓은 공간 필요 (1.5배)
                    korean_count = len([c for c in cell_text if ord(c) >= 0xAC00 and ord(c) <= 0xD7A3])
                    english_count = len(cell_text) - korean_count
                    weighted_length = english_count + (korean_count * 1.8)
                    max_lengths[j] = max(max_lengths[j], weighted_length)
            
            # 최소 너비 보장 (열당 최소 5%)
            min_proportion = 0.05
            
            # 총 가중치 길이
            total_length = sum(max_lengths)
            if total_length == 0:
                # 모두 빈 셀이면 균등 분배
                equal_width = total_table_width // col_count
                for j in range(col_count):
                    ppt_table.columns[j].width = equal_width
                return
            
            # 각 열에 비례적으로 너비 할당
            for j in range(col_count):
                proportion = max_lengths[j] / total_length
                # 최소 너비 보장
                proportion = max(proportion, min_proportion)
                col_width = int(total_table_width * proportion)
                ppt_table.columns[j].width = col_width
                
            logger.debug(f"열 너비 자동 조정 완료: {[ppt_table.columns[j].width for j in range(col_count)]}")
        
        except Exception as e:
            logger.debug(f"열 너비 조정 실패 (무시): {e}")
    
    def _create_evidence_table_slide(self, evidence_div: Tag, section_title: str) -> None:
        """Evidence 테이블 슬라이드 생성 (원본 헤더 유지, 링크 포함)"""
        slide = self.prs.slides.add_slide(self.prs.slide_layouts[6])
        
        # 제목
        title_box = slide.shapes.add_textbox(
            self.margin_left, self.margin_top - Inches(0.2),
            self.content_width, Inches(0.4)
        )
        title_frame = title_box.text_frame
        title_frame.text = section_title
        title_para = title_frame.paragraphs[0]
        title_para.font.size = Pt(18)
        title_para.font.bold = True
        title_para.font.color.rgb = self.colors['primary_red']
        
        # Evidence 행들 찾기
        evidence_rows = evidence_div.find_all('div', class_='evidence-row')
        
        if not evidence_rows:
            # evidence-cell로도 시도
            evidence_rows = evidence_div.find_all('div', class_='evidence-cell')
        
        if not evidence_rows:
            return
        
        # 최대 10개 행만 처리 (슬라이드에 맞게)
        max_rows = min(len(evidence_rows), 10)
        
        # 테이블 데이터 추출
        table_data = []
        link_data = []  # 각 셀의 링크 정보 저장 [(row, col, url), ...]
        
        # 헤더 추출 (원본 유지)
        header_div = evidence_div.find('div', class_='evidence-header')
        if header_div:
            headers = [elem.strip() for elem in header_div.stripped_strings]
            # 원본 헤더 사용 (최대 8개 열)
            if len(headers) > 0:
                table_data.append(headers[:8])
        
        # 데이터 행 추출
        for row_idx, row in enumerate(evidence_rows[:max_rows]):
            row_texts = []
            text_elements = row.find_all('div', class_='evidence-text')
            
            for col_idx, elem in enumerate(text_elements[:8]):
                # 링크 확인
                link = elem.find('a')
                if link:
                    link_text = link.get_text(strip=True)
                    link_url = link.get('href', '')
                    row_texts.append(link_text)
                    if link_url:
                        link_data.append((len(table_data), col_idx, link_url))
                else:
                    text = self._clean_text(elem.get_text(strip=True))
                    # 너무 긴 텍스트는 자르기
                    if len(text) > 80:
                        text = text[:77] + "..."
                    row_texts.append(text)
            
            if row_texts:
                table_data.append(row_texts)
        
        if len(table_data) <= 1:
            return
        
        # 열 수 통일
        max_cols = max(len(row) for row in table_data)
        for row in table_data:
            while len(row) < max_cols:
                row.append('')
        
        # 테이블 생성
        try:
            table_top = self.margin_top + Inches(0.4)
            table_height = self.slide_height - table_top - self.margin_bottom
            
            ppt_table = slide.shapes.add_table(
                len(table_data), max_cols,
                self.margin_left, table_top,
                self.content_width, table_height
            ).table
            
            # 데이터 채우기
            for i, row_data in enumerate(table_data):
                for j, cell_data in enumerate(row_data):
                    if j >= max_cols:
                        continue
                    cell = ppt_table.cell(i, j)
                    cell.text = str(cell_data)
                    cell.vertical_anchor = MSO_ANCHOR.MIDDLE
                    
                    # 셀 여백
                    cell.margin_left = Pt(3)
                    cell.margin_right = Pt(3)
                    cell.margin_top = Pt(2)
                    cell.margin_bottom = Pt(2)
                    
                    # 배경색 없음
                    cell.fill.background()
                    
                    for paragraph in cell.text_frame.paragraphs:
                        paragraph.font.size = Pt(7)
                        
                        if i == 0:  # 헤더
                            paragraph.font.bold = True
                            paragraph.font.color.rgb = self.colors['black']
                            paragraph.alignment = PP_ALIGN.CENTER
                            cell.text_frame.word_wrap = False
                        else:
                            paragraph.font.color.rgb = self.colors['gray_800']
                            cell.text_frame.word_wrap = True
                            # 문서 제목, AI 요약 열은 왼쪽 정렬 (2, 5번 인덱스)
                            if j in [2, 5]:
                                paragraph.alignment = PP_ALIGN.LEFT
                            else:
                                paragraph.alignment = PP_ALIGN.CENTER
                            
                            # Link 열인 경우 파란색으로 표시
                            if cell_data == 'Link':
                                paragraph.font.color.rgb = RGBColor(0, 102, 204)
                                paragraph.font.underline = True
            
            # 하이퍼링크 추가
            for row_idx, col_idx, url in link_data:
                try:
                    cell = ppt_table.cell(row_idx, col_idx)
                    # PowerPoint에서는 테이블 셀에 직접 하이퍼링크를 추가하기 어려움
                    # 대신 텍스트에 밑줄과 파란색을 유지
                    for paragraph in cell.text_frame.paragraphs:
                        for run in paragraph.runs:
                            run.font.color.rgb = RGBColor(0, 102, 204)
                            run.font.underline = True
                except:
                    pass
            
            # 테두리 적용
            self._apply_academic_table_borders(ppt_table, 1, len(table_data), max_cols)
        
        except Exception as e:
            logger.error(f"Evidence 테이블 생성 실패: {e}")
    
    def _add_table_to_slide(
        self, 
        slide, 
        table_elem: Tag, 
        left: float, 
        top: float, 
        width: float, 
        height: float
    ) -> None:
        """HTML 테이블을 PowerPoint 테이블로 변환하여 추가"""
        
        # 테이블 데이터 추출
        rows_data = []
        col_widths_html = []
        header_count = 0
        merge_info = []  # [(row_idx, col_idx, colspan, rowspan), ...]
        
        # thead와 tbody 모두 처리
        thead = table_elem.find('thead')
        tbody = table_elem.find('tbody')
        
        if thead:
            header_rows = thead.find_all('tr')
            for idx, tr in enumerate(header_rows):
                cells = tr.find_all(['th', 'td'])
                row_data = []
                col_idx = 0
                for cell in cells:
                    text = cell.get_text(strip=True)
                    colspan = int(cell.get('colspan', 1))
                    rowspan = int(cell.get('rowspan', 1))
                    
                    row_data.append(text)
                    # colspan이 있으면 빈 셀 추가
                    for _ in range(colspan - 1):
                        row_data.append('')
                    
                    # 머지 정보 저장
                    if colspan > 1 or rowspan > 1:
                        merge_info.append((len(rows_data), col_idx, colspan, rowspan))
                    
                    col_idx += colspan
                
                rows_data.append(row_data)
                header_count += 1
                
                # 첫 헤더 행에서 width 추출
                if idx == 0 and not col_widths_html:
                    col_widths_html = self._extract_column_widths(cells)
        
        if tbody:
            body_rows = tbody.find_all('tr')
            for idx, tr in enumerate(body_rows):
                cells = tr.find_all(['th', 'td'])
                row_data = []
                col_idx = 0
                for cell in cells:
                    text = cell.get_text(strip=True)
                    colspan = int(cell.get('colspan', 1))
                    rowspan = int(cell.get('rowspan', 1))
                    
                    row_data.append(text)
                    # colspan이 있으면 빈 셀 추가
                    for _ in range(colspan - 1):
                        row_data.append('')
                    
                    # 머지 정보 저장
                    if colspan > 1 or rowspan > 1:
                        merge_info.append((len(rows_data), col_idx, colspan, rowspan))
                    
                    col_idx += colspan
                
                rows_data.append(row_data)
                
                # thead가 없으면 첫 행에서 width 추출
                if not thead and idx == 0 and not col_widths_html:
                    col_widths_html = self._extract_column_widths(cells)
        
        # 테이블이 비어있으면 리턴
        if not rows_data:
            return
        
        # 열 개수 결정
        max_cols = max(len(row) for row in rows_data)
        
        # 모든 행의 열 개수를 동일하게 맞춤
        for row in rows_data:
            while len(row) < max_cols:
                row.append("")
        
        # PowerPoint 테이블 생성
        try:
            ppt_table = slide.shapes.add_table(
                len(rows_data), max_cols,
                left, top, width, height
            ).table
            
            # 테이블 데이터 채우기 (논문 스타일)
            for i, row_data in enumerate(rows_data):
                for j, cell_data in enumerate(row_data):
                    cell = ppt_table.cell(i, j)
                    cell.text = str(cell_data)
                    
                    # ✨ 수직 중앙 정렬 (MIDDLE)
                    cell.vertical_anchor = MSO_ANCHOR.MIDDLE
                    
                    # 셀 여백 최소화
                    cell.margin_left = Pt(4)
                    cell.margin_right = Pt(4)
                    cell.margin_top = Pt(2)
                    cell.margin_bottom = Pt(2)
                    
                    # 논문 스타일: 배경색 없음
                    cell.fill.background()
                    
                    # 텍스트 포맷 설정
                    for paragraph in cell.text_frame.paragraphs:
                        paragraph.font.size = Pt(9)
                        
                        # 헤더 행 스타일
                        if i < header_count or i == 0:
                            paragraph.font.bold = True
                            paragraph.font.color.rgb = self.colors['black']
                            paragraph.alignment = PP_ALIGN.CENTER
                            cell.text_frame.word_wrap = False  # 헤더 줄바꿈 방지
                        else:
                            paragraph.font.color.rgb = self.colors['gray_800']
                            paragraph.alignment = PP_ALIGN.CENTER
                            cell.text_frame.word_wrap = True  # 데이터는 줄바꿈 허용
            
            # 논문 스타일 테두리 적용
            self._apply_academic_table_borders(ppt_table, header_count, len(rows_data), max_cols)
            
            # HTML width 속성을 기반으로 열 너비 조정
            if col_widths_html and any(w is not None for w in col_widths_html):
                self._apply_html_column_widths(ppt_table, col_widths_html, width)
            else:
                # 자동 조정
                self._adjust_column_widths(ppt_table, rows_data)
            
            # 셀 머지 적용
            for row_idx, col_idx, colspan, rowspan in merge_info:
                try:
                    start_cell = ppt_table.cell(row_idx, col_idx)
                    end_row = row_idx + rowspan - 1
                    end_col = col_idx + colspan - 1
                    
                    # 범위 체크
                    if end_row < len(rows_data) and end_col < max_cols:
                        end_cell = ppt_table.cell(end_row, end_col)
                        start_cell.merge(end_cell)
                        
                        # 머지된 셀 중앙 정렬
                        for paragraph in start_cell.text_frame.paragraphs:
                            paragraph.alignment = PP_ALIGN.CENTER
                except Exception as merge_err:
                    logger.debug(f"셀 머지 실패: {merge_err}")
            
        except Exception as e:
            logger.error(f"테이블 추가 실패: {e}")
    
    def _clean_text(self, text: str) -> str:
        """텍스트 정리 (불필요한 공백, 특수문자 제거)"""
        # 연속된 공백을 하나로
        text = re.sub(r'\s+', ' ', text)
        # 앞뒤 공백 제거
        text = text.strip()
        return text
    
    def _capture_element_screenshot(self, selector: str, title: str, index: int = 0) -> None:
        """
        Playwright를 사용하여 HTML 요소를 스크린샷으로 캡처하고 슬라이드에 추가
        
        Args:
            selector: CSS 선택자 (예: '.SeqViewerApp')
            title: 슬라이드 제목
            index: 여러 요소 중 캡처할 요소의 인덱스 (0부터 시작)
        """
        if not self.html_path:
            logger.warning(f"HTML 경로가 설정되지 않아 스크린샷을 캡처할 수 없습니다: {selector}")
            return
        
        try:
            from playwright.sync_api import sync_playwright
            import tempfile
            
            logger.info(f"Playwright로 요소 캡처 중: {selector} (index={index})")
            
            with sync_playwright() as p:
                # Chromium 브라우저 실행
                browser = p.chromium.launch(headless=True)
                
                # 페이지 생성 (넉넉한 뷰포트 크기)
                page = browser.new_page(viewport={'width': 1400, 'height': 900})
                
                # HTML 파일 로드
                file_url = f"file://{self.html_path}"
                page.goto(file_url, wait_until='networkidle')
                
                # JavaScript 렌더링 대기
                page.wait_for_timeout(2000)
                
                # 요소 찾기
                elements = page.locator(selector)
                count = elements.count()
                
                if count == 0:
                    logger.warning(f"요소를 찾을 수 없습니다: {selector}")
                    browser.close()
                    return
                
                if index >= count:
                    logger.warning(f"인덱스가 범위를 벗어났습니다: {index} >= {count}")
                    browser.close()
                    return
                
                # 특정 인덱스의 요소 선택
                element = elements.nth(index)
                
                # 요소가 보이도록 스크롤
                element.scroll_into_view_if_needed()
                page.wait_for_timeout(500)
                
                # 임시 파일에 스크린샷 저장
                with tempfile.NamedTemporaryFile(suffix='.png', delete=False) as tmp_file:
                    screenshot_path = tmp_file.name
                
                element.screenshot(path=screenshot_path)
                logger.info(f"스크린샷 저장 완료: {screenshot_path}")
                
                browser.close()
                
                # 스크린샷을 슬라이드에 추가
                self._create_screenshot_slide(screenshot_path, title)
                
                # 임시 파일 삭제
                import os
                os.unlink(screenshot_path)
                
        except ImportError:
            logger.error("Playwright가 설치되지 않았습니다. 'pip install playwright && playwright install chromium'을 실행하세요.")
        except Exception as e:
            logger.error(f"요소 스크린샷 캡처 실패: {e}")
    
    def _create_screenshot_slide(self, image_path: str, title: str) -> None:
        """
        스크린샷 이미지를 슬라이드에 추가
        
        Args:
            image_path: 이미지 파일 경로
            title: 슬라이드 제목
        """
        try:
            from PIL import Image
            
            # 새 슬라이드 추가
            blank_layout = self.prs.slide_layouts[6]  # 빈 레이아웃
            slide = self.prs.slides.add_slide(blank_layout)
            
            # 제목 추가
            title_top = Inches(0.3)
            title_shape = slide.shapes.add_textbox(
                self.margin_left,
                title_top,
                self.content_width,
                Inches(0.5)
            )
            tf = title_shape.text_frame
            tf.word_wrap = True
            p = tf.paragraphs[0]
            p.text = title
            p.font.size = Pt(18)
            p.font.bold = True
            p.font.color.rgb = self.colors['gray_900']
            
            # 이미지 크기 확인
            with Image.open(image_path) as img:
                img_width, img_height = img.size
            
            # 사용 가능한 영역 계산
            content_top = Inches(1.0)
            available_width = self.content_width
            available_height = self.slide_height - content_top - self.margin_bottom
            
            # 비율 유지하며 크기 조정
            width_ratio = available_width / img_width
            height_ratio = available_height / img_height
            scale = min(width_ratio, height_ratio)
            
            # 최종 크기 (EMU로 변환)
            final_width = int(img_width * scale)
            final_height = int(img_height * scale)
            
            # 중앙 정렬
            left = self.margin_left + (available_width - final_width) / 2
            top = content_top + (available_height - final_height) / 2
            
            # 이미지 추가
            slide.shapes.add_picture(
                image_path,
                left,
                top,
                final_width,
                final_height
            )
            
            logger.info(f"스크린샷 슬라이드 추가 완료: {title}")
            
        except Exception as e:
            logger.error(f"스크린샷 슬라이드 생성 실패: {e}")


def convert_html_to_pptx(html_path: Path, output_path: Path) -> None:
    """
    HTML 파일을 PPTX로 변환하는 편의 함수
    
    Args:
        html_path: 입력 HTML 파일 경로
        output_path: 출력 PPTX 파일 경로
    
    Example:
        >>> from pathlib import Path
        >>> html_path = Path("input.html")
        >>> output_path = Path("output.pptx")
        >>> convert_html_to_pptx(html_path, output_path)
    """
    converter = HtmlToPptxConverter()
    converter.convert(html_path, output_path)


if __name__ == "__main__":
    # 테스트 실행
    import sys
    
    if len(sys.argv) < 3:
        print("사용법: python html_to_pptx.py <input.html> <output.pptx>")
        sys.exit(1)
    
    input_path = Path(sys.argv[1])
    output_path = Path(sys.argv[2])
    
    convert_html_to_pptx(input_path, output_path)
    print(f"변환 완료: {output_path}")
