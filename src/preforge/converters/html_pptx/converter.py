"""
HTML(.html) 문서를 PowerPoint(.pptx)로 변환하는 컨버터

모듈화된 구조로 HTML 구조를 분석하여 섹션별로 슬라이드를 생성합니다.
"""
import logging
import tempfile
import os
from pathlib import Path
from typing import List, Optional

from bs4 import BeautifulSoup, Tag
from pptx import Presentation
from pptx.util import Inches, Pt

from .config import (
    SlideConfig, 
    TableConfig, 
    ColorPalette,
    DEFAULT_SLIDE_CONFIG,
    DEFAULT_TABLE_CONFIG,
    DEFAULT_COLORS
)
from .style_utils import TextUtils
from .table_builder import TableBuilder, TableDataExtractor
from .slide_factory import (
    TitleSlideBuilder,
    ContentSlideBuilder,
    TableSlideBuilder,
    ImageSlideBuilder,
    EvidenceSlideBuilder
)

logger = logging.getLogger(__name__)


class HtmlToPptxConverter:
    """HTML을 PowerPoint로 변환하는 컨버터"""
    
    def __init__(
        self,
        slide_config: SlideConfig = None,
        table_config: TableConfig = None,
        colors: ColorPalette = None
    ):
        """
        컨버터 초기화
        
        Args:
            slide_config: 슬라이드 레이아웃 설정
            table_config: 테이블 설정
            colors: 색상 팔레트
        """
        self.slide_config = slide_config or DEFAULT_SLIDE_CONFIG
        self.table_config = table_config or DEFAULT_TABLE_CONFIG
        self.colors = colors or DEFAULT_COLORS
        
        self.prs: Optional[Presentation] = None
        self.html_path: Optional[Path] = None
        
        # 슬라이드 빌더들 (convert 시 초기화)
        self._title_builder: Optional[TitleSlideBuilder] = None
        self._content_builder: Optional[ContentSlideBuilder] = None
        self._table_builder: Optional[TableSlideBuilder] = None
        self._image_builder: Optional[ImageSlideBuilder] = None
        self._evidence_builder: Optional[EvidenceSlideBuilder] = None
    
    def convert(self, html_path: Path, output_path: Path) -> None:
        """
        HTML 파일을 PPTX로 변환
        
        Args:
            html_path: 입력 HTML 파일 경로
            output_path: 출력 PPTX 파일 경로
        """
        logger.info(f"HTML -> PPTX 변환 시작: {html_path} -> {output_path}")
        
        self.html_path = Path(html_path).absolute()
        
        # HTML 파일 읽기
        with open(html_path, 'r', encoding='utf-8') as f:
            html_content = f.read()
        
        soup = BeautifulSoup(html_content, 'lxml')
        
        # 프레젠테이션 초기화
        self.prs = Presentation()
        self.prs.slide_width = self.slide_config.width
        self.prs.slide_height = self.slide_config.height
        
        # 슬라이드 빌더 초기화
        self._init_builders()
        
        # 슬라이드 생성
        self._create_title_slide(soup)
        self._create_analysis_summary_slides(soup)
        self._process_main_content(soup)
        
        # 저장
        self.prs.save(str(output_path))
        logger.info(f"변환 완료: {output_path} (총 {len(self.prs.slides)}개 슬라이드)")
    
    def _init_builders(self) -> None:
        """슬라이드 빌더 초기화"""
        self._title_builder = TitleSlideBuilder(
            self.prs, self.slide_config, self.colors
        )
        self._content_builder = ContentSlideBuilder(
            self.prs, self.slide_config, self.colors
        )
        self._table_builder = TableSlideBuilder(
            self.prs, self.slide_config, self.colors,
            self.table_config.max_rows_per_slide
        )
        self._image_builder = ImageSlideBuilder(
            self.prs, self.slide_config, self.colors
        )
        self._evidence_builder = EvidenceSlideBuilder(
            self.prs, self.slide_config, self.colors
        )
    
    def _create_title_slide(self, soup: BeautifulSoup) -> None:
        """타이틀 슬라이드 생성"""
        title_elem = soup.find('div', class_='header-title')
        subtitle_elem = soup.find('div', class_='header-subtitle')
        
        title = title_elem.get_text(strip=True) if title_elem else "GeneSeq Vista AI Agent"
        subtitle = subtitle_elem.get_text(strip=True) if subtitle_elem else ""
        
        self._title_builder.create(title, subtitle)
    
    def _create_analysis_summary_slides(self, soup: BeautifulSoup) -> None:
        """Analysis Summary 섹션 슬라이드 생성"""
        analysis_div = soup.find('div', class_='analysis-summary')
        if not analysis_div:
            return
        
        summary_sections = analysis_div.find_all('div', class_='summary-section')
        
        for idx, section in enumerate(summary_sections):
            header = section.find('div', class_='section-header')
            header_text = header.get_text(strip=True) if header else f"Summary {idx + 1}"
            
            table_elem = section.find('table')
            if table_elem:
                self._table_builder.create_from_html(table_elem, header_text)
    
    def _process_main_content(self, soup: BeautifulSoup) -> None:
        """메인 컨텐츠 처리"""
        content_container = soup.find('div', class_='content-container')
        if not content_container:
            return
        
        gene_title_elem = content_container.find('h1', class_='gene-title')
        main_title = gene_title_elem.get_text(strip=True) if gene_title_elem else "Gene Analysis"
        
        gene_sections = content_container.find_all('div', class_='gene-section')
        seq_viewer_index = 0
        
        for idx, gene_section in enumerate(gene_sections, 1):
            section_title = self._get_section_title(gene_section, idx)
            
            # SeqViewerApp 스크린샷 캡처
            seq_viewers = gene_section.find_all('div', class_='SeqViewerApp', recursive=False)
            for sv_idx, _ in enumerate(seq_viewers):
                viewer_title = f"{section_title} - Sequence Viewer"
                if len(seq_viewers) > 1:
                    viewer_title += f" ({sv_idx + 1})"
                self._capture_element_screenshot('.SeqViewerApp', viewer_title, seq_viewer_index)
                seq_viewer_index += 1
            
            # 이미지 처리
            self._process_images(gene_section, section_title, main_title)
            
            # 테이블 처리
            self._process_tables(gene_section, section_title, main_title)
            
            # 하위 섹션 처리
            self._process_subsections(gene_section, main_title)
        
        # gene-section 외부의 독립적인 h3 섹션 처리 (3.3, 3.4 등)
        self._process_standalone_h3_sections(content_container, main_title)
        
        # Evidence 테이블 처리
        self._process_evidence_tables(content_container)
    
    def _get_section_title(self, gene_section: Tag, default_idx: int) -> str:
        """섹션 제목 추출"""
        subsection_title = gene_section.find('h2', class_='subsection-title')
        if subsection_title:
            return subsection_title.get_text(strip=True)
        return f"Section {default_idx}"
    
    def _process_images(self, gene_section: Tag, section_title: str, main_title: str) -> None:
        """이미지 처리"""
        image_placeholder = gene_section.find('div', class_='image-placeholder')
        if image_placeholder:
            img = image_placeholder.find('img')
            if img and img.get('src', '').startswith('data:image'):
                self._image_builder.create_from_base64(img, section_title)
    
    def _process_tables(self, gene_section: Tag, section_title: str, main_title: str) -> None:
        """테이블 처리 (동적 그룹화)"""
        tables = gene_section.find_all('table', class_='data-table')
        if not tables:
            return
        
        table_infos = [
            {'table': t, 'rows': len(t.find_all('tr'))} 
            for t in tables
        ]
        
        # 동적 그룹화
        title_space = Inches(0.85)
        table_gap = Inches(0.3)
        available_height = (
            self.slide_config.height - 
            self.slide_config.margin_top - 
            self.slide_config.margin_bottom - 
            title_space
        )
        row_height = Inches(0.28)
        
        i = 0
        while i < len(table_infos):
            current_group = [table_infos[i]]
            current_height = row_height * table_infos[i]['rows']
            
            while i + 1 < len(table_infos):
                next_rows = table_infos[i + 1]['rows']
                next_table_height = row_height * next_rows + table_gap
                remaining_space = available_height - current_height
                
                if next_table_height <= remaining_space:
                    i += 1
                    current_group.append(table_infos[i])
                    current_height += next_table_height
                else:
                    break
            
            if len(current_group) == 1:
                self._table_builder.create_from_html(
                    current_group[0]['table'], section_title, main_title
                )
            else:
                self._create_combined_table_slide(
                    [info['table'] for info in current_group],
                    section_title, main_title
                )
            
            i += 1
    
    def _create_combined_table_slide(
        self, 
        tables: List[Tag], 
        section_title: str, 
        main_title: str
    ) -> None:
        """여러 테이블을 하나의 슬라이드에 합침"""
        slide = self.prs.slides.add_slide(self.prs.slide_layouts[6])
        
        # 제목 추가
        if main_title:
            main_box = slide.shapes.add_textbox(
                self.slide_config.margin_left, Inches(0.1),
                self.slide_config.content_width, Inches(0.25)
            )
            main_frame = main_box.text_frame
            main_frame.text = main_title
            main_frame.paragraphs[0].font.size = Pt(12)
            main_frame.paragraphs[0].font.color.rgb = self.colors['gray_600']
        
        title_top = Inches(0.35) if main_title else self.slide_config.margin_top - Inches(0.1)
        title_box = slide.shapes.add_textbox(
            self.slide_config.margin_left, title_top,
            self.slide_config.content_width, Inches(0.4)
        )
        title_frame = title_box.text_frame
        title_frame.text = section_title
        title_para = title_frame.paragraphs[0]
        title_para.font.size = Pt(18)
        title_para.font.bold = True
        title_para.font.color.rgb = self.colors['primary_red']
        
        # 테이블 추가
        current_top = title_top + Inches(0.5)
        available_height = self.slide_config.height - current_top - self.slide_config.margin_bottom
        table_builder = TableBuilder(colors=self.colors)
        
        for table_idx, table_elem in enumerate(tables):
            extractor = TableDataExtractor(table_elem).extract()
            rows = len(extractor.rows_data)
            
            table_height = min(Inches(0.25) * rows, available_height * 0.4)
            
            if table_idx > 0:
                current_top += Inches(0.2)
            
            table_builder.create_table(
                slide,
                extractor.rows_data,
                len(extractor.header_rows),
                extractor.col_widths_html,
                self.slide_config.margin_left, current_top,
                self.slide_config.content_width, table_height,
                extractor.merge_info,
                extractor.cell_styles
            )
            
            current_top += table_height + Inches(0.1)
    
    def _process_subsections(self, gene_section: Tag, main_title: str) -> None:
        """하위 섹션(h3) 처리 - 옵션 A: 모든 h3 제목과 테이블을 하나의 슬라이드에"""
        h3_sections = gene_section.find_all('h3')
        if not h3_sections:
            return
        
        # h3와 관련 테이블을 수집
        h3_table_pairs = []
        for h3 in h3_sections:
            h3_title = h3.get_text(strip=True)
            next_table = h3.find_next('table')
            if next_table:
                # 테이블이 현재 gene_section 내에 있는지 확인
                table_parent = next_table.parent
                while table_parent and table_parent != gene_section:
                    table_parent = table_parent.parent
                if table_parent == gene_section:
                    h3_table_pairs.append({'h3_title': h3_title, 'table': next_table})
        
        if not h3_table_pairs:
            return
        
        # 하나의 슬라이드에 모든 h3와 테이블을 표시 (옵션 A)
        self._create_h3_combined_slide(h3_table_pairs, main_title)
    
    def _create_h3_combined_slide(
        self,
        h3_table_pairs: List[dict],
        main_title: str
    ) -> None:
        """h3 제목들과 테이블들을 하나의 슬라이드에 합침"""
        slide = self.prs.slides.add_slide(self.prs.slide_layouts[6])
        
        # 메인 타이틀
        if main_title:
            main_box = slide.shapes.add_textbox(
                self.slide_config.margin_left, Inches(0.1),
                self.slide_config.content_width, Inches(0.25)
            )
            main_frame = main_box.text_frame
            main_frame.text = main_title
            main_frame.paragraphs[0].font.size = Pt(12)
            main_frame.paragraphs[0].font.color.rgb = self.colors['gray_600']
        
        current_top = Inches(0.35) if main_title else self.slide_config.margin_top
        table_builder = TableBuilder(colors=self.colors)
        
        # 각 h3+테이블 추가
        for idx, pair in enumerate(h3_table_pairs):
            h3_title = pair['h3_title']
            table_elem = pair['table']
            
            # h3 제목 추가
            h3_box = slide.shapes.add_textbox(
                self.slide_config.margin_left, current_top,
                self.slide_config.content_width, Inches(0.3)
            )
            h3_frame = h3_box.text_frame
            h3_frame.text = h3_title
            h3_para = h3_frame.paragraphs[0]
            h3_para.font.size = Pt(14)
            h3_para.font.bold = True
            h3_para.font.color.rgb = self.colors['primary_red']
            
            current_top += Inches(0.35)
            
            # 테이블 추가
            extractor = TableDataExtractor(table_elem).extract()
            rows = len(extractor.rows_data)
            
            # 남은 공간 계산
            remaining_height = self.slide_config.height - current_top - self.slide_config.margin_bottom
            # 최대 높이 제한
            table_height = min(Inches(0.25) * rows, remaining_height * 0.4)
            
            table_builder.create_table(
                slide,
                extractor.rows_data,
                len(extractor.header_rows),
                extractor.col_widths_html,
                self.slide_config.margin_left, current_top,
                self.slide_config.content_width, table_height,
                extractor.merge_info,
                extractor.cell_styles
            )
            
            current_top += table_height + Inches(0.3)
    
    def _process_standalone_h3_sections(self, content_container: Tag, main_title: str) -> None:
        """
        gene-section 외부에 있는 독립적인 h3 섹션 처리 (예: 3.3, 3.4 섹션)
        content-container의 직접 자식인 h3 요소를 찾아 처리
        """
        # content-container의 직접 자식 중 h3 요소 찾기
        standalone_h3_elements = content_container.find_all('h3', recursive=False)
        
        for h3 in standalone_h3_elements:
            h3_title = h3.get_text(strip=True)
            
            # h3 다음에 오는 테이블 또는 reference-card 찾기
            next_sibling = h3.find_next_sibling()
            
            if next_sibling:
                # 테이블 처리 (예: 3.3 주요기관 권고 현황)
                if next_sibling.name == 'table':
                    self._table_builder.create_from_html(next_sibling, h3_title, main_title)
                
                # reference-card 처리 (예: 3.4 주요 문헌)
                elif next_sibling.name == 'div' and 'reference-card' in next_sibling.get('class', []):
                    self._create_reference_slide(next_sibling, h3_title, main_title)
    
    def _create_reference_slide(self, reference_div: Tag, title: str, main_title: str = "") -> None:
        """Reference card 슬라이드 생성"""
        slide = self.prs.slides.add_slide(self.prs.slide_layouts[6])
        
        # 메인 타이틀
        if main_title:
            main_box = slide.shapes.add_textbox(
                self.slide_config.margin_left, Inches(0.1),
                self.slide_config.content_width, Inches(0.25)
            )
            main_frame = main_box.text_frame
            main_frame.text = main_title
            main_frame.paragraphs[0].font.size = Pt(12)
            main_frame.paragraphs[0].font.color.rgb = self.colors['gray_600']
        
        # 섹션 타이틀
        title_top = Inches(0.35) if main_title else self.slide_config.margin_top
        title_box = slide.shapes.add_textbox(
            self.slide_config.margin_left, title_top,
            self.slide_config.content_width, Inches(0.4)
        )
        title_frame = title_box.text_frame
        title_frame.text = title
        title_para = title_frame.paragraphs[0]
        title_para.font.size = Pt(18)
        title_para.font.bold = True
        title_para.font.color.rgb = self.colors['primary_red']
        
        # Reference 내용 추출
        content_top = title_top + Inches(0.5)
        references = reference_div.find_all('div', class_='reference-item')
        
        if references:
            content_text = []
            for idx, ref in enumerate(references, 1):
                ref_text = ref.get_text(strip=True)
                content_text.append(f"{idx}. {ref_text}")
            
            content = "\n\n".join(content_text)
        else:
            # reference-item이 없으면 전체 텍스트 추출
            content = reference_div.get_text(strip=True)
        
        # 내용 텍스트박스
        content_box = slide.shapes.add_textbox(
            self.slide_config.margin_left, content_top,
            self.slide_config.content_width, 
            self.slide_config.height - content_top - self.slide_config.margin_bottom
        )
        content_frame = content_box.text_frame
        content_frame.word_wrap = True
        content_frame.text = content
        
        for paragraph in content_frame.paragraphs:
            paragraph.font.size = Pt(10)
            paragraph.font.color.rgb = self.colors['gray_800']
            paragraph.line_spacing = 1.3
    
    def _process_evidence_tables(self, content_container: Tag) -> None:
        """Evidence 테이블 처리"""
        all_subsections = content_container.find_all('h2', class_='subsection-title')
        
        for subsection in all_subsections:
            section_title = subsection.get_text(strip=True)
            
            next_elem = subsection.find_next_sibling()
            while next_elem:
                if next_elem.name == 'h2':
                    break
                if next_elem.get('class') and 'gene-section' in next_elem.get('class', []):
                    break
                
                if next_elem.name == 'div' and 'evidence-table' in next_elem.get('class', []):
                    self._evidence_builder.create(next_elem, section_title)
                
                next_elem = next_elem.find_next_sibling()
    
    def _capture_element_screenshot(self, selector: str, title: str, index: int = 0) -> None:
        """Playwright를 사용하여 HTML 요소 스크린샷 캡처"""
        if not self.html_path:
            logger.warning(f"HTML 경로가 설정되지 않아 스크린샷을 캡처할 수 없습니다")
            return
        
        try:
            from playwright.sync_api import sync_playwright
            
            logger.info(f"Playwright로 요소 캡처 중: {selector} (index={index})")
            
            with sync_playwright() as p:
                browser = p.chromium.launch(headless=True)
                page = browser.new_page(viewport={'width': 1400, 'height': 900})
                
                file_url = f"file://{self.html_path}"
                page.goto(file_url, wait_until='networkidle')
                page.wait_for_timeout(2000)
                
                elements = page.locator(selector)
                count = elements.count()
                
                if count == 0 or index >= count:
                    browser.close()
                    return
                
                element = elements.nth(index)
                element.scroll_into_view_if_needed()
                page.wait_for_timeout(500)
                
                with tempfile.NamedTemporaryFile(suffix='.png', delete=False) as tmp_file:
                    screenshot_path = tmp_file.name
                
                element.screenshot(path=screenshot_path)
                browser.close()
                
                self._image_builder.create_from_file(screenshot_path, title)
                os.unlink(screenshot_path)
                
        except ImportError:
            logger.error("Playwright가 설치되지 않았습니다.")
        except Exception as e:
            logger.error(f"요소 스크린샷 캡처 실패: {e}")


def convert_html_to_pptx(html_path: Path, output_path: Path) -> None:
    """
    HTML 파일을 PPTX로 변환하는 편의 함수
    
    Args:
        html_path: 입력 HTML 파일 경로
        output_path: 출력 PPTX 파일 경로
    """
    converter = HtmlToPptxConverter()
    converter.convert(html_path, output_path)


if __name__ == "__main__":
    import sys
    
    if len(sys.argv) < 3:
        print("사용법: python converter.py <input.html> <output.pptx>")
        sys.exit(1)
    
    input_path = Path(sys.argv[1])
    output_path = Path(sys.argv[2])
    
    convert_html_to_pptx(input_path, output_path)
    print(f"변환 완료: {output_path}")
