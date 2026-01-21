"""
PowerPoint(.pptx)를 Word(.docx)로 변환하는 컨버터

슬라이드 구조를 유지하면서 Word 문서 형식으로 변환합니다.
전처리 파싱을 통해 슬라이드 구조를 분석하고, 제목/목차/본문을 구분합니다.
"""
from pathlib import Path
from typing import List, Optional, Dict, Any, Tuple
from io import BytesIO
from dataclasses import dataclass, field
import logging
import re

try:
    from PIL import Image
    HAS_PIL = True
except ImportError:
    HAS_PIL = False

from docx import Document as DocxDocument
from docx.shared import Inches, Pt, Cm, RGBColor, Twips
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.style import WD_STYLE_TYPE
from docx.enum.section import WD_ORIENT
from docx.oxml.ns import qn, nsmap
from docx.oxml import OxmlElement
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.util import Emu

logger = logging.getLogger(__name__)

# XML 호환되지 않는 제어 문자 패턴
INVALID_XML_CHARS_RE = re.compile(r'[\x00-\x08\x0b\x0c\x0e-\x1f\x7f-\x9f]')

# 특수 기호 대체 맵핑
SPECIAL_CHAR_MAP = {
    '\uf0d8': '▶',  # 화살표
    '\uf0fc': '✓',  # 체크마크
    '\uf0a7': '•',  # 불릿
    '\uf0b7': '•',  # 불릿
    '\uf076': '★',  # 별
    '\uf0e0': '→',  # 화살표
    '\uf0e8': '●',  # 원
    '\uf0ab': '◆',  # 다이아몬드
    '\uf0a8': '○',  # 빈 원
    '\uf02d': '–',  # 대시
    '\uf0b2': '■',  # 사각형
    '\uf06c': '◎',  # 이중 원
    '\uf0d7': '▼',  # 아래 화살표
    '\uf0de': '▲',  # 위 화살표
    '\uf0a0': ' ',  # 특수 공백
    '\uf020': ' ',  # 특수 공백2
    '\uf06e': '♦',  # 다이아몬드2
}

# 강조할 키워드 목록
HIGHLIGHT_KEYWORDS = [
    'Pathogen', 'Pathogens',
    'Symptom', 'Symptoms', 
    'Diagnosis', 'Diagnostic',
    'Treatment', 'Treatments',
    'Disease', 'Diseases',
    'Epidemiology',
    'Prevention',
    'Transmission',
    'Clinical Features',
]


@dataclass
class GridCell:
    """그리드 셀 정보"""
    row: int
    col: int
    rowspan: int = 1
    colspan: int = 1
    content_type: str = 'empty'  # 'text', 'table', 'image', 'mixed', 'empty'
    left: int = 0  # EMU 단위 좌표
    top: int = 0
    width: int = 0
    height: int = 0
    shapes: List[Any] = field(default_factory=list)  # 셀에 포함된 shape들


@dataclass
class GridLayout:
    """슬라이드 그리드 레이아웃"""
    rows: int = 1
    cols: int = 1
    cells: List[GridCell] = field(default_factory=list)
    row_heights: List[int] = field(default_factory=list)  # EMU 단위
    col_widths: List[int] = field(default_factory=list)  # EMU 단위
    
    def get_cell(self, row: int, col: int) -> Optional[GridCell]:
        """특정 위치의 셀 반환"""
        for cell in self.cells:
            if cell.row == row and cell.col == col:
                return cell
        return None
    
    def to_dict(self) -> Dict[str, Any]:
        """그리드 정보를 딕셔너리로 변환 (파싱 결과 저장용)"""
        return {
            'rows': self.rows,
            'cols': self.cols,
            'row_heights': self.row_heights,
            'col_widths': self.col_widths,
            'cells': [
                {
                    'row': cell.row,
                    'col': cell.col,
                    'rowspan': cell.rowspan,
                    'colspan': cell.colspan,
                    'content_type': cell.content_type,
                    'shape_count': len(cell.shapes),
                }
                for cell in self.cells
            ]
        }


@dataclass
class SlideContent:
    """파싱된 슬라이드 콘텐츠"""
    slide_index: int
    slide_type: str  # 'title', 'toc', 'content'
    slide: Any = None  # 슬라이드 객체 참조 (테이블 내 이미지 처리용)
    title: Optional[str] = None
    subtitle: Optional[str] = None
    author: Optional[str] = None
    date: Optional[str] = None
    texts: List[Dict[str, Any]] = field(default_factory=list)
    tables: List[Dict[str, Any]] = field(default_factory=list)
    images: List[Dict[str, Any]] = field(default_factory=list)
    toc_items: List[str] = field(default_factory=list)
    section_title: Optional[str] = None  # 목차의 대제목 (예: "1. 질병")
    grid_layout: Optional[GridLayout] = None  # 그리드 레이아웃 정보


@dataclass
class ParsedPresentation:
    """파싱된 프레젠테이션 전체 구조"""
    title: Optional[str] = None
    author: Optional[str] = None
    date: Optional[str] = None
    slides: List[SlideContent] = field(default_factory=list)
    toc_slides: List[int] = field(default_factory=list)  # 목차 슬라이드 인덱스
    section_titles: Dict[int, str] = field(default_factory=dict)  # 슬라이드 인덱스별 섹션 제목


def sanitize_text(text: str) -> str:
    """
    XML 호환되지 않는 문자를 제거하고 특수 기호를 대체합니다.
    """
    if not text:
        return text
    
    # 특수 기호 대체
    for old_char, new_char in SPECIAL_CHAR_MAP.items():
        text = text.replace(old_char, new_char)
    
    # 제어 문자 제거
    text = INVALID_XML_CHARS_RE.sub('', text)
    
    return text


def is_highlight_keyword(text: str) -> bool:
    """강조해야 할 키워드인지 확인 (줄바꿈 제거 후 검사)"""
    # 줄바꿈을 공백으로 대체하여 확인
    text_normalized = ' '.join(text.split()).lower()
    for keyword in HIGHLIGHT_KEYWORDS:
        if keyword.lower() in text_normalized:
            return True
    return False


def normalize_text_for_highlighting(text: str) -> str:
    """강조 처리를 위해 텍스트 정규화 (줄바꿈 제거)"""
    return ' '.join(text.split())


def is_page_number(text: str) -> bool:
    """
    텍스트가 페이지 번호인지 확인
    
    페이지 번호 패턴:
    - 숫자만 있는 경우 (1, 2, 3, ...)
    - 짧은 숫자 (1-3자리)
    """
    text = text.strip()
    if not text:
        return False
    
    # 숫자만 있고 1-3자리인 경우 페이지 번호로 간주
    if text.isdigit() and len(text) <= 3:
        return True
    
    return False


class PptxToDocxConverter:
    """PowerPoint를 Word 문서로 변환하는 클래스"""
    
    def __init__(
        self,
        include_images: bool = True,
        include_tables: bool = True,
        include_notes: bool = False,
        landscape_after_toc: bool = True,
        image_max_width_inches: float = 8.0,
        highlight_keywords: bool = True,
    ):
        """
        Args:
            include_images: 이미지 포함 여부
            include_tables: 테이블 포함 여부
            include_notes: 발표자 노트 포함 여부
            landscape_after_toc: 목차 이후 가로 레이아웃 적용
            image_max_width_inches: 이미지 최대 너비 (인치)
            highlight_keywords: 키워드 강조 여부
        """
        self.include_images = include_images
        self.include_tables = include_tables
        self.include_notes = include_notes
        self.landscape_after_toc = landscape_after_toc
        self.image_max_width_inches = image_max_width_inches
        self.highlight_keywords = highlight_keywords
        
        # 현재 섹션 제목 추적
        self._current_section_title = None
        self._processed_section_titles = set()
    
    def convert(
        self,
        pptx_path: Path,
        output_path: Optional[Path] = None,
    ) -> Path:
        """
        PPTX 파일을 DOCX로 변환
        """
        pptx_path = Path(pptx_path)
        
        if not pptx_path.exists():
            raise FileNotFoundError(f"파일을 찾을 수 없습니다: {pptx_path}")
        
        if pptx_path.suffix.lower() not in [".pptx", ".ppt"]:
            raise ValueError(f"지원하지 않는 파일 형식: {pptx_path.suffix}")
        
        if output_path is None:
            output_path = pptx_path.with_suffix(".docx")
        else:
            output_path = Path(output_path)
        
        # 상태 초기화
        self._processed_section_titles = set()
        
        # 1단계: PPTX 전처리 파싱
        prs = Presentation(pptx_path)
        parsed = self._preprocess_presentation(prs)
        
        # 2단계: DOCX 생성
        doc = DocxDocument()
        self._setup_document_styles(doc)
        self._copy_metadata(prs, doc)
        
        # 3단계: 콘텐츠 변환
        self._convert_parsed_content(doc, parsed, prs)
        
        # 저장
        doc.save(output_path)
        logger.info(f"변환 완료: {pptx_path} -> {output_path}")
        
        return output_path
    
    def _preprocess_presentation(self, prs: Presentation) -> ParsedPresentation:
        """프레젠테이션 전처리 파싱"""
        parsed = ParsedPresentation()
        
        for slide_idx, slide in enumerate(prs.slides, 1):
            slide_content = self._parse_slide(slide, slide_idx)
            parsed.slides.append(slide_content)
            
            # 제목 슬라이드 정보 저장
            if slide_content.slide_type == 'title':
                parsed.title = slide_content.title
                parsed.author = slide_content.author
                parsed.date = slide_content.date
            
            # 목차 슬라이드 인덱스 저장
            if slide_content.slide_type == 'toc':
                parsed.toc_slides.append(slide_idx)
            
            # 섹션 제목 저장
            if slide_content.section_title:
                parsed.section_titles[slide_idx] = slide_content.section_title
        
        return parsed
    
    def _parse_slide(self, slide: Any, slide_idx: int) -> SlideContent:
        """단일 슬라이드 파싱"""
        title_text = self._get_slide_title(slide)
        
        # 슬라이드 타입 결정
        slide_type = self._determine_slide_type(slide, slide_idx, title_text)
        
        content = SlideContent(
            slide_index=slide_idx,
            slide_type=slide_type,
            slide=slide,  # 슬라이드 참조 저장
            title=title_text,
        )
        
        if slide_type == 'title':
            self._parse_title_slide(slide, content)
        elif slide_type == 'toc':
            self._parse_toc_slide(slide, content)
        else:
            self._parse_content_slide(slide, content)
        
        return content
    
    def _determine_slide_type(
        self, 
        slide: Any, 
        slide_idx: int, 
        title_text: Optional[str]
    ) -> str:
        """슬라이드 타입 결정"""
        # 첫 번째 슬라이드는 보통 제목
        if slide_idx == 1:
            return 'title'
        
        # 제목에 '목차'가 포함되면 목차
        if title_text:
            title_lower = title_text.lower()
            if '목차' in title_lower or 'contents' in title_lower or 'index' in title_lower:
                return 'toc'
        
        # 제목이 없거나 두 번째 슬라이드인 경우, 본문에서 목차 키워드 확인
        if slide_idx == 2 or not title_text:
            all_text = self._get_all_slide_text(slide).lower()
            if '[목차]' in all_text or '목차' in all_text[:20]:
                return 'toc'
        
        return 'content'
    
    def _get_all_slide_text(self, slide: Any) -> str:
        """슬라이드의 모든 텍스트 추출"""
        texts = []
        for shape in slide.shapes:
            if hasattr(shape, 'text_frame'):
                texts.append(shape.text_frame.text)
            elif hasattr(shape, 'text'):
                texts.append(shape.text)
        return ' '.join(texts)
    
    def _parse_title_slide(self, slide: Any, content: SlideContent):
        """제목 슬라이드 파싱"""
        texts = []
        
        for shape in slide.shapes:
            if not hasattr(shape, 'text_frame'):
                continue
            
            text = sanitize_text(shape.text_frame.text.strip())
            if not text:
                continue
            
            # 제목은 이미 추출됨
            if slide.shapes.title and shape == slide.shapes.title:
                continue
            
            texts.append(text)
        
        # 텍스트에서 날짜, 작성자 추출 시도
        for text in texts:
            # 날짜 패턴 (예: 2020.10.27, 2020-10-27)
            date_pattern = r'\d{4}[.\-/]\s?\d{1,2}[.\-/]\s?\d{1,2}'
            if re.search(date_pattern, text):
                content.date = text
            # 팀/작성자 패턴
            elif '팀' in text or 'Team' in text or any(c in text for c in ['김', '이', '박', '최', '정']):
                content.author = text
            else:
                content.subtitle = text
    
    def _parse_toc_slide(self, slide: Any, content: SlideContent):
        """목차 슬라이드 파싱"""
        toc_items = []
        
        for shape in slide.shapes:
            if slide.shapes.title and shape == slide.shapes.title:
                continue
            
            if hasattr(shape, 'text_frame'):
                for para in shape.text_frame.paragraphs:
                    text = sanitize_text(para.text.strip())
                    if text and len(text) > 1:
                        # "[목차]" 텍스트는 제목으로 사용하고 목록에서 제외
                        text_lower = text.lower()
                        if '[목차]' in text_lower or text_lower == '목차':
                            if not content.title:
                                content.title = text
                        else:
                            toc_items.append(text)
        
        # 제목이 없으면 기본값 설정
        if not content.title:
            content.title = "[목차]"
        
        content.toc_items = toc_items
    
    def _parse_content_slide(self, slide: Any, content: SlideContent):
        """일반 콘텐츠 슬라이드 파싱"""
        title_shape = slide.shapes.title
        
        # 섹션 제목 추출 (목차의 대제목과 매칭)
        if content.title:
            # 숫자로 시작하는 제목은 섹션 제목으로 간주
            if re.match(r'^\d+[\.\s]', content.title):
                content.section_title = content.title
        
        # shape들 수집
        shapes_data = []
        for shape in slide.shapes:
            if title_shape and shape == title_shape:
                continue
            
            shape_info = self._get_shape_info(shape)
            if shape_info:
                shapes_data.append(shape_info)
        
        # 위치 정렬
        shapes_data.sort(key=lambda x: (x['top'], x['left']))
        
        for shape_info in shapes_data:
            if shape_info['type'] == 'text':
                content.texts.append(shape_info)
            elif shape_info['type'] == 'table':
                content.tables.append(shape_info)
            elif shape_info['type'] == 'image':
                content.images.append(shape_info)
            elif shape_info['type'] == 'group':
                # 그룹 내부 처리
                self._parse_group_shape(shape_info['shape'], content)
        
        # 그리드 레이아웃 분석
        content.grid_layout = self._analyze_grid_layout(slide, title_shape)
    
    def _parse_group_shape(self, group_shape: Any, content: SlideContent):
        """그룹 shape 파싱"""
        for sub_shape in group_shape.shapes:
            shape_info = self._get_shape_info(sub_shape)
            if shape_info:
                if shape_info['type'] == 'text':
                    content.texts.append(shape_info)
                elif shape_info['type'] == 'table':
                    content.tables.append(shape_info)
                elif shape_info['type'] == 'image':
                    content.images.append(shape_info)
                elif shape_info['type'] == 'group':
                    self._parse_group_shape(sub_shape, content)
    
    def _analyze_grid_layout(
        self, 
        slide: Any, 
        title_shape: Any = None
    ) -> GridLayout:
        """
        슬라이드의 그리드 레이아웃을 분석합니다.
        
        shape들의 위치를 기반으로 행/열 구조를 파악하고
        각 셀에 어떤 콘텐츠가 있는지 매핑합니다.
        
        Args:
            slide: 슬라이드 객체
            title_shape: 제목 shape (그리드에서 제외)
            
        Returns:
            GridLayout: 분석된 그리드 레이아웃
        """
        # 모든 shape 수집 (제목 제외)
        shapes = []
        for shape in slide.shapes:
            if title_shape and shape == title_shape:
                continue
            
            # 로고, Confidential 등 슬라이드 우측 상단 요소는 제외
            shape_info = self._extract_shape_bounds(shape)
            if shape_info:
                shapes.append(shape_info)
        
        if not shapes:
            return GridLayout(rows=1, cols=1)
        
        # 1. Y좌표 기준으로 행 경계 찾기
        y_coords = sorted(set(
            [s['top'] for s in shapes] + [s['bottom'] for s in shapes]
        ))
        
        # 슬라이드 크기 기반 동적 임계값 계산 (전체 크기의 ~15%)
        # 일반적인 슬라이드: 9144000 x 6858000 EMU (10" x 7.5")
        y_range = max(y_coords) - min(y_coords) if y_coords else 1
        x_range = 0
        
        # 2. X좌표 기준으로 열 경계 찾기
        x_coords = sorted(set(
            [s['left'] for s in shapes] + [s['right'] for s in shapes]
        ))
        x_range = max(x_coords) - min(x_coords) if x_coords else 1
        
        # 임계값: 슬라이드 크기의 15% 정도로 설정
        # 2열 이상으로 나뉘려면 최소 ~50% 이상 간격이 있어야 함
        y_threshold = max(y_range // 6, 500000)  # 최소 ~0.5인치
        x_threshold = max(x_range // 6, 500000)  # 최소 ~0.5인치
        
        row_boundaries = self._find_cluster_boundaries(y_coords, y_threshold)
        col_boundaries = self._find_cluster_boundaries(x_coords, x_threshold)
        
        # 3. 행/열 수 결정
        num_rows = len(row_boundaries) - 1 if len(row_boundaries) > 1 else 1
        num_cols = len(col_boundaries) - 1 if len(col_boundaries) > 1 else 1
        
        # 4. 그리드 셀 생성 및 shape 매핑
        cells = []
        row_heights = []
        col_widths = []
        
        # 행 높이 계산
        for i in range(num_rows):
            if i + 1 < len(row_boundaries):
                row_heights.append(row_boundaries[i + 1] - row_boundaries[i])
            else:
                row_heights.append(0)
        
        # 열 너비 계산
        for i in range(num_cols):
            if i + 1 < len(col_boundaries):
                col_widths.append(col_boundaries[i + 1] - col_boundaries[i])
            else:
                col_widths.append(0)
        
        # 각 셀에 shape 매핑
        for row_idx in range(num_rows):
            row_top = row_boundaries[row_idx] if row_idx < len(row_boundaries) else 0
            row_bottom = row_boundaries[row_idx + 1] if row_idx + 1 < len(row_boundaries) else row_top
            
            for col_idx in range(num_cols):
                col_left = col_boundaries[col_idx] if col_idx < len(col_boundaries) else 0
                col_right = col_boundaries[col_idx + 1] if col_idx + 1 < len(col_boundaries) else col_left
                
                cell = GridCell(
                    row=row_idx,
                    col=col_idx,
                    left=col_left,
                    top=row_top,
                    width=col_right - col_left,
                    height=row_bottom - row_top,
                )
                
                # 이 셀에 속하는 shape들 찾기
                cell_shapes = []
                content_types = set()
                
                for s in shapes:
                    # shape 중심이 셀 영역에 있는지 확인
                    center_x = (s['left'] + s['right']) // 2
                    center_y = (s['top'] + s['bottom']) // 2
                    
                    if (col_left <= center_x < col_right and 
                        row_top <= center_y < row_bottom):
                        cell_shapes.append(s['shape'])
                        content_types.add(s['type'])
                
                cell.shapes = cell_shapes
                
                # 콘텐츠 타입 결정
                if not cell_shapes:
                    cell.content_type = 'empty'
                elif len(content_types) == 1:
                    cell.content_type = list(content_types)[0]
                else:
                    cell.content_type = 'mixed'
                
                cells.append(cell)
        
        # 5. 셀 병합 분석 (인접한 셀이 같은 shape를 공유하는 경우)
        cells = self._detect_cell_spans(cells, num_rows, num_cols)
        
        return GridLayout(
            rows=num_rows,
            cols=num_cols,
            cells=cells,
            row_heights=row_heights,
            col_widths=col_widths,
        )
    
    def _extract_shape_bounds(self, shape: Any) -> Optional[Dict[str, Any]]:
        """shape의 경계 및 타입 정보 추출"""
        try:
            left = shape.left if hasattr(shape, 'left') else 0
            top = shape.top if hasattr(shape, 'top') else 0
            width = shape.width if hasattr(shape, 'width') else 0
            height = shape.height if hasattr(shape, 'height') else 0
            
            # 매우 작은 shape는 제외 (로고, 페이지 번호 등)
            if width < 100000 and height < 100000:  # ~1cm 미만
                return None
            
            # 타입 결정
            shape_type = 'text'
            if shape.has_table:
                shape_type = 'table'
            elif shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
                shape_type = 'image'
            elif shape.shape_type == MSO_SHAPE_TYPE.GROUP:
                shape_type = 'group'
            elif hasattr(shape, 'text_frame'):
                shape_type = 'text'
            
            return {
                'shape': shape,
                'left': left,
                'top': top,
                'right': left + width,
                'bottom': top + height,
                'type': shape_type,
            }
        except Exception:
            return None
    
    def _find_cluster_boundaries(
        self, 
        coords: List[int], 
        threshold: int = 200000  # ~2cm EMU
    ) -> List[int]:
        """
        좌표들을 클러스터링하여 그리드 경계 찾기
        
        가까운 좌표들은 같은 경계로 병합
        """
        if not coords:
            return [0]
        
        coords = sorted(set(coords))
        boundaries = [coords[0]]
        
        for coord in coords[1:]:
            # 이전 경계와 충분히 떨어져 있으면 새 경계로 추가
            if coord - boundaries[-1] > threshold:
                boundaries.append(coord)
        
        # 마지막 좌표가 경계에 없으면 추가
        if coords[-1] - boundaries[-1] > threshold:
            boundaries.append(coords[-1])
        
        return boundaries
    
    def _detect_cell_spans(
        self, 
        cells: List[GridCell], 
        num_rows: int, 
        num_cols: int
    ) -> List[GridCell]:
        """
        셀 병합(span) 감지
        
        하나의 shape가 여러 셀에 걸쳐있는 경우 셀 병합으로 처리
        """
        # 각 shape가 어떤 셀들에 나타나는지 추적
        shape_cells = {}  # shape_id -> [(row, col), ...]
        
        for cell in cells:
            for shape in cell.shapes:
                shape_id = id(shape)
                if shape_id not in shape_cells:
                    shape_cells[shape_id] = []
                shape_cells[shape_id].append((cell.row, cell.col))
        
        # shape가 여러 셀에 걸쳐있는 경우 처리
        merged_cells = set()  # (row, col) -> 병합된 셀들
        
        for shape_id, positions in shape_cells.items():
            if len(positions) > 1:
                # 최소/최대 row, col 찾기
                min_row = min(p[0] for p in positions)
                max_row = max(p[0] for p in positions)
                min_col = min(p[1] for p in positions)
                max_col = max(p[1] for p in positions)
                
                # 병합 범위에 있는 셀들 표시
                for r in range(min_row, max_row + 1):
                    for c in range(min_col, max_col + 1):
                        if (r, c) != (min_row, min_col):
                            merged_cells.add((r, c))
        
        # 병합된 셀 업데이트
        result_cells = []
        for cell in cells:
            if (cell.row, cell.col) in merged_cells:
                # 이 셀은 다른 셀에 병합됨 - 건너뛰기
                continue
            
            # rowspan, colspan 계산
            rowspan = 1
            colspan = 1
            
            # 같은 shape를 가진 인접 셀 찾기
            for shape in cell.shapes:
                shape_id = id(shape)
                if shape_id in shape_cells:
                    positions = shape_cells[shape_id]
                    if len(positions) > 1:
                        min_row = min(p[0] for p in positions)
                        max_row = max(p[0] for p in positions)
                        min_col = min(p[1] for p in positions)
                        max_col = max(p[1] for p in positions)
                        
                        if cell.row == min_row and cell.col == min_col:
                            rowspan = max(rowspan, max_row - min_row + 1)
                            colspan = max(colspan, max_col - min_col + 1)
            
            cell.rowspan = rowspan
            cell.colspan = colspan
            result_cells.append(cell)
        
        return result_cells

    def _convert_parsed_content(
        self, 
        doc: DocxDocument, 
        parsed: ParsedPresentation,
        prs: Presentation
    ):
        """파싱된 콘텐츠를 DOCX로 변환"""
        toc_end_idx = max(parsed.toc_slides) if parsed.toc_slides else 1
        
        for slide_content in parsed.slides:
            # 콘텐츠가 없는 슬라이드는 건너뛰기 (공백 페이지 방지)
            if self._is_empty_slide(slide_content):
                continue
            
            if slide_content.slide_type == 'title':
                self._create_title_page(doc, slide_content, parsed)
                doc.add_page_break()
                
            elif slide_content.slide_type == 'toc':
                self._create_toc_page(doc, slide_content)
                
                # 목차 마지막이면 구역 나누기 후 가로 레이아웃
                if (self.landscape_after_toc and 
                    slide_content.slide_index == toc_end_idx):
                    self._add_landscape_section(doc)
                else:
                    doc.add_page_break()
                
            else:
                self._create_content_page(doc, slide_content, prs)
                
                # 마지막 슬라이드가 아니면 페이지 나누기
                if slide_content.slide_index < len(parsed.slides):
                    doc.add_page_break()
    
    def _is_empty_slide(self, content: SlideContent) -> bool:
        """슬라이드에 표시할 콘텐츠가 없는지 확인"""
        # 제목 슬라이드나 목차 슬라이드는 항상 포함
        if content.slide_type in ('title', 'toc'):
            return False
        
        # 제목, 텍스트, 테이블, 이미지 모두 없으면 빈 슬라이드
        has_title = bool(content.title and content.title.strip())
        has_texts = any(t.get('shape') and hasattr(t['shape'], 'text_frame') 
                       and t['shape'].text_frame.text.strip() for t in content.texts)
        has_tables = len(content.tables) > 0
        has_images = len(content.images) > 0
        
        return not (has_title or has_texts or has_tables or has_images)
    
    def _create_title_page(
        self, 
        doc: DocxDocument, 
        content: SlideContent,
        parsed: ParsedPresentation
    ):
        """제목 페이지 생성 (페이지 가득 차게)"""
        # 상단 여백을 위한 빈 단락들
        for _ in range(8):
            doc.add_paragraph()
        
        # 메인 제목
        if content.title:
            title_para = doc.add_paragraph()
            title_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
            title_run = title_para.add_run(sanitize_text(content.title))
            title_run.font.size = Pt(36)
            title_run.font.bold = True
            title_run.font.color.rgb = RGBColor(0, 51, 102)
        
        # 부제목
        if content.subtitle:
            doc.add_paragraph()
            subtitle_para = doc.add_paragraph()
            subtitle_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
            subtitle_run = subtitle_para.add_run(sanitize_text(content.subtitle))
            subtitle_run.font.size = Pt(18)
            subtitle_run.font.color.rgb = RGBColor(100, 100, 100)
        
        # 하단 여백
        for _ in range(6):
            doc.add_paragraph()
        
        # 구분선
        self._add_horizontal_line(doc, color='003366')
        
        doc.add_paragraph()
        
        # 날짜
        if content.date:
            date_para = doc.add_paragraph()
            date_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
            date_run = date_para.add_run(sanitize_text(content.date))
            date_run.font.size = Pt(14)
            date_run.font.color.rgb = RGBColor(80, 80, 80)
        
        # 작성자
        if content.author:
            author_para = doc.add_paragraph()
            author_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
            author_run = author_para.add_run(sanitize_text(content.author))
            author_run.font.size = Pt(12)
            author_run.font.color.rgb = RGBColor(100, 100, 100)
    
    def _create_toc_page(self, doc: DocxDocument, content: SlideContent):
        """목차 페이지 생성"""
        # 목차 제목
        if content.title:
            heading = doc.add_heading(sanitize_text(content.title), level=1)
            heading.alignment = WD_ALIGN_PARAGRAPH.CENTER
        
        doc.add_paragraph()
        
        # 목차 항목들
        for idx, item in enumerate(content.toc_items, 1):
            item_text = sanitize_text(item)
            
            # 숫자가 이미 있는지 확인
            if re.match(r'^\d+[\.\s]', item_text):
                # 이미 번호가 있으면 그대로 사용
                para = doc.add_paragraph(item_text, style='List Number')
            else:
                # 번호 추가
                para = doc.add_paragraph(f"{idx}. {item_text}")
            
            para.paragraph_format.space_after = Pt(8)
            para.paragraph_format.left_indent = Cm(1)
            
            for run in para.runs:
                run.font.size = Pt(12)
    
    def _create_content_page(
        self, 
        doc: DocxDocument, 
        content: SlideContent,
        prs: Presentation
    ):
        """일반 콘텐츠 페이지 생성"""
        # 섹션 제목 처리 (목차의 대제목은 한 번만)
        if content.section_title:
            if content.section_title not in self._processed_section_titles:
                # 첫 번째 등장: 섹션 제목 표시
                self._processed_section_titles.add(content.section_title)
                heading = doc.add_heading(sanitize_text(content.section_title), level=1)
                self._add_horizontal_line(doc)
            else:
                # 이미 등장한 섹션: 소제목만 표시 (있으면)
                pass
        elif content.title:
            # 일반 제목
            heading = doc.add_heading(sanitize_text(content.title), level=1)
            self._add_horizontal_line(doc)
        
        # 그리드 레이아웃이 있고 2열 이상인 경우 그리드 기반 렌더링
        grid = content.grid_layout
        if grid and grid.cols >= 2 and self._should_use_grid_layout(grid):
            self._create_grid_based_content(doc, content, grid, prs)
            return
        
        # 기존 방식: 모든 콘텐츠를 위치 기반으로 정렬하여 순서대로 출력
        all_items = []
        
        for text_info in content.texts:
            all_items.append(('text', text_info['top'], text_info['left'], text_info))
        
        if self.include_tables:
            for table_info in content.tables:
                all_items.append(('table', table_info['top'], table_info['left'], table_info))
        
        if self.include_images:
            for image_info in content.images:
                all_items.append(('image', image_info['top'], image_info['left'], image_info))
        
        # 테이블 내부 이미지 좌표 수집 (중복 방지용)
        table_image_positions = set()
        # 테이블 영역 수집 (테이블 내 캡션 텍스트 중복 방지용)
        table_regions = []
        if self.include_tables and content.slide:
            for table_info in content.tables:
                table_shape = table_info['shape']
                # 테이블 영역 저장
                table_regions.append({
                    'left': table_shape.left,
                    'right': table_shape.left + table_shape.width,
                    'top': table_shape.top,
                    'bottom': table_shape.top + table_shape.height,
                })
                
                img_map = self._find_table_cell_images(
                    content.slide, table_shape, table_shape.table
                )
                for img_list in img_map.values():
                    for img_data in img_list:
                        # 이미지 위치를 저장 (추후 독립 이미지에서 제외)
                        table_image_positions.add(id(img_data.get('blob', b'')))
        
        # 위치 기반 정렬 (위에서 아래, 왼쪽에서 오른쪽)
        all_items.sort(key=lambda x: (x[1], x[2]))
        
        for item_type, _, _, item_info in all_items:
            if item_type == 'text':
                # 테이블 영역과 겹치는 텍스트는 건너뛰기 (테이블 캡션으로 이미 추가됨)
                shape = item_info['shape']
                if self._is_shape_in_table_region(shape, table_regions):
                    continue
                self._add_text_from_shape(doc, shape)
            elif item_type == 'table':
                self._add_table_from_shape(
                    doc, item_info['shape'], prs, content.slide
                )
            elif item_type == 'image':
                self._add_image_from_shape(doc, item_info['shape'])
    
    def _should_use_grid_layout(self, grid: GridLayout) -> bool:
        """
        그리드 레이아웃을 사용해야 하는지 결정
        
        2열 이상이고, 의미 있는 콘텐츠가 나란히 배치된 경우에만 사용
        """
        if grid.cols < 2:
            return False
        
        # 같은 행에 서로 다른 열에 콘텐츠가 있는 경우가 있어야 함
        for row_idx in range(grid.rows):
            cols_with_content = []
            for cell in grid.cells:
                if cell.row == row_idx and cell.content_type != 'empty':
                    cols_with_content.append(cell.col)
            
            # 같은 행에 2개 이상의 열에 콘텐츠가 있으면 그리드 사용
            if len(set(cols_with_content)) >= 2:
                return True
        
        return False
    
    def _create_grid_based_content(
        self, 
        doc: DocxDocument, 
        content: SlideContent, 
        grid: GridLayout,
        prs: Presentation
    ):
        """
        그리드 레이아웃을 기반으로 DOCX 콘텐츠 생성
        
        DOCX 테이블을 사용하여 2열 레이아웃 등을 구현
        """
        # 그리드 정보 주석 추가 (디버깅용, 옵션)
        # doc.add_paragraph(f"[Grid: {grid.rows}x{grid.cols}]")
        
        # 행별로 처리
        for row_idx in range(grid.rows):
            row_cells = [c for c in grid.cells if c.row == row_idx]
            row_cells.sort(key=lambda c: c.col)
            
            # 이 행의 열 수 (colspan 고려)
            non_empty_cells = [c for c in row_cells if c.content_type != 'empty']
            
            if not non_empty_cells:
                continue
            
            # 단일 셀 (전체 너비)
            if len(non_empty_cells) == 1:
                cell = non_empty_cells[0]
                self._render_grid_cell_content(doc, cell, prs, content.slide)
            else:
                # 다중 열: DOCX 테이블로 레이아웃
                self._create_layout_table(doc, non_empty_cells, prs, content.slide, grid)
        
        doc.add_paragraph()
    
    def _create_layout_table(
        self, 
        doc: DocxDocument, 
        cells: List[GridCell], 
        prs: Presentation,
        slide: Any,
        grid: GridLayout
    ):
        """
        레이아웃용 테이블 생성 (테두리 없음)
        
        그리드 셀들을 DOCX 테이블의 셀로 매핑
        """
        num_cols = len(cells)
        layout_table = doc.add_table(rows=1, cols=num_cols)
        layout_table.alignment = WD_TABLE_ALIGNMENT.CENTER
        
        # 테두리 제거 (레이아웃용)
        self._remove_table_borders(layout_table)
        
        # 열 너비 설정 (그리드 비율 기반)
        total_width = sum(grid.col_widths) if grid.col_widths else 1
        
        for idx, cell in enumerate(cells):
            doc_cell = layout_table.rows[0].cells[idx]
            
            # 셀 너비 설정
            if grid.col_widths and cell.col < len(grid.col_widths):
                # colspan 고려한 너비
                cell_width = 0
                for c in range(cell.col, min(cell.col + cell.colspan, len(grid.col_widths))):
                    cell_width += grid.col_widths[c]
                
                # 가로 레이아웃 기준 대략적인 너비 (총 10인치 정도)
                width_ratio = cell_width / total_width if total_width > 0 else 0.5
                doc_cell.width = Inches(10 * width_ratio)
            
            # 셀 콘텐츠 렌더링
            self._render_cell_shapes(doc_cell, cell.shapes, prs, slide)
        
        doc.add_paragraph()
    
    def _remove_table_borders(self, table: Any):
        """테이블 테두리 제거 (레이아웃 테이블용)"""
        try:
            tbl = table._tbl
            tbl_pr = tbl.tblPr if tbl.tblPr is not None else OxmlElement('w:tblPr')
            tbl_borders = OxmlElement('w:tblBorders')
            
            for border_name in ['top', 'left', 'bottom', 'right', 'insideH', 'insideV']:
                border = OxmlElement(f'w:{border_name}')
                border.set(qn('w:val'), 'nil')
                tbl_borders.append(border)
            
            tbl_pr.append(tbl_borders)
            if tbl.tblPr is None:
                tbl.insert(0, tbl_pr)
        except Exception as e:
            logger.debug(f"테이블 테두리 제거 실패: {e}")
    
    def _render_grid_cell_content(
        self, 
        doc: DocxDocument, 
        cell: GridCell, 
        prs: Presentation,
        slide: Any
    ):
        """단일 그리드 셀 콘텐츠 렌더링 (전체 너비)"""
        for shape in cell.shapes:
            if shape.has_table:
                self._add_table_from_shape(doc, shape, prs, slide)
            elif shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
                self._add_image_from_shape(doc, shape)
            elif hasattr(shape, 'text_frame') and shape.text_frame.text.strip():
                self._add_text_from_shape(doc, shape)
    
    def _render_cell_shapes(
        self, 
        doc_cell: Any, 
        shapes: List[Any], 
        prs: Presentation,
        slide: Any
    ):
        """DOCX 테이블 셀에 shape들 렌더링"""
        for shape in shapes:
            try:
                if shape.has_table:
                    # 테이블을 셀에 중첩하기는 복잡하므로 텍스트만 추출
                    self._add_table_text_to_cell(doc_cell, shape.table)
                elif shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
                    self._add_image_to_cell(doc_cell, shape)
                elif hasattr(shape, 'text_frame') and shape.text_frame.text.strip():
                    self._add_text_to_cell(doc_cell, shape)
            except Exception as e:
                logger.debug(f"셀 shape 렌더링 실패: {e}")
    
    def _add_table_text_to_cell(self, doc_cell: Any, ppt_table: Any):
        """PPT 테이블 내용을 DOCX 셀에 텍스트로 추가"""
        for row in ppt_table.rows:
            row_texts = []
            for cell in row.cells:
                text = sanitize_text(cell.text.strip())
                if text:
                    row_texts.append(text)
            
            if row_texts:
                para = doc_cell.add_paragraph(' | '.join(row_texts))
                para.paragraph_format.space_after = Pt(2)
    
    def _add_image_to_cell(self, doc_cell: Any, shape: Any):
        """이미지를 DOCX 셀에 추가"""
        try:
            image = shape.image
            image_bytes = image.blob
            
            # 크롭 정보 적용
            crop_info = self._get_image_crop_info(shape)
            if crop_info and HAS_PIL:
                image_bytes = self._apply_image_crop(image_bytes, crop_info)
            
            # 셀 내 이미지 크기 제한 (최대 3인치)
            width = shape.width
            width_inches = min(width / 914400, 3.0)
            
            para = doc_cell.add_paragraph()
            para.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run = para.add_run()
            run.add_picture(BytesIO(image_bytes), width=Inches(width_inches))
        except Exception as e:
            logger.debug(f"셀 이미지 추가 실패: {e}")
    
    def _add_text_to_cell(self, doc_cell: Any, shape: Any):
        """텍스트를 DOCX 셀에 추가"""
        if not hasattr(shape, 'text_frame'):
            return
        
        for paragraph in shape.text_frame.paragraphs:
            text = sanitize_text(paragraph.text.strip())
            if not text:
                continue
            
            # 페이지 번호 제외
            if is_page_number(text):
                continue
            
            para = doc_cell.add_paragraph(text)
            para.paragraph_format.space_after = Pt(2)
            
            # 키워드 강조
            if self.highlight_keywords and is_highlight_keyword(text):
                for run in para.runs:
                    self._apply_highlight_style(run)

    def _is_shape_in_table_region(
        self, 
        shape: Any, 
        table_regions: List[Dict[str, int]]
    ) -> bool:
        """
        shape가 테이블 영역과 겹치는지 확인
        (짧은 캡션 텍스트만 제외 대상)
        """
        if not table_regions:
            return False
        
        # 텍스트가 있는 shape만 확인
        if not hasattr(shape, 'text_frame'):
            return False
        
        text = shape.text_frame.text.strip()
        # 짧은 캡션 텍스트만 제외 (100자 미만)
        if len(text) >= 100:
            return False
        
        shape_left = shape.left if hasattr(shape, 'left') else 0
        shape_top = shape.top if hasattr(shape, 'top') else 0
        shape_right = shape_left + (shape.width if hasattr(shape, 'width') else 0)
        shape_bottom = shape_top + (shape.height if hasattr(shape, 'height') else 0)
        
        for region in table_regions:
            # 테이블 영역과 겹치는지 확인
            if (shape_left < region['right'] and shape_right > region['left'] and
                shape_top < region['bottom'] and shape_bottom > region['top']):
                return True
        
        return False

    def _add_text_from_shape(self, doc: DocxDocument, shape: Any):
        """shape에서 텍스트 추출하여 추가"""
        if not hasattr(shape, 'text_frame'):
            if hasattr(shape, 'text') and shape.text.strip():
                text = sanitize_text(shape.text.strip())
                # 페이지 번호 제외
                if text and not is_page_number(text):
                    para = doc.add_paragraph(text)
                    self._apply_keyword_highlighting(para)
            return
        
        for paragraph in shape.text_frame.paragraphs:
            text = sanitize_text(paragraph.text.strip())
            if not text:
                continue
            
            # 페이지 번호 제외
            if is_page_number(text):
                continue
            
            level = paragraph.level if hasattr(paragraph, 'level') else 0
            has_bullet = self._has_bullet(paragraph)
            
            # 키워드 강조 여부 확인
            is_keyword = self.highlight_keywords and is_highlight_keyword(text)
            
            # 불릿이 있거나 들여쓰기가 있으면 리스트 스타일
            if has_bullet or level > 0:
                # 불릿 문자가 텍스트에 없으면 추가
                if not text.startswith(('•', '-', '▶', '●', '○', '■', '◆')):
                    para = doc.add_paragraph(f"• {text}")
                else:
                    para = doc.add_paragraph(text)
                para.paragraph_format.left_indent = Cm(max(level, 1) * 0.75)
                if is_keyword:
                    for run in para.runs:
                        self._apply_highlight_style(run)
            else:
                para = doc.add_paragraph()
                run = para.add_run(text)
                if is_keyword:
                    self._apply_highlight_style(run)
            
            self._copy_paragraph_style(paragraph, para)
            para.paragraph_format.space_after = Pt(4)
    
    def _has_bullet(self, paragraph: Any) -> bool:
        """PPT 단락에 불릿이 있는지 확인"""
        try:
            # pptx의 paragraph 객체에서 bullet 확인
            if hasattr(paragraph, '_pPr') and paragraph._pPr is not None:
                # XML에서 buNone이 아니면 불릿이 있음
                pPr = paragraph._pPr
                # buChar, buAutoNum 등이 있으면 불릿
                bu_elements = pPr.findall('.//' + qn('a:buChar'))
                bu_elements += pPr.findall('.//' + qn('a:buAutoNum'))
                bu_elements += pPr.findall('.//' + qn('a:buBlip'))
                return len(bu_elements) > 0
            
            # 텍스트가 불릿 문자로 시작하면 불릿으로 간주
            text = paragraph.text.strip() if hasattr(paragraph, 'text') else ''
            bullet_chars = ('•', '-', '–', '▶', '●', '○', '■', '◆', '★', '✓', '*')
            return text.startswith(bullet_chars)
        except Exception:
            return False
    
    def _add_table_from_shape(
        self, 
        doc: DocxDocument, 
        shape: Any, 
        prs: Presentation,
        slide: Any = None
    ):
        """테이블 추가 (셀 병합 및 이미지 포함)"""
        ppt_table = shape.table
        row_count = len(ppt_table.rows)
        col_count = len(ppt_table.columns)
        
        if row_count == 0 or col_count == 0:
            return
        
        # 테이블 영역 내 이미지 찾기
        cell_image_map = {}  # (row, col) -> image_bytes
        if slide and self.include_images:
            cell_image_map = self._find_table_cell_images(slide, shape, ppt_table)
        
        # 셀 병합 정보 수집
        merge_info = self._get_table_merge_info(ppt_table)
        
        doc_table = doc.add_table(rows=row_count, cols=col_count)
        doc_table.style = 'Table Grid'
        doc_table.alignment = WD_TABLE_ALIGNMENT.CENTER
        
        # 먼저 셀 병합 처리
        self._apply_cell_merges(doc_table, merge_info)
        
        # 병합된 셀 추적 (spanned 셀은 건너뛰기)
        processed_cells = set()
        
        for row_idx, row in enumerate(ppt_table.rows):
            for col_idx, cell in enumerate(row.cells):
                cell_key = (row_idx, col_idx)
                
                # 이미 처리된 셀(병합된 하위 셀)은 건너뛰기
                if cell_key in processed_cells:
                    continue
                
                # 병합된 셀인지 확인
                if cell.is_spanned and not cell.is_merge_origin:
                    continue
                
                doc_cell = doc_table.rows[row_idx].cells[col_idx]
                
                # 셀 내 이미지 확인 및 추가
                has_image = False
                if cell_key in cell_image_map:
                    for img_data in cell_image_map[cell_key]:
                        try:
                            para = doc_cell.add_paragraph()
                            para.alignment = WD_ALIGN_PARAGRAPH.CENTER
                            run = para.add_run()
                            run.add_picture(BytesIO(img_data['blob']), width=Inches(1.5))
                            has_image = True
                            
                            # 캡션 추가
                            caption = img_data.get('caption')
                            if caption:
                                cap_para = doc_cell.add_paragraph()
                                cap_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
                                cap_run = cap_para.add_run(caption)
                                cap_run.font.size = Pt(9)
                                cap_run.font.italic = True
                                
                        except Exception as e:
                            logger.debug(f"셀 이미지 추가 실패: {e}")
                
                # 텍스트 추가
                cell_text = sanitize_text(cell.text.strip())
                if cell_text:
                    if has_image:
                        # 이미지가 있으면 새 단락에 텍스트 추가
                        para = doc_cell.add_paragraph(cell_text)
                    else:
                        doc_cell.text = cell_text
                    
                    # 키워드 강조
                    if self.highlight_keywords and is_highlight_keyword(cell_text):
                        for para in doc_cell.paragraphs:
                            for run in para.runs:
                                self._apply_highlight_style(run)
                
                # 첫 번째 행은 헤더로 굵게
                if row_idx == 0:
                    for para in doc_cell.paragraphs:
                        for run in para.runs:
                            run.font.bold = True
        
        doc.add_paragraph()
    
    def _get_table_merge_info(self, ppt_table: Any) -> List[Dict[str, Any]]:
        """PPTX 테이블에서 셀 병합 정보 추출"""
        merge_info = []
        
        for row_idx in range(len(ppt_table.rows)):
            for col_idx in range(len(ppt_table.columns)):
                cell = ppt_table.rows[row_idx].cells[col_idx]
                
                if cell.is_merge_origin:
                    # 병합 시작 셀 - colspan과 rowspan 계산
                    colspan = 1
                    rowspan = 1
                    
                    # colspan 계산 (오른쪽으로 병합된 셀 수)
                    for c in range(col_idx + 1, len(ppt_table.columns)):
                        if ppt_table.rows[row_idx].cells[c].is_spanned:
                            # 같은 원본 셀에서 병합된 것인지 확인
                            colspan += 1
                        else:
                            break
                    
                    # rowspan 계산 (아래로 병합된 셀 수)
                    for r in range(row_idx + 1, len(ppt_table.rows)):
                        if ppt_table.rows[r].cells[col_idx].is_spanned:
                            rowspan += 1
                        else:
                            break
                    
                    merge_info.append({
                        'row': row_idx,
                        'col': col_idx,
                        'rowspan': rowspan,
                        'colspan': colspan,
                    })
        
        return merge_info
    
    def _apply_cell_merges(self, doc_table: Any, merge_info: List[Dict[str, Any]]):
        """DOCX 테이블에 셀 병합 적용"""
        for merge in merge_info:
            row = merge['row']
            col = merge['col']
            rowspan = merge['rowspan']
            colspan = merge['colspan']
            
            try:
                # 시작 셀
                start_cell = doc_table.rows[row].cells[col]
                
                # 끝 셀 (병합 범위의 오른쪽 아래)
                end_row = row + rowspan - 1
                end_col = col + colspan - 1
                
                if end_row < len(doc_table.rows) and end_col < len(doc_table.rows[0].cells):
                    end_cell = doc_table.rows[end_row].cells[end_col]
                    start_cell.merge(end_cell)
                    
            except Exception as e:
                logger.debug(f"셀 병합 실패 ({row},{col}): {e}")
    
    def _find_merge_origin(
        self, 
        ppt_table: Any, 
        row: int, 
        col: int
    ) -> Tuple[int, int]:
        """
        주어진 셀이 병합된 경우 merge_origin 셀 좌표 반환
        
        Args:
            ppt_table: PPTX 테이블 객체
            row: 행 인덱스
            col: 열 인덱스
            
        Returns:
            (origin_row, origin_col): merge_origin 셀 좌표 (병합되지 않은 경우 원래 좌표)
        """
        try:
            cell = ppt_table.rows[row].cells[col]
            
            # 이미 merge_origin이거나 병합되지 않은 셀
            if cell.is_merge_origin or not cell.is_spanned:
                return (row, col)
            
            # is_spanned인 경우 위쪽/왼쪽으로 merge_origin 찾기
            # 먼저 같은 열에서 위쪽으로 검색
            for r in range(row - 1, -1, -1):
                check_cell = ppt_table.rows[r].cells[col]
                if check_cell.is_merge_origin:
                    return (r, col)
                elif not check_cell.is_spanned:
                    break
            
            # 같은 행에서 왼쪽으로 검색
            for c in range(col - 1, -1, -1):
                check_cell = ppt_table.rows[row].cells[c]
                if check_cell.is_merge_origin:
                    return (row, c)
                elif not check_cell.is_spanned:
                    break
            
            # 대각선으로 검색 (왼쪽 위)
            for offset in range(1, max(row, col) + 1):
                r, c = row - offset, col - offset
                if r >= 0 and c >= 0:
                    check_cell = ppt_table.rows[r].cells[c]
                    if check_cell.is_merge_origin:
                        return (r, c)
            
        except Exception as e:
            logger.debug(f"merge_origin 찾기 실패 ({row},{col}): {e}")
        
        return (row, col)

    def _find_table_cell_images(
        self, 
        slide: Any, 
        table_shape: Any, 
        table: Any,
        include_side_images: bool = True
    ) -> Dict[Tuple[int, int], List[Dict[str, Any]]]:
        """
        테이블 영역 내 이미지를 찾아 셀별로 매핑
        
        Args:
            slide: 슬라이드 객체
            table_shape: 테이블 shape
            table: 테이블 객체
            include_side_images: 테이블 옆(오른쪽) 이미지도 포함할지 여부
        """
        cell_image_map = {}
        
        try:
            # 각 열의 절대 위치 계산
            col_positions = [table_shape.left]
            for i in range(len(table.columns)):
                col_positions.append(col_positions[-1] + table.columns[i].width)
            
            # 각 행의 절대 위치 계산
            row_positions = [table_shape.top]
            for i in range(len(table.rows)):
                row_positions.append(row_positions[-1] + table.rows[i].height)
            
            # 테이블 영역
            table_left = table_shape.left
            table_right = col_positions[-1]
            table_top = table_shape.top
            table_bottom = row_positions[-1]
            
            # 슬라이드의 모든 이미지 찾기
            images_to_check = self._collect_images_from_slide(slide)
            
            # 테이블 옆 이미지 캡션 정보 수집
            side_image_captions = {}
            if include_side_images:
                side_image_captions = self._find_side_image_captions(
                    slide, table_right, table_top, table_bottom
                )
            
            # 테이블 내부 및 인접 영역의 캡션 수집 (이미지 아래 텍스트)
            inner_captions = self._find_table_inner_captions(
                slide, table_left, table_right, table_top, table_bottom
            )
            # side_image_captions와 inner_captions 병합
            all_captions = {**side_image_captions, **inner_captions}
            
            # 각 이미지가 테이블 셀에 속하는지 확인
            for img_shape in images_to_check:
                try:
                    img_center_x = img_shape.left + img_shape.width // 2
                    img_center_y = img_shape.top + img_shape.height // 2
                    
                    # 테이블 내부 이미지
                    col = -1
                    for i in range(len(col_positions) - 1):
                        if col_positions[i] <= img_center_x < col_positions[i + 1]:
                            col = i
                            break
                    
                    row = -1
                    for i in range(len(row_positions) - 1):
                        if row_positions[i] <= img_center_y < row_positions[i + 1]:
                            row = i
                            break
                    
                    if row >= 0 and col >= 0:
                        # 병합된 셀인 경우 merge_origin으로 리다이렉트
                        actual_row, actual_col = self._find_merge_origin(
                            table, row, col
                        )
                        
                        # 테이블 내부 이미지
                        cell_key = (actual_row, actual_col)
                        if cell_key not in cell_image_map:
                            cell_image_map[cell_key] = []
                        
                        # ImageWithPosition 또는 일반 shape 처리
                        actual_shape = img_shape.shape if hasattr(img_shape, 'shape') else img_shape
                        crop_info = self._get_image_crop_info(actual_shape)
                        image_blob = actual_shape.image.blob
                        if crop_info and HAS_PIL:
                            image_blob = self._apply_image_crop(image_blob, crop_info)
                        
                        # 테이블 내부 이미지에도 캡션 찾기
                        caption = self._find_caption_for_image(img_shape, all_captions)
                        
                        cell_image_map[cell_key].append({
                            'blob': image_blob,
                            'ext': actual_shape.image.ext,
                            'width': img_shape.width,
                            'height': img_shape.height,
                            'caption': caption,
                        })
                    
                    elif include_side_images and img_center_x > table_right:
                        # 테이블 오른쪽의 이미지 - 같은 높이의 행에 매핑
                        if table_top <= img_center_y <= table_bottom:
                            # 이미지 중심이 어느 행과 겹치는지 확인
                            for i in range(len(row_positions) - 1):
                                if row_positions[i] <= img_center_y < row_positions[i + 1]:
                                    row = i
                                    break
                            
                            if row >= 0:
                                # 마지막 열에 매핑 (오른쪽 이미지용 열)
                                col = len(table.columns) - 1
                                cell_key = (row, col)
                                
                                if cell_key not in cell_image_map:
                                    cell_image_map[cell_key] = []
                                
                                actual_shape = img_shape.shape if hasattr(img_shape, 'shape') else img_shape
                                crop_info = self._get_image_crop_info(actual_shape)
                                image_blob = actual_shape.image.blob
                                if crop_info and HAS_PIL:
                                    image_blob = self._apply_image_crop(image_blob, crop_info)
                                
                                # 캡션 찾기
                                caption = self._find_caption_for_image(
                                    img_shape, all_captions
                                )
                                
                                cell_image_map[cell_key].append({
                                    'blob': image_blob,
                                    'ext': actual_shape.image.ext,
                                    'width': img_shape.width,
                                    'height': img_shape.height,
                                    'caption': caption,
                                })
                                
                except Exception as e:
                    logger.debug(f"이미지 위치 확인 실패: {e}")
                    
        except Exception as e:
            logger.debug(f"테이블 셀 이미지 찾기 실패: {e}")
        
        return cell_image_map
    
    def _collect_images_from_slide(self, slide: Any) -> List[Any]:
        """슬라이드에서 모든 이미지 shape 수집 (그룹 포함, 절대 좌표 계산)"""
        
        @dataclass
        class ImageWithPosition:
            """이미지와 절대 위치 정보"""
            shape: Any
            abs_left: int
            abs_top: int
            abs_width: int
            abs_height: int
            
            @property
            def left(self):
                return self.abs_left
            
            @property
            def top(self):
                return self.abs_top
            
            @property
            def width(self):
                return self.abs_width
            
            @property
            def height(self):
                return self.abs_height
            
            @property
            def image(self):
                return self.shape.image
        
        images = []
        
        def collect_recursive(shape, parent_left=0, parent_top=0):
            if shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
                # 절대 좌표 계산
                abs_left = (shape.left if hasattr(shape, 'left') else 0) + parent_left
                abs_top = (shape.top if hasattr(shape, 'top') else 0) + parent_top
                
                images.append(ImageWithPosition(
                    shape=shape,
                    abs_left=abs_left,
                    abs_top=abs_top,
                    abs_width=shape.width if hasattr(shape, 'width') else 0,
                    abs_height=shape.height if hasattr(shape, 'height') else 0,
                ))
            elif shape.shape_type == MSO_SHAPE_TYPE.GROUP:
                try:
                    # 그룹의 좌표를 누적하여 하위 shape에 전달
                    group_left = (shape.left if hasattr(shape, 'left') else 0) + parent_left
                    group_top = (shape.top if hasattr(shape, 'top') else 0) + parent_top
                    for sub_shape in shape.shapes:
                        collect_recursive(sub_shape, group_left, group_top)
                except Exception:
                    pass
        
        for shape in slide.shapes:
            collect_recursive(shape)
        
        return images
    
    def _find_side_image_captions(
        self, 
        slide: Any, 
        table_right: int, 
        table_top: int, 
        table_bottom: int
    ) -> Dict[Tuple[int, int], str]:
        """
        테이블 오른쪽 영역에서 이미지 캡션 텍스트 수집 (그룹 내 텍스트 포함)
        
        Returns:
            Dict[(top, bottom), caption_text]: 위치별 캡션 텍스트
        """
        captions = {}
        
        def collect_text_recursive(shape, parent_left=0, parent_top=0):
            """재귀적으로 텍스트 수집"""
            abs_left = (shape.left if hasattr(shape, 'left') else 0) + parent_left
            abs_top = (shape.top if hasattr(shape, 'top') else 0) + parent_top
            
            if hasattr(shape, 'text_frame') and shape.text_frame.text.strip():
                text = shape.text_frame.text.strip()
                # 페이지 번호 제외, 짧은 캡션만
                if not is_page_number(text) and len(text) < 100:
                    shape_height = shape.height if hasattr(shape, 'height') else 0
                    captions[(abs_top, abs_top + shape_height)] = sanitize_text(text)
            
            if shape.shape_type == MSO_SHAPE_TYPE.GROUP:
                try:
                    for sub_shape in shape.shapes:
                        collect_text_recursive(sub_shape, abs_left, abs_top)
                except Exception:
                    pass
        
        for shape in slide.shapes:
            # 테이블은 건너뛰기
            if shape.has_table:
                continue
            
            shape_left = shape.left if hasattr(shape, 'left') else 0
            # 테이블 오른쪽에 있는 요소만
            if shape_left >= table_right * 0.8:
                collect_text_recursive(shape)
        
        return captions
    
    def _find_table_inner_captions(
        self, 
        slide: Any, 
        table_left: int, 
        table_right: int, 
        table_top: int, 
        table_bottom: int
    ) -> Dict[Tuple[int, int], str]:
        """
        테이블 영역 내부에 있는 이미지 캡션 텍스트 수집
        (테이블 셀 내용이 아닌, 테이블 위에 떠있는 TEXT_BOX)
        
        Returns:
            Dict[(top, bottom), caption_text]: 위치별 캡션 텍스트
        """
        captions = {}
        
        def collect_text_recursive(shape, parent_left=0, parent_top=0):
            """재귀적으로 텍스트 수집"""
            abs_left = (shape.left if hasattr(shape, 'left') else 0) + parent_left
            abs_top = (shape.top if hasattr(shape, 'top') else 0) + parent_top
            abs_height = shape.height if hasattr(shape, 'height') else 0
            
            if hasattr(shape, 'text_frame') and shape.text_frame.text.strip():
                text = shape.text_frame.text.strip()
                # 페이지 번호 제외, 짧은 캡션만
                if not is_page_number(text) and len(text) < 100:
                    captions[(abs_top, abs_top + abs_height)] = sanitize_text(text)
            
            if shape.shape_type == MSO_SHAPE_TYPE.GROUP:
                try:
                    for sub_shape in shape.shapes:
                        collect_text_recursive(sub_shape, abs_left, abs_top)
                except Exception:
                    pass
        
        for shape in slide.shapes:
            # 테이블은 건너뛰기
            if shape.has_table:
                continue
            
            shape_left = shape.left if hasattr(shape, 'left') else 0
            shape_top = shape.top if hasattr(shape, 'top') else 0
            shape_right = shape_left + (shape.width if hasattr(shape, 'width') else 0)
            shape_bottom = shape_top + (shape.height if hasattr(shape, 'height') else 0)
            
            # 테이블 영역과 겹치는 요소 (테이블 내부 또는 인접)
            if (shape_left < table_right and shape_right > table_left and
                shape_top < table_bottom and shape_bottom > table_top):
                collect_text_recursive(shape)
        
        return captions
    
    def _find_caption_for_image(
        self, 
        img_shape: Any, 
        captions: Dict[Tuple[int, int], str]
    ) -> Optional[str]:
        """
        이미지와 가장 가까운 캡션 찾기 (이미지 아래 또는 약간 겹치는 위치)
        """
        if not captions:
            return None
        
        img_top = img_shape.top
        img_bottom = img_shape.top + img_shape.height
        img_center_x = img_shape.left + img_shape.width // 2
        img_height = img_shape.height
        
        best_caption = None
        best_distance = float('inf')
        
        for (cap_top, cap_bottom), text in captions.items():
            # 캡션이 이미지 하단 근처에 있는지 확인
            # - 캡션이 이미지 하단 50% 이하에서 시작하거나
            # - 캡션이 이미지 바로 아래에 있는 경우
            img_lower_half = img_top + img_height // 2
            
            if cap_top >= img_lower_half:
                # 이미지 하단과 캡션 상단 사이의 거리
                distance = abs(cap_top - img_bottom)
                
                # 거리가 이미지 높이의 50% 이내인 경우
                if distance < img_height // 2 and distance < best_distance:
                    best_distance = distance
                    best_caption = text
        
        return best_caption
    
    def _add_image_from_shape(self, doc: DocxDocument, shape: Any):
        """이미지 추가 (크롭 정보 적용)"""
        try:
            image = shape.image
            image_bytes = image.blob
            
            # 크롭 정보 가져오기
            crop_info = self._get_image_crop_info(shape)
            
            # 크롭이 있으면 이미지 자르기
            if crop_info and HAS_PIL:
                image_bytes = self._apply_image_crop(image_bytes, crop_info)
            
            width = shape.width
            height = shape.height
            
            # EMU to Inches
            width_inches = width / 914400
            height_inches = height / 914400
            
            # 최대 너비 제한
            if width_inches > self.image_max_width_inches:
                scale = self.image_max_width_inches / width_inches
                width_inches = self.image_max_width_inches
                height_inches = height_inches * scale
            
            para = doc.add_paragraph()
            para.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run = para.add_run()
            run.add_picture(BytesIO(image_bytes), width=Inches(width_inches))
            
            doc.add_paragraph()
            
        except Exception as e:
            logger.warning(f"이미지 추가 실패: {e}")
    
    def _get_image_crop_info(self, shape: Any) -> Optional[Dict[str, float]]:
        """
        PPTX shape에서 이미지 크롭 정보 추출
        
        PPTX에서 이미지 크롭은 a:srcRect 요소에 저장됨:
        - l (left): 왼쪽 크롭 비율 (1/100000 단위)
        - t (top): 위쪽 크롭 비율
        - r (right): 오른쪽 크롭 비율
        - b (bottom): 아래쪽 크롭 비율
        """
        try:
            # shape의 XML 요소 접근
            spTree = shape._element
            
            # a:srcRect 요소 찾기 (blipFill 내부)
            nsmap_a = {'a': 'http://schemas.openxmlformats.org/drawingml/2006/main'}
            srcRect = spTree.find('.//a:srcRect', nsmap_a)
            
            if srcRect is None:
                return None
            
            # 크롭 값 추출 (1/100000 단위 -> 0~1 비율)
            crop = {
                'left': int(srcRect.get('l', '0')) / 100000,
                'top': int(srcRect.get('t', '0')) / 100000,
                'right': int(srcRect.get('r', '0')) / 100000,
                'bottom': int(srcRect.get('b', '0')) / 100000,
            }
            
            # 크롭이 없으면 None 반환
            if all(v == 0 for v in crop.values()):
                return None
            
            logger.debug(f"이미지 크롭 정보: {crop}")
            return crop
            
        except Exception as e:
            logger.debug(f"크롭 정보 추출 실패: {e}")
            return None
    
    def _apply_image_crop(
        self, 
        image_bytes: bytes, 
        crop: Dict[str, float]
    ) -> bytes:
        """
        PIL을 사용하여 이미지에 크롭 적용
        
        Args:
            image_bytes: 원본 이미지 바이트
            crop: 크롭 정보 (left, top, right, bottom 비율)
        
        Returns:
            크롭된 이미지 바이트
        """
        try:
            # 이미지 열기
            img = Image.open(BytesIO(image_bytes))
            orig_width, orig_height = img.size
            
            # 크롭 영역 계산 (픽셀 단위)
            left = int(orig_width * crop['left'])
            top = int(orig_height * crop['top'])
            right = int(orig_width * (1 - crop['right']))
            bottom = int(orig_height * (1 - crop['bottom']))
            
            # 유효성 검사
            if left >= right or top >= bottom:
                logger.debug("크롭 영역이 유효하지 않음")
                return image_bytes
            
            # 이미지 크롭
            cropped_img = img.crop((left, top, right, bottom))
            
            # 바이트로 변환
            output = BytesIO()
            # 원본 포맷 유지 (지원되지 않으면 PNG)
            img_format = img.format if img.format else 'PNG'
            if img_format.upper() == 'JPEG':
                # RGBA면 RGB로 변환
                if cropped_img.mode == 'RGBA':
                    cropped_img = cropped_img.convert('RGB')
                cropped_img.save(output, format='JPEG', quality=95)
            else:
                cropped_img.save(output, format=img_format)
            
            logger.debug(
                f"이미지 크롭 완료: {orig_width}x{orig_height} -> "
                f"{right-left}x{bottom-top}"
            )
            
            return output.getvalue()
            
        except Exception as e:
            logger.warning(f"이미지 크롭 실패, 원본 사용: {e}")
            return image_bytes
    
    def _apply_highlight_style(self, run: Any):
        """강조 스타일 적용"""
        run.font.bold = True
        run.font.color.rgb = RGBColor(0, 102, 153)  # 진한 청록색
        run.font.size = Pt(12)
    
    def _apply_keyword_highlighting(self, para: Any):
        """단락 내 키워드 강조"""
        if not self.highlight_keywords:
            return
        
        for run in para.runs:
            if is_highlight_keyword(run.text):
                self._apply_highlight_style(run)
    
    def _add_landscape_section(self, doc: DocxDocument):
        """가로 방향 새 섹션 추가"""
        # 새 섹션 시작
        new_section = doc.add_section()
        
        # 가로 방향 설정
        new_section.orientation = WD_ORIENT.LANDSCAPE
        
        # 페이지 크기 조정 (A4 가로)
        new_section.page_width = Cm(29.7)
        new_section.page_height = Cm(21.0)
        
        # 여백 설정
        new_section.top_margin = Cm(1.5)
        new_section.bottom_margin = Cm(1.5)
        new_section.left_margin = Cm(2)
        new_section.right_margin = Cm(2)
    
    def _setup_document_styles(self, doc: DocxDocument):
        """문서 스타일 설정"""
        # 기본 섹션 여백
        for section in doc.sections:
            section.top_margin = Cm(2)
            section.bottom_margin = Cm(2)
            section.left_margin = Cm(2.5)
            section.right_margin = Cm(2.5)
        
        styles = doc.styles
        
        # Heading 1 스타일
        heading1 = styles['Heading 1']
        heading1.font.size = Pt(18)
        heading1.font.bold = True
        heading1.font.color.rgb = RGBColor(0, 51, 102)
        heading1.paragraph_format.space_before = Pt(12)
        heading1.paragraph_format.space_after = Pt(6)
        
        # Heading 2 스타일
        heading2 = styles['Heading 2']
        heading2.font.size = Pt(14)
        heading2.font.bold = True
        heading2.font.color.rgb = RGBColor(0, 102, 153)
        heading2.paragraph_format.space_before = Pt(10)
        heading2.paragraph_format.space_after = Pt(4)
        
        # 강조 스타일 (Intense Emphasis)
        try:
            emphasis = styles['Intense Emphasis']
            emphasis.font.bold = True
            emphasis.font.color.rgb = RGBColor(0, 102, 153)
        except KeyError:
            pass
    
    def _copy_metadata(self, prs: Presentation, doc: DocxDocument):
        """메타데이터 복사"""
        src = prs.core_properties
        dst = doc.core_properties
        
        if src.title:
            dst.title = src.title
        if src.author:
            dst.author = src.author
        if src.subject:
            dst.subject = src.subject
        if src.keywords:
            dst.keywords = src.keywords
        if src.comments:
            dst.comments = src.comments
    
    def _get_slide_title(self, slide: Any) -> Optional[str]:
        """슬라이드 제목 추출"""
        if slide.shapes.title and slide.shapes.title.text.strip():
            return sanitize_text(slide.shapes.title.text.strip())
        return None
    
    def _get_shape_info(self, shape: Any, parent_top: int = 0, parent_left: int = 0) -> Optional[dict]:
        """shape의 타입과 위치 정보 추출"""
        try:
            top = (shape.top if hasattr(shape, 'top') else 0) + parent_top
            left = (shape.left if hasattr(shape, 'left') else 0) + parent_left
        except Exception:
            top = parent_top
            left = parent_left
        
        shape_type = None
        
        if shape.shape_type == MSO_SHAPE_TYPE.GROUP:
            shape_type = 'group'
        elif shape.has_table:
            shape_type = 'table'
        elif hasattr(shape, 'image'):
            shape_type = 'image'
        elif hasattr(shape, 'text_frame') and shape.text_frame.text.strip():
            shape_type = 'text'
        elif hasattr(shape, 'text') and shape.text.strip():
            shape_type = 'text'
        
        if shape_type:
            return {
                'shape': shape,
                'type': shape_type,
                'top': top,
                'left': left,
            }
        
        return None
    
    def _copy_paragraph_style(self, src_paragraph: Any, dst_paragraph: Any):
        """단락 스타일 복사"""
        try:
            if src_paragraph.runs:
                src_run = src_paragraph.runs[0]
                if dst_paragraph.runs:
                    dst_run = dst_paragraph.runs[0]
                    
                    if src_run.font.bold:
                        dst_run.font.bold = True
                    if src_run.font.italic:
                        dst_run.font.italic = True
                    if src_run.font.size:
                        size_pt = src_run.font.size.pt if src_run.font.size else 11
                        dst_run.font.size = Pt(min(size_pt, 14))
        except Exception as e:
            logger.debug(f"스타일 복사 실패: {e}")
    
    def _add_horizontal_line(self, doc: DocxDocument, color: str = '003366'):
        """구분선 추가"""
        para = doc.add_paragraph()
        para.paragraph_format.space_before = Pt(0)
        para.paragraph_format.space_after = Pt(6)
        
        p = para._p
        pPr = p.get_or_add_pPr()
        pBdr = OxmlElement('w:pBdr')
        bottom = OxmlElement('w:bottom')
        bottom.set(qn('w:val'), 'single')
        bottom.set(qn('w:sz'), '6')
        bottom.set(qn('w:space'), '1')
        bottom.set(qn('w:color'), color)
        pBdr.append(bottom)
        pPr.append(pBdr)
