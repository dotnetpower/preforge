"""
문서 기본 클래스 및 데이터 모델
"""
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, List, Optional
from enum import Enum


class DocumentType(Enum):
    """지원하는 문서 타입"""
    DOCX = "docx"
    PPTX = "pptx"
    XLSX = "xlsx"
    PDF = "pdf"
    HTML = "html"
    MARKDOWN = "markdown"
    TXT = "txt"
    UNKNOWN = "unknown"


@dataclass
class DocumentMetadata:
    """문서 메타데이터"""
    title: Optional[str] = None
    author: Optional[str] = None
    created_at: Optional[datetime] = None
    modified_at: Optional[datetime] = None
    subject: Optional[str] = None
    keywords: Optional[List[str]] = None
    page_count: Optional[int] = None
    word_count: Optional[int] = None
    properties: Dict[str, Any] = field(default_factory=dict)


@dataclass
class TextContent:
    """텍스트 컨텐츠"""
    text: str
    level: int = 0  # 제목 레벨 (0: 본문, 1~6: 제목)
    style: Optional[str] = None
    page_number: Optional[int] = None
    position: Optional[int] = None  # top 위치 (절대 좌표)
    left: Optional[int] = None  # left 위치 (절대 좌표)  # top 위치 (절대 좌표)
    left: Optional[int] = None  # left 위치 (절대 좌표)


@dataclass
class CellImage:
    """테이블 셀 내 이미지 정보"""
    row: int
    col: int
    data: bytes
    format: str
    width: int
    height: int
    embed_id: Optional[str] = None  # 이미지 관계 ID (중복 체크용)


@dataclass
class CellMerge:
    """셀 병합 정보"""
    row: int
    col: int
    colspan: int = 1  # 가로 병합 (gridSpan)
    rowspan: int = 1  # 세로 병합 (vMerge)
    is_merged: bool = False  # 병합된 셀의 일부인지 (표시 안함)


@dataclass
class TableContent:
    """테이블 컨텐츠"""
    headers: List[str]
    rows: List[List[str]]
    caption: Optional[str] = None
    page_number: Optional[int] = None
    cell_images: List[CellImage] = field(default_factory=list)
    cell_merges: List[CellMerge] = field(default_factory=list)  # 셀 병합 정보


@dataclass
class ImageContent:
    """이미지 컨텐츠"""
    data: bytes
    format: str
    width: Optional[int] = None
    height: Optional[int] = None
    caption: Optional[str] = None
    page_number: Optional[int] = None
    position: Optional[int] = None  # top 위치 (절대 좌표)
    left: Optional[int] = None  # left 위치 (절대 좌표)


@dataclass
class GridCell:
    """그리드 셀 정보"""
    row: int  # 행 번호 (0-based)
    col: int  # 열 번호 (0-based)
    top: int  # 상단 위치 (EMU)
    left: int  # 좌측 위치 (EMU)
    width: int  # 너비 (EMU)
    height: int  # 높이 (EMU)
    content_ids: List[str] = field(default_factory=list)  # 포함된 컨텐츠 ID 리스트
    color: Optional[str] = None  # 시각화용 색상
    colspan: int = 1  # 열 병합 (1이면 병합 없음)
    rowspan: int = 1  # 행 병합 (1이면 병합 없음)


@dataclass
class PageLayout:
    """페이지 레이아웃 정보"""
    page_number: int
    rows: int  # 행 개수 (1-3)
    cols: int  # 열 개수 (1-3)
    grid_cells: List[GridCell] = field(default_factory=list)
    slide_width: int = 9144000  # 표준 16:9 슬라이드 너비 (EMU)
    slide_height: int = 5143500  # 표준 16:9 슬라이드 높이 (EMU)


@dataclass
class Document:
    """문서 객체"""
    file_path: Path
    doc_type: DocumentType
    metadata: DocumentMetadata = field(default_factory=DocumentMetadata)
    text_contents: List[TextContent] = field(default_factory=list)
    tables: List[TableContent] = field(default_factory=list)
    images: List[ImageContent] = field(default_factory=list)
    page_layouts: List[PageLayout] = field(default_factory=list)  # 페이지별 레이아웃 정보
    raw_content: Optional[Any] = None
    
    @property
    def full_text(self) -> str:
        """모든 텍스트를 하나의 문자열로 반환"""
        return "\n".join(tc.text for tc in self.text_contents)
    
    @property
    def headings(self) -> List[TextContent]:
        """제목만 추출"""
        return [tc for tc in self.text_contents if tc.level > 0]
    
    @property
    def body_text(self) -> str:
        """본문 텍스트만 반환"""
        return "\n".join(tc.text for tc in self.text_contents if tc.level == 0)
    
    def to_dict(self) -> Dict[str, Any]:
        """문서를 딕셔너리로 변환"""
        return {
            "file_path": str(self.file_path),
            "doc_type": self.doc_type.value,
            "metadata": self.metadata.__dict__,
            "text_count": len(self.text_contents),
            "table_count": len(self.tables),
            "image_count": len(self.images),
            "full_text": self.full_text[:500],  # 처음 500자만
        }
