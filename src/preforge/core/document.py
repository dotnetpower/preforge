"""
Document base class and data models
"""
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, List, Optional
from enum import Enum


class DocumentType(Enum):
    """Supported document types"""
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
    """Document metadata"""
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
    """Text content"""
    text: str
    level: int = 0  # Heading level (0: body, 1~6: heading)
    style: Optional[str] = None
    page_number: Optional[int] = None
    position: Optional[int] = None  # top position (absolute coordinate)
    left: Optional[int] = None  # left position (absolute coordinate)


@dataclass
class CellImage:
    """Image information within table cell"""
    row: int
    col: int
    data: bytes
    format: str
    width: int
    height: int
    embed_id: Optional[str] = None  # Image relationship ID (for duplicate check)


@dataclass
class CellMerge:
    """Cell merge information"""
    row: int
    col: int
    colspan: int = 1  # Horizontal merge (gridSpan)
    rowspan: int = 1  # Vertical merge (vMerge)
    is_merged: bool = False  # Whether part of merged cell (not displayed)


@dataclass
class TableContent:
    """Table content"""
    headers: List[str]
    rows: List[List[str]]
    caption: Optional[str] = None
    page_number: Optional[int] = None
    cell_images: List[CellImage] = field(default_factory=list)
    cell_merges: List[CellMerge] = field(default_factory=list)  # Cell merge information


@dataclass
class ImageContent:
    """Image content"""
    data: bytes
    format: str
    width: Optional[int] = None
    height: Optional[int] = None
    caption: Optional[str] = None
    page_number: Optional[int] = None
    position: Optional[int] = None  # top position (absolute coordinate)
    left: Optional[int] = None  # left position (absolute coordinate)


@dataclass
class GridCell:
    """Grid cell information"""
    row: int  # Row number (0-based)
    col: int  # Column number (0-based)
    top: int  # Top position (EMU)
    left: int  # Left position (EMU)
    width: int  # Width (EMU)
    height: int  # Height (EMU)
    content_ids: List[str] = field(default_factory=list)  # List of included content IDs
    color: Optional[str] = None  # Color for visualization
    colspan: int = 1  # Column merge (1 means no merge)
    rowspan: int = 1  # Row merge (1 means no merge)


@dataclass
class PageLayout:
    """Page layout information"""
    page_number: int
    rows: int  # Number of rows (1-3)
    cols: int  # Number of columns (1-3)
    grid_cells: List[GridCell] = field(default_factory=list)
    slide_width: int = 9144000  # Standard 16:9 slide width (EMU)
    slide_height: int = 5143500  # Standard 16:9 slide height (EMU)


@dataclass
class Document:
    """Document object"""
    file_path: Path
    doc_type: DocumentType
    metadata: DocumentMetadata = field(default_factory=DocumentMetadata)
    text_contents: List[TextContent] = field(default_factory=list)
    tables: List[TableContent] = field(default_factory=list)
    images: List[ImageContent] = field(default_factory=list)
    page_layouts: List[PageLayout] = field(default_factory=list)  # Page layout information
    raw_content: Optional[Any] = None
    
    @property
    def full_text(self) -> str:
        """Return all text as a single string"""
        return "\n".join(tc.text for tc in self.text_contents)
    
    @property
    def headings(self) -> List[TextContent]:
        """Extract only headings"""
        return [tc for tc in self.text_contents if tc.level > 0]
    
    @property
    def body_text(self) -> str:
        """Return only body text"""
        return "\n".join(tc.text for tc in self.text_contents if tc.level == 0)
    
    def to_dict(self) -> Dict[str, Any]:
        """Convert document to dictionary"""
        return {
            "file_path": str(self.file_path),
            "doc_type": self.doc_type.value,
            "metadata": self.metadata.__dict__,
            "text_count": len(self.text_contents),
            "table_count": len(self.tables),
            "image_count": len(self.images),
            "full_text": self.full_text[:500],  # First 500 characters only
        }
