"""
PDF document parser
"""
from pathlib import Path
from typing import List
import pdfplumber
from pypdf import PdfReader
import logging

from ..core.document import (
    Document,
    DocumentType,
    DocumentMetadata,
    TextContent,
    TableContent,
    ImageContent,
)
from ..core.parser import BaseParser

logger = logging.getLogger(__name__)


class PdfParser(BaseParser):
    """PDF document parser"""
    
    @property
    def supported_extensions(self) -> List[str]:
        return [".pdf"]
    
    @property
    def document_type(self) -> DocumentType:
        return DocumentType.PDF
    
    def parse(self, file_path: Path) -> Document:
        """Parse PDF document"""
        self.validate_file(file_path)
        
        # Extract metadata using pypdf
        reader = PdfReader(file_path)
        metadata = self._extract_metadata(reader)
        
        # Extract text and tables using pdfplumber (more accurate)
        with pdfplumber.open(file_path) as pdf:
            text_contents = self._extract_text(pdf)
            tables = self._extract_tables(pdf)
        
        # Extract images
        images = self._extract_images(reader)
        
        return Document(
            file_path=file_path,
            doc_type=self.document_type,
            metadata=metadata,
            text_contents=text_contents,
            tables=tables,
            images=images,
            raw_content=reader,
        )
    
    def _extract_metadata(self, reader: PdfReader) -> DocumentMetadata:
        """Extract metadata"""
        meta = reader.metadata
        
        if not meta:
            return DocumentMetadata(page_count=len(reader.pages))
        
        return DocumentMetadata(
            title=meta.get("/Title"),
            author=meta.get("/Author"),
            subject=meta.get("/Subject"),
            keywords=meta.get("/Keywords", "").split(",") if meta.get("/Keywords") else None,
            page_count=len(reader.pages),
            properties={
                "creator": meta.get("/Creator"),
                "producer": meta.get("/Producer"),
                "creation_date": meta.get("/CreationDate"),
                "modification_date": meta.get("/ModDate"),
            }
        )
    
    def _extract_text(self, pdf: pdfplumber.PDF) -> List[TextContent]:
        """Extract text (coordinate-based, estimate heading level from font size)"""
        text_contents = []
        
        for page_num, page in enumerate(pdf.pages, 1):
            # Page height (for coordinate conversion)
            page_height = page.height
            
            # Extract character-by-character to utilize font info
            chars = page.chars
            if not chars:
                # If no chars, use basic text extraction
                text = page.extract_text()
                if text and text.strip():
                    text_contents.append(
                        TextContent(
                            text=text,
                            level=0,
                            page_number=page_num,
                            position=0,
                        )
                    )
                continue
            
            # Group by line (based on y coordinate)
            lines = {}
            for char in chars:
                # Convert PDF coordinate system (bottom-left origin) -> top-based
                y = page_height - char['top']
                x = char['x0']
                
                # Group as same line (y coordinate difference < 2)
                line_key = int(y / 2)
                if line_key not in lines:
                    lines[line_key] = {
                        'chars': [],
                        'y': y,
                        'x_min': x,
                        'font_size': char.get('size', 12),
                    }
                
                lines[line_key]['chars'].append(char)
                lines[line_key]['x_min'] = min(lines[line_key]['x_min'], x)
                lines[line_key]['font_size'] = max(lines[line_key]['font_size'], char.get('size', 12))
            
            # Convert lines to text and estimate heading level from font size
            for line_key in sorted(lines.keys()):
                line_info = lines[line_key]
                text = ''.join(c['text'] for c in line_info['chars']).strip()
                
                if not text:
                    continue
                
                # Estimate heading level from font size
                font_size = line_info['font_size']
                level = 0
                
                if font_size >= 18:
                    level = 1  # H1
                elif font_size >= 16:
                    level = 2  # H2
                elif font_size >= 14:
                    level = 3  # H3
                elif font_size >= 13:
                    level = 4  # H4
                
                # Short text + large font = likely a heading
                if len(text) < 100 and font_size > 12 and level == 0:
                    level = 5
                
                text_contents.append(
                    TextContent(
                        text=text,
                        level=level,
                        page_number=page_num,
                        position=int(line_info['y']),
                        left=int(line_info['x_min']),
                    )
                )
        
        return text_contents
    
    def _extract_tables(self, pdf: pdfplumber.PDF) -> List[TableContent]:
        """Extract tables"""
        tables = []
        
        for page_num, page in enumerate(pdf.pages, 1):
            page_tables = page.extract_tables()
            
            if not page_tables:
                continue
            
            for table_data in page_tables:
                if not table_data or len(table_data) < 2:
                    continue
                
                # First row is considered header
                headers = [str(cell).strip() if cell else "" for cell in table_data[0]]
                
                # Extract remaining rows as data
                rows = []
                for row in table_data[1:]:
                    row_data = [str(cell).strip() if cell else "" for cell in row]
                    rows.append(row_data)
                
                tables.append(
                    TableContent(
                        headers=headers,
                        rows=rows,
                        page_number=page_num,
                    )
                )
        
        return tables
    
    def _extract_images(self, reader: PdfReader) -> List[ImageContent]:
        """Extract images"""
        images = []
        
        for page_num, page in enumerate(reader.pages, 1):
            if "/XObject" not in page["/Resources"]:
                continue
            
            xobjects = page["/Resources"]["/XObject"].get_object()
            
            for obj_name in xobjects:
                obj = xobjects[obj_name]
                
                if obj["/Subtype"] != "/Image":
                    continue
                
                try:
                    # Extract image data
                    data = obj.get_data()
                    
                    # Extract image format
                    if "/Filter" in obj:
                        filter_type = obj["/Filter"]
                        if filter_type == "/DCTDecode":
                            image_format = "jpeg"
                        elif filter_type == "/FlateDecode":
                            image_format = "png"
                        else:
                            image_format = "unknown"
                    else:
                        image_format = "unknown"
                    
                    width = obj.get("/Width")
                    height = obj.get("/Height")
                    
                    images.append(
                        ImageContent(
                            data=data,
                            format=image_format,
                            width=width,
                            height=height,
                            page_number=page_num,
                        )
                    )
                except Exception:
                    continue
        
        return images
