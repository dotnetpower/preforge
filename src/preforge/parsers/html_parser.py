"""
HTML document parser
"""
from pathlib import Path
from typing import List
import base64
import logging
from bs4 import BeautifulSoup
from urllib.parse import urljoin, urlparse
import requests

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


class HtmlParser(BaseParser):
    """HTML document parser"""
    
    @property
    def supported_extensions(self) -> List[str]:
        return [".html", ".htm"]
    
    @property
    def document_type(self) -> DocumentType:
        return DocumentType.HTML
    
    def parse(self, file_path: Path) -> Document:
        """Parse HTML document"""
        self.validate_file(file_path)
        
        with open(file_path, "r", encoding="utf-8") as f:
            html_content = f.read()
        
        soup = BeautifulSoup(html_content, "lxml")
        
        # Extract metadata
        metadata = self._extract_metadata(soup)
        
        # Extract text
        text_contents = self._extract_text(soup)
        
        # Extract tables
        tables = self._extract_tables(soup)
        
        # Extract images
        images = self._extract_images(soup, file_path)
        
        return Document(
            file_path=file_path,
            doc_type=self.document_type,
            metadata=metadata,
            text_contents=text_contents,
            tables=tables,
            images=images,
            raw_content=soup,
        )
    
    def _extract_metadata(self, soup: BeautifulSoup) -> DocumentMetadata:
        """Extract metadata"""
        title = soup.find("title")
        
        # Extract info from meta tags
        meta_author = soup.find("meta", attrs={"name": "author"})
        meta_description = soup.find("meta", attrs={"name": "description"})
        meta_keywords = soup.find("meta", attrs={"name": "keywords"})
        
        return DocumentMetadata(
            title=title.text if title else None,
            author=meta_author.get("content") if meta_author else None,
            subject=meta_description.get("content") if meta_description else None,
            keywords=meta_keywords.get("content").split(",") if meta_keywords else None,
        )
    
    def _extract_text(self, soup: BeautifulSoup) -> List[TextContent]:
        """Extract text (preserve semantic tag structure, track position)"""
        text_contents = []
        position = 0
        
        # Remove script, style tags
        for element in soup.find_all(['script', 'style']):
            element.decompose()
        
        # Process by semantic tags
        semantic_tags = {
            'header': 'Header',
            'nav': 'Navigation',
            'main': 'Main',
            'article': 'Article',
            'section': 'Section',
            'aside': 'Sidebar',
            'footer': 'Footer',
        }
        
        for tag_name, tag_label in semantic_tags.items():
            for element in soup.find_all(tag_name):
                # Mark semantic tag boundaries
                text_contents.append(
                    TextContent(
                        text=f"=== {tag_label} ===",
                        level=0,
                        style=tag_name,
                        position=position,
                    )
                )
                position += 100
        
        # Extract heading tags (h1 ~ h6)
        for level in range(1, 7):
            for header in soup.find_all(f"h{level}"):
                text = header.get_text().strip()
                if text:
                    text_contents.append(
                        TextContent(
                            text=text,
                            level=level,
                            style=f"h{level}",
                            position=position,
                        )
                    )
                    position += 100
        
        # Paragraph extraction
        for para in soup.find_all("p"):
            text = para.get_text().strip()
            if text:
                text_contents.append(
                    TextContent(
                        text=text,
                        level=0,
                        style="p",
                        position=position,
                    )
                )
                position += 50
        
        # List items
        for item in soup.find_all("li"):
            text = item.get_text().strip()
            if text:
                text_contents.append(
                    TextContent(
                        text=f"• {text}",
                        level=0,
                        style="li",
                        position=position,
                    )
                )
                position += 30
        
        # blockquote
        for quote in soup.find_all("blockquote"):
            text = quote.get_text().strip()
            if text:
                text_contents.append(
                    TextContent(
                        text=f"> {text}",
                        level=0,
                        style="blockquote",
                        position=position,
                    )
                )
                position += 50
        
        return text_contents
    
    def _extract_tables(self, soup: BeautifulSoup) -> List[TableContent]:
        """Extract tables (colspan, rowspan supported)"""
        tables = []
        
        for table in soup.find_all("table"):
            headers = []
            rows = []
            
            # Extract header (thead or first tr)
            thead = table.find("thead")
            if thead:
                header_row = thead.find("tr")
                if header_row:
                    for th in header_row.find_all(["th", "td"]):
                        text = th.get_text().strip().replace('\n', '<br>')
                        # colspan handling
                        colspan = int(th.get('colspan', 1))
                        headers.extend([text] * colspan)
            else:
                # If no thead, treat first row as header
                first_row = table.find("tr")
                if first_row:
                    for cell in first_row.find_all(["th", "td"]):
                        text = cell.get_text().strip().replace('\n', '<br>')
                        colspan = int(cell.get('colspan', 1))
                        headers.extend([text] * colspan)
            
            # Extract data rows
            tbody = table.find("tbody")
            row_elements = tbody.find_all("tr") if tbody else table.find_all("tr")[1:]
            
            for row in row_elements:
                row_data = []
                for cell in row.find_all(["td", "th"]):
                    text = cell.get_text().strip().replace('\n', '<br>')
                    colspan = int(cell.get('colspan', 1))
                    row_data.extend([text] * colspan)
                
                if row_data:
                    rows.append(row_data)
            
            if headers or rows:
                # Extract caption
                caption = table.find("caption")
                caption_text = caption.get_text().strip() if caption else None
                
                tables.append(
                    TableContent(
                        headers=headers if headers else [],
                        rows=rows,
                        caption=caption_text,
                    )
                )
        
        return tables
        
        return tables
    
    def _extract_images(self, soup: BeautifulSoup, file_path: Path) -> List[ImageContent]:
        """Extract images (local/remote/Base64 supported)"""
        images = []
        position = 0
        
        for img in soup.find_all("img"):
            src = img.get("src")
            if not src:
                continue
            
            try:
                image_data = None
                image_format = "unknown"
                
                # Base64 inline image
                if src.startswith("data:"):
                    # data:image/png;base64,iVBORw0KG...
                    parts = src.split(",", 1)
                    if len(parts) == 2:
                        header = parts[0]
                        data_part = parts[1]
                        
                        # Extract image format
                        if "image/" in header:
                            image_format = header.split("image/")[1].split(";")[0]
                        
                        # Base64 decode
                        image_data = base64.b64decode(data_part)
                
                # Remote URL (optional download)
                elif src.startswith(("http://", "https://")):
                    logger.info(f"Remote image found (skipping download): {src}")
                    # Uncomment below to enable download
                    # try:
                    #     response = requests.get(src, timeout=5)
                    #     if response.status_code == 200:
                    #         image_data = response.content
                    #         image_format = src.split('.')[-1].split('?')[0]
                    # except:
                    #     pass
                
                # Local file path
                else:
                    img_path = file_path.parent / src
                    if img_path.exists():
                        with open(img_path, "rb") as f:
                            image_data = f.read()
                        image_format = img_path.suffix[1:] if img_path.suffix else "unknown"
                    else:
                        logger.warning(f"Image file not found: {img_path}")
                
                if image_data:
                    # Extract image dimensions
                    width = img.get("width", 0)
                    height = img.get("height", 0)
                    if width:
                        width = int(width) if str(width).isdigit() else 0
                    if height:
                        height = int(height) if str(height).isdigit() else 0
                    
                    # alt text
                    caption = img.get("alt")
                    
                    images.append(
                        ImageContent(
                            data=image_data,
                            format=image_format,
                            width=width,
                            height=height,
                            caption=caption,
                            position=position,
                        )
                    )
            except Exception as e:
                logger.warning(f"Failed to extract image: {src}, error: {e}")
                continue
            
            position += 100
        
        return images

