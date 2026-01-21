"""
HTML 문서 파서
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
    """HTML 문서 파서"""
    
    @property
    def supported_extensions(self) -> List[str]:
        return [".html", ".htm"]
    
    @property
    def document_type(self) -> DocumentType:
        return DocumentType.HTML
    
    def parse(self, file_path: Path) -> Document:
        """HTML 문서 파싱"""
        self.validate_file(file_path)
        
        with open(file_path, "r", encoding="utf-8") as f:
            html_content = f.read()
        
        soup = BeautifulSoup(html_content, "lxml")
        
        # 메타데이터 추출
        metadata = self._extract_metadata(soup)
        
        # 텍스트 추출
        text_contents = self._extract_text(soup)
        
        # 테이블 추출
        tables = self._extract_tables(soup)
        
        # 이미지 추출
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
        """메타데이터 추출"""
        title = soup.find("title")
        
        # meta 태그에서 정보 추출
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
        """텍스트 추출 (시맨틱 태그 구조 보존, 위치 추적)"""
        text_contents = []
        position = 0
        
        # script, style 태그 제거
        for element in soup.find_all(['script', 'style']):
            element.decompose()
        
        # 시맨틱 태그별로 처리
        semantic_tags = {
            'header': '헤더',
            'nav': '네비게이션',
            'main': '메인',
            'article': '아티클',
            'section': '섹션',
            'aside': '사이드바',
            'footer': '푸터',
        }
        
        for tag_name, tag_label in semantic_tags.items():
            for element in soup.find_all(tag_name):
                # 시맨틱 태그 구분 표시
                text_contents.append(
                    TextContent(
                        text=f"=== {tag_label} ===",
                        level=0,
                        style=tag_name,
                        position=position,
                    )
                )
                position += 100
        
        # 제목 태그 추출 (h1 ~ h6)
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
        
        # 단락 추출
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
        
        # 리스트 항목
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
        """테이블 추출 (colspan, rowspan 지원)"""
        tables = []
        
        for table in soup.find_all("table"):
            headers = []
            rows = []
            
            # 헤더 추출 (thead 또는 첫 번째 tr)
            thead = table.find("thead")
            if thead:
                header_row = thead.find("tr")
                if header_row:
                    for th in header_row.find_all(["th", "td"]):
                        text = th.get_text().strip().replace('\n', '<br>')
                        # colspan 처리
                        colspan = int(th.get('colspan', 1))
                        headers.extend([text] * colspan)
            else:
                # thead가 없으면 첫 번째 행을 헤더로 간주
                first_row = table.find("tr")
                if first_row:
                    for cell in first_row.find_all(["th", "td"]):
                        text = cell.get_text().strip().replace('\n', '<br>')
                        colspan = int(cell.get('colspan', 1))
                        headers.extend([text] * colspan)
            
            # 데이터 행 추출
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
                # 캡션 추출
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
        """이미지 추출 (로컬/원격/Base64 지원)"""
        images = []
        position = 0
        
        for img in soup.find_all("img"):
            src = img.get("src")
            if not src:
                continue
            
            try:
                image_data = None
                image_format = "unknown"
                
                # Base64 인라인 이미지
                if src.startswith("data:"):
                    # data:image/png;base64,iVBORw0KG...
                    parts = src.split(",", 1)
                    if len(parts) == 2:
                        header = parts[0]
                        data_part = parts[1]
                        
                        # 이미지 포맷 추출
                        if "image/" in header:
                            image_format = header.split("image/")[1].split(";")[0]
                        
                        # Base64 디코딩
                        image_data = base64.b64decode(data_part)
                
                # 원격 URL (선택적 다운로드)
                elif src.startswith(("http://", "https://")):
                    logger.info(f"원격 이미지 발견 (다운로드 생략): {src}")
                    # 다운로드를 원하면 아래 주석 해제
                    # try:
                    #     response = requests.get(src, timeout=5)
                    #     if response.status_code == 200:
                    #         image_data = response.content
                    #         image_format = src.split('.')[-1].split('?')[0]
                    # except:
                    #     pass
                
                # 로컬 파일 경로
                else:
                    img_path = file_path.parent / src
                    if img_path.exists():
                        with open(img_path, "rb") as f:
                            image_data = f.read()
                        image_format = img_path.suffix[1:] if img_path.suffix else "unknown"
                    else:
                        logger.warning(f"이미지 파일을 찾을 수 없음: {img_path}")
                
                if image_data:
                    # 이미지 크기 추출
                    width = img.get("width", 0)
                    height = img.get("height", 0)
                    if width:
                        width = int(width) if str(width).isdigit() else 0
                    if height:
                        height = int(height) if str(height).isdigit() else 0
                    
                    # alt 텍스트
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
                logger.warning(f"이미지 추출 실패: {src}, 오류: {e}")
                continue
            
            position += 100
        
        return images

