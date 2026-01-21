"""
Word 문서(.docx) 파서
"""
from pathlib import Path
from typing import List, Optional
import docx
from docx.document import Document as DocxDocument
from docx.table import Table as DocxTable
from docx.oxml.text.paragraph import CT_P
from docx.oxml.table import CT_Tbl
from docx.text.paragraph import Paragraph
from docx.oxml.ns import qn
import logging

from ..core.document import (
    Document,
    DocumentType,
    DocumentMetadata,
    TextContent,
    TableContent,
    ImageContent,
    CellImage,
    CellMerge,
)
from ..core.parser import BaseParser

logger = logging.getLogger(__name__)


class DocxParser(BaseParser):
    """Word 문서 파서"""
    
    def __init__(self):
        super().__init__()
        self._numbering_counters = {}  # numId별 카운터 추적
    
    def _get_paragraph_number(self, paragraph: Paragraph) -> Optional[str]:
        """단락의 자동 번호 매기기 번호를 가져옵니다"""
        try:
            if paragraph._element.pPr is None:
                return None
            
            numPr = paragraph._element.pPr.numPr
            if numPr is None:
                return None
            
            # 번호 매기기 레벨 및 ID 가져오기
            ilvl_element = numPr.ilvl
            numId_element = numPr.numId
            
            if ilvl_element is None or numId_element is None:
                return None
            
            ilvl = ilvl_element.val
            numId = numId_element.val
            
            # numId가 0이면 번호 매기기가 없음
            if numId == 0:
                return None
            
            # numId별로 카운터 추적
            counter_key = (numId, ilvl)
            if counter_key not in self._numbering_counters:
                self._numbering_counters[counter_key] = 0
            
            self._numbering_counters[counter_key] += 1
            counter = self._numbering_counters[counter_key]
            
            # 레벨에 따라 번호 형식 결정
            if ilvl == 0:
                return f"{counter}."
            elif ilvl == 1:
                # 상위 레벨 카운터 가져오기
                parent_key = (numId, ilvl - 1)
                parent_counter = self._numbering_counters.get(parent_key, 1)
                return f"{parent_counter}.{counter}"
            elif ilvl == 2:
                # 2단계 상위 카운터
                parent_key = (numId, ilvl - 1)
                grandparent_key = (numId, ilvl - 2)
                parent_counter = self._numbering_counters.get(parent_key, 1)
                grandparent_counter = self._numbering_counters.get(grandparent_key, 1)
                return f"{grandparent_counter}.{parent_counter}.{counter}"
            else:
                return f"[{counter}]"
            
        except Exception as e:
            logger.debug(f"번호 매기기 정보 추출 실패: {e}")
            return None
    
    @property
    def supported_extensions(self) -> List[str]:
        return [".docx", ".doc"]
    
    @property
    def document_type(self) -> DocumentType:
        return DocumentType.DOCX
    
    def parse(self, file_path: Path) -> Document:
        """Word 문서 파싱"""
        self.validate_file(file_path)
        
        docx_doc = docx.Document(file_path)
        
        # 메타데이터 추출
        metadata = self._extract_metadata(docx_doc)
        
        # 텍스트 추출
        text_contents = self._extract_text(docx_doc)
        
        # 테이블 추출
        tables = self._extract_tables(docx_doc)
        
        # 이미지 추출
        images = self._extract_images(docx_doc)
        
        return Document(
            file_path=file_path,
            doc_type=self.document_type,
            metadata=metadata,
            text_contents=text_contents,
            tables=tables,
            images=images,
            raw_content=docx_doc,
        )
    
    def _extract_metadata(self, doc: DocxDocument) -> DocumentMetadata:
        """메타데이터 추출"""
        core_props = doc.core_properties
        
        # 페이지 수 계산 (섹션 기반 추정)
        page_count = len(doc.sections) if doc.sections else None
        
        return DocumentMetadata(
            title=core_props.title,
            author=core_props.author,
            created_at=core_props.created,
            modified_at=core_props.modified,
            subject=core_props.subject,
            keywords=core_props.keywords.split(",") if core_props.keywords else None,
            page_count=page_count,
            properties={
                "category": core_props.category,
                "comments": core_props.comments,
                "language": core_props.language,
                "section_count": len(doc.sections),
            }
        )
    
    def _extract_text(self, doc: DocxDocument) -> List[TextContent]:
        """텍스트 추출 (섹션, 머리글, 바닥글, 텍스트 박스 포함)"""
        text_contents = []
        current_page = 1
        
        # 섹션별로 처리
        for section_idx, section in enumerate(doc.sections, 1):
            # 머리글 추출
            if section.header:
                for para in section.header.paragraphs:
                    if para.text.strip():
                        text_contents.append(
                            TextContent(
                                text=f"[머리글] {para.text}",
                                level=0,
                                style="Header",
                                page_number=current_page,
                            )
                        )
            
            # 섹션 구분 표시
            if section_idx > 1:
                current_page = section_idx  # 섹션 변경 시 페이지 증가
                text_contents.append(
                    TextContent(
                        text=f"--- 섹션 {section_idx} ---",
                        level=0,
                        style="SectionBreak",
                        page_number=current_page,
                    )
                )
        
        # 본문 단락 추출 (위치 정보 포함)
        position = 0
        current_page = 1  # 페이지 초기화
        
        for element in doc.element.body:
            if isinstance(element, CT_P):
                para = Paragraph(element, doc)
                
                # 페이지 브레이크 확인 (텍스트가 없어도 체크)
                has_page_break = False
                try:
                    page_breaks = para._element.findall('.//' + qn('w:br'))
                    for br in page_breaks:
                        if br.get(qn('w:type')) == 'page':
                            current_page += 1
                            position += 2000  # 페이지 브레이크 후 위치 증가
                            has_page_break = True
                            break
                except:
                    pass
                
                if has_page_break:
                    continue
                
                # 빈 단락도 position 증가 (이미지 전용 단락 때문)
                if not para.text.strip():
                    position += 1000
                    continue
                
                # 스타일에서 제목 레벨 판단
                level = 0
                style_name = para.style.name if para.style else ""
                
                if "Heading" in style_name:
                    try:
                        level = int(style_name.split()[-1])
                    except (ValueError, IndexError):
                        level = 1
                
                # Drawing 객체(텍스트 박스, 도형) 확인
                has_drawing = False
                try:
                    if para._element.findall('.//' + qn('w:drawing')):
                        has_drawing = True
                        style_name = "Drawing"
                except:
                    pass
                
                # 자동 번호 매기기 확인
                number_prefix = self._get_paragraph_number(para)
                text = para.text
                if number_prefix:
                    text = f"{number_prefix} {text}"
                
                text_contents.append(
                    TextContent(
                        text=text,
                        level=level,
                        style=style_name,
                        page_number=current_page,
                        position=position,
                    )
                )
                position += 1000  # 단락 간 간격
            
            elif isinstance(element, CT_Tbl):
                # 테이블은 별도 처리되므로 위치만 증가
                position += 5000
        
        # 바닥글 추출
        for section_idx, section in enumerate(doc.sections, 1):
            if section.footer:
                for para in section.footer.paragraphs:
                    if para.text.strip():
                        text_contents.append(
                            TextContent(
                                text=f"[바닥글] {para.text}",
                                level=0,
                                style="Footer",
                                page_number=section_idx,
                            )
                        )
        
        return text_contents
    
    def _extract_tables(self, doc: DocxDocument) -> List[TableContent]:
        """테이블 추출 (병합 셀, 중첩 테이블, 셀 이미지 지원)"""
        tables = []
        current_page = 1
        
        # 페이지 브레이크 추적을 위해 전체 문서 순회
        table_page_map = {}  # {table_idx: page_number}
        table_idx = 0
        
        for element in doc.element.body:
            if isinstance(element, CT_P):
                para = Paragraph(element, doc)
                
                # 페이지 브레이크 확인
                try:
                    page_breaks = para._element.findall('.//' + qn('w:br'))
                    for br in page_breaks:
                        if br.get(qn('w:type')) == 'page':
                            current_page += 1
                            break
                except:
                    pass
            
            elif isinstance(element, CT_Tbl):
                table_page_map[table_idx] = current_page
                table_idx += 1
        
        for table_idx, table in enumerate(doc.tables):
            if not table.rows:
                continue
            
            # 첫 번째 행을 헤더로 간주
            headers = []
            header_merges = []
            for col_idx, cell in enumerate(table.rows[0].cells):
                headers.append(cell.text.strip().replace('\n', '<br>'))
                
                # 헤더의 colspan 추출
                try:
                    tc = cell._element
                    tcPr = tc.find(qn('w:tcPr'))
                    if tcPr is not None:
                        gridSpan = tcPr.find(qn('w:gridSpan'))
                        if gridSpan is not None:
                            colspan = int(gridSpan.get(qn('w:val'), 1))
                            if colspan > 1:
                                header_merges.append(CellMerge(
                                    row=0,
                                    col=col_idx,
                                    colspan=colspan,
                                    rowspan=1,
                                    is_merged=False
                                ))
                except:
                    pass
            
            # 나머지 행을 데이터로 추출
            rows = []
            cell_images = []
            cell_merges = header_merges.copy()
            seen_image_ids = set()  # 중복 이미지 방지
            
            # 1단계: 모든 행 데이터와 colspan 수집
            all_rows_data = []
            all_colspan_data = {}  # {(row, col): colspan}
            all_vmerge_data = {}  # {(row, col): 'restart' or 'continue'}
            
            for row_idx, row in enumerate(table.rows[1:], start=1):
                row_data = []
                for col_idx, cell in enumerate(row.cells):
                    cell_text = cell.text.strip().replace('\n', '<br>')
                    row_data.append(cell_text)
                    
                    # 셀 속성 수집
                    try:
                        tc = cell._element
                        tcPr = tc.find(qn('w:tcPr'))
                        if tcPr is not None:
                            # gridSpan (colspan)
                            gridSpan = tcPr.find(qn('w:gridSpan'))
                            if gridSpan is not None:
                                colspan = int(gridSpan.get(qn('w:val'), 1))
                                if colspan > 1:
                                    all_colspan_data[(row_idx, col_idx)] = colspan
                            
                            # vMerge (rowspan)
                            vMerge = tcPr.find(qn('w:vMerge'))
                            if vMerge is not None:
                                val = vMerge.get(qn('w:val'))
                                if val == 'restart':
                                    all_vmerge_data[(row_idx, col_idx)] = 'restart'
                                else:
                                    all_vmerge_data[(row_idx, col_idx)] = 'continue'
                    except:
                        pass
                
                all_rows_data.append(row_data)
            
            # 2단계: vMerge 정보로 rowspan 계산
            vmerge_spans = {}  # {(start_row, col): rowspan}
            for col_idx in range(len(table.rows[0].cells)):
                current_start = None
                for row_idx in range(1, len(table.rows)):
                    if (row_idx, col_idx) in all_vmerge_data:
                        if all_vmerge_data[(row_idx, col_idx)] == 'restart':
                            # 이전 병합 종료 및 새로운 시작
                            if current_start is not None:
                                span = row_idx - current_start
                                if span > 1:
                                    vmerge_spans[(current_start, col_idx)] = span
                            current_start = row_idx
                        # 'continue'인 경우는 계속
                    else:
                        # vMerge 없음 - 이전 병합 종료
                        if current_start is not None:
                            span = row_idx - current_start
                            if span > 1:
                                vmerge_spans[(current_start, col_idx)] = span
                            current_start = None
                
                # 마지막까지 병합 중이면 종료
                if current_start is not None:
                    span = len(table.rows) - current_start
                    if span > 1:
                        vmerge_spans[(current_start, col_idx)] = span
            
            # 3단계: CellMerge 객체 생성
            for (row, col), colspan in all_colspan_data.items():
                cell_merges.append(CellMerge(
                    row=row,
                    col=col,
                    colspan=colspan,
                    rowspan=1,
                    is_merged=False
                ))
            
            for (row, col), rowspan in vmerge_spans.items():
                # 해당 셀에 이미 colspan이 있는지 확인
                existing = None
                for merge in cell_merges:
                    if merge.row == row and merge.col == col:
                        existing = merge
                        break
                
                if existing:
                    existing.rowspan = rowspan
                else:
                    cell_merges.append(CellMerge(
                        row=row,
                        col=col,
                        colspan=1,
                        rowspan=rowspan,
                        is_merged=False
                    ))
            
            # 병합된 셀의 일부 표시
            for (row, col), status in all_vmerge_data.items():
                if status == 'continue':
                    cell_merges.append(CellMerge(
                        row=row,
                        col=col,
                        colspan=1,
                        rowspan=1,
                        is_merged=True
                    ))
            
            # 4단계: 셀 이미지 추출
            rows = all_rows_data
            for row_idx, row in enumerate(table.rows[1:], start=1):
                for col_idx, cell in enumerate(row.cells):
                    try:
                        for para in cell.paragraphs:
                            for run in para.runs:
                                if hasattr(run, '_element'):
                                    # qn() 함수 사용
                                    drawings = run._element.findall('.//' + qn('w:drawing'))
                                    
                                    for drawing in drawings:
                                        # 이미지 추출 시도
                                        try:
                                            blips = drawing.findall('.//' + qn('a:blip'))
                                            
                                            # 모든 고유 blip을 추출 (중복 제거)
                                            for blip in blips:
                                                embed_id = blip.get(qn('r:embed'))
                                                if embed_id and embed_id not in seen_image_ids:
                                                    seen_image_ids.add(embed_id)
                                                    try:
                                                        # document part를 통해 관계 찾기
                                                        image_part = doc.part.rels[embed_id].target_part
                                                        cell_images.append(
                                                            CellImage(
                                                                row=row_idx,
                                                                col=col_idx,
                                                                data=image_part.blob,
                                                                format=image_part.content_type.split('/')[-1],
                                                                width=0,
                                                                height=0,
                                                                embed_id=embed_id,
                                                            )
                                                        )
                                                    except KeyError:
                                                        pass
                                        except Exception as e:
                                            logger.debug(f"테이블 셀 이미지 추출 실패: {e}")
                                            continue
                    except Exception as e:
                        logger.debug(f"테이블 셀 처리 중 오류: {e}")
                        pass
            
            tables.append(
                TableContent(
                    headers=headers,
                    rows=rows,
                    cell_images=cell_images,
                    cell_merges=cell_merges,
                    page_number=table_page_map.get(table_idx),
                )
            )
        
        return tables
    
    def _extract_images(self, doc: DocxDocument) -> List[ImageContent]:
        """이미지 추출 (Drawing 객체, 플로팅 이미지 포함)"""
        images = []
        position = 0
        current_page = 1
        
        # 관계(relationships)를 통한 이미지 접근
        image_rels = {}
        for rel_id, rel in doc.part.rels.items():
            if "image" in rel.target_ref:
                image_rels[rel_id] = rel.target_part
        
        # 본문의 모든 Drawing 객체 탐색
        for element in doc.element.body:
            if isinstance(element, CT_P):
                para = Paragraph(element, doc)
                
                # 페이지 브레이크 확인
                has_page_break = False
                try:
                    page_breaks = para._element.findall('.//' + qn('w:br'))
                    for br in page_breaks:
                        if br.get(qn('w:type')) == 'page':
                            current_page += 1
                            position += 2000
                            has_page_break = True
                            break
                except:
                    pass
                
                if has_page_break:
                    continue
                
                # 단락 내 Drawing 검색
                for run in para.runs:
                    try:
                        if hasattr(run, '_element'):
                            # Inline 이미지 - qn() 함수 사용
                            drawings = run._element.findall('.//' + qn('w:drawing'))
                            
                            for drawing in drawings:
                                blips = drawing.findall('.//' + qn('a:blip'))
                                
                                for blip in blips:
                                    embed_id = blip.get(qn('r:embed'))
                                    if embed_id and embed_id in image_rels:
                                        try:
                                            image_part = image_rels[embed_id]
                                            
                                            # 이미지 크기 추출 시도
                                            width = 0
                                            height = 0
                                            extents = drawing.findall('.//' + qn('wp:extent'))
                                            if extents:
                                                width = int(extents[0].get('cx', 0))
                                                height = int(extents[0].get('cy', 0))
                                            
                                            images.append(
                                                ImageContent(
                                                    data=image_part.blob,
                                                    format=image_part.content_type.split("/")[-1],
                                                    width=width,
                                                    height=height,
                                                    page_number=current_page,
                                                    position=position,
                                                )
                                            )
                                        except Exception as e:
                                            logger.warning(f"이미지 추출 실패: {e}")
                                            continue
                    except Exception as e:
                        logger.debug(f"Drawing 처리 중 오류: {e}")
                        pass
                
                position += 1000
            
            elif isinstance(element, CT_Tbl):
                position += 5000
        
        return images
