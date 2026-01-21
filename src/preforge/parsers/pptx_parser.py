"""
PowerPoint 문서(.pptx) 파서
"""
import json
from pathlib import Path
from typing import Any, Dict, List, Optional
from pptx import Presentation
from pptx.util import Inches

from ..core.document import (
    Document,
    DocumentType,
    DocumentMetadata,
    TextContent,
    TableContent,
    ImageContent,
    CellMerge,
    PageLayout,
    GridCell,
)
from ..core.parser import BaseParser


class PptxParser(BaseParser):
    """PowerPoint 문서 파서"""

    def __init__(self, layout_overrides_path: Optional[Path] = None) -> None:
        self._layout_overrides = self._load_layout_overrides(layout_overrides_path)
    
    @property
    def supported_extensions(self) -> List[str]:
        return [".pptx", ".ppt"]
    
    @property
    def document_type(self) -> DocumentType:
        return DocumentType.PPTX
    
    def parse(self, file_path: Path) -> Document:
        """PowerPoint 문서 파싱"""
        self.validate_file(file_path)
        
        prs = Presentation(file_path)
        
        # 메타데이터 추출
        metadata = self._extract_metadata(prs)
        
        # 텍스트 추출
        text_contents = self._extract_text(prs)
        
        # 테이블 추출
        tables = self._extract_tables(prs)
        
        # 이미지 추출
        images = self._extract_images(prs)
        
        # 페이지 레이아웃 분석
        page_layouts = self._analyze_page_layouts(prs, text_contents, tables, images)
        
        return Document(
            file_path=file_path,
            doc_type=self.document_type,
            metadata=metadata,
            text_contents=text_contents,
            tables=tables,
            images=images,
            page_layouts=page_layouts,
            raw_content=prs,
        )

    def _load_layout_overrides(
        self, layout_overrides_path: Optional[Path]
    ) -> Dict[int, Dict[str, Any]]:
        if layout_overrides_path is None:
            layout_overrides_path = Path.cwd() / "private" / "layout_overrides.json"

        if not layout_overrides_path.exists():
            return {}

        try:
            payload = json.loads(layout_overrides_path.read_text(encoding="utf-8"))
        except (OSError, json.JSONDecodeError):
            return {}

        pages = payload.get("pages", {}) if isinstance(payload, dict) else {}
        overrides: Dict[int, Dict[str, Any]] = {}
        for page_str, config in pages.items():
            try:
                page_num = int(page_str)
            except (TypeError, ValueError):
                continue
            if isinstance(config, dict) and "rows" in config and "cols" in config:
                overrides[page_num] = config
        return overrides
    
    def _extract_metadata(self, prs: Presentation) -> DocumentMetadata:
        """메타데이터 추출"""
        core_props = prs.core_properties
        
        return DocumentMetadata(
            title=core_props.title,
            author=core_props.author,
            created_at=core_props.created,
            modified_at=core_props.modified,
            subject=core_props.subject,
            keywords=core_props.keywords.split(",") if core_props.keywords else None,
            page_count=len(prs.slides),
            properties={
                "category": core_props.category,
                "comments": core_props.comments,
                "language": core_props.language,
                "slide_count": len(prs.slides),
            }
        )
    
    def _extract_text_from_shape(self, shape, slide_idx: int, text_contents: list, is_title: bool = False, parent_top: int = 0, parent_left: int = 0):
        """사해이프에서 텍스트를 재귀적으로 추출 (GROUP 지원, 절대 좌표 계산)"""
        from pptx.enum.shapes import MSO_SHAPE_TYPE
        
        # 제목 shape는 이미 처리했으므로 스킵
        if is_title:
            return
        
        # 현재 shape의 top + 부모의 누적 top = 절대 위치
        try:
            shape_top = (shape.top if hasattr(shape, 'top') else 0) + parent_top
            shape_left = (shape.left if hasattr(shape, 'left') else 0) + parent_left
        except:
            shape_top = parent_top
            shape_left = parent_left
        
        # GROUP인 경우 내부 shape들을 재귀적으로 처리
        if shape.shape_type == MSO_SHAPE_TYPE.GROUP:
            for sub_shape in shape.shapes:
                self._extract_text_from_shape(sub_shape, slide_idx, text_contents, False, shape_top, shape_left)
        # 텍스트가 있는 shape 처리
        elif hasattr(shape, "text"):
            text = shape.text.strip()
            if text:
                text_contents.append(
                    TextContent(
                        text=text,
                        level=0,
                        page_number=slide_idx,
                        position=shape_top,
                        left=shape_left,
                    )
                )
    
    def _extract_text(self, prs: Presentation) -> List[TextContent]:
        """텍스트 추출"""
        text_contents = []
        
        for slide_idx, slide in enumerate(prs.slides, 1):
            # 슬라이드 제목 추출
            title_shape = None
            if slide.shapes.title:
                title_shape = slide.shapes.title
                text_contents.append(
                    TextContent(
                        text=title_shape.text,
                        level=1,  # 슬라이드 제목은 레벨 1
                        style="Title",
                        page_number=slide_idx,
                    )
                )
            
            # shape들을 위치 기준으로 정렬 (top 우선, 같은 줄이면 left)
            # 제목을 제외한 shape들만 정렬
            shapes_to_process = []
            for shape in slide.shapes:
                is_title = (title_shape is not None and shape == title_shape)
                if not is_title:
                    shapes_to_process.append(shape)
            
            # 위치 정보가 있는 shape들을 top, left 기준으로 정렬
            def get_position(shape):
                try:
                    return (shape.top, shape.left)
                except:
                    # 위치 정보가 없는 경우 큰 값으로 (맨 뒤로)
                    return (999999999, 999999999)
            
            shapes_to_process.sort(key=get_position)
            
            # 정렬된 순서대로 텍스트 추출 (GROUP 포함, 재귀적)
            for shape in shapes_to_process:
                self._extract_text_from_shape(shape, slide_idx, text_contents, False, parent_top=0, parent_left=0)
        
        return text_contents
    
    def _extract_tables(self, prs: Presentation) -> List[TableContent]:
        """테이블 추출"""
        from pptx.enum.shapes import MSO_SHAPE_TYPE
        from preforge.core.document import CellImage
        
        tables = []
        
        for slide_idx, slide in enumerate(prs.slides, 1):
            # 슬라이드의 모든 테이블 shape 찾기
            table_shapes = [s for s in slide.shapes if s.has_table]
            
            for table_shape in table_shapes:
                table = table_shape.table
                
                # 첫 번째 행을 헤더로 간주 (병합된 셀도 포함)
                headers = []
                for col_idx, cell in enumerate(table.rows[0].cells):
                    if cell.is_spanned:
                        # 병합된 셀인 경우, 왼쪽으로 가면서 merge_origin 찾기
                        for prev_col_idx in range(col_idx - 1, -1, -1):
                            prev_cell = table.rows[0].cells[prev_col_idx]
                            if prev_cell.is_merge_origin or not prev_cell.is_spanned:
                                # 병합된 셀임을 표시하거나 빈 문자열
                                headers.append("")
                                break
                        else:
                            headers.append("")
                    else:
                        headers.append(cell.text.strip())
                
                # 나머지 행을 데이터로 추출 (병합된 셀 처리)
                rows = []
                for row_idx in range(1, len(table.rows)):
                    row = table.rows[row_idx]
                    row_data = []
                    last_value = {}  # 각 열의 마지막 병합 시작 값 추적
                    
                    for col_idx, cell in enumerate(row.cells):
                        if cell.is_spanned:
                            # 병합된 셀인 경우, 같은 열의 이전 병합 시작 값을 찾음
                            # 위쪽으로 올라가면서 merge_origin인 셀 찾기
                            for prev_row_idx in range(row_idx - 1, -1, -1):
                                prev_cell = table.rows[prev_row_idx].cells[col_idx]
                                if prev_cell.is_merge_origin or not prev_cell.is_spanned:
                                    row_data.append(prev_cell.text.strip())
                                    break
                            else:
                                row_data.append("")
                        else:
                            row_data.append(cell.text.strip())
                    
                    rows.append(row_data)
                
                # 테이블 셀 내 이미지 찾기
                cell_images = self._find_images_in_table(slide, table_shape, table)
                
                # 셀 병합 정보 추출
                cell_merges = []
                for row_idx in range(len(table.rows)):
                    for col_idx in range(len(table.columns)):
                        cell = table.rows[row_idx].cells[col_idx]
                        if cell.is_merge_origin:
                            # 병합 시작 셀 - colspan과 rowspan 계산
                            colspan = 1
                            rowspan = 1
                            
                            # colspan 계산 (오른쪽으로 병합된 셀 수)
                            for c in range(col_idx + 1, len(table.columns)):
                                if table.rows[row_idx].cells[c].is_spanned:
                                    # 같은 행에서 오른쪽으로 spanned인지 확인
                                    colspan += 1
                                else:
                                    break
                            
                            # rowspan 계산 (아래로 병합된 셀 수)
                            for r in range(row_idx + 1, len(table.rows)):
                                if table.rows[r].cells[col_idx].is_spanned:
                                    rowspan += 1
                                else:
                                    break
                            
                            cell_merges.append(CellMerge(
                                row=row_idx,
                                col=col_idx,
                                colspan=colspan,
                                rowspan=rowspan,
                                is_merged=False
                            ))
                        elif cell.is_spanned:
                            # 병합된 셀의 일부 (표시하지 않음)
                            cell_merges.append(CellMerge(
                                row=row_idx,
                                col=col_idx,
                                colspan=1,
                                rowspan=1,
                                is_merged=True
                            ))
                
                tables.append(
                    TableContent(
                        headers=headers,
                        rows=rows,
                        page_number=slide_idx,
                        cell_images=cell_images,
                        cell_merges=cell_merges,
                    )
                )
        
        return tables
    
    def _find_images_in_table(self, slide, table_shape, table) -> List:
        """테이블 내부의 이미지 찾기"""
        from pptx.enum.shapes import MSO_SHAPE_TYPE
        from preforge.core.document import CellImage
        
        cell_images = []
        
        # 각 열의 절대 위치 계산
        col_positions = [table_shape.left]
        for i in range(len(table.columns)):
            col_positions.append(col_positions[-1] + table.columns[i].width)
        
        # 각 행의 절대 위치 계산
        row_positions = [table_shape.top]
        for i in range(len(table.rows)):
            row_positions.append(row_positions[-1] + table.rows[i].height)
        
        # 슬라이드의 모든 이미지 찾기 (직접 이미지 + 그룹 내 이미지)
        images_to_check = []
        for shape in slide.shapes:
            if shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
                images_to_check.append(shape)
            elif shape.shape_type == MSO_SHAPE_TYPE.GROUP:
                for sub_shape in shape.shapes:
                    if sub_shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
                        images_to_check.append(sub_shape)
        
        # 각 이미지가 테이블 셀에 속하는지 확인
        for img in images_to_check:
            img_center_x = img.left + img.width // 2
            img_center_y = img.top + img.height // 2
            
            # 어느 열에 속하는지 찾기
            col = -1
            for i in range(len(col_positions) - 1):
                if col_positions[i] <= img_center_x < col_positions[i + 1]:
                    col = i
                    break
            
            # 어느 행에 속하는지 찾기
            row = -1
            for i in range(len(row_positions) - 1):
                if row_positions[i] <= img_center_y < row_positions[i + 1]:
                    row = i
                    break
            
            # 테이블 내부에 있는 경우에만 추가
            if row >= 0 and col >= 0:
                try:
                    cell_images.append(
                        CellImage(
                            row=row,
                            col=col,
                            data=img.image.blob,
                            format=img.image.ext,
                            width=img.width,
                            height=img.height,
                        )
                    )
                except Exception:
                    # 이미지 추출 실패 시 무시
                    pass
        
        return cell_images
    
    def _is_image_in_table(self, img, tables_info):
        """이미지가 테이블 내부에 있는지 확인"""
        img_center_x = img.left + img.width // 2
        img_center_y = img.top + img.height // 2
        
        for table_info in tables_info:
            table_shape = table_info['shape']
            table = table_info['table']
            
            # 각 열의 절대 위치 계산
            col_positions = [table_shape.left]
            for i in range(len(table.columns)):
                col_positions.append(col_positions[-1] + table.columns[i].width)
            
            # 각 행의 절대 위치 계산
            row_positions = [table_shape.top]
            for i in range(len(table.rows)):
                row_positions.append(row_positions[-1] + table.rows[i].height)
            
            # 이미지가 테이블 영역 안에 있는지 확인
            if (col_positions[0] <= img_center_x < col_positions[-1] and
                row_positions[0] <= img_center_y < row_positions[-1]):
                return True
        
        return False
    
    def _extract_images(self, prs: Presentation) -> List[ImageContent]:
        """이미지 추출 (재귀적으로 중첩 그룹 탐색, 테이블 내부 이미지는 제외)"""
        from pptx.enum.shapes import MSO_SHAPE_TYPE
        
        images = []
        
        def extract_from_shape(shape, slide_idx, tables_info, parent_top=0, parent_left=0):
            """재귀적으로 shape에서 이미지 추출 (절대 좌표 계산)"""
            # 현재 shape의 top + 부모의 누적 top = 절대 위치
            shape_top = (shape.top if hasattr(shape, 'top') else 0) + parent_top
            shape_left = (shape.left if hasattr(shape, 'left') else 0) + parent_left
            
            if shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
                try:
                    # 테이블 내부의 이미지는 제외
                    if not self._is_image_in_table(shape, tables_info):
                        image = shape.image
                        images.append(
                            ImageContent(
                                data=image.blob,
                                format=image.ext,
                                width=shape.width,
                                height=shape.height,
                                page_number=slide_idx,
                                position=shape_top,
                                left=shape_left,
                            )
                        )
                except Exception:
                    pass
            elif shape.shape_type == MSO_SHAPE_TYPE.GROUP:
                # GROUP인 경우 하위 shape들을 재귀적으로 탐색
                # 하위 shape들에게 현재까지의 누적 top을 전달
                try:
                    for sub_shape in shape.shapes:
                        extract_from_shape(sub_shape, slide_idx, tables_info, shape_top, shape_left)
                except Exception:
                    pass
        
        for slide_idx, slide in enumerate(prs.slides, 1):
            # 현재 슬라이드의 모든 테이블 정보 수집
            tables_info = []
            for shape in slide.shapes:
                if shape.has_table:
                    tables_info.append({'shape': shape, 'table': shape.table})
            
            # 이미지 추출 (테이블 정보 전달)
            for shape in slide.shapes:
                extract_from_shape(shape, slide_idx, tables_info, parent_top=0)
        
        return images
    
    def _analyze_page_layouts(
        self, 
        prs: Presentation, 
        text_contents: List[TextContent],
        tables: List[TableContent],
        images: List[ImageContent]
    ) -> List[PageLayout]:
        """페이지별 그리드 레이아웃 분석"""
        from pptx.enum.shapes import MSO_SHAPE_TYPE
        
        page_layouts = []
        
        # 색상 팔레트 (시각화용)
        colors = [
            '#FFE5E5', '#E5F3FF', '#E5FFE5', '#FFF5E5', '#F5E5FF', '#E5FFFF',
            '#FFD4D4', '#D4E8FF', '#D4FFD4', '#FFEBD4', '#EBD4FF', '#D4FFFF'
        ]
        
        for slide_idx, slide in enumerate(prs.slides, 1):
            # 슬라이드 크기 (EMU 단위)
            slide_width = prs.slide_width
            slide_height = prs.slide_height
            
            # 하단 90% 이상은 페이지 번호/바닥글로 간주
            footer_threshold = slide_height * 90 // 100
            # 상단 15% 미만은 제목 영역으로 간주
            header_threshold = slide_height * 15 // 100
            
            # 페이지의 모든 컨텐츠 수집 (shapes 직접 분석)
            content_items = []
            from pptx.enum.shapes import MSO_SHAPE_TYPE
            
            for shape in slide.shapes:
                top = shape.top
                # 푸터 영역 제외
                if top >= footer_threshold:
                    continue
                
                if shape.has_table:
                    content_items.append({
                        'type': 'table',
                        'id': f'table_{len(content_items)}',
                        'top': top,
                        'left': shape.left,
                        'width': shape.width,
                        'height': shape.height,
                    })
                elif shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
                    content_items.append({
                        'type': 'image',
                        'id': f'image_{len(content_items)}',
                        'top': top,
                        'left': shape.left,
                        'width': shape.width,
                        'height': shape.height,
                    })
                elif shape.has_text_frame and shape.text_frame.text.strip():
                    # 헤더 영역의 텍스트는 레이아웃 감지에서 제외
                    if top < header_threshold:
                        continue
                    content_items.append({
                        'type': 'text',
                        'id': f'text_{len(content_items)}',
                        'top': top,
                        'left': shape.left,
                        'width': shape.width,
                        'height': shape.height,
                    })
            
            override = self._layout_overrides.get(slide_idx)
            if override:
                rows, cols, grid_cells = self._build_layout_from_override(
                    override,
                    content_items,
                    slide_width,
                    slide_height,
                    colors,
                )
            elif not content_items:
                # 컨텐츠가 없으면 1x1 그리드로 설정
                rows, cols = 1, 1
                grid_cells = [
                    GridCell(
                        row=0, col=0,
                        top=0, left=0,
                        width=slide_width, height=slide_height,
                        color=colors[0]
                    )
                ]
            else:
                # 그리드 분석: 컨텐츠 위치 기반으로 행/열 결정
                rows, cols, grid_cells = self._detect_grid_layout(
                    content_items, slide_width, slide_height, colors
                )
            
            layout = PageLayout(
                page_number=slide_idx,
                rows=rows,
                cols=cols,
                slide_width=slide_width,
                slide_height=slide_height,
                grid_cells=grid_cells
            )
            page_layouts.append(layout)
        
        return page_layouts

    def _build_layout_from_override(
        self,
        override: Dict[str, Any],
        content_items: List[dict],
        slide_width: int,
        slide_height: int,
        colors: List[str],
    ) -> tuple:
        rows = int(override.get("rows", 1))
        cols = int(override.get("cols", 1))
        row_colspans = override.get("row_colspans")

        if not row_colspans or len(row_colspans) != rows:
            row_colspans = [[1] * cols for _ in range(rows)]

        row_height = slide_height / rows
        col_width = slide_width / cols

        row_boundaries = [int(round(row_height * r)) for r in range(rows + 1)]
        col_boundaries = [int(round(col_width * c)) for c in range(cols + 1)]

        grid_cells: List[GridCell] = []
        color_idx = 0

        for r in range(rows):
            row_top = row_boundaries[r]
            row_bottom = row_boundaries[r + 1]
            row_height_actual = row_bottom - row_top
            col_index = 0

            for span in row_colspans[r]:
                span = int(span)
                left = col_boundaries[col_index]
                right = col_boundaries[min(col_index + span, cols)]
                width = right - left

                cell = GridCell(
                    row=r,
                    col=col_index,
                    top=row_top,
                    left=left,
                    width=width,
                    height=row_height_actual,
                    content_ids=[],
                    color=colors[color_idx % len(colors)],
                    colspan=span,
                )
                grid_cells.append(cell)
                color_idx += 1
                col_index += span

        for item in content_items:
            item_center_x = item['left'] + item['width'] // 2
            item_center_y = item['top'] + item['height'] // 2

            row_idx = min(max(int(item_center_y // row_height), 0), rows - 1)
            col_idx = min(max(int(item_center_x // col_width), 0), cols - 1)

            for cell in grid_cells:
                if cell.row != row_idx:
                    continue
                if cell.col <= col_idx < cell.col + cell.colspan:
                    cell.content_ids.append(item['id'])
                    break

        return rows, cols, grid_cells
    
    def _detect_grid_layout(
        self, 
        content_items: List[dict], 
        slide_width: int, 
        slide_height: int,
        colors: List[str]
    ) -> tuple:
        """컨텐츠 배치를 기반으로 그리드 레이아웃 감지
        
        원칙:
        1. 행은 최소화 (대부분 1행)
        2. 열은 좌우에 명확히 분리된 요소가 있을 때만 2열
        3. 대칭 레이아웃은 1열로 처리
        """
        
        if not content_items:
            return 1, 1, []
        
        # 헤더 영역(상단 15%) 제외한 본문 요소만 분석
        header_threshold = slide_height * 15 // 100
        body_items = [item for item in content_items if item['top'] > header_threshold]
        
        if not body_items:
            return self._create_single_cell_layout(content_items, slide_width, slide_height, colors)
        
        mid_x = slide_width // 2
        
        # 좌우 분류
        left_items = [item for item in body_items 
                     if (item['left'] + item['width'] // 2) < mid_x]
        right_items = [item for item in body_items 
                      if (item['left'] + item['width'] // 2) >= mid_x]
        
        # 열 결정
        cols = 1
        
        if left_items and right_items:
            # 양쪽에 요소가 있는 경우
            
            # 엄격한 대칭 패턴 확인 (양쪽 요소 수가 동일하고 위치도 대칭)
            if self._is_symmetric_layout(left_items, right_items, slide_height):
                cols = 1
            else:
                # 기본적으로 좌우에 요소가 있으면 2열
                cols = 2
        
        # 행 결정: 기본적으로 1행
        rows = 1
        
        return self._build_grid_cells(
            content_items, rows, cols, slide_width, slide_height, colors
        )
    
    def _is_symmetric_layout(
        self, 
        left_items: List[dict], 
        right_items: List[dict], 
        slide_height: int
    ) -> bool:
        """좌우 대칭 레이아웃인지 확인 (목차, 그리드 등)
        
        조건:
        1. 양쪽에 각각 3개 이상 요소가 있어야 함
        2. 양쪽 요소 수 차이가 2개 이하
        3. y 위치가 비슷한 쌍이 많아야 함
        """
        # 양쪽에 충분한 요소가 있어야 대칭 판단
        if len(left_items) < 3 or len(right_items) < 3:
            return False
        
        # 양쪽 요소 수 차이가 너무 크면 대칭 아님
        if abs(len(left_items) - len(right_items)) > 2:
            return False
        
        # y 위치 매칭 확인
        left_tops = sorted([item['top'] for item in left_items])
        right_tops = sorted([item['top'] for item in right_items])
        
        # 비슷한 y 위치의 쌍 찾기
        matches = 0
        used_right = set()
        for lt in left_tops:
            for i, rt in enumerate(right_tops):
                if i not in used_right and abs(lt - rt) < slide_height * 0.08:
                    matches += 1
                    used_right.add(i)
                    break
        
        # 70% 이상 매칭되면 대칭
        min_items = min(len(left_tops), len(right_tops))
        return matches >= min_items * 0.7
        for lt, rt in zip(left_tops, right_tops):
            if abs(lt - rt) > slide_height * 0.08:  # 8% 허용
                return False
        
        return True
    
    def _create_single_cell_layout(
        self,
        content_items: List[dict],
        slide_width: int,
        slide_height: int,
        colors: List[str]
    ) -> tuple:
        """1x1 그리드 생성"""
        cell = GridCell(
            row=0,
            col=0,
            top=0,
            left=0,
            width=slide_width,
            height=slide_height,
            content_ids=[item['id'] for item in content_items],
            color=colors[0]
        )
        return 1, 1, [cell]
    
    def _build_grid_cells(
        self,
        content_items: List[dict],
        rows: int,
        cols: int,
        slide_width: int,
        slide_height: int,
        colors: List[str]
    ) -> tuple:
        """행/열 수에 맞는 그리드 셀 생성"""
        row_height = slide_height // rows
        col_width = slide_width // cols
        
        grid_cells = []
        color_idx = 0
        
        for r in range(rows):
            row_top = r * row_height
            row_bottom = (r + 1) * row_height if r < rows - 1 else slide_height
            actual_row_height = row_bottom - row_top
            
            for c in range(cols):
                col_left = c * col_width
                col_right = (c + 1) * col_width if c < cols - 1 else slide_width
                actual_col_width = col_right - col_left
                
                # 이 셀에 속하는 컨텐츠 찾기
                cell_contents = []
                for item in content_items:
                    item_center_x = item['left'] + item['width'] // 2
                    item_center_y = item['top'] + item['height'] // 2
                    
                    if (row_top <= item_center_y < row_bottom and
                        col_left <= item_center_x < col_right):
                        cell_contents.append(item['id'])
                
                cell = GridCell(
                    row=r,
                    col=c,
                    top=row_top,
                    left=col_left,
                    width=actual_col_width,
                    height=actual_row_height,
                    content_ids=cell_contents,
                    color=colors[color_idx % len(colors)]
                )
                grid_cells.append(cell)
                color_idx += 1
        
        return rows, cols, grid_cells
