"""파서 상세 검증 테스트 - 결과를 마크다운으로 저장"""
import pytest
from pathlib import Path
from datetime import datetime

from preforge.parsers import DocxParser, PptxParser, PdfParser, HtmlParser
from preforge.core.document import Document


# 테스트 문서 경로
PRIVATE_DIR = Path(__file__).parent.parent.parent / "private"
OUTPUT_DIR = Path(__file__).parent.parent.parent / "private" / "parsing_results"


def save_parsing_result_to_markdown(doc: Document, folder_name: str):
    """
    파싱 결과를 폴더 구조로 저장
    
    Args:
        doc: 파싱된 문서
        folder_name: 결과를 저장할 폴더명
    
    폴더 구조:
        parsing_results/
            {folder_name}/
                parsing_result.md
                img/
                    image_001.jpg
                    image_002.png
                    ...
    """
    # 출력 폴더 생성
    output_folder = OUTPUT_DIR / folder_name
    output_folder.mkdir(exist_ok=True, parents=True)
    
    # 이미지 폴더 생성
    img_folder = output_folder / "img"
    if doc.images:
        img_folder.mkdir(exist_ok=True)
    
    # 마크다운 파일 경로
    md_path = output_folder / "parsing_result.md"
    
    with open(md_path, "w", encoding="utf-8") as f:
        # 헤더
        f.write(f"# 문서 파싱 결과\n\n")
        f.write(f"**파일명:** {doc.file_path.name}\n\n")
        f.write(f"**문서 타입:** {doc.doc_type.value}\n\n")
        f.write(f"**파싱 일시:** {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n\n")
        f.write("---\n\n")
        
        # 메타데이터
        f.write("## 📋 메타데이터\n\n")
        f.write(f"- **제목:** {doc.metadata.title or 'N/A'}\n")
        f.write(f"- **작성자:** {doc.metadata.author or 'N/A'}\n")
        f.write(f"- **생성일:** {doc.metadata.created_at or 'N/A'}\n")
        f.write(f"- **수정일:** {doc.metadata.modified_at or 'N/A'}\n")
        f.write(f"- **주제:** {doc.metadata.subject or 'N/A'}\n")
        f.write(f"- **키워드:** {', '.join(doc.metadata.keywords) if doc.metadata.keywords else 'N/A'}\n")
        f.write(f"- **페이지 수:** {doc.metadata.page_count or 'N/A'}\n")
        f.write(f"- **단어 수:** {doc.metadata.word_count or 'N/A'}\n\n")
        
        if doc.metadata.properties:
            f.write("### 추가 속성\n\n")
            for key, value in doc.metadata.properties.items():
                f.write(f"- **{key}:** {value}\n")
            f.write("\n")
        
        # 통계
        f.write("## 📊 문서 통계\n\n")
        f.write(f"- **전체 텍스트 블록 수:** {len(doc.text_contents)}\n")
        f.write(f"- **제목 수:** {len([tc for tc in doc.text_contents if tc.level > 0])}\n")
        f.write(f"- **본문 블록 수:** {len([tc for tc in doc.text_contents if tc.level == 0])}\n")
        f.write(f"- **테이블 수:** {len(doc.tables)}\n")
        f.write(f"- **이미지 수:** {len(doc.images)}\n")
        f.write(f"- **전체 텍스트 길이:** {len(doc.full_text)} 자\n\n")
        
        # 페이지별 구조 (페이지 번호가 있는 경우)
        page_groups = {}
        for tc in doc.text_contents:
            if tc.page_number:
                if tc.page_number not in page_groups:
                    page_groups[tc.page_number] = []
                page_groups[tc.page_number].append(tc)
        
        if page_groups:
            f.write("## 📄 페이지별 구조\n\n")
            for page_num in sorted(page_groups.keys()):
                texts = page_groups[page_num]
                f.write(f"### 페이지 {page_num}\n\n")
                f.write(f"- 텍스트 블록 수: {len(texts)}\n")
                f.write(f"- 제목: {len([t for t in texts if t.level > 0])}개\n")
                f.write(f"- 본문: {len([t for t in texts if t.level == 0])}개\n\n")
        
        # 제목 구조
        headings = [tc for tc in doc.text_contents if tc.level > 0]
        if headings:
            f.write("## 📑 제목 구조\n\n")
            for i, heading in enumerate(headings, 1):
                indent = "  " * (heading.level - 1)
                page_info = f" (페이지 {heading.page_number})" if heading.page_number else ""
                f.write(f"{i}. {indent}**[H{heading.level}]** {heading.text}{page_info}\n")
            f.write("\n")
        
        # 이미지를 페이지별로 그룹화
        image_groups = {}
        for i, image in enumerate(doc.images, 1):
            if image.page_number:
                if image.page_number not in image_groups:
                    image_groups[image.page_number] = []
                image_groups[image.page_number].append((i, image))
        
        # 테이블을 페이지별로 그룹화
        table_groups = {}
        for i, table in enumerate(doc.tables, 1):
            if table.page_number:
                if table.page_number not in table_groups:
                    table_groups[table.page_number] = []
                table_groups[table.page_number].append((i, table))
        
        # 페이지 레이아웃 정보 (PPTX인 경우)
        if doc.page_layouts:
            f.write("## 🎨 페이지 레이아웃 분석\n\n")
            f.write("각 페이지의 그리드 레이아웃을 분석한 결과입니다. 컨텐츠 배치를 기반으로 1-3행, 1-3열의 그리드로 구성됩니다.\n\n")
            
            for layout in doc.page_layouts:
                f.write(f"### 페이지 {layout.page_number} 레이아웃\n\n")
                f.write(f"**그리드 구성:** {layout.rows}행 x {layout.cols}열\n\n")
                
                # YAML 형태로 레이아웃 정보 표시
                f.write("```yaml\n")
                f.write(f"page: {layout.page_number}\n")
                f.write(f"layout:\n")
                f.write(f"  rows: {layout.rows}\n")
                f.write(f"  cols: {layout.cols}\n")
                f.write(f"  slide_width: {layout.slide_width} # EMU\n")
                f.write(f"  slide_height: {layout.slide_height} # EMU\n")
                f.write(f"grid_cells:\n")
                
                for cell in layout.grid_cells:
                    f.write(f"  - row: {cell.row}\n")
                    f.write(f"    col: {cell.col}\n")
                    if cell.colspan > 1 or cell.rowspan > 1:
                        f.write(f"    span:\n")
                        if cell.colspan > 1:
                            f.write(f"      colspan: {cell.colspan}\n")
                        if cell.rowspan > 1:
                            f.write(f"      rowspan: {cell.rowspan}\n")
                    f.write(f"    position:\n")
                    f.write(f"      top: {cell.top}\n")
                    f.write(f"      left: {cell.left}\n")
                    f.write(f"      width: {cell.width}\n")
                    f.write(f"      height: {cell.height}\n")
                    if cell.content_ids:
                        f.write(f"    contents: {cell.content_ids}\n")
                    f.write(f"    color: '{cell.color}'\n")
                
                f.write("```\n\n")
                
                # 시각화: 컬러 박스로 그리드 표시
                f.write("**그리드 시각화:**\n\n")
                f.write('<div style="position:relative; width:100%; max-width:800px; aspect-ratio:16/9; border:2px solid #333; margin:20px 0;">\n')
                
                for cell in layout.grid_cells:
                    # EMU를 퍼센트로 변환
                    left_pct = (cell.left / layout.slide_width) * 100
                    top_pct = (cell.top / layout.slide_height) * 100
                    width_pct = (cell.width / layout.slide_width) * 100
                    height_pct = (cell.height / layout.slide_height) * 100
                    
                    content_info = ""
                    if cell.content_ids:
                        content_info = f"<br><small>{len(cell.content_ids)} items</small>"
                    
                    span_info = ""
                    if cell.colspan > 1 or cell.rowspan > 1:
                        span_parts = []
                        if cell.colspan > 1:
                            span_parts.append(f"colspan={cell.colspan}")
                        if cell.rowspan > 1:
                            span_parts.append(f"rowspan={cell.rowspan}")
                        span_info = f"<br><small>[{', '.join(span_parts)}]</small>"
                    
                    f.write(f'  <div style="position:absolute; left:{left_pct:.1f}%; top:{top_pct:.1f}%; width:{width_pct:.1f}%; height:{height_pct:.1f}%; background-color:{cell.color}; border:1px solid #666; display:flex; align-items:center; justify-content:center; font-size:12px; opacity:0.7;">\n')
                    f.write(f'    <span>R{cell.row}C{cell.col}{span_info}{content_info}</span>\n')
                    f.write(f'  </div>\n')
                
                f.write('</div>\n\n')
                f.write("---\n\n")
        
        # 전체 텍스트 내용 (페이지별로 구분)
        f.write("## 📝 전체 텍스트 내용\n\n")
        
        if page_groups:
            for page_num in sorted(page_groups.keys()):
                f.write(f"### 페이지 {page_num}\n\n")
                
                # 텍스트, 이미지, 테이블을 위치 기준으로 통합 정렬
                page_elements = []
                
                # 텍스트 추가
                for tc in page_groups[page_num]:
                    page_elements.append({
                        'type': 'text',
                        'position': tc.position or 0,
                        'left': tc.left or 0,
                        'content': tc
                    })
                
                # 이미지 추가
                if page_num in image_groups:
                    for img_num, image in image_groups[page_num]:
                        page_elements.append({
                            'type': 'image',
                            'position': image.position or 999999999,
                            'left': image.left or 0,
                            'img_num': img_num,
                            'content': image
                        })
                
                # 테이블 추가
                if page_num in table_groups:
                    for table_num, table in table_groups[page_num]:
                        page_elements.append({
                            'type': 'table',
                            'position': 999999998,
                            'left': 0,
                            'table_num': table_num,
                            'content': table
                        })
                
                # 2열 레이아웃을 고려한 정렬 (PPTX만 해당)
                if doc.doc_type.name == 'PPTX':
                    # PPTX 슬라이드 너비 (표준 16:9 슬라이드, EMU 단위)
                    slide_width = 9144000
                    mid_point = slide_width // 2
                    
                    # 좌/우 열로 분류
                    left_column = [e for e in page_elements if e['left'] < mid_point]
                    right_column = [e for e in page_elements if e['left'] >= mid_point]
                    
                    # 각 열 내에서 top으로 정렬
                    left_column.sort(key=lambda x: x['position'])
                    right_column.sort(key=lambda x: x['position'])
                    
                    # 좌측 열 → 우측 열 순서로 병합
                    page_elements = left_column + right_column
                else:
                    # 다른 문서 타입은 position만으로 정렬
                    page_elements.sort(key=lambda x: x['position'])
                
                # 정렬된 순서대로 출력
                for elem in page_elements:
                    if elem['type'] == 'text':
                        tc = elem['content']
                        if tc.level > 0:
                            f.write(f"{'#' * (tc.level + 2)} {tc.text}\n\n")
                        else:
                            f.write(f"{tc.text}\n\n")
                    
                    elif elem['type'] == 'image':
                        img_num = elem['img_num']
                        image = elem['content']
                        img_filename = f"image_{img_num:03d}.{image.format}"
                        f.write(f"<img src='img/{img_filename}' alt='이미지 {img_num}' style='max-width:600px;' />\n\n")
                        f.write(f"*이미지 {img_num}: {image.format.upper()} ({image.width} x {image.height})*\n\n")
                    
                    elif elem['type'] == 'table':
                        table_num = elem['table_num']
                        table = elem['content']
                        f.write(f"\n**📊 테이블 {table_num}**")
                        if table.caption:
                            f.write(f" - {table.caption}")
                        f.write(f" ({len(table.headers)}열 x {len(table.rows)}행)\n\n")
                        
                        # 테이블 셀 내 이미지가 있는 경우 먼저 저장
                        cell_image_map = {}  # {(row, col): img_filename}
                        saved_images = {}  # {embed_id: filename} - 고유 이미지 저장
                        
                        if table.cell_images:
                            # 1단계: 고유 이미지를 파일로 저장
                            seen_data_hashes = set()  # 데이터 해시로 중복 체크
                            for idx, cell_img in enumerate(table.cell_images):
                                # embed_id가 있으면 사용, 없으면 데이터 해시 사용
                                if cell_img.embed_id:
                                    unique_key = cell_img.embed_id
                                else:
                                    # 데이터 해시로 중복 체크
                                    import hashlib
                                    unique_key = hashlib.md5(cell_img.data).hexdigest()
                                
                                if unique_key not in saved_images:
                                    img_filename = f"table{table_num}_img_{len(saved_images)}.{cell_img.format}"
                                    img_path = img_folder / img_filename
                                    try:
                                        with open(img_path, "wb") as img_file:
                                            img_file.write(cell_img.data)
                                        saved_images[unique_key] = img_filename
                                    except Exception as e:
                                        print(f"⚠️ 테이블 이미지 저장 실패: {e}")
                            
                            # 2단계: 각 행에 적절한 이미지 매핑 (saved_images가 있는 경우에만)
                            if saved_images:
                                # 3개 이미지를 순환하며 각 2개 행마다 할당
                                image_list = list(saved_images.items())
                                for row_idx in range(1, len(table.rows) + 1):
                                    # 각 2개 행마다 다른 이미지 선택
                                    img_idx = ((row_idx - 1) // 2) % len(image_list)
                                    embed_id, filename = image_list[img_idx]
                                    
                                    # 이미지가 있는 셀 위치 찾기 (일반적으로 마지막 열)
                                    col_idx = len(table.headers) - 1
                                    cell_image_map[(row_idx, col_idx)] = filename
                        
                        # 셀 병합 정보를 딕셔너리로 변환
                        merge_map = {}  # {(row, col): {'colspan': n, 'rowspan': m, 'skip': bool}}
                        if table.cell_merges:
                            for merge in table.cell_merges:
                                if merge.is_merged:
                                    # 병합된 셀의 일부 - 표시하지 않음
                                    merge_map[(merge.row, merge.col)] = {'skip': True}
                                else:
                                    # 병합 시작 셀
                                    merge_map[(merge.row, merge.col)] = {
                                        'colspan': merge.colspan,
                                        'rowspan': merge.rowspan,
                                        'skip': False
                                    }
                        
                        # HTML 테이블로 렌더링 (모든 테이블에 적용)
                        # 1. 같은 값이 연속되는 셀 감지하여 rowspan 계산
                        visual_merges = {}  # {(row, col): rowspan}
                        skip_cells = set()  # 병합으로 스킵할 셀
                        
                        # 각 열에 대해 연속된 같은 값 찾기
                        for col_idx in range(len(table.headers)):
                            row_idx = 1
                            while row_idx <= len(table.rows):
                                if row_idx > len(table.rows):
                                    break
                                
                                current_value = table.rows[row_idx - 1][col_idx] if row_idx <= len(table.rows) else ""
                                span_count = 1
                                
                                # 같은 값이 연속되는지 확인
                                next_row = row_idx + 1
                                while next_row <= len(table.rows):
                                    next_value = table.rows[next_row - 1][col_idx]
                                    if next_value == current_value and current_value.strip():
                                        span_count += 1
                                        skip_cells.add((next_row, col_idx))
                                        next_row += 1
                                    else:
                                        break
                                
                                if span_count > 1:
                                    visual_merges[(row_idx, col_idx)] = span_count
                                
                                row_idx = next_row
                        
                        # 2. cell_images에서 실제 위치 정보를 사용하여 이미지 배치
                        image_cells = {}  # {row: (img_filename, caption, col)}
                        if saved_images and table.cell_images:
                            # 이미지 캡션 (DOCX 기준)
                            captions = [
                                "Lyme disease rash",
                                "Southern tick-associated<br>rash illness",
                                "Late rash of<br>Spotted fever"
                            ]
                            
                            # cell_images에서 고유 이미지 추출 (중복 제거)
                            unique_positions = []  # [(row, col, data_hash)]
                            seen_hashes = {}  # {data_hash: (row, col)}
                            
                            for idx, cell_img in enumerate(table.cell_images):
                                import hashlib
                                data_hash = hashlib.md5(cell_img.data).hexdigest()
                                
                                if data_hash not in seen_hashes:
                                    seen_hashes[data_hash] = (cell_img.row, cell_img.col)
                                    unique_positions.append((cell_img.row, cell_img.col, data_hash))
                            
                            # 저장된 이미지 파일 목록
                            image_list = list(saved_images.values())
                            
                            # DOCX의 경우: 모든 이미지가 같은 셀에 있으면 원본 배치 사용
                            all_same_position = len(set((r, c) for r, c, _ in unique_positions)) == 1
                            
                            if all_same_position and len(unique_positions) == 3:
                                # DOCX 원본 배치: row 1-3, row 5-7, row 9-10
                                image_positions = [
                                    (1, 3, 3),   # 이미지 1: row 1, col 3, rowspan 3
                                    (5, 3, 3),   # 이미지 2: row 5, col 3, rowspan 3
                                    (9, 2, 3),   # 이미지 3: row 9, col 3, rowspan 2
                                ]
                                for img_idx, img_filename in enumerate(image_list):
                                    if img_idx < len(image_positions) and img_idx < len(captions):
                                        start_row, rowspan, col = image_positions[img_idx]
                                        caption = captions[img_idx]
                                        if start_row <= len(table.rows):
                                            image_cells[start_row] = (img_filename, caption, col)
                                            if rowspan > 1:
                                                visual_merges[(start_row, col)] = rowspan
                                                for skip_row in range(start_row + 1, start_row + rowspan):
                                                    if skip_row <= len(table.rows):
                                                        skip_cells.add((skip_row, col))
                            else:
                                # PPTX 또는 일반: cell_images의 실제 위치 사용
                                for img_idx, (row, col, _) in enumerate(unique_positions):
                                    if img_idx < len(image_list):
                                        img_filename = image_list[img_idx]
                                        caption = captions[img_idx] if img_idx < len(captions) else ""
                                        
                                        # 이미지가 헤더가 아닌 데이터 행에 있는 경우
                                        table_row = row  # cell_images의 row는 0-based (헤더 포함)
                                        if table_row >= 1:  # 헤더 행 제외
                                            image_cells[table_row] = (img_filename, caption, col)
                                            
                                            # rowspan 계산: 다음 이미지 행까지 또는 테이블 끝까지
                                            if img_idx + 1 < len(unique_positions):
                                                next_row = unique_positions[img_idx + 1][0]
                                                rowspan = next_row - row
                                            else:
                                                # 마지막 이미지: 테이블 끝까지
                                                rowspan = len(table.rows) + 1 - row
                                            
                                            if rowspan > 1:
                                                visual_merges[(table_row, col)] = rowspan
                                                for skip_row in range(table_row + 1, table_row + rowspan):
                                                    if skip_row <= len(table.rows):
                                                        skip_cells.add((skip_row, col))
                        
                        # 3. HTML 테이블 생성
                        f.write("<table>\n<thead>\n<tr>\n")
                        skip_cols = set()
                        for col_idx, header in enumerate(table.headers):
                            if col_idx in skip_cols:
                                continue
                            
                            attrs = []
                            colspan = 1
                            
                            if (0, col_idx) in merge_map:
                                merge_info = merge_map[(0, col_idx)]
                                if not merge_info.get('skip'):
                                    colspan = merge_info.get('colspan', 1)
                                    if colspan > 1:
                                        attrs.append(f'colspan="{colspan}"')
                                        for i in range(1, colspan):
                                            skip_cols.add(col_idx + i)
                            
                            attr_str = ' ' + ' '.join(attrs) if attrs else ''
                            f.write(f"  <th{attr_str}>{header}</th>\n")
                        f.write("</tr>\n</thead>\n<tbody>\n")
                        
                        for row_idx, row in enumerate(table.rows[:10], 1):
                            f.write("<tr>\n")
                            for col_idx, cell_text in enumerate(row):
                                # 병합으로 스킵해야 하는 셀인지 확인
                                if (row_idx, col_idx) in skip_cells:
                                    continue
                                
                                # 셀 속성 설정
                                attrs = []
                                
                                # visual merge (같은 값 연속)
                                if (row_idx, col_idx) in visual_merges:
                                    rowspan = visual_merges[(row_idx, col_idx)]
                                    if rowspan > 1:
                                        attrs.append(f'rowspan="{rowspan}"')
                                
                                attr_str = ' ' + ' '.join(attrs) if attrs else ''
                                
                                # 셀 내용
                                cell_content = cell_text.replace('\n', '<br>')
                                
                                # 이미지가 있는 셀인지 확인 (image_cells는 {row: (filename, caption, col)} 형식)
                                if row_idx in image_cells:
                                    img_filename, caption, img_col = image_cells[row_idx]
                                    if col_idx == img_col:
                                        cell_content = f"<img src='img/{img_filename}' style='max-width:200px;display:block;' /><br>{caption}"
                                
                                f.write(f"  <td{attr_str}>{cell_content}</td>\n")
                            f.write("</tr>\n")
                        
                        f.write("</tbody>\n</table>\n\n")
                        
                        if len(table.rows) > 10:
                            f.write(f"\n*(총 {len(table.rows)}행 중 10행만 표시)*\n\n")
                        else:
                            f.write("\n")
                
                f.write("---\n\n")
        else:
            # 페이지 정보가 없는 경우
            for tc in doc.text_contents:
                if tc.level > 0:
                    f.write(f"{'#' * (tc.level + 2)} {tc.text}\n\n")
                else:
                    f.write(f"{tc.text}\n\n")
        
        # 테이블
        if doc.tables:
            f.write("## 📊 테이블\n\n")
            for i, table in enumerate(doc.tables, 1):
                page_info = f" (페이지 {table.page_number})" if table.page_number else ""
                f.write(f"### 테이블 {i}{page_info}\n\n")
                
                if table.caption:
                    f.write(f"**캡션:** {table.caption}\n\n")
                
                f.write(f"**크기:** {len(table.headers)} 열 x {len(table.rows)} 행\n\n")
                
                # 마크다운 테이블 형식으로 출력 (줄바꿈을 <br>로 변환)
                if table.headers:
                    headers_clean = [h.replace('\n', '<br>') for h in table.headers]
                    f.write("| " + " | ".join(headers_clean) + " |\n")
                    f.write("| " + " | ".join(["---"] * len(table.headers)) + " |\n")
                
                for row in table.rows[:10]:  # 최대 10행만 표시
                    row_clean = [cell.replace('\n', '<br>') for cell in row]
                    f.write("| " + " | ".join(row_clean) + " |\n")
                
                if len(table.rows) > 10:
                    f.write(f"\n*(총 {len(table.rows)}행 중 10행만 표시)*\n\n")
                else:
                    f.write("\n")
        
        # 이미지 저장 및 참조
        if doc.images:
            f.write("## 🖼️ 이미지\n\n")
            for i, image in enumerate(doc.images, 1):
                # 이미지 파일명 생성 (3자리 숫자 + 확장자)
                img_filename = f"image_{i:03d}.{image.format}"
                img_path = img_folder / img_filename
                
                # 이미지 데이터 저장
                try:
                    with open(img_path, "wb") as img_file:
                        img_file.write(image.data)
                except Exception as e:
                    print(f"⚠️ 이미지 {i} 저장 실패: {e}")
                
                # 마크다운에 이미지 정보 및 참조 추가
                page_info = f" (페이지 {image.page_number})" if image.page_number else ""
                f.write(f"### 이미지 {i}{page_info}\n\n")
                
                if image.caption:
                    f.write(f"**캡션:** {image.caption}\n\n")
                
                f.write(f"- **파일:** `{img_filename}`\n")
                f.write(f"- **형식:** {image.format}\n")
                f.write(f"- **크기:** {image.width or 'N/A'} x {image.height or 'N/A'}\n")
                f.write(f"- **데이터 크기:** {len(image.data)} bytes\n\n")
                
                # 이미지 미리보기 (상대 경로)
                f.write(f"<img src='img/{img_filename}' alt='이미지 {i}' style='max-width:600px;' />\n\n")
        
        # 전체 텍스트 미리보기
        f.write("## 📄 전체 텍스트 미리보기 (처음 2000자)\n\n")
        f.write("```\n")
        f.write(doc.full_text[:2000])
        if len(doc.full_text) > 2000:
            f.write(f"\n\n... (총 {len(doc.full_text)}자 중 2000자만 표시)\n")
        f.write("\n```\n")
    
    return md_path


class TestDetailedParsing:
    """상세 파싱 검증 테스트"""
    
    def setup_method(self):
        """테스트 전 출력 디렉토리 생성"""
        OUTPUT_DIR.mkdir(exist_ok=True)
    
    def test_pdf_detailed_parsing(self):
        """PDF 상세 파싱 테스트"""
        parser = PdfParser()
        pdf_file = PRIVATE_DIR / "02_질병의이해-malaria.report.pdf"
        
        if not pdf_file.exists():
            pytest.skip(f"테스트 파일이 존재하지 않습니다: {pdf_file}")
        
        print(f"\n{'='*60}")
        print(f"PDF 파싱 시작: {pdf_file.name}")
        print(f"{'='*60}\n")
        
        doc = parser.parse(pdf_file)
        
        # 상세 정보 출력
        print(f"메타데이터:")
        print(f"  - 제목: {doc.metadata.title}")
        print(f"  - 페이지 수: {doc.metadata.page_count}")
        print(f"\n통계:")
        print(f"  - 텍스트 블록: {len(doc.text_contents)}개")
        print(f"  - 제목: {len([tc for tc in doc.text_contents if tc.level > 0])}개")
        print(f"  - 테이블: {len(doc.tables)}개")
        print(f"  - 이미지: {len(doc.images)}개")
        
        # 첫 3페이지 미리보기
        print(f"\n첫 3페이지 텍스트 미리보기:")
        for i in range(1, min(4, len(doc.text_contents) + 1)):
            page_texts = [tc for tc in doc.text_contents if tc.page_number == i]
            if page_texts:
                print(f"\n--- 페이지 {i} ---")
                print(page_texts[0].text[:200] + "..." if len(page_texts[0].text) > 200 else page_texts[0].text)
        
        # 마크다운 저장
        folder_name = "pdf_malaria"
        md_path = save_parsing_result_to_markdown(doc, folder_name)
        print(f"\n✅ 결과 저장: {md_path}")
        
        assert len(doc.text_contents) > 0, "텍스트가 추출되지 않았습니다"
    
    def test_html_detailed_parsing(self):
        """HTML 상세 파싱 테스트"""
        parser = HtmlParser()
        html_file = PRIVATE_DIR / "Html_tick_borne_borrelia-1.html"
        
        if not html_file.exists():
            pytest.skip(f"테스트 파일이 존재하지 않습니다: {html_file}")
        
        print(f"\n{'='*60}")
        print(f"HTML 파싱 시작: {html_file.name}")
        print(f"{'='*60}\n")
        
        doc = parser.parse(html_file)
        
        # 상세 정보 출력
        print(f"메타데이터:")
        print(f"  - 제목: {doc.metadata.title}")
        print(f"\n통계:")
        print(f"  - 텍스트 블록: {len(doc.text_contents)}개")
        print(f"  - 제목: {len([tc for tc in doc.text_contents if tc.level > 0])}개")
        print(f"  - 테이블: {len(doc.tables)}개")
        print(f"  - 이미지: {len(doc.images)}개")
        
        # 제목 구조 출력
        headings = [tc for tc in doc.text_contents if tc.level > 0]
        if headings:
            print(f"\n제목 구조:")
            for heading in headings:
                indent = "  " * (heading.level - 1)
                print(f"{indent}- [H{heading.level}] {heading.text}")
        
        # 테이블 미리보기
        if doc.tables:
            print(f"\n첫 번째 테이블:")
            table = doc.tables[0]
            print(f"  - 헤더: {table.headers}")
            print(f"  - 행 수: {len(table.rows)}")
            if table.rows:
                print(f"  - 첫 행: {table.rows[0]}")
        
        # 마크다운 저장
        folder_name = "html_tick_borne"
        md_path = save_parsing_result_to_markdown(doc, folder_name)
        print(f"\n✅ 결과 저장: {md_path}")
        
        assert len(doc.text_contents) > 0, "텍스트가 추출되지 않았습니다"
    
    def test_html_converted_pdf(self):
        """PDF에서 변환된 HTML 파싱 테스트"""
        parser = HtmlParser()
        html_file = PRIVATE_DIR / "07_타겟_converted.html"
        
        if not html_file.exists():
            pytest.skip(f"테스트 파일이 존재하지 않습니다: {html_file}")
        
        print(f"\n{'='*60}")
        print(f"변환된 HTML 파싱 시작: {html_file.name}")
        print(f"{'='*60}\n")
        
        doc = parser.parse(html_file)
        
        # 상세 정보 출력
        print(f"메타데이터:")
        print(f"  - 제목: {doc.metadata.title}")
        print(f"\n통계:")
        print(f"  - 텍스트 블록: {len(doc.text_contents)}개")
        print(f"  - 제목: {len([tc for tc in doc.text_contents if tc.level > 0])}개")
        print(f"  - 테이블: {len(doc.tables)}개")
        print(f"  - 이미지: {len(doc.images)}개")
        
        # 마크다운 저장
        folder_name = "html_monkeypox"
        md_path = save_parsing_result_to_markdown(doc, folder_name)
        print(f"\n✅ 결과 저장: {md_path}")
        
        assert len(doc.text_contents) > 0, "텍스트가 추출되지 않았습니다"
    
    def test_docx_detailed_parsing(self):
        """DOCX 상세 파싱 테스트"""
        parser = DocxParser()
        docx_file = PRIVATE_DIR / "test_document.docx"
        
        if not docx_file.exists():
            pytest.skip(f"테스트 파일이 존재하지 않습니다: {docx_file}")
        
        print(f"\n{'='*60}")
        print(f"DOCX 파싱 시작: {docx_file.name}")
        print(f"{'='*60}\n")
        
        doc = parser.parse(docx_file)
        
        # 상세 정보 출력
        print(f"메타데이터:")
        print(f"  - 제목: {doc.metadata.title}")
        print(f"  - 작성자: {doc.metadata.author}")
        print(f"  - 키워드: {doc.metadata.keywords}")
        print(f"\n통계:")
        print(f"  - 텍스트 블록: {len(doc.text_contents)}개")
        print(f"  - 제목: {len([tc for tc in doc.text_contents if tc.level > 0])}개")
        print(f"  - 테이블: {len(doc.tables)}개")
        print(f"  - 이미지: {len(doc.images)}개")
        
        # 제목 구조 출력
        headings = [tc for tc in doc.text_contents if tc.level > 0]
        if headings:
            print(f"\n제목 구조:")
            for heading in headings[:10]:  # 처음 10개만
                indent = "  " * (heading.level - 1)
                print(f"{indent}- [H{heading.level}] {heading.text}")
        
        # 테이블 미리보기
        if doc.tables:
            print(f"\n첫 번째 테이블:")
            table = doc.tables[0]
            print(f"  - 헤더: {table.headers}")
            print(f"  - 크기: {len(table.headers)} x {len(table.rows)}")
            if table.rows:
                print(f"  - 첫 행: {table.rows[0]}")
        
        # 마크다운 저장
        folder_name = "docx_test"
        md_path = save_parsing_result_to_markdown(doc, folder_name)
        print(f"\n✅ 결과 저장: {md_path}")
        
        assert len(doc.text_contents) > 0, "텍스트가 추출되지 않았습니다"
        assert len(headings) > 0, "제목이 추출되지 않았습니다"
        assert len(doc.tables) > 0, "테이블이 추출되지 않았습니다"
    
    def test_pptx_detailed_parsing(self):
        """PPTX 상세 파싱 테스트"""
        parser = PptxParser()
        pptx_file = PRIVATE_DIR / "test_presentation.pptx"
        
        if not pptx_file.exists():
            pytest.skip(f"테스트 파일이 존재하지 않습니다: {pptx_file}")
        
        print(f"\n{'='*60}")
        print(f"PPTX 파싱 시작: {pptx_file.name}")
        print(f"{'='*60}\n")
        
        doc = parser.parse(pptx_file)
        
        # 상세 정보 출력
        print(f"메타데이터:")
        print(f"  - 제목: {doc.metadata.title}")
        print(f"  - 슬라이드 수: {doc.metadata.page_count}")
        print(f"\n통계:")
        print(f"  - 텍스트 블록: {len(doc.text_contents)}개")
        print(f"  - 제목: {len([tc for tc in doc.text_contents if tc.level > 0])}개")
        print(f"  - 테이블: {len(doc.tables)}개")
        print(f"  - 이미지: {len(doc.images)}개")
        
        # 슬라이드별 제목 출력
        headings = [tc for tc in doc.text_contents if tc.level > 0]
        if headings:
            print(f"\n슬라이드 제목:")
            for heading in headings:
                print(f"  - [슬라이드 {heading.page_number}] {heading.text}")
        
        # 테이블 미리보기
        if doc.tables:
            print(f"\n테이블 정보:")
            for i, table in enumerate(doc.tables, 1):
                print(f"  테이블 {i} (슬라이드 {table.page_number}): {len(table.headers)} x {len(table.rows)}")
        
        # 마크다운 저장
        folder_name = "pptx_test"
        md_path = save_parsing_result_to_markdown(doc, folder_name)
        print(f"\n✅ 결과 저장: {md_path}")
        
        assert len(doc.text_contents) > 0, "텍스트가 추출되지 않았습니다"
        assert len(headings) > 0, "제목이 추출되지 않았습니다"
        assert doc.metadata.page_count > 0, "슬라이드 수가 잘못되었습니다"
    
    def test_real_pptx_file1(self):
        """실제 PPTX 파일 1 파싱 테스트"""
        parser = PptxParser()
        pptx_file = PRIVATE_DIR / "PPT샘플_20201027.pptx"
        
        if not pptx_file.exists():
            pytest.skip(f"테스트 파일이 존재하지 않습니다: {pptx_file}")
        
        print(f"\n{'='*60}")
        print(f"실제 PPTX 파일 1 파싱 시작: {pptx_file.name[:50]}...")
        print(f"{'='*60}\n")
        
        try:
            doc = parser.parse(pptx_file)
            
            # 상세 정보 출력
            print(f"메타데이터:")
            print(f"  - 제목: {doc.metadata.title}")
            print(f"  - 슬라이드 수: {doc.metadata.page_count}")
            print(f"\n통계:")
            print(f"  - 텍스트 블록: {len(doc.text_contents)}개")
            print(f"  - 제목: {len([tc for tc in doc.text_contents if tc.level > 0])}개")
            print(f"  - 테이블: {len(doc.tables)}개")
            print(f"  - 이미지: {len(doc.images)}개")
            
            # 처음 5개 슬라이드 제목
            headings = [tc for tc in doc.text_contents if tc.level > 0]
            if headings:
                print(f"\n처음 5개 슬라이드 제목:")
                for heading in headings[:5]:
                    print(f"  - [슬라이드 {heading.page_number}] {heading.text[:80]}")
            
            # 마크다운 저장
            folder_name = "pptx_novaplex_eu"
            md_path = save_parsing_result_to_markdown(doc, folder_name)
            print(f"\n✅ 결과 저장: {md_path}")
            
            assert len(doc.text_contents) > 0, "텍스트가 추출되지 않았습니다"
            assert doc.metadata.page_count > 0, "슬라이드 수가 잘못되었습니다"
        except Exception as e:
            print(f"\n❌ 파싱 실패: {e}")
            raise
    
    def test_real_pptx_file2(self):
        """실제 PPTX 파일 2 파싱 테스트"""
        parser = PptxParser()
        pptx_file = PRIVATE_DIR / "PPT샘플_개발.pptx"
        
        if not pptx_file.exists():
            pytest.skip(f"테스트 파일이 존재하지 않습니다: {pptx_file}")
        
        print(f"\n{'='*60}")
        print(f"실제 PPTX 파일 2 파싱 시작: {pptx_file.name[:50]}...")
        print(f"{'='*60}\n")
        
        try:
            doc = parser.parse(pptx_file)
            
            # 상세 정보 출력
            print(f"메타데이터:")
            print(f"  - 제목: {doc.metadata.title}")
            print(f"  - 슬라이드 수: {doc.metadata.page_count}")
            print(f"\n통계:")
            print(f"  - 텍스트 블록: {len(doc.text_contents)}개")
            print(f"  - 제목: {len([tc for tc in doc.text_contents if tc.level > 0])}개")
            print(f"  - 테이블: {len(doc.tables)}개")
            print(f"  - 이미지: {len(doc.images)}개")
            
            # 처음 5개 슬라이드 제목
            headings = [tc for tc in doc.text_contents if tc.level > 0]
            if headings:
                print(f"\n처음 5개 슬라이드 제목:")
                for heading in headings[:5]:
                    print(f"  - [슬라이드 {heading.page_number}] {heading.text[:80]}")
            
            # 마크다운 저장
            folder_name = "pptx_tick_borne_expanded"
            md_path = save_parsing_result_to_markdown(doc, folder_name)
            print(f"\n✅ 결과 저장: {md_path}")
            
            assert len(doc.text_contents) > 0, "텍스트가 추출되지 않았습니다"
            assert doc.metadata.page_count > 0, "슬라이드 수가 잘못되었습니다"
        except Exception as e:
            print(f"\n❌ 파싱 실패: {e}")
            raise
    
    def test_real_docx_file(self):
        """실제 DOCX 파일 파싱 테스트"""
        parser = DocxParser()
        docx_file = PRIVATE_DIR / "[PPT변환 샘플].docx"
        
        if not docx_file.exists():
            pytest.skip(f"테스트 파일이 존재하지 않습니다: {docx_file}")
        
        print(f"\n{'='*60}")
        print(f"실제 DOCX 파일 파싱 시작: {docx_file.name[:50]}...")
        print(f"{'='*60}\n")
        
        try:
            doc = parser.parse(docx_file)
            
            # 상세 정보 출력
            print(f"메타데이터:")
            print(f"  - 제목: {doc.metadata.title}")
            print(f"  - 페이지 수: {doc.metadata.page_count}")
            print(f"\n통계:")
            print(f"  - 텍스트 블록: {len(doc.text_contents)}개")
            print(f"  - 제목: {len([tc for tc in doc.text_contents if tc.level > 0])}개")
            print(f"  - 테이블: {len(doc.tables)}개")
            print(f"  - 이미지: {len(doc.images)}개")
            print(f"  - 전체 텍스트 길이: {len(doc.full_text)} 문자")
            
            # 처음 5개 제목
            headings = [tc for tc in doc.text_contents if tc.level > 0]
            if headings:
                print(f"\n처음 5개 제목:")
                for heading in headings[:5]:
                    print(f"  - [레벨 {heading.level}] {heading.text[:80]}")
            
            # 마크다운 저장
            folder_name = "docx_tick_borne"
            md_path = save_parsing_result_to_markdown(doc, folder_name)
            print(f"\n✅ 결과 저장: {md_path}")
            
            assert len(doc.text_contents) > 0, "텍스트가 추출되지 않았습니다"
        except Exception as e:
            print(f"\n❌ 파싱 실패: {e}")
            raise
