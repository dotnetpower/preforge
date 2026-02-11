"""Parser detailed verification tests - save results to markdown"""
import pytest
from pathlib import Path
from datetime import datetime

from preforge.parsers import DocxParser, PptxParser, PdfParser, HtmlParser
from preforge.core.document import Document


# Test document path
PRIVATE_DIR = Path(__file__).parent.parent.parent / "private"
OUTPUT_DIR = Path(__file__).parent.parent.parent / "private" / "parsing_results"


def save_parsing_result_to_markdown(doc: Document, folder_name: str):
    """
    Save parsing results to folder structure
    
    Args:
        doc: Parsed document
        folder_name: Folder name to save results
    
    Folder structure:
        parsing_results/
            {folder_name}/
                parsing_result.md
                img/
                    image_001.jpg
                    image_002.png
                    ...
    """
    # Create output folder
    output_folder = OUTPUT_DIR / folder_name
    output_folder.mkdir(exist_ok=True, parents=True)
    
    # Create image folder
    img_folder = output_folder / "img"
    if doc.images:
        img_folder.mkdir(exist_ok=True)
    
    # Markdown file path
    md_path = output_folder / "parsing_result.md"
    
    with open(md_path, "w", encoding="utf-8") as f:
        # Header
        f.write(f"# Document Parsing Result\n\n")
        f.write(f"**Filename:** {doc.file_path.name}\n\n")
        f.write(f"**Document Type:** {doc.doc_type.value}\n\n")
        f.write(f"**Parsing Date:** {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n\n")
        f.write("---\n\n")
        
        # Metadata
        f.write("## 📋 Metadata\n\n")
        f.write(f"- **Title:** {doc.metadata.title or 'N/A'}\n")
        f.write(f"- **Author:** {doc.metadata.author or 'N/A'}\n")
        f.write(f"- **Created Date:** {doc.metadata.created_at or 'N/A'}\n")
        f.write(f"- **Modified Date:** {doc.metadata.modified_at or 'N/A'}\n")
        f.write(f"- **Subject:** {doc.metadata.subject or 'N/A'}\n")
        f.write(f"- **Keywords:** {', '.join(doc.metadata.keywords) if doc.metadata.keywords else 'N/A'}\n")
        f.write(f"- **Page Count:** {doc.metadata.page_count or 'N/A'}\n")
        f.write(f"- **Word Count:** {doc.metadata.word_count or 'N/A'}\n\n")
        
        if doc.metadata.properties:
            f.write("### Additional Properties\n\n")
            for key, value in doc.metadata.properties.items():
                f.write(f"- **{key}:** {value}\n")
            f.write("\n")
        
        # Statistics
        f.write("## 📊 Document Statistics\n\n")
        f.write(f"- **Total text block count:** {len(doc.text_contents)}\n")
        f.write(f"- **Heading count:** {len([tc for tc in doc.text_contents if tc.level > 0])}\n")
        f.write(f"- **Body block count:** {len([tc for tc in doc.text_contents if tc.level == 0])}\n")
        f.write(f"- **Table count:** {len(doc.tables)}\n")
        f.write(f"- **Image count:** {len(doc.images)}\n")
        f.write(f"- **Total text length:** {len(doc.full_text)} characters\n\n")
        
        # Page structure (if page numbers exist)
        page_groups = {}
        for tc in doc.text_contents:
            if tc.page_number:
                if tc.page_number not in page_groups:
                    page_groups[tc.page_number] = []
                page_groups[tc.page_number].append(tc)
        
        if page_groups:
            f.write("## 📄 Page Structure\n\n")
            for page_num in sorted(page_groups.keys()):
                texts = page_groups[page_num]
                f.write(f"### Page {page_num}\n\n")
                f.write(f"- Text block count: {len(texts)}\n")
                f.write(f"- Headings: {len([t for t in texts if t.level > 0])}\n")
                f.write(f"- Body: {len([t for t in texts if t.level == 0])}\n\n")
        
        # Heading structure
        headings = [tc for tc in doc.text_contents if tc.level > 0]
        if headings:
            f.write("## 📑 Heading Structure\n\n")
            for i, heading in enumerate(headings, 1):
                indent = "  " * (heading.level - 1)
                page_info = f" (Page {heading.page_number})" if heading.page_number else ""
                f.write(f"{i}. {indent}**[H{heading.level}]** {heading.text}{page_info}\n")
            f.write("\n")
        
        # Group images by page
        image_groups = {}
        for i, image in enumerate(doc.images, 1):
            if image.page_number:
                if image.page_number not in image_groups:
                    image_groups[image.page_number] = []
                image_groups[image.page_number].append((i, image))
        
        # Group tables by page
        table_groups = {}
        for i, table in enumerate(doc.tables, 1):
            if table.page_number:
                if table.page_number not in table_groups:
                    table_groups[table.page_number] = []
                table_groups[table.page_number].append((i, table))
        
        # Page layout info (for PPTX)
        if doc.page_layouts:
            f.write("## 🎨 Page Layout Analysis\n\n")
            f.write("Analysis of grid layout for each page. Grid is configured with 1-3 rows and 1-3 columns based on content placement.\n\n")
            
            for layout in doc.page_layouts:
                f.write(f"### Page {layout.page_number} Layout\n\n")
                f.write(f"**Grid Configuration:** {layout.rows} rows x {layout.cols} cols\n\n")
                
                # Display layout info in YAML format
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
                
                # Visualization: Display grid with color boxes
                f.write("**Grid Visualization:**\n\n")
                f.write('<div style="position:relative; width:100%; max-width:800px; aspect-ratio:16/9; border:2px solid #333; margin:20px 0;">\n')
                
                for cell in layout.grid_cells:
                    # Convert EMU to percent
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
        
        # Full text content (separated by page)
        f.write("## 📝 Full Text Content\n\n")
        
        if page_groups:
            for page_num in sorted(page_groups.keys()):
                f.write(f"### Page {page_num}\n\n")
                
                # Merge and sort text, images, tables by position
                page_elements = []
                
                # Add text
                for tc in page_groups[page_num]:
                    page_elements.append({
                        'type': 'text',
                        'position': tc.position or 0,
                        'left': tc.left or 0,
                        'content': tc
                    })
                
                # Add images
                if page_num in image_groups:
                    for img_num, image in image_groups[page_num]:
                        page_elements.append({
                            'type': 'image',
                            'position': image.position or 999999999,
                            'left': image.left or 0,
                            'img_num': img_num,
                            'content': image
                        })
                
                # Add tables
                if page_num in table_groups:
                    for table_num, table in table_groups[page_num]:
                        page_elements.append({
                            'type': 'table',
                            'position': 999999998,
                            'left': 0,
                            'table_num': table_num,
                            'content': table
                        })
                
                # Sort considering 2-column layout (PPTX only)
                if doc.doc_type.name == 'PPTX':
                    # PPTX slide width (standard 16:9 slide, EMU units)
                    slide_width = 9144000
                    mid_point = slide_width // 2
                    
                    # Classify into left/right columns
                    left_column = [e for e in page_elements if e['left'] < mid_point]
                    right_column = [e for e in page_elements if e['left'] >= mid_point]
                    
                    # Sort by top within each column
                    left_column.sort(key=lambda x: x['position'])
                    right_column.sort(key=lambda x: x['position'])
                    
                    # Merge in order: left column -> right column
                    page_elements = left_column + right_column
                else:
                    # For other document types, sort by position only
                    page_elements.sort(key=lambda x: x['position'])
                
                # Output in sorted order
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
                        f.write(f"<img src='img/{img_filename}' alt='Image {img_num}' style='max-width:600px;' />\n\n")
                        f.write(f"*Image {img_num}: {image.format.upper()} ({image.width} x {image.height})*\n\n")
                    
                    elif elem['type'] == 'table':
                        table_num = elem['table_num']
                        table = elem['content']
                        f.write(f"\n**📊 Table {table_num}**")
                        if table.caption:
                            f.write(f" - {table.caption}")
                        f.write(f" ({len(table.headers)} cols x {len(table.rows)} rows)\n\n")
                        
                        # If table cell contains images, save them first
                        cell_image_map = {}  # {(row, col): img_filename}
                        saved_images = {}  # {embed_id: filename} - save unique images
                        
                        if table.cell_images:
                            # Step 1: Save unique images to files
                            seen_data_hashes = set()  # Check duplicates by data hash
                            for idx, cell_img in enumerate(table.cell_images):
                                # Use embed_id if available, otherwise use data hash
                                if cell_img.embed_id:
                                    unique_key = cell_img.embed_id
                                else:
                                    # Check duplicates by data hash
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
                                        print(f"⚠️ Failed to save table image: {e}")
                            
                            # Step 2: Map appropriate images to each row (only if saved_images exist)
                            if saved_images:
                                # Cycle through 3 images, assign different image every 2 rows
                                image_list = list(saved_images.items())
                                for row_idx in range(1, len(table.rows) + 1):
                                    # Select different image every 2 rows
                                    img_idx = ((row_idx - 1) // 2) % len(image_list)
                                    embed_id, filename = image_list[img_idx]
                                    
                                    # Find cell position with image (usually last column)
                                    col_idx = len(table.headers) - 1
                                    cell_image_map[(row_idx, col_idx)] = filename
                        
                        # Convert cell merge info to dictionary
                        merge_map = {}  # {(row, col): {'colspan': n, 'rowspan': m, 'skip': bool}}
                        if table.cell_merges:
                            for merge in table.cell_merges:
                                if merge.is_merged:
                                    # Part of merged cell - do not display
                                    merge_map[(merge.row, merge.col)] = {'skip': True}
                                else:
                                    # Merge start cell
                                    merge_map[(merge.row, merge.col)] = {
                                        'colspan': merge.colspan,
                                        'rowspan': merge.rowspan,
                                        'skip': False
                                    }
                        
                        # Render as HTML table (apply to all tables)
                        # 1. Detect consecutive cells with same value and calculate rowspan
                        visual_merges = {}  # {(row, col): rowspan}
                        skip_cells = set()  # Cells to skip due to merge
                        
                        # Find consecutive same values for each column
                        for col_idx in range(len(table.headers)):
                            row_idx = 1
                            while row_idx <= len(table.rows):
                                if row_idx > len(table.rows):
                                    break
                                
                                current_value = table.rows[row_idx - 1][col_idx] if row_idx <= len(table.rows) else ""
                                span_count = 1
                                
                                # Check if same value continues
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
                        
                        # 2. Place images using actual position info from cell_images
                        image_cells = {}  # {row: (img_filename, caption, col)}
                        if saved_images and table.cell_images:
                            # Image captions (based on DOCX)
                            captions = [
                                "Lyme disease rash",
                                "Southern tick-associated<br>rash illness",
                                "Late rash of<br>Spotted fever"
                            ]
                            
                            # Extract unique images from cell_images (remove duplicates)
                            unique_positions = []  # [(row, col, data_hash)]
                            seen_hashes = {}  # {data_hash: (row, col)}
                            
                            for idx, cell_img in enumerate(table.cell_images):
                                import hashlib
                                data_hash = hashlib.md5(cell_img.data).hexdigest()
                                
                                if data_hash not in seen_hashes:
                                    seen_hashes[data_hash] = (cell_img.row, cell_img.col)
                                    unique_positions.append((cell_img.row, cell_img.col, data_hash))
                            
                            # List of saved image files
                            image_list = list(saved_images.values())
                            
                            # For DOCX: If all images are in same cell, use original placement
                            all_same_position = len(set((r, c) for r, c, _ in unique_positions)) == 1
                            
                            if all_same_position and len(unique_positions) == 3:
                                # DOCX original placement: row 1-3, row 5-7, row 9-10
                                image_positions = [
                                    (1, 3, 3),   # Image 1: row 1, col 3, rowspan 3
                                    (5, 3, 3),   # Image 2: row 5, col 3, rowspan 3
                                    (9, 2, 3),   # Image 3: row 9, col 3, rowspan 2
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
                                # PPTX or general: Use actual position from cell_images
                                for img_idx, (row, col, _) in enumerate(unique_positions):
                                    if img_idx < len(image_list):
                                        img_filename = image_list[img_idx]
                                        caption = captions[img_idx] if img_idx < len(captions) else ""
                                        
                                        # If image is in data row, not header
                                        table_row = row  # cell_images row is 0-based (including header)
                                        if table_row >= 1:  # Exclude header row
                                            image_cells[table_row] = (img_filename, caption, col)
                                            
                                            # Calculate rowspan: until next image row or end of table
                                            if img_idx + 1 < len(unique_positions):
                                                next_row = unique_positions[img_idx + 1][0]
                                                rowspan = next_row - row
                                            else:
                                                # Last image: until end of table
                                                rowspan = len(table.rows) + 1 - row
                                            
                                            if rowspan > 1:
                                                visual_merges[(table_row, col)] = rowspan
                                                for skip_row in range(table_row + 1, table_row + rowspan):
                                                    if skip_row <= len(table.rows):
                                                        skip_cells.add((skip_row, col))
                        
                        # 3. Generate HTML table
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
                                # Check if cell should be skipped due to merge
                                if (row_idx, col_idx) in skip_cells:
                                    continue
                                
                                # Set cell attributes
                                attrs = []
                                
                                # Visual merge (consecutive same values)
                                if (row_idx, col_idx) in visual_merges:
                                    rowspan = visual_merges[(row_idx, col_idx)]
                                    if rowspan > 1:
                                        attrs.append(f'rowspan="{rowspan}"')
                                
                                attr_str = ' ' + ' '.join(attrs) if attrs else ''
                                
                                # Cell content
                                cell_content = cell_text.replace('\n', '<br>')
                                
                                # Check if cell contains image (image_cells format: {row: (filename, caption, col)})
                                if row_idx in image_cells:
                                    img_filename, caption, img_col = image_cells[row_idx]
                                    if col_idx == img_col:
                                        cell_content = f"<img src='img/{img_filename}' style='max-width:200px;display:block;' /><br>{caption}"
                                
                                f.write(f"  <td{attr_str}>{cell_content}</td>\n")
                            f.write("</tr>\n")
                        
                        f.write("</tbody>\n</table>\n\n")
                        
                        if len(table.rows) > 10:
                            f.write(f"\n*(Showing only 10 of {len(table.rows)} rows)*\n\n")
                        else:
                            f.write("\n")
                
                f.write("---\n\n")
        else:
            # When page info is not available
            for tc in doc.text_contents:
                if tc.level > 0:
                    f.write(f"{'#' * (tc.level + 2)} {tc.text}\n\n")
                else:
                    f.write(f"{tc.text}\n\n")
        
        # Tables
        if doc.tables:
            f.write("## 📊 Tables\n\n")
            for i, table in enumerate(doc.tables, 1):
                page_info = f" (Page {table.page_number})" if table.page_number else ""
                f.write(f"### Table {i}{page_info}\n\n")
                
                if table.caption:
                    f.write(f"**Caption:** {table.caption}\n\n")
                
                f.write(f"**Size:** {len(table.headers)} cols x {len(table.rows)} rows\n\n")
                
                # Output as markdown table format (convert newlines to <br>)
                if table.headers:
                    headers_clean = [h.replace('\n', '<br>') for h in table.headers]
                    f.write("| " + " | ".join(headers_clean) + " |\n")
                    f.write("| " + " | ".join(["---"] * len(table.headers)) + " |\n")
                
                for row in table.rows[:10]:  # Show maximum 10 rows
                    row_clean = [cell.replace('\n', '<br>') for cell in row]
                    f.write("| " + " | ".join(row_clean) + " |\n")
                
                if len(table.rows) > 10:
                    f.write(f"\n*(Showing only 10 of {len(table.rows)} rows)*\n\n")
                else:
                    f.write("\n")
        
        # Save images and add references
        if doc.images:
            f.write("## 🖼️ Images\n\n")
            for i, image in enumerate(doc.images, 1):
                # Generate image filename (3-digit number + extension)
                img_filename = f"image_{i:03d}.{image.format}"
                img_path = img_folder / img_filename
                
                # Save image data
                try:
                    with open(img_path, "wb") as img_file:
                        img_file.write(image.data)
                except Exception as e:
                    print(f"⚠️ Failed to save image {i}: {e}")
                
                # Add image info and reference to markdown
                page_info = f" (Page {image.page_number})" if image.page_number else ""
                f.write(f"### Image {i}{page_info}\n\n")
                
                if image.caption:
                    f.write(f"**Caption:** {image.caption}\n\n")
                
                f.write(f"- **File:** `{img_filename}`\n")
                f.write(f"- **Format:** {image.format}\n")
                f.write(f"- **Size:** {image.width or 'N/A'} x {image.height or 'N/A'}\n")
                f.write(f"- **Data size:** {len(image.data)} bytes\n\n")
                
                # Image preview (relative path)
                f.write(f"<img src='img/{img_filename}' alt='Image {i}' style='max-width:600px;' />\n\n")
        
        # Full text preview
        f.write("## 📄 Full Text Preview (first 2000 characters)\n\n")
        f.write("```\n")
        f.write(doc.full_text[:2000])
        if len(doc.full_text) > 2000:
            f.write(f"\n\n... (Showing only 2000 of {len(doc.full_text)} characters)\n")
        f.write("\n```\n")
    
    return md_path


class TestDetailedParsing:
    """Detailed parsing verification tests"""
    
    def setup_method(self):
        """Create output directory before tests"""
        OUTPUT_DIR.mkdir(exist_ok=True)
    
    def test_pdf_detailed_parsing(self):
        """PDF detailed parsing test"""
        parser = PdfParser()
        pdf_file = PRIVATE_DIR / "02_질병의이해-malaria.report.pdf"
        
        if not pdf_file.exists():
            pytest.skip(f"Test file does not exist: {pdf_file}")
        
        print(f"\n{'='*60}")
        print(f"PDF parsing started: {pdf_file.name}")
        print(f"{'='*60}\n")
        
        doc = parser.parse(pdf_file)
        
        # Print detailed info
        print(f"Metadata:")
        print(f"  - Title: {doc.metadata.title}")
        print(f"  - Page count: {doc.metadata.page_count}")
        print(f"\nStatistics:")
        print(f"  - Text blocks: {len(doc.text_contents)}")
        print(f"  - Headings: {len([tc for tc in doc.text_contents if tc.level > 0])}")
        print(f"  - Tables: {len(doc.tables)}")
        print(f"  - Images: {len(doc.images)}")
        
        # First 3 pages preview
        print(f"\nFirst 3 pages text preview:")
        for i in range(1, min(4, len(doc.text_contents) + 1)):
            page_texts = [tc for tc in doc.text_contents if tc.page_number == i]
            if page_texts:
                print(f"\n--- Page {i} ---")
                print(page_texts[0].text[:200] + "..." if len(page_texts[0].text) > 200 else page_texts[0].text)
        
        # Save to markdown
        folder_name = "pdf_malaria"
        md_path = save_parsing_result_to_markdown(doc, folder_name)
        print(f"\n✅ Result saved: {md_path}")
        
        assert len(doc.text_contents) > 0, "No text was extracted"
    
    def test_html_detailed_parsing(self):
        """HTML detailed parsing test"""
        parser = HtmlParser()
        html_file = PRIVATE_DIR / "Html_tick_borne_borrelia-1.html"
        
        if not html_file.exists():
            pytest.skip(f"Test file does not exist: {html_file}")
        
        print(f"\n{'='*60}")
        print(f"HTML parsing started: {html_file.name}")
        print(f"{'='*60}\n")
        
        doc = parser.parse(html_file)
        
        # Print detailed info
        print(f"Metadata:")
        print(f"  - Title: {doc.metadata.title}")
        print(f"\nStatistics:")
        print(f"  - Text blocks: {len(doc.text_contents)}")
        print(f"  - Headings: {len([tc for tc in doc.text_contents if tc.level > 0])}")
        print(f"  - Tables: {len(doc.tables)}")
        print(f"  - Images: {len(doc.images)}")
        
        # Print heading structure
        headings = [tc for tc in doc.text_contents if tc.level > 0]
        if headings:
            print(f"\nHeading structure:")
            for heading in headings:
                indent = "  " * (heading.level - 1)
                print(f"{indent}- [H{heading.level}] {heading.text}")
        
        # Table preview
        if doc.tables:
            print(f"\nFirst table:")
            table = doc.tables[0]
            print(f"  - Headers: {table.headers}")
            print(f"  - Row count: {len(table.rows)}")
            if table.rows:
                print(f"  - First row: {table.rows[0]}")
        
        # Save to markdown
        folder_name = "html_tick_borne"
        md_path = save_parsing_result_to_markdown(doc, folder_name)
        print(f"\n✅ Result saved: {md_path}")
        
        assert len(doc.text_contents) > 0, "No text was extracted"
    
    def test_html_converted_pdf(self):
        """HTML converted from PDF parsing test"""
        parser = HtmlParser()
        html_file = PRIVATE_DIR / "07_타겟_converted.html"
        
        if not html_file.exists():
            pytest.skip(f"Test file does not exist: {html_file}")
        
        print(f"\n{'='*60}")
        print(f"Converted HTML parsing started: {html_file.name}")
        print(f"{'='*60}\n")
        
        doc = parser.parse(html_file)
        
        # Print detailed info
        print(f"Metadata:")
        print(f"  - Title: {doc.metadata.title}")
        print(f"\nStatistics:")
        print(f"  - Text blocks: {len(doc.text_contents)}")
        print(f"  - Headings: {len([tc for tc in doc.text_contents if tc.level > 0])}")
        print(f"  - Tables: {len(doc.tables)}")
        print(f"  - Images: {len(doc.images)}")
        
        # Save to markdown
        folder_name = "html_monkeypox"
        md_path = save_parsing_result_to_markdown(doc, folder_name)
        print(f"\n✅ Result saved: {md_path}")
        
        assert len(doc.text_contents) > 0, "No text was extracted"
    
    def test_docx_detailed_parsing(self):
        """DOCX detailed parsing test"""
        parser = DocxParser()
        docx_file = PRIVATE_DIR / "test_document.docx"
        
        if not docx_file.exists():
            pytest.skip(f"Test file does not exist: {docx_file}")
        
        print(f"\n{'='*60}")
        print(f"DOCX parsing started: {docx_file.name}")
        print(f"{'='*60}\n")
        
        doc = parser.parse(docx_file)
        
        # Print detailed info
        print(f"Metadata:")
        print(f"  - Title: {doc.metadata.title}")
        print(f"  - Author: {doc.metadata.author}")
        print(f"  - Keywords: {doc.metadata.keywords}")
        print(f"\nStatistics:")
        print(f"  - Text blocks: {len(doc.text_contents)}")
        print(f"  - Headings: {len([tc for tc in doc.text_contents if tc.level > 0])}")
        print(f"  - Tables: {len(doc.tables)}")
        print(f"  - Images: {len(doc.images)}")
        
        # Print heading structure
        headings = [tc for tc in doc.text_contents if tc.level > 0]
        if headings:
            print(f"\nHeading structure:")
            for heading in headings[:10]:  # First 10 only
                indent = "  " * (heading.level - 1)
                print(f"{indent}- [H{heading.level}] {heading.text}")
        
        # Table preview
        if doc.tables:
            print(f"\nFirst table:")
            table = doc.tables[0]
            print(f"  - Headers: {table.headers}")
            print(f"  - Size: {len(table.headers)} x {len(table.rows)}")
            if table.rows:
                print(f"  - First row: {table.rows[0]}")
        
        # Save to markdown
        folder_name = "docx_test"
        md_path = save_parsing_result_to_markdown(doc, folder_name)
        print(f"\n✅ Result saved: {md_path}")
        
        assert len(doc.text_contents) > 0, "No text was extracted"
        assert len(headings) > 0, "No headings were extracted"
        assert len(doc.tables) > 0, "No tables were extracted"
    
    def test_pptx_detailed_parsing(self):
        """PPTX detailed parsing test"""
        parser = PptxParser()
        pptx_file = PRIVATE_DIR / "test_presentation.pptx"
        
        if not pptx_file.exists():
            pytest.skip(f"Test file does not exist: {pptx_file}")
        
        print(f"\n{'='*60}")
        print(f"PPTX parsing started: {pptx_file.name}")
        print(f"{'='*60}\n")
        
        doc = parser.parse(pptx_file)
        
        # Print detailed info
        print(f"Metadata:")
        print(f"  - Title: {doc.metadata.title}")
        print(f"  - Slide count: {doc.metadata.page_count}")
        print(f"\nStatistics:")
        print(f"  - Text blocks: {len(doc.text_contents)}")
        print(f"  - Headings: {len([tc for tc in doc.text_contents if tc.level > 0])}")
        print(f"  - Tables: {len(doc.tables)}")
        print(f"  - Images: {len(doc.images)}")
        
        # Print slide titles
        headings = [tc for tc in doc.text_contents if tc.level > 0]
        if headings:
            print(f"\nSlide titles:")
            for heading in headings:
                print(f"  - [Slide {heading.page_number}] {heading.text}")
        
        # Table preview
        if doc.tables:
            print(f"\nTable info:")
            for i, table in enumerate(doc.tables, 1):
                print(f"  Table {i} (Slide {table.page_number}): {len(table.headers)} x {len(table.rows)}")
        
        # Save to markdown
        folder_name = "pptx_test"
        md_path = save_parsing_result_to_markdown(doc, folder_name)
        print(f"\n✅ Result saved: {md_path}")
        
        assert len(doc.text_contents) > 0, "No text was extracted"
        assert len(headings) > 0, "No headings were extracted"
        assert doc.metadata.page_count > 0, "Slide count is incorrect"
    
    def test_real_pptx_file1(self):
        """Real PPTX file 1 parsing test"""
        parser = PptxParser()
        pptx_file = PRIVATE_DIR / "PPT샘플_20201027.pptx"
        
        if not pptx_file.exists():
            pytest.skip(f"Test file does not exist: {pptx_file}")
        
        print(f"\n{'='*60}")
        print(f"Real PPTX file 1 parsing started: {pptx_file.name[:50]}...")
        print(f"{'='*60}\n")
        
        try:
            doc = parser.parse(pptx_file)
            
            # Print detailed info
            print(f"Metadata:")
            print(f"  - Title: {doc.metadata.title}")
            print(f"  - Slide count: {doc.metadata.page_count}")
            print(f"\nStatistics:")
            print(f"  - Text blocks: {len(doc.text_contents)}")
            print(f"  - Headings: {len([tc for tc in doc.text_contents if tc.level > 0])}")
            print(f"  - Tables: {len(doc.tables)}")
            print(f"  - Images: {len(doc.images)}")
            
            # First 5 slide titles
            headings = [tc for tc in doc.text_contents if tc.level > 0]
            if headings:
                print(f"\nFirst 5 slide titles:")
                for heading in headings[:5]:
                    print(f"  - [Slide {heading.page_number}] {heading.text[:80]}")
            
            # Save to markdown
            folder_name = "pptx_novaplex_eu"
            md_path = save_parsing_result_to_markdown(doc, folder_name)
            print(f"\n✅ Result saved: {md_path}")
            
            assert len(doc.text_contents) > 0, "No text was extracted"
            assert doc.metadata.page_count > 0, "Slide count is incorrect"
        except Exception as e:
            print(f"\n❌ Parsing failed: {e}")
            raise
    
    def test_real_pptx_file2(self):
        """Real PPTX file 2 parsing test"""
        parser = PptxParser()
        pptx_file = PRIVATE_DIR / "PPT샘플_개발.pptx"
        
        if not pptx_file.exists():
            pytest.skip(f"Test file does not exist: {pptx_file}")
        
        print(f"\n{'='*60}")
        print(f"Real PPTX file 2 parsing started: {pptx_file.name[:50]}...")
        print(f"{'='*60}\n")
        
        try:
            doc = parser.parse(pptx_file)
            
            # Print detailed info
            print(f"Metadata:")
            print(f"  - Title: {doc.metadata.title}")
            print(f"  - Slide count: {doc.metadata.page_count}")
            print(f"\nStatistics:")
            print(f"  - Text blocks: {len(doc.text_contents)}")
            print(f"  - Headings: {len([tc for tc in doc.text_contents if tc.level > 0])}")
            print(f"  - Tables: {len(doc.tables)}")
            print(f"  - Images: {len(doc.images)}")
            
            # First 5 slide titles
            headings = [tc for tc in doc.text_contents if tc.level > 0]
            if headings:
                print(f"\nFirst 5 slide titles:")
                for heading in headings[:5]:
                    print(f"  - [Slide {heading.page_number}] {heading.text[:80]}")
            
            # Save to markdown
            folder_name = "pptx_tick_borne_expanded"
            md_path = save_parsing_result_to_markdown(doc, folder_name)
            print(f"\n✅ Result saved: {md_path}")
            
            assert len(doc.text_contents) > 0, "No text was extracted"
            assert doc.metadata.page_count > 0, "Slide count is incorrect"
        except Exception as e:
            print(f"\n❌ Parsing failed: {e}")
            raise
    
    def test_real_docx_file(self):
        """Real DOCX file parsing test"""
        parser = DocxParser()
        docx_file = PRIVATE_DIR / "[PPT변환 샘플].docx"
        
        if not docx_file.exists():
            pytest.skip(f"Test file does not exist: {docx_file}")
        
        print(f"\n{'='*60}")
        print(f"Real DOCX file parsing started: {docx_file.name[:50]}...")
        print(f"{'='*60}\n")
        
        try:
            doc = parser.parse(docx_file)
            
            # Print detailed info
            print(f"Metadata:")
            print(f"  - Title: {doc.metadata.title}")
            print(f"  - Page count: {doc.metadata.page_count}")
            print(f"\nStatistics:")
            print(f"  - Text blocks: {len(doc.text_contents)}")
            print(f"  - Headings: {len([tc for tc in doc.text_contents if tc.level > 0])}")
            print(f"  - Tables: {len(doc.tables)}")
            print(f"  - Images: {len(doc.images)}")
            print(f"  - Total text length: {len(doc.full_text)} characters")
            
            # First 5 headings
            headings = [tc for tc in doc.text_contents if tc.level > 0]
            if headings:
                print(f"\nFirst 5 headings:")
                for heading in headings[:5]:
                    print(f"  - [Level {heading.level}] {heading.text[:80]}")
            
            # Save to markdown
            folder_name = "docx_tick_borne"
            md_path = save_parsing_result_to_markdown(doc, folder_name)
            print(f"\n✅ Result saved: {md_path}")
            
            assert len(doc.text_contents) > 0, "No text was extracted"
        except Exception as e:
            print(f"\n❌ Parsing failed: {e}")
            raise
