"""
Slide creation factory module

Provides functionality to create various types of slides (title, table, image, etc.).
"""
import logging
import base64
from io import BytesIO
from typing import List, Optional, Any
from bs4 import Tag
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.dml.color import RGBColor

try:
    from PIL import Image
    HAS_PIL = True
except ImportError:
    HAS_PIL = False

from .config import SlideConfig, ColorPalette, DEFAULT_SLIDE_CONFIG, DEFAULT_COLORS
from .table_builder import TableBuilder, TableDataExtractor
from .style_utils import TextUtils

logger = logging.getLogger(__name__)


class SlideFactory:
    """Slide creation factory"""
    
    def __init__(
        self,
        presentation: Presentation,
        slide_config: SlideConfig = None,
        colors: ColorPalette = None
    ):
        self.prs = presentation
        self.config = slide_config or DEFAULT_SLIDE_CONFIG
        self.colors = colors or DEFAULT_COLORS
        self.table_builder = TableBuilder(colors=colors)
    
    def _get_blank_slide(self):
        """Create blank layout slide"""
        return self.prs.slides.add_slide(self.prs.slide_layouts[6])
    
    def _add_title(
        self, 
        slide, 
        text: str, 
        font_size: int = Pt(20),
        top: float = None,
        bold: bool = True,
        color: RGBColor = None
    ) -> float:
        """Add title to slide"""
        top = top if top is not None else self.config.margin_top - Inches(0.2)
        color = color or self.colors['primary_red']
        
        title_box = slide.shapes.add_textbox(
            self.config.margin_left, top,
            self.config.content_width, Inches(0.5)
        )
        title_frame = title_box.text_frame
        title_frame.text = text
        title_para = title_frame.paragraphs[0]
        title_para.font.size = font_size
        title_para.font.bold = bold
        title_para.font.color.rgb = color
        
        return top + Inches(0.5)
    
    def _add_subtitle(
        self, 
        slide, 
        text: str, 
        font_size: int = Pt(12),
        top: float = None,
        color: RGBColor = None
    ) -> float:
        """Add subtitle to slide"""
        top = top if top is not None else Inches(0.1)
        color = color or self.colors['gray_600']
        
        subtitle_box = slide.shapes.add_textbox(
            self.config.margin_left, top,
            self.config.content_width, Inches(0.25)
        )
        subtitle_frame = subtitle_box.text_frame
        subtitle_frame.text = text
        subtitle_para = subtitle_frame.paragraphs[0]
        subtitle_para.font.size = font_size
        subtitle_para.font.color.rgb = color
        
        return top + Inches(0.3)


class TitleSlideBuilder(SlideFactory):
    """Title slide builder"""
    
    def create(self, title: str, subtitle: str = "") -> Any:
        """Create title slide"""
        slide = self._get_blank_slide()
        
        # Background
        background = slide.shapes.add_shape(
            MSO_SHAPE.RECTANGLE,
            0, 0,
            self.config.width, self.config.height
        )
        background.fill.solid()
        background.fill.fore_color.rgb = self.colors['gray_50']
        background.line.fill.background()
        
        # Top red box
        header_box = slide.shapes.add_shape(
            MSO_SHAPE.RECTANGLE,
            0, Inches(2),
            self.config.width, Inches(1.5)
        )
        header_box.fill.solid()
        header_box.fill.fore_color.rgb = self.colors['primary_red']
        header_box.line.fill.background()
        
        # Title text
        title_frame = header_box.text_frame
        title_frame.text = title
        title_frame.paragraphs[0].alignment = PP_ALIGN.CENTER
        title_frame.paragraphs[0].font.size = Pt(40)
        title_frame.paragraphs[0].font.bold = True
        title_frame.paragraphs[0].font.color.rgb = self.colors['white']
        
        # Subtitle
        if subtitle:
            subtitle_box = slide.shapes.add_textbox(
                Inches(1), Inches(4),
                self.config.width - Inches(2), Inches(1.5)
            )
            subtitle_frame = subtitle_box.text_frame
            subtitle_frame.text = subtitle
            subtitle_frame.paragraphs[0].alignment = PP_ALIGN.CENTER
            subtitle_frame.paragraphs[0].font.size = Pt(16)
            subtitle_frame.paragraphs[0].font.color.rgb = self.colors['gray_800']
            subtitle_frame.word_wrap = True
        
        return slide


class ContentSlideBuilder(SlideFactory):
    """General content slide builder"""
    
    def create_with_text(
        self, 
        title: str, 
        content: str,
        subtitle: str = None
    ) -> Any:
        """Create text content slide"""
        slide = self._get_blank_slide()
        
        y_position = self.config.margin_top - Inches(0.2)
        
        # Title
        y_position = self._add_title(slide, title, Pt(32), y_position)
        
        # Subtitle
        if subtitle:
            subtitle_box = slide.shapes.add_textbox(
                self.config.margin_left, y_position,
                self.config.content_width, Inches(0.3)
            )
            subtitle_frame = subtitle_box.text_frame
            subtitle_frame.text = subtitle
            subtitle_para = subtitle_frame.paragraphs[0]
            subtitle_para.font.size = Pt(20)
            subtitle_para.font.bold = True
            subtitle_para.font.color.rgb = self.colors['gray_800']
            y_position += Inches(0.4)
        
        # Body
        text_box = slide.shapes.add_textbox(
            self.config.margin_left, y_position + Inches(0.2),
            self.config.content_width, Inches(4)
        )
        text_frame = text_box.text_frame
        text_frame.text = content
        text_frame.word_wrap = True
        
        for paragraph in text_frame.paragraphs:
            paragraph.font.size = Pt(12)
            paragraph.font.color.rgb = self.colors['gray_800']
            paragraph.line_spacing = 1.5
        
        return slide


class TableSlideBuilder(SlideFactory):
    """Table slide builder"""
    
    def __init__(
        self,
        presentation: Presentation,
        slide_config: SlideConfig = None,
        colors: ColorPalette = None,
        max_rows_per_slide: int = 8
    ):
        super().__init__(presentation, slide_config, colors)
        self.max_rows_per_slide = max_rows_per_slide
    
    def create_from_html(
        self, 
        table_elem: Tag, 
        title: str,
        main_title: str = ""
    ) -> List[Any]:
        """Create slides from HTML table (auto-split)"""
        # Extract table data first
        extractor = TableDataExtractor(table_elem).extract()
        
        # If splitting is needed - check before slide creation
        header_count = len(extractor.header_rows)
        body_count = len(extractor.body_rows)
        
        y_position = self.config.margin_top - Inches(0.2)
        if main_title:
            y_position = Inches(0.35)
        table_top = y_position + Inches(0.5)
        table_height = self.config.height - table_top - self.config.margin_bottom
        
        if body_count > self.max_rows_per_slide:
            # Create separate slides when splitting
            return self._create_split_table_slides(
                extractor, title, main_title, table_top, table_height
            )
        
        # Create single slide
        slide = self._get_blank_slide()
        
        # Main title
        if main_title:
            self._add_subtitle(slide, main_title, Pt(12), Inches(0.1))
        
        # Section title
        self._add_title(slide, title, Pt(18), y_position)
        
        # Display in card style if key-value table
        if extractor.is_key_value_table():
            self._add_key_value_cards(
                slide, table_elem,
                self.config.margin_left, table_top,
                self.config.content_width, table_height
            )
            return [slide]
        
        # Single slide table
        self.table_builder.create_table(
            slide,
            extractor.rows_data,
            header_count,
            extractor.col_widths_html,
            self.config.margin_left, table_top,
            self.config.content_width, table_height,
            extractor.merge_info,
            extractor.cell_styles
        )
        
        return [slide]
    
    def _create_split_table_slides(
        self,
        extractor: TableDataExtractor,
        title: str,
        main_title: str,
        table_top: float,
        table_height: float
    ) -> List[Any]:
        """Create split table slides"""
        slides = []
        header_count = len(extractor.header_rows)
        body_rows = extractor.body_rows
        
        num_chunks = (len(body_rows) + self.max_rows_per_slide - 1) // self.max_rows_per_slide
        
        for chunk_idx in range(num_chunks):
            start_idx = chunk_idx * self.max_rows_per_slide
            end_idx = min(start_idx + self.max_rows_per_slide, len(body_rows))
            
            chunk_data = extractor.header_rows + body_rows[start_idx:end_idx]
            
            # Filter merge_info for this chunk
            chunk_merge_info = []
            for row_idx, col_idx, colspan, rowspan in extractor.merge_info:
                if row_idx < header_count:
                    chunk_merge_info.append((row_idx, col_idx, colspan, rowspan))
                elif row_idx - header_count >= start_idx and row_idx - header_count < end_idx:
                    new_row_idx = header_count + (row_idx - header_count - start_idx)
                    chunk_merge_info.append((new_row_idx, col_idx, colspan, rowspan))
            
            # Filter cell_styles for this chunk
            chunk_cell_styles = {}
            for (r, c), styles in extractor.cell_styles.items():
                if r < header_count:
                    chunk_cell_styles[(r, c)] = styles
                elif r - header_count >= start_idx and r - header_count < end_idx:
                    new_r = header_count + (r - header_count - start_idx)
                    chunk_cell_styles[(new_r, c)] = styles
            
            # Create slide
            slide = self._get_blank_slide()
            
            # Title (add "(continued N)" after first slide)
            slide_title = title if chunk_idx == 0 else f"{title} (continued {chunk_idx + 1})"
            self._add_title(slide, slide_title, Pt(18), Inches(0.1))
            
            # Create table
            self.table_builder.create_table(
                slide,
                chunk_data,
                header_count,
                extractor.col_widths_html,
                self.config.margin_left, Inches(0.6),
                self.config.content_width, self.config.height - Inches(0.9),
                chunk_merge_info,
                chunk_cell_styles
            )
            
            slides.append(slide)
        
        return slides
    
    def _add_key_value_cards(
        self,
        slide,
        table_elem: Tag,
        left: float,
        top: float,
        width: float,
        height: float
    ) -> List[Any]:
        """Display key-value table as card style"""
        tbody = table_elem.find('tbody')
        if not tbody:
            return []
        
        rows = tbody.find_all('tr')
        if not rows:
            return []
        
        label_bg_color = RGBColor(55, 65, 81)
        value_bg_color = RGBColor(249, 250, 251)
        border_color = RGBColor(209, 213, 219)
        
        y_position = top
        card_height = Inches(1.3)
        card_spacing = Inches(0.12)
        label_width = Inches(1.6)
        
        for tr in rows:
            cells = tr.find_all(['th', 'td'])
            if len(cells) < 2:
                continue
            
            label = TextUtils.clean_text(cells[0].get_text(strip=True))
            value = TextUtils.clean_text(cells[1].get_text(strip=True))
            
            # Label area
            label_shape = slide.shapes.add_shape(
                MSO_SHAPE.RECTANGLE,
                left, y_position,
                label_width, card_height
            )
            label_shape.fill.solid()
            label_shape.fill.fore_color.rgb = label_bg_color
            label_shape.line.fill.background()
            
            label_tf = label_shape.text_frame
            label_tf.word_wrap = True
            label_tf.paragraphs[0].text = label
            label_tf.paragraphs[0].font.size = Pt(13)
            label_tf.paragraphs[0].font.bold = True
            label_tf.paragraphs[0].font.color.rgb = RGBColor(255, 255, 255)
            label_tf.paragraphs[0].alignment = PP_ALIGN.CENTER
            label_shape.text_frame.margin_left = Pt(10)
            label_shape.text_frame.margin_right = Pt(10)
            label_shape.text_frame.margin_top = Pt(10)
            label_shape.text_frame.margin_bottom = Pt(10)
            
            # Value area
            value_left = left + label_width
            value_width = width - label_width
            
            value_shape = slide.shapes.add_shape(
                MSO_SHAPE.RECTANGLE,
                value_left, y_position,
                value_width, card_height
            )
            value_shape.fill.solid()
            value_shape.fill.fore_color.rgb = value_bg_color
            value_shape.line.color.rgb = border_color
            value_shape.line.width = Pt(1)
            
            value_tf = value_shape.text_frame
            value_tf.word_wrap = True
            value_tf.paragraphs[0].text = value
            value_tf.paragraphs[0].font.size = Pt(11)
            value_tf.paragraphs[0].font.color.rgb = self.colors['gray_800']
            value_tf.paragraphs[0].alignment = PP_ALIGN.LEFT
            value_shape.text_frame.margin_left = Pt(12)
            value_shape.text_frame.margin_right = Pt(12)
            value_shape.text_frame.margin_top = Pt(10)
            value_shape.text_frame.margin_bottom = Pt(10)
            
            y_position += card_height + card_spacing
        
        return [slide]


class ImageSlideBuilder(SlideFactory):
    """Image slide builder"""
    
    def create_from_base64(self, img_tag: Tag, title: str) -> Optional[Any]:
        """Create slide from Base64 image"""
        if not HAS_PIL:
            logger.warning("Cannot create image slide because PIL is not installed")
            return None
        
        src = img_tag.get('src', '')
        if not src.startswith('data:image'):
            return None
        
        try:
            header, data = src.split(',', 1)
            img_bytes = base64.b64decode(data)
            pil_img = Image.open(BytesIO(img_bytes))
            img_width, img_height = pil_img.size
            
            slide = self._get_blank_slide()
            
            # Title
            self._add_title(slide, title or "Analysis Chart", Pt(20), Inches(0.2))
            
            # Calculate image size
            available_width = self.config.content_width
            available_height = self.config.height - Inches(1.2)
            
            img_ratio = img_width / img_height
            available_ratio = available_width / available_height
            
            if img_ratio > available_ratio:
                final_width = available_width
                final_height = available_width / img_ratio
            else:
                final_height = available_height
                final_width = available_height * img_ratio
            
            # Center alignment
            img_left = self.config.margin_left + (available_width - final_width) / 2
            img_top = Inches(0.7) + (available_height - final_height) / 2
            
            # Save and add image
            img_stream = BytesIO()
            pil_img.save(img_stream, format='PNG')
            img_stream.seek(0)
            
            slide.shapes.add_picture(
                img_stream,
                img_left, img_top,
                final_width, final_height
            )
            
            logger.info(f"Image slide created: {title} ({img_width}x{img_height})")
            return slide
            
        except Exception as e:
            logger.error(f"Failed to create image slide: {e}")
            return None
    
    def create_from_file(self, image_path: str, title: str) -> Optional[Any]:
        """Create image slide from file"""
        if not HAS_PIL:
            return None
        
        try:
            pil_img = Image.open(image_path)
            img_width, img_height = pil_img.size
            
            slide = self._get_blank_slide()
            
            self._add_title(slide, title, Pt(20), Inches(0.2))
            
            available_width = self.config.content_width
            available_height = self.config.height - Inches(1.2)
            
            img_ratio = img_width / img_height
            available_ratio = available_width / available_height
            
            if img_ratio > available_ratio:
                final_width = available_width
                final_height = available_width / img_ratio
            else:
                final_height = available_height
                final_width = available_height * img_ratio
            
            img_left = self.config.margin_left + (available_width - final_width) / 2
            img_top = Inches(0.7) + (available_height - final_height) / 2
            
            slide.shapes.add_picture(
                image_path,
                img_left, img_top,
                final_width, final_height
            )
            
            return slide
            
        except Exception as e:
            logger.error(f"Failed to create image slide: {e}")
            return None


class SectionSlideBuilder(SlideFactory):
    """Section divider slide builder (intermediate title)"""
    
    def create(self, title: str, subtitle: str = "") -> Any:
        """Create section divider slide"""
        slide = self._get_blank_slide()
        
        # Background
        background = slide.shapes.add_shape(
            MSO_SHAPE.RECTANGLE,
            0, 0,
            self.config.width, self.config.height
        )
        background.fill.solid()
        background.fill.fore_color.rgb = self.colors['gray_50']
        background.line.fill.background()
        
        # Center section title box
        title_height = Inches(1.2)
        title_box = slide.shapes.add_shape(
            MSO_SHAPE.RECTANGLE,
            0, (self.config.height - title_height) / 2,
            self.config.width, title_height
        )
        title_box.fill.solid()
        title_box.fill.fore_color.rgb = self.colors['primary_red']
        title_box.line.fill.background()
        
        # Section title text
        title_tf = title_box.text_frame
        title_tf.text = title
        title_tf.paragraphs[0].alignment = PP_ALIGN.CENTER
        title_tf.paragraphs[0].font.size = Pt(32)
        title_tf.paragraphs[0].font.bold = True
        title_tf.paragraphs[0].font.color.rgb = self.colors['white']
        
        # Subtitle
        if subtitle:
            subtitle_box = slide.shapes.add_textbox(
                Inches(1), (self.config.height + title_height) / 2 + Inches(0.2),
                self.config.width - Inches(2), Inches(0.8)
            )
            subtitle_frame = subtitle_box.text_frame
            subtitle_frame.text = subtitle
            subtitle_frame.paragraphs[0].alignment = PP_ALIGN.CENTER
            subtitle_frame.paragraphs[0].font.size = Pt(16)
            subtitle_frame.paragraphs[0].font.color.rgb = self.colors['gray_600']
            subtitle_frame.word_wrap = True
        
        return slide


class ReferenceCardSlideBuilder(SlideFactory):
    """Reference card slide builder"""
    
    def create(self, reference_card: Tag, section_title: str = "") -> Optional[Any]:
        """Create slide from reference card"""
        slide = self._get_blank_slide()
        
        y_position = self.config.margin_top - Inches(0.2)
        
        # Section title (small text)
        if section_title:
            self._add_subtitle(slide, section_title, Pt(11), Inches(0.1))
            y_position = Inches(0.35)
        
        # Extract reference title (from reference-number)
        ref_number = reference_card.find('div', class_='reference-number')
        if ref_number:
            ref_title = ref_number.get_text(strip=True)
        else:
            ref_title = "Reference"
        
        # Title (Reference title)
        title_box = slide.shapes.add_textbox(
            self.config.margin_left, y_position,
            self.config.content_width, Inches(0.6)
        )
        title_frame = title_box.text_frame
        title_frame.word_wrap = True
        title_frame.text = ref_title
        title_para = title_frame.paragraphs[0]
        title_para.font.size = Pt(16)
        title_para.font.bold = True
        title_para.font.color.rgb = self.colors['primary_red']
        
        y_position += Inches(0.7)
        
        # Meta information (journal, date, etc.)
        ref_meta = reference_card.find('div', class_='reference-meta')
        if ref_meta:
            meta_text = ref_meta.get_text(strip=True)
            meta_box = slide.shapes.add_textbox(
                self.config.margin_left, y_position,
                self.config.content_width, Inches(0.3)
            )
            meta_frame = meta_box.text_frame
            meta_frame.text = meta_text
            meta_para = meta_frame.paragraphs[0]
            meta_para.font.size = Pt(10)
            meta_para.font.color.rgb = self.colors['gray_600']
            y_position += Inches(0.4)
        
        # Summary content
        ref_summary = reference_card.find('div', class_='reference-summary')
        if ref_summary:
            # Clean summary text
            summary_text = ref_summary.get_text(separator='\n', strip=True)
            # Length limit
            if len(summary_text) > 1500:
                summary_text = summary_text[:1500] + "..."
            
            summary_box = slide.shapes.add_textbox(
                self.config.margin_left, y_position,
                self.config.content_width,
                self.config.height - y_position - self.config.margin_bottom
            )
            summary_frame = summary_box.text_frame
            summary_frame.word_wrap = True
            summary_frame.text = summary_text
            
            for paragraph in summary_frame.paragraphs:
                paragraph.font.size = Pt(9)
                paragraph.font.color.rgb = self.colors['gray_800']
                paragraph.line_spacing = 1.3
        
        return slide


class EvidenceSlideBuilder(SlideFactory):
    """Evidence table slide builder"""
    
    def create(self, evidence_div: Tag, title: str) -> Optional[Any]:
        """Create evidence table slide"""
        slide = self._get_blank_slide()
        
        self._add_title(slide, title, Pt(18))
        
        # Find evidence-cell inside evidence-row or directly
        evidence_row = evidence_div.find('div', class_='evidence-row')
        if evidence_row:
            # Use evidence-cells inside evidence-row as data rows
            data_rows = evidence_row.find_all('div', class_='evidence-cell')
        else:
            # If evidence-row not found, find evidence-cell directly
            data_rows = evidence_div.find_all('div', class_='evidence-cell')
        
        if not data_rows:
            return slide
        
        max_rows = min(len(data_rows), 10)
        
        # Extract table data
        table_data = []
        link_data = []
        
        # Extract header
        header_div = evidence_div.find('div', class_='evidence-header')
        if header_div:
            headers = [elem.strip() for elem in header_div.stripped_strings]
            if headers:
                table_data.append(headers[:8])
        
        # Extract data rows
        for row_idx, row in enumerate(data_rows[:max_rows]):
            row_texts = []
            text_elements = row.find_all('div', class_='evidence-text')
            
            for col_idx, elem in enumerate(text_elements[:8]):
                link = elem.find('a')
                if link:
                    link_text = link.get_text(strip=True)
                    link_url = link.get('href', '')
                    row_texts.append(link_text)
                    if link_url:
                        link_data.append((len(table_data), col_idx, link_url))
                else:
                    text = TextUtils.clean_text(elem.get_text(strip=True))
                    text = TextUtils.truncate_text(text, 80)
                    row_texts.append(text)
            
            if row_texts:
                table_data.append(row_texts)
        
        if len(table_data) <= 1:
            return slide
        
        # Normalize column count
        max_cols = max(len(row) for row in table_data)
        for row in table_data:
            while len(row) < max_cols:
                row.append('')
        
        # Create table - adjust height based on row count
        table_top = self.config.margin_top + Inches(0.4)
        max_table_height = self.config.height - table_top - self.config.margin_bottom
        
        # Calculate row height (header: 0.3 inch, data: 0.4 inch)
        num_rows = len(table_data)
        row_height = Inches(0.4)
        calculated_height = row_height * num_rows
        
        # Use smaller of calculated height and max height
        table_height = min(calculated_height, max_table_height)
        
        try:
            ppt_table = slide.shapes.add_table(
                len(table_data), max_cols,
                self.config.margin_left, table_top,
                self.config.content_width, table_height
            ).table
            
            # Fill data
            for i, row_data in enumerate(table_data):
                for j, cell_data in enumerate(row_data):
                    if j >= max_cols:
                        continue
                    
                    cell = ppt_table.cell(i, j)
                    cell.text = str(cell_data)
                    cell.vertical_anchor = MSO_ANCHOR.MIDDLE
                    
                    cell.margin_left = Pt(3)
                    cell.margin_right = Pt(3)
                    cell.margin_top = Pt(2)
                    cell.margin_bottom = Pt(2)
                    
                    cell.fill.background()
                    
                    for paragraph in cell.text_frame.paragraphs:
                        paragraph.font.size = Pt(7)
                        
                        if i == 0:
                            paragraph.font.bold = True
                            paragraph.font.color.rgb = self.colors['black']
                            paragraph.alignment = PP_ALIGN.CENTER
                            cell.text_frame.word_wrap = False
                        else:
                            paragraph.font.color.rgb = self.colors['gray_800']
                            cell.text_frame.word_wrap = True
                            
                            if j in [2, 5]:
                                paragraph.alignment = PP_ALIGN.LEFT
                            else:
                                paragraph.alignment = PP_ALIGN.CENTER
                            
                            if cell_data == 'Link':
                                paragraph.font.color.rgb = self.colors['link_blue']
                                paragraph.font.underline = True
            
            # Apply link styles
            for row_idx, col_idx, url in link_data:
                try:
                    cell = ppt_table.cell(row_idx, col_idx)
                    for paragraph in cell.text_frame.paragraphs:
                        for run in paragraph.runs:
                            run.font.color.rgb = self.colors['link_blue']
                            run.font.underline = True
                except Exception:
                    pass
            
            # Apply borders
            from .table_builder import TableBorderStyler
            border_styler = TableBorderStyler(colors=self.colors)
            border_styler.apply_academic_borders(ppt_table, 1, len(table_data), max_cols)
            
        except Exception as e:
            logger.error(f"Failed to create evidence table: {e}")
        
        return slide
