"""
Style extraction and application utilities

Provides functionality to extract styles from HTML elements and apply them to PowerPoint elements.
"""
import re
from typing import Dict, Any, Optional, List
from bs4 import Tag
from pptx.dml.color import RGBColor


class StyleExtractor:
    """Class that extracts style information from HTML elements"""
    
    @staticmethod
    def extract_cell_styles(cell_elem: Tag) -> Dict[str, Any]:
        """
        Extract styles (Bold, Color, etc.) from HTML cell
        
        Args:
            cell_elem: BeautifulSoup Tag object
            
        Returns:
            Style information dictionary {'bold': bool, 'color': RGBColor, 'background': RGBColor, 'link': str}
        """
        styles = {
            'bold': False,
            'color': None,
            'background': None,
            'link': None,
        }
        
        style_attr = cell_elem.get('style', '')
        
        # Extract color
        color_match = re.search(
            r'color:\s*(#[a-fA-F0-9]{6}|#[a-fA-F0-9]{3}|rgb\([^)]+\))', 
            style_attr
        )
        if color_match:
            styles['color'] = StyleExtractor.parse_color(color_match.group(1))
        
        # Extract background-color
        bg_match = re.search(
            r'background(?:-color)?:\s*(#[a-fA-F0-9]{6}|#[a-fA-F0-9]{3}|rgb\([^)]+\))', 
            style_attr
        )
        if bg_match:
            styles['background'] = StyleExtractor.parse_color(bg_match.group(1))
        
        # Check font-weight
        if 'font-weight' in style_attr:
            weight_match = re.search(r'font-weight:\s*(\w+)', style_attr)
            if weight_match:
                weight = weight_match.group(1)
                if weight in ('bold', '700', '800', '900'):
                    styles['bold'] = True
        
        # Check for bold tags (b, strong)
        if cell_elem.find(['b', 'strong']):
            styles['bold'] = True
        
        # Check for color styles in inner elements (span, etc.)
        colored_elem = cell_elem.find(style=True)
        if colored_elem and not styles['color']:
            inner_style = colored_elem.get('style', '')
            inner_color = re.search(
                r'color:\s*(#[a-fA-F0-9]{6}|#[a-fA-F0-9]{3}|rgb\([^)]+\))', 
                inner_style
            )
            if inner_color:
                styles['color'] = StyleExtractor.parse_color(inner_color.group(1))
        
        # Check for link
        link = cell_elem.find('a')
        if link:
            styles['link'] = link.get('href', '')
        
        return styles
    
    @staticmethod
    def parse_color(color_str: str) -> Optional[RGBColor]:
        """
        Convert color string to RGBColor
        
        Args:
            color_str: Color string in '#rrggbb', '#rgb', 'rgb(r, g, b)' format
            
        Returns:
            RGBColor object or None
        """
        if not color_str:
            return None
        
        try:
            # hex color (#rrggbb or #rgb)
            if color_str.startswith('#'):
                hex_color = color_str[1:]
                if len(hex_color) == 3:
                    hex_color = ''.join([c * 2 for c in hex_color])
                r = int(hex_color[0:2], 16)
                g = int(hex_color[2:4], 16)
                b = int(hex_color[4:6], 16)
                return RGBColor(r, g, b)
            
            # rgb(r, g, b)
            if color_str.startswith('rgb'):
                match = re.search(
                    r'rgb\s*\(\s*(\d+)\s*,\s*(\d+)\s*,\s*(\d+)\s*\)', 
                    color_str
                )
                if match:
                    r = int(match.group(1))
                    g = int(match.group(2))
                    b = int(match.group(3))
                    return RGBColor(r, g, b)
        except Exception:
            pass
        
        return None
    
    @staticmethod
    def extract_column_widths(cells: List[Tag]) -> List[Optional[int]]:
        """
        Extract width attribute from HTML table cells
        
        Args:
            cells: List of BeautifulSoup Tag objects
            
        Returns:
            List of cell widths (in pixels, None if not specified)
        """
        widths = []
        for cell in cells:
            width = None
            
            # Extract width from style attribute
            style = cell.get('style', '')
            if 'width:' in style:
                match = re.search(r'width:\s*(\d+)(?:px|%)?', style)
                if match:
                    width = int(match.group(1))
            
            # Check width attribute directly
            elif cell.get('width'):
                try:
                    width = int(cell.get('width').replace('px', '').replace('%', ''))
                except ValueError:
                    pass
            
            widths.append(width)
        
        return widths


class TextUtils:
    """Text processing utilities"""
    
    @staticmethod
    def clean_text(text: str) -> str:
        """
        Clean text (remove unnecessary whitespace)
        
        Args:
            text: Original text
            
        Returns:
            Cleaned text
        """
        if not text:
            return ""
        # Multiple whitespace to single
        text = re.sub(r'\s+', ' ', text)
        return text.strip()
    
    @staticmethod
    def extract_cell_text_with_formatting(cell_elem) -> str:
        """
        Extract text from HTML cell while preserving bullets and line breaks
        
        Args:
            cell_elem: BeautifulSoup Tag object (td or th)
            
        Returns:
            Text with formatting preserved
        """
        from bs4 import NavigableString
        
        if not cell_elem:
            return ""
        
        result_parts = []
        
        def process_element(elem, bullet_prefix=""):
            """Process element recursively"""
            if isinstance(elem, NavigableString):
                text = str(elem).strip()
                if text:
                    result_parts.append(text)
                return
            
            tag_name = getattr(elem, 'name', None)
            
            if tag_name == 'br':
                result_parts.append('\n')
            elif tag_name == 'ul':
                # Process li inside ul
                for li in elem.find_all('li', recursive=False):
                    result_parts.append('\n• ')
                    for child in li.children:
                        process_element(child)
            elif tag_name == 'ol':
                # Process li inside ol (numbered)
                for idx, li in enumerate(elem.find_all('li', recursive=False), 1):
                    result_parts.append(f'\n{idx}. ')
                    for child in li.children:
                        process_element(child)
            elif tag_name == 'li':
                # Standalone li outside ul/ol
                result_parts.append('\n• ')
                for child in elem.children:
                    process_element(child)
            elif tag_name in ('p', 'div'):
                # Add line break for block elements
                for child in elem.children:
                    process_element(child)
                if result_parts and not result_parts[-1].endswith('\n'):
                    result_parts.append('\n')
            else:
                # Process children for other elements
                if hasattr(elem, 'children'):
                    for child in elem.children:
                        process_element(child)
        
        # Process all children of the cell
        for child in cell_elem.children:
            process_element(child)
        
        # Clean up result
        text = ''.join(result_parts)
        # Clean up consecutive line breaks
        text = re.sub(r'\n\s*\n', '\n', text)
        # Remove leading/trailing whitespace and line breaks
        text = text.strip()
        # Multiple spaces to single (excluding line breaks)
        text = re.sub(r'[ \t]+', ' ', text)
        
        return text
    
    @staticmethod
    def truncate_text(text: str, max_length: int = 100, suffix: str = "...") -> str:
        """
        Truncate text to maximum length
        
        Args:
            text: Original text
            max_length: Maximum length
            suffix: Suffix to append when truncated
            
        Returns:
            Truncated text
        """
        if len(text) <= max_length:
            return text
        return text[:max_length - len(suffix)] + suffix
