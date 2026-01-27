"""
스타일 추출 및 적용 유틸리티

HTML 요소에서 스타일을 추출하고 PowerPoint 요소에 적용하는 기능을 제공합니다.
"""
import re
from typing import Dict, Any, Optional, List
from bs4 import Tag
from pptx.dml.color import RGBColor


class StyleExtractor:
    """HTML 요소에서 스타일 정보를 추출하는 클래스"""
    
    @staticmethod
    def extract_cell_styles(cell_elem: Tag) -> Dict[str, Any]:
        """
        HTML 셀에서 스타일(Bold, Color 등) 추출
        
        Args:
            cell_elem: BeautifulSoup Tag 객체
            
        Returns:
            스타일 정보 딕셔너리 {'bold': bool, 'color': RGBColor, 'background': RGBColor, 'link': str}
        """
        styles = {
            'bold': False,
            'color': None,
            'background': None,
            'link': None,
        }
        
        style_attr = cell_elem.get('style', '')
        
        # color 추출
        color_match = re.search(
            r'color:\s*(#[a-fA-F0-9]{6}|#[a-fA-F0-9]{3}|rgb\([^)]+\))', 
            style_attr
        )
        if color_match:
            styles['color'] = StyleExtractor.parse_color(color_match.group(1))
        
        # background-color 추출
        bg_match = re.search(
            r'background(?:-color)?:\s*(#[a-fA-F0-9]{6}|#[a-fA-F0-9]{3}|rgb\([^)]+\))', 
            style_attr
        )
        if bg_match:
            styles['background'] = StyleExtractor.parse_color(bg_match.group(1))
        
        # font-weight 확인
        if 'font-weight' in style_attr:
            weight_match = re.search(r'font-weight:\s*(\w+)', style_attr)
            if weight_match:
                weight = weight_match.group(1)
                if weight in ('bold', '700', '800', '900'):
                    styles['bold'] = True
        
        # 내부의 bold 태그 확인 (b, strong)
        if cell_elem.find(['b', 'strong']):
            styles['bold'] = True
        
        # 내부의 색상 스타일 확인 (span 등)
        colored_elem = cell_elem.find(style=True)
        if colored_elem and not styles['color']:
            inner_style = colored_elem.get('style', '')
            inner_color = re.search(
                r'color:\s*(#[a-fA-F0-9]{6}|#[a-fA-F0-9]{3}|rgb\([^)]+\))', 
                inner_style
            )
            if inner_color:
                styles['color'] = StyleExtractor.parse_color(inner_color.group(1))
        
        # 링크 확인
        link = cell_elem.find('a')
        if link:
            styles['link'] = link.get('href', '')
        
        return styles
    
    @staticmethod
    def parse_color(color_str: str) -> Optional[RGBColor]:
        """
        색상 문자열을 RGBColor로 변환
        
        Args:
            color_str: '#rrggbb', '#rgb', 'rgb(r, g, b)' 형식의 색상 문자열
            
        Returns:
            RGBColor 객체 또는 None
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
        HTML 테이블 셀에서 width 속성 추출
        
        Args:
            cells: BeautifulSoup Tag 리스트
            
        Returns:
            각 셀의 너비 리스트 (픽셀 단위, None이면 지정되지 않음)
        """
        widths = []
        for cell in cells:
            width = None
            
            # style 속성에서 width 추출
            style = cell.get('style', '')
            if 'width:' in style:
                match = re.search(r'width:\s*(\d+)(?:px|%)?', style)
                if match:
                    width = int(match.group(1))
            
            # width 속성 직접 확인
            elif cell.get('width'):
                try:
                    width = int(cell.get('width').replace('px', '').replace('%', ''))
                except ValueError:
                    pass
            
            widths.append(width)
        
        return widths


class TextUtils:
    """텍스트 처리 유틸리티"""
    
    @staticmethod
    def clean_text(text: str) -> str:
        """
        텍스트 정리 (불필요한 공백 제거)
        
        Args:
            text: 원본 텍스트
            
        Returns:
            정리된 텍스트
        """
        if not text:
            return ""
        # 여러 공백을 하나로
        text = re.sub(r'\s+', ' ', text)
        return text.strip()
    
    @staticmethod
    def extract_cell_text_with_formatting(cell_elem) -> str:
        """
        HTML 셀에서 bullet, linebreak를 유지하며 텍스트 추출
        
        Args:
            cell_elem: BeautifulSoup Tag 객체 (td 또는 th)
            
        Returns:
            포맷팅이 유지된 텍스트
        """
        from bs4 import NavigableString
        
        if not cell_elem:
            return ""
        
        result_parts = []
        
        def process_element(elem, bullet_prefix=""):
            """재귀적으로 요소 처리"""
            if isinstance(elem, NavigableString):
                text = str(elem).strip()
                if text:
                    result_parts.append(text)
                return
            
            tag_name = getattr(elem, 'name', None)
            
            if tag_name == 'br':
                result_parts.append('\n')
            elif tag_name == 'ul':
                # ul 내부의 li 처리
                for li in elem.find_all('li', recursive=False):
                    result_parts.append('\n• ')
                    for child in li.children:
                        process_element(child)
            elif tag_name == 'ol':
                # ol 내부의 li 처리 (숫자)
                for idx, li in enumerate(elem.find_all('li', recursive=False), 1):
                    result_parts.append(f'\n{idx}. ')
                    for child in li.children:
                        process_element(child)
            elif tag_name == 'li':
                # ul/ol 외부의 단독 li
                result_parts.append('\n• ')
                for child in elem.children:
                    process_element(child)
            elif tag_name in ('p', 'div'):
                # 블록 요소는 줄바꿈 추가
                for child in elem.children:
                    process_element(child)
                if result_parts and not result_parts[-1].endswith('\n'):
                    result_parts.append('\n')
            else:
                # 다른 요소는 자식들 처리
                if hasattr(elem, 'children'):
                    for child in elem.children:
                        process_element(child)
        
        # 셀의 모든 자식 요소 처리
        for child in cell_elem.children:
            process_element(child)
        
        # 결과 정리
        text = ''.join(result_parts)
        # 연속된 줄바꿈 정리
        text = re.sub(r'\n\s*\n', '\n', text)
        # 앞뒤 공백/줄바꿈 제거
        text = text.strip()
        # 여러 공백을 하나로 (줄바꿈 제외)
        text = re.sub(r'[ \t]+', ' ', text)
        
        return text
    
    @staticmethod
    def truncate_text(text: str, max_length: int = 100, suffix: str = "...") -> str:
        """
        텍스트를 최대 길이로 자르기
        
        Args:
            text: 원본 텍스트
            max_length: 최대 길이
            suffix: 잘린 경우 붙일 접미사
            
        Returns:
            잘린 텍스트
        """
        if len(text) <= max_length:
            return text
        return text[:max_length - len(suffix)] + suffix
