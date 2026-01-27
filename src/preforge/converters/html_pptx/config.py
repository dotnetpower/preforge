"""
HTML to PPTX 변환 설정 및 상수

슬라이드 크기, 색상, 여백 등 공통 설정을 관리합니다.
"""
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from dataclasses import dataclass
from typing import Dict


@dataclass
class SlideConfig:
    """슬라이드 레이아웃 설정"""
    width: int = Inches(10)
    height: int = Inches(7.5)
    margin_left: int = Inches(0.3)
    margin_right: int = Inches(0.3)
    margin_top: int = Inches(0.5)
    margin_bottom: int = Inches(0.3)
    
    @property
    def content_width(self) -> int:
        return self.width - self.margin_left - self.margin_right
    
    @property
    def content_height(self) -> int:
        return self.height - self.margin_top - self.margin_bottom


@dataclass
class TableConfig:
    """테이블 설정"""
    max_rows_per_slide: int = 8
    min_row_height: int = Inches(0.22)
    row_height_estimate: int = Inches(0.28)
    header_font_size: int = Pt(9)
    body_font_size: int = Pt(8)
    small_font_size: int = Pt(7)
    cell_margin: int = Pt(4)
    cell_margin_vertical: int = Pt(2)


@dataclass
class BorderConfig:
    """테두리 설정"""
    thick_line: int = Pt(1.5)
    thin_line: int = Pt(0.5)
    no_line: int = Pt(0)


class ColorPalette:
    """색상 팔레트"""
    
    def __init__(self):
        self._colors: Dict[str, RGBColor] = {
            'primary_red': RGBColor(220, 38, 38),      # #dc2626
            'primary_red_light': RGBColor(254, 242, 242),  # #fef2f2
            'primary_red_dark': RGBColor(153, 27, 27),     # #991b1b
            'gray_50': RGBColor(249, 250, 251),
            'gray_100': RGBColor(243, 244, 246),
            'gray_200': RGBColor(229, 231, 235),
            'gray_600': RGBColor(75, 85, 99),
            'gray_800': RGBColor(31, 41, 55),
            'gray_900': RGBColor(17, 24, 39),
            'white': RGBColor(255, 255, 255),
            'black': RGBColor(0, 0, 0),
            'link_blue': RGBColor(0, 102, 204),
            'gray_line': RGBColor(200, 200, 200),
        }
    
    def __getitem__(self, key: str) -> RGBColor:
        return self._colors.get(key, self._colors['black'])
    
    def get(self, key: str, default: RGBColor = None) -> RGBColor:
        return self._colors.get(key, default or self._colors['black'])


# 기본 설정 인스턴스
DEFAULT_SLIDE_CONFIG = SlideConfig()
DEFAULT_TABLE_CONFIG = TableConfig()
DEFAULT_BORDER_CONFIG = BorderConfig()
DEFAULT_COLORS = ColorPalette()
