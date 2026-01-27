"""
HTML(.html) 문서를 PowerPoint(.pptx)로 변환하는 컨버터

이 모듈은 하위 호환성을 위해 유지됩니다.
실제 구현은 html_pptx 패키지에 있습니다.
"""
from .html_pptx import HtmlToPptxConverter
from .html_pptx.converter import convert_html_to_pptx

# 하위 호환성을 위한 re-export
__all__ = ['HtmlToPptxConverter', 'convert_html_to_pptx']
