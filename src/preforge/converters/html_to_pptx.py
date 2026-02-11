"""
Converter for HTML(.html) documents to PowerPoint(.pptx)

This module is maintained for backward compatibility.
Actual implementation is in the html_pptx package.
"""
from .html_pptx import HtmlToPptxConverter
from .html_pptx.converter import convert_html_to_pptx

# Re-export for backward compatibility
__all__ = ['HtmlToPptxConverter', 'convert_html_to_pptx']
