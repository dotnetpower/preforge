"""Parsers module"""
from .docx_parser import DocxParser
from .pptx_parser import PptxParser
from .pdf_parser import PdfParser
from .html_parser import HtmlParser

__all__ = [
    "DocxParser",
    "PptxParser",
    "PdfParser",
    "HtmlParser",
]
