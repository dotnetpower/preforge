"""
preforge - Document Parsing and Analysis Library

A Python library for reading and converting various document formats
(docx, pptx, xlsx, html, md, pdf, etc.), with document analysis features
utilizing NER models and AI Agents.
"""

__version__ = "0.1.0"
__author__ = "dotnetpower"

from .core.document import Document
from .core.parser import BaseParser
from .core.extractor import BaseExtractor
from .converters import PptxToDocxConverter

__all__ = [
    "Document",
    "BaseParser",
    "BaseExtractor",
    "PptxToDocxConverter",
]
