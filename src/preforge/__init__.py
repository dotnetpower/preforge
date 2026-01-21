"""
preforge - 문서 파싱 및 분석 라이브러리

각종 문서(docx, pptx, xlsx, html, md, pdf 등)를 읽고 변환하며,
NER 모델과 AI Agent를 활용한 문서 분석 기능을 제공합니다.
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
