"""
데이터 추출 기본 클래스
"""
from abc import ABC, abstractmethod
from typing import Any, List

from .document import Document, TextContent, TableContent, ImageContent


class BaseExtractor(ABC):
    """모든 추출기의 기본 인터페이스"""
    
    @abstractmethod
    def extract(self, source: Any) -> Any:
        """
        소스에서 데이터 추출
        
        Args:
            source: 추출할 데이터 소스
            
        Returns:
            추출된 데이터
        """
        pass


class TextExtractor(BaseExtractor):
    """텍스트 추출기 기본 클래스"""
    
    @abstractmethod
    def extract(self, source: Any) -> List[TextContent]:
        """텍스트 컨텐츠 추출"""
        pass


class TableExtractor(BaseExtractor):
    """테이블 추출기 기본 클래스"""
    
    @abstractmethod
    def extract(self, source: Any) -> List[TableContent]:
        """테이블 컨텐츠 추출"""
        pass


class ImageExtractor(BaseExtractor):
    """이미지 추출기 기본 클래스"""
    
    @abstractmethod
    def extract(self, source: Any) -> List[ImageContent]:
        """이미지 컨텐츠 추출"""
        pass


class MetadataExtractor(BaseExtractor):
    """메타데이터 추출기 기본 클래스"""
    
    @abstractmethod
    def extract(self, source: Any) -> dict:
        """메타데이터 추출"""
        pass
