"""
Base class for data extraction
"""
from abc import ABC, abstractmethod
from typing import Any, List

from .document import Document, TextContent, TableContent, ImageContent


class BaseExtractor(ABC):
    """Base interface for all extractors"""
    
    @abstractmethod
    def extract(self, source: Any) -> Any:
        """
        Extract data from source
        
        Args:
            source: Data source to extract from
            
        Returns:
            Extracted data
        """
        pass


class TextExtractor(BaseExtractor):
    """Base class for text extractors"""
    
    @abstractmethod
    def extract(self, source: Any) -> List[TextContent]:
        """Extract text content"""
        pass


class TableExtractor(BaseExtractor):
    """Base class for table extractors"""
    
    @abstractmethod
    def extract(self, source: Any) -> List[TableContent]:
        """Extract table content"""
        pass


class ImageExtractor(BaseExtractor):
    """Base class for image extractors"""
    
    @abstractmethod
    def extract(self, source: Any) -> List[ImageContent]:
        """Extract image content"""
        pass


class MetadataExtractor(BaseExtractor):
    """Base class for metadata extractors"""
    
    @abstractmethod
    def extract(self, source: Any) -> dict:
        """Extract metadata"""
        pass
