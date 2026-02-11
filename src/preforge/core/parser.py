"""
Parser interface definition
"""
from abc import ABC, abstractmethod
from pathlib import Path
from typing import List

from .document import Document, DocumentType


class BaseParser(ABC):
    """Base interface for all parsers"""
    
    @property
    @abstractmethod
    def supported_extensions(self) -> List[str]:
        """List of file extensions supported by this parser"""
        pass
    
    @property
    @abstractmethod
    def document_type(self) -> DocumentType:
        """Document type processed by this parser"""
        pass
    
    @abstractmethod
    def parse(self, file_path: Path) -> Document:
        """
        Parse file and convert to Document object
        
        Args:
            file_path: Path to the file to parse
            
        Returns:
            Document: Parsed document object
            
        Raises:
            FileNotFoundError: If file does not exist
            ValueError: If file format is not supported
            Exception: If error occurs during parsing
        """
        pass
    
    def can_parse(self, file_path: Path) -> bool:
        """
        Check if this parser can parse the specified file
        
        Args:
            file_path: Path to the file to check
            
        Returns:
            bool: Whether the file can be parsed
        """
        if not file_path.exists():
            return False
        
        extension = file_path.suffix.lower()
        return extension in self.supported_extensions
    
    def validate_file(self, file_path: Path) -> None:
        """
        Validate file
        
        Args:
            file_path: Path to the file to validate
            
        Raises:
            FileNotFoundError: If file does not exist
            ValueError: If file format is not supported
        """
        if not file_path.exists():
            raise FileNotFoundError(f"File not found: {file_path}")
        
        if not self.can_parse(file_path):
            raise ValueError(
                f"Unsupported file format. "
                f"Supported formats: {', '.join(self.supported_extensions)}"
            )
