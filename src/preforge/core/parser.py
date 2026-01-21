"""
파서 인터페이스 정의
"""
from abc import ABC, abstractmethod
from pathlib import Path
from typing import List

from .document import Document, DocumentType


class BaseParser(ABC):
    """모든 파서의 기본 인터페이스"""
    
    @property
    @abstractmethod
    def supported_extensions(self) -> List[str]:
        """이 파서가 지원하는 파일 확장자 목록"""
        pass
    
    @property
    @abstractmethod
    def document_type(self) -> DocumentType:
        """이 파서가 처리하는 문서 타입"""
        pass
    
    @abstractmethod
    def parse(self, file_path: Path) -> Document:
        """
        파일을 파싱하여 Document 객체로 변환
        
        Args:
            file_path: 파싱할 파일 경로
            
        Returns:
            Document: 파싱된 문서 객체
            
        Raises:
            FileNotFoundError: 파일이 존재하지 않는 경우
            ValueError: 지원하지 않는 파일 형식인 경우
            Exception: 파싱 중 오류 발생
        """
        pass
    
    def can_parse(self, file_path: Path) -> bool:
        """
        이 파서가 해당 파일을 파싱할 수 있는지 확인
        
        Args:
            file_path: 확인할 파일 경로
            
        Returns:
            bool: 파싱 가능 여부
        """
        if not file_path.exists():
            return False
        
        extension = file_path.suffix.lower()
        return extension in self.supported_extensions
    
    def validate_file(self, file_path: Path) -> None:
        """
        파일 유효성 검사
        
        Args:
            file_path: 검사할 파일 경로
            
        Raises:
            FileNotFoundError: 파일이 존재하지 않는 경우
            ValueError: 지원하지 않는 파일 형식인 경우
        """
        if not file_path.exists():
            raise FileNotFoundError(f"파일을 찾을 수 없습니다: {file_path}")
        
        if not self.can_parse(file_path):
            raise ValueError(
                f"지원하지 않는 파일 형식입니다. "
                f"지원 형식: {', '.join(self.supported_extensions)}"
            )
