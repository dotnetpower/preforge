"""
HTML(.html) 문서를 PDF(.pdf)로 변환하는 컨버터

Playwright(Chromium)를 사용하여 브라우저와 동일한 품질의 PDF를 생성합니다.
"""
from pathlib import Path
from typing import Optional
import logging

try:
    from playwright.sync_api import sync_playwright
    HAS_PLAYWRIGHT = True
except ImportError:
    HAS_PLAYWRIGHT = False

logger = logging.getLogger(__name__)


class HtmlToPdfConverter:
    """HTML을 PDF로 변환하는 컨버터 (Chromium 기반)"""
    
    def __init__(self):
        if not HAS_PLAYWRIGHT:
            raise ImportError(
                "playwright가 설치되지 않았습니다. "
                "pip install playwright && playwright install chromium 으로 설치하세요."
            )
    
    def convert(
        self, 
        input_path: Path, 
        output_path: Optional[Path] = None,
        format: str = "A4",
        print_background: bool = True,
        margin_top: str = "10mm",
        margin_bottom: str = "10mm",
        margin_left: str = "10mm",
        margin_right: str = "10mm",
    ) -> Path:
        """
        HTML 파일을 PDF로 변환 (브라우저 Print to PDF와 동일)
        
        Args:
            input_path: 입력 HTML 파일 경로
            output_path: 출력 PDF 파일 경로 (None이면 같은 위치에 .pdf로 저장)
            format: 용지 크기 (A4, Letter 등)
            print_background: 배경색/이미지 포함 여부
            margin_*: 여백 설정
            
        Returns:
            생성된 PDF 파일 경로
        """
        input_path = Path(input_path).absolute()
        
        if not input_path.exists():
            raise FileNotFoundError(f"입력 파일을 찾을 수 없습니다: {input_path}")
        
        if output_path is None:
            output_path = input_path.with_suffix('.pdf')
        else:
            output_path = Path(output_path).absolute()
        
        # 출력 디렉토리 생성
        output_path.parent.mkdir(parents=True, exist_ok=True)
        
        logger.info(f"HTML → PDF 변환 시작: {input_path}")
        
        try:
            with sync_playwright() as p:
                # Chromium 브라우저 실행
                browser = p.chromium.launch()
                page = browser.new_page()
                
                # HTML 파일 로드
                page.goto(f"file://{input_path}", wait_until="networkidle")
                
                # PDF 생성 (브라우저 Print to PDF와 동일)
                page.pdf(
                    path=str(output_path),
                    format=format,
                    print_background=print_background,
                    margin={
                        "top": margin_top,
                        "bottom": margin_bottom,
                        "left": margin_left,
                        "right": margin_right,
                    }
                )
                
                browser.close()
            
            logger.info(f"PDF 변환 완료: {output_path}")
            return output_path
            
        except Exception as e:
            logger.error(f"PDF 변환 실패: {e}")
            raise
    
    def convert_string(
        self, 
        html_string: str, 
        output_path: Path,
        format: str = "A4",
        print_background: bool = True,
    ) -> Path:
        """
        HTML 문자열을 PDF로 변환
        
        Args:
            html_string: HTML 문자열
            output_path: 출력 PDF 파일 경로
            format: 용지 크기
            print_background: 배경 포함 여부
            
        Returns:
            생성된 PDF 파일 경로
        """
        output_path = Path(output_path).absolute()
        output_path.parent.mkdir(parents=True, exist_ok=True)
        
        try:
            with sync_playwright() as p:
                browser = p.chromium.launch()
                page = browser.new_page()
                
                # HTML 문자열 설정
                page.set_content(html_string, wait_until="networkidle")
                
                # PDF 생성
                page.pdf(
                    path=str(output_path),
                    format=format,
                    print_background=print_background,
                    margin={"top": "10mm", "bottom": "10mm", "left": "10mm", "right": "10mm"}
                )
                
                browser.close()
            
            return output_path
            
        except Exception as e:
            logger.error(f"PDF 변환 실패: {e}")
            raise


def convert_html_to_pdf(
    html_path: Path, 
    output_path: Optional[Path] = None
) -> Path:
    """
    HTML 파일을 PDF로 변환하는 편의 함수
    
    Args:
        html_path: 입력 HTML 파일 경로
        output_path: 출력 PDF 파일 경로
        
    Returns:
        생성된 PDF 파일 경로
    """
    converter = HtmlToPdfConverter()
    return converter.convert(html_path, output_path)
