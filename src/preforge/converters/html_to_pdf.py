"""
HTML(.html) to PDF(.pdf) converter

Uses Playwright (Chromium) to generate PDFs with browser-quality output.
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
    """HTML to PDF converter (Chromium-based)"""
    
    def __init__(self):
        if not HAS_PLAYWRIGHT:
            raise ImportError(
                "playwright is not installed. "
                "Install with: pip install playwright && playwright install chromium"
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
        Convert HTML file to PDF (equivalent to browser Print to PDF)
        
        Args:
            input_path: Input HTML file path
            output_path: Output PDF file path (if None, saves as .pdf in same location)
            format: Paper size (A4, Letter, etc.)
            print_background: Whether to include background colors/images
            margin_*: Margin settings
            
        Returns:
            Generated PDF file path
        """
        input_path = Path(input_path).absolute()
        
        if not input_path.exists():
            raise FileNotFoundError(f"Input file not found: {input_path}")
        
        if output_path is None:
            output_path = input_path.with_suffix('.pdf')
        else:
            output_path = Path(output_path).absolute()
        
        # Create output directory
        output_path.parent.mkdir(parents=True, exist_ok=True)
        
        logger.info(f"HTML -> PDF conversion started: {input_path}")
        
        try:
            with sync_playwright() as p:
                # Launch Chromium browser
                browser = p.chromium.launch()
                page = browser.new_page()
                
                # Load HTML file
                page.goto(f"file://{input_path}", wait_until="networkidle")
                
                # Generate PDF (equivalent to browser Print to PDF)
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
            
            logger.info(f"PDF conversion complete: {output_path}")
            return output_path
            
        except Exception as e:
            logger.error(f"PDF conversion failed: {e}")
            raise
    
    def convert_string(
        self, 
        html_string: str, 
        output_path: Path,
        format: str = "A4",
        print_background: bool = True,
    ) -> Path:
        """
        Convert HTML string to PDF
        
        Args:
            html_string: HTML string
            output_path: Output PDF file path
            format: Paper size
            print_background: Whether to include background
            
        Returns:
            Generated PDF file path
        """
        output_path = Path(output_path).absolute()
        output_path.parent.mkdir(parents=True, exist_ok=True)
        
        try:
            with sync_playwright() as p:
                browser = p.chromium.launch()
                page = browser.new_page()
                
                # Set HTML string
                page.set_content(html_string, wait_until="networkidle")
                
                # Generate PDF
                page.pdf(
                    path=str(output_path),
                    format=format,
                    print_background=print_background,
                    margin={"top": "10mm", "bottom": "10mm", "left": "10mm", "right": "10mm"}
                )
                
                browser.close()
            
            return output_path
            
        except Exception as e:
            logger.error(f"PDF conversion failed: {e}")
            raise


def convert_html_to_pdf(
    html_path: Path, 
    output_path: Optional[Path] = None
) -> Path:
    """
    Convenience function to convert HTML file to PDF
    
    Args:
        html_path: Input HTML file path
        output_path: Output PDF file path
        
    Returns:
        Generated PDF file path
    """
    converter = HtmlToPdfConverter()
    return converter.convert(html_path, output_path)
