"""
HTML to PPTX converter tests
"""
import pytest
from pathlib import Path
import tempfile
import shutil

from preforge.converters.html_to_pptx import HtmlToPptxConverter, convert_html_to_pptx


class TestHtmlToPptxConverter:
    """HtmlToPptxConverter tests"""
    
    @pytest.fixture
    def temp_dir(self):
        """Create temporary directory"""
        temp_path = Path(tempfile.mkdtemp())
        yield temp_path
        shutil.rmtree(temp_path)
    
    @pytest.fixture
    def sample_html(self, temp_dir):
        """Create sample HTML file"""
        html_content = """
        <!DOCTYPE html>
        <html lang="en">
        <head>
            <meta charset="UTF-8">
            <title>Test Document</title>
        </head>
        <body>
            <div class="app-header">
                <div class="header-title">Test Title</div>
                <div class="header-subtitle">Test Subtitle</div>
            </div>
            
            <div class="analysis-summary">
                <div class="summary-section">
                    <div class="section-header">Summary</div>
                    <table class="data-table">
                        <tbody>
                            <tr>
                                <td>Item1</td>
                                <td>Value1</td>
                            </tr>
                            <tr>
                                <td>Item2</td>
                                <td>Value2</td>
                            </tr>
                        </tbody>
                    </table>
                </div>
            </div>
            
            <div class="gene-section">
                <h2 class="gene-title">Gene ABC</h2>
                <div class="background-text">
                    This is a background description.
                </div>
            </div>
        </body>
        </html>
        """
        
        html_path = temp_dir / "test.html"
        html_path.write_text(html_content, encoding='utf-8')
        return html_path
    
    def test_converter_initialization(self):
        """Converter initialization test"""
        converter = HtmlToPptxConverter()
        assert converter is not None
        assert converter.colors is not None
        assert 'primary_red' in converter.colors
    
    def test_convert_basic_html(self, sample_html, temp_dir):
        """Basic HTML conversion test"""
        output_path = temp_dir / "output.pptx"
        
        converter = HtmlToPptxConverter()
        converter.convert(sample_html, output_path)
        
        # Verify output file was created
        assert output_path.exists()
        assert output_path.stat().st_size > 0
    
    def test_convert_html_to_pptx_function(self, sample_html, temp_dir):
        """Convenience function test"""
        output_path = temp_dir / "output2.pptx"
        
        convert_html_to_pptx(sample_html, output_path)
        
        # Verify output file was created
        assert output_path.exists()
        assert output_path.stat().st_size > 0
    
    def test_convert_real_file(self, temp_dir):
        """Real file conversion test (only if file exists)"""
        real_file = Path("private/07_타겟_converted.html")
        
        if not real_file.exists():
            pytest.skip("Real HTML file not found")
        
        output_path = temp_dir / "real_output.pptx"
        
        convert_html_to_pptx(real_file, output_path)
        
        # Verify output file was created
        assert output_path.exists()
        assert output_path.stat().st_size > 0
    
    def test_clean_text(self):
        """Text cleaning function test"""
        converter = HtmlToPptxConverter()
        
        # Remove consecutive whitespace
        text = "Hello    World"
        cleaned = converter._clean_text(text)
        assert cleaned == "Hello World"
        
        # Remove leading/trailing whitespace
        text = "  Hello World  "
        cleaned = converter._clean_text(text)
        assert cleaned == "Hello World"
        
        # Convert newlines to spaces
        text = "Hello\n\nWorld"
        cleaned = converter._clean_text(text)
        assert cleaned == "Hello World"


if __name__ == "__main__":
    pytest.main([__file__, "-v"])
