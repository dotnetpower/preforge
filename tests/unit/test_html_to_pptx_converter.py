"""
HTML to PPTX converter 테스트
"""
import pytest
from pathlib import Path
import tempfile
import shutil

from preforge.converters.html_to_pptx import HtmlToPptxConverter, convert_html_to_pptx


class TestHtmlToPptxConverter:
    """HtmlToPptxConverter 테스트"""
    
    @pytest.fixture
    def temp_dir(self):
        """임시 디렉토리 생성"""
        temp_path = Path(tempfile.mkdtemp())
        yield temp_path
        shutil.rmtree(temp_path)
    
    @pytest.fixture
    def sample_html(self, temp_dir):
        """샘플 HTML 파일 생성"""
        html_content = """
        <!DOCTYPE html>
        <html lang="ko">
        <head>
            <meta charset="UTF-8">
            <title>테스트 문서</title>
        </head>
        <body>
            <div class="app-header">
                <div class="header-title">테스트 타이틀</div>
                <div class="header-subtitle">테스트 부제목</div>
            </div>
            
            <div class="analysis-summary">
                <div class="summary-section">
                    <div class="section-header">요약</div>
                    <table class="data-table">
                        <tbody>
                            <tr>
                                <td>항목1</td>
                                <td>값1</td>
                            </tr>
                            <tr>
                                <td>항목2</td>
                                <td>값2</td>
                            </tr>
                        </tbody>
                    </table>
                </div>
            </div>
            
            <div class="gene-section">
                <h2 class="gene-title">Gene ABC</h2>
                <div class="background-text">
                    이것은 배경 설명입니다.
                </div>
            </div>
        </body>
        </html>
        """
        
        html_path = temp_dir / "test.html"
        html_path.write_text(html_content, encoding='utf-8')
        return html_path
    
    def test_converter_initialization(self):
        """컨버터 초기화 테스트"""
        converter = HtmlToPptxConverter()
        assert converter is not None
        assert converter.colors is not None
        assert 'primary_red' in converter.colors
    
    def test_convert_basic_html(self, sample_html, temp_dir):
        """기본 HTML 변환 테스트"""
        output_path = temp_dir / "output.pptx"
        
        converter = HtmlToPptxConverter()
        converter.convert(sample_html, output_path)
        
        # 출력 파일이 생성되었는지 확인
        assert output_path.exists()
        assert output_path.stat().st_size > 0
    
    def test_convert_html_to_pptx_function(self, sample_html, temp_dir):
        """편의 함수 테스트"""
        output_path = temp_dir / "output2.pptx"
        
        convert_html_to_pptx(sample_html, output_path)
        
        # 출력 파일이 생성되었는지 확인
        assert output_path.exists()
        assert output_path.stat().st_size > 0
    
    def test_convert_real_file(self, temp_dir):
        """실제 파일 변환 테스트 (존재하는 경우에만)"""
        real_file = Path("private/07_타겟_converted.html")
        
        if not real_file.exists():
            pytest.skip("실제 HTML 파일이 없습니다")
        
        output_path = temp_dir / "real_output.pptx"
        
        convert_html_to_pptx(real_file, output_path)
        
        # 출력 파일이 생성되었는지 확인
        assert output_path.exists()
        assert output_path.stat().st_size > 0
    
    def test_clean_text(self):
        """텍스트 정리 기능 테스트"""
        converter = HtmlToPptxConverter()
        
        # 연속된 공백 제거
        text = "Hello    World"
        cleaned = converter._clean_text(text)
        assert cleaned == "Hello World"
        
        # 앞뒤 공백 제거
        text = "  Hello World  "
        cleaned = converter._clean_text(text)
        assert cleaned == "Hello World"
        
        # 줄바꿈을 공백으로
        text = "Hello\n\nWorld"
        cleaned = converter._clean_text(text)
        assert cleaned == "Hello World"


if __name__ == "__main__":
    pytest.main([__file__, "-v"])
