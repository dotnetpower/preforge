"""파서 단위 테스트"""
import pytest
from pathlib import Path

from preforge.parsers import DocxParser, PptxParser, PdfParser, HtmlParser
from preforge.core.document import DocumentType


# 테스트 문서 경로 - 프로젝트 루트의 private 폴더
PRIVATE_DIR = Path(__file__).parent.parent.parent / "private"


class TestDocxParser:
    """Word 문서 파서 테스트"""
    
    def test_parse_docx(self):
        """DOCX 파일 파싱 테스트"""
        parser = DocxParser()
        docx_file = PRIVATE_DIR / "[PPT변환 샘플].docx"
        
        if not docx_file.exists():
            pytest.skip(f"테스트 파일이 존재하지 않습니다: {docx_file}")
        
        doc = parser.parse(docx_file)
        
        assert doc is not None
        assert doc.doc_type == DocumentType.DOCX
        assert doc.file_path == docx_file
        assert len(doc.text_contents) > 0
        
        print(f"\n=== DOCX 파싱 결과 ===")
        print(f"파일: {doc.file_path.name}")
        print(f"제목: {doc.metadata.title}")
        print(f"작성자: {doc.metadata.author}")
        print(f"텍스트 개수: {len(doc.text_contents)}")
        print(f"테이블 개수: {len(doc.tables)}")
        print(f"이미지 개수: {len(doc.images)}")
        print(f"\n첫 3개 텍스트:")
        for i, tc in enumerate(doc.text_contents[:3], 1):
            print(f"{i}. [레벨 {tc.level}] {tc.text[:100]}...")


class TestPptxParser:
    """PowerPoint 파서 테스트"""
    
    def test_parse_pptx(self):
        """PPTX 파일 파싱 테스트"""
        parser = PptxParser()
        pptx_file = PRIVATE_DIR / "PPT샘플_20201027.pptx"
        
        if not pptx_file.exists():
            pytest.skip(f"테스트 파일이 존재하지 않습니다: {pptx_file}")
        
        doc = parser.parse(pptx_file)
        
        assert doc is not None
        assert doc.doc_type == DocumentType.PPTX
        assert doc.file_path == pptx_file
        assert len(doc.text_contents) > 0
        
        print(f"\n=== PPTX 파싱 결과 ===")
        print(f"파일: {doc.file_path.name}")
        print(f"제목: {doc.metadata.title}")
        print(f"슬라이드 수: {doc.metadata.page_count}")
        print(f"텍스트 개수: {len(doc.text_contents)}")
        print(f"테이블 개수: {len(doc.tables)}")
        print(f"이미지 개수: {len(doc.images)}")
        print(f"\n첫 5개 슬라이드 제목:")
        for tc in doc.headings[:5]:
            print(f"- [슬라이드 {tc.page_number}] {tc.text}")


class TestPdfParser:
    """PDF 파서 테스트"""
    
    def test_parse_pdf(self):
        """PDF 파일 파싱 테스트"""
        parser = PdfParser()
        pdf_file = PRIVATE_DIR / "02_질병의이해-malaria.report.pdf"
        
        if not pdf_file.exists():
            pytest.skip(f"테스트 파일이 존재하지 않습니다: {pdf_file}")
        
        doc = parser.parse(pdf_file)
        
        assert doc is not None
        assert doc.doc_type == DocumentType.PDF
        assert doc.file_path == pdf_file
        assert len(doc.text_contents) > 0
        
        print(f"\n=== PDF 파싱 결과 ===")
        print(f"파일: {doc.file_path.name}")
        print(f"제목: {doc.metadata.title}")
        print(f"페이지 수: {doc.metadata.page_count}")
        print(f"텍스트 개수: {len(doc.text_contents)}")
        print(f"테이블 개수: {len(doc.tables)}")
        print(f"이미지 개수: {len(doc.images)}")
        print(f"\n첫 페이지 텍스트 미리보기:")
        if doc.text_contents:
            print(doc.text_contents[0].text[:300] + "...")


class TestHtmlParser:
    """HTML 파서 테스트"""
    
    def test_parse_html(self):
        """HTML 파일 파싱 테스트"""
        parser = HtmlParser()
        html_file = PRIVATE_DIR / "Html_tick_borne_borrelia-1.html"
        
        if not html_file.exists():
            pytest.skip(f"테스트 파일이 존재하지 않습니다: {html_file}")
        
        doc = parser.parse(html_file)
        
        assert doc is not None
        assert doc.doc_type == DocumentType.HTML
        assert doc.file_path == html_file
        
        print(f"\n=== HTML 파싱 결과 ===")
        print(f"파일: {doc.file_path.name}")
        print(f"제목: {doc.metadata.title}")
        print(f"텍스트 개수: {len(doc.text_contents)}")
        print(f"테이블 개수: {len(doc.tables)}")
        print(f"이미지 개수: {len(doc.images)}")
        
        # 제목 추출
        headings = [tc for tc in doc.text_contents if tc.level > 0]
        print(f"\n제목 개수: {len(headings)}")
        if headings:
            print("첫 3개 제목:")
            for tc in headings[:3]:
                print(f"- [H{tc.level}] {tc.text}")


class TestParserIntegration:
    """파서 통합 테스트"""
    
    def test_all_parsers(self):
        """모든 파서 통합 테스트"""
        parsers = {
            "PDF": (PdfParser(), "02_질병의이해-malaria.report.pdf"),
            "HTML": (HtmlParser(), "Html_tick_borne_borrelia-1.html"),
        }
        
        print("\n" + "="*60)
        print("전체 파서 통합 테스트")
        print("="*60)
        
        for doc_type, (parser, filename) in parsers.items():
            file_path = PRIVATE_DIR / filename
            
            if not file_path.exists():
                print(f"\n[SKIP] {doc_type}: 파일 없음 - {filename}")
                continue
            
            try:
                doc = parser.parse(file_path)
                print(f"\n[OK] {doc_type}")
                print(f"  - 파일: {filename}")
                print(f"  - 텍스트: {len(doc.text_contents)}개")
                print(f"  - 테이블: {len(doc.tables)}개")
                print(f"  - 이미지: {len(doc.images)}개")
                
                if doc.tables:
                    print(f"  - 첫 번째 테이블 크기: {len(doc.tables[0].headers)} x {len(doc.tables[0].rows)}")
                
            except Exception as e:
                print(f"\n[FAIL] {doc_type}: {str(e)}")
                raise
