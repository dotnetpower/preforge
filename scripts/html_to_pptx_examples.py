#!/usr/bin/env python3
"""
HTML to PPTX 변환 예제 스크립트

다양한 사용 방법을 보여줍니다.
"""
import sys
import logging
from pathlib import Path

# 프로젝트 루트를 path에 추가
project_root = Path(__file__).parent.parent
sys.path.insert(0, str(project_root / "src"))

# 프로젝트 모듈 import
from preforge.converters import HtmlToPptxConverter, convert_html_to_pptx

# 로깅 설정
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
)

logger = logging.getLogger(__name__)


def example1_simple_conversion():
    """예제 1: 간단한 변환 (편의 함수 사용)"""
    print("\n=== 예제 1: 간단한 변환 ===")
    
    input_path = Path("private/07_타겟_converted.html")
    output_path = Path("private/example1_output.pptx")
    
    if not input_path.exists():
        print(f"❌ 입력 파일이 없습니다: {input_path}")
        return
    
    # 편의 함수로 변환
    convert_html_to_pptx(input_path, output_path)
    
    print(f"✅ 변환 완료: {output_path}")


def example2_converter_class():
    """예제 2: Converter 클래스 직접 사용"""
    print("\n=== 예제 2: Converter 클래스 사용 ===")
    
    input_path = Path("private/07_타겟_converted.html")
    output_path = Path("private/example2_output.pptx")
    
    if not input_path.exists():
        print(f"❌ 입력 파일이 없습니다: {input_path}")
        return
    
    # Converter 인스턴스 생성
    converter = HtmlToPptxConverter()
    
    # 색상 커스터마이징 (선택적)
    converter.colors['primary_red'] = converter.colors['primary_red']
    
    # 변환 실행
    converter.convert(input_path, output_path)
    
    print(f"✅ 변환 완료: {output_path}")
    print(f"   슬라이드 수: {len(converter.prs.slides)}")


def example3_batch_conversion():
    """예제 3: 여러 파일 일괄 변환"""
    print("\n=== 예제 3: 일괄 변환 ===")
    
    # private 폴더의 모든 HTML 파일 찾기
    input_dir = Path("private")
    output_dir = Path("private/pptx_output")
    
    if not input_dir.exists():
        print(f"❌ 입력 폴더가 없습니다: {input_dir}")
        return
    
    # 출력 폴더 생성
    output_dir.mkdir(parents=True, exist_ok=True)
    
    # HTML 파일 찾기
    html_files = list(input_dir.glob("*.html"))
    
    if not html_files:
        print(f"❌ HTML 파일을 찾을 수 없습니다: {input_dir}")
        return
    
    print(f"발견된 HTML 파일: {len(html_files)}개")
    
    # 각 파일 변환
    success_count = 0
    for html_file in html_files:
        try:
            output_file = output_dir / f"{html_file.stem}.pptx"
            print(f"\n변환 중: {html_file.name} -> {output_file.name}")
            
            convert_html_to_pptx(html_file, output_file)
            success_count += 1
            
            print(f"  ✅ 완료")
            
        except Exception as e:
            print(f"  ❌ 실패: {e}")
    
    print(f"\n총 {success_count}/{len(html_files)}개 파일 변환 완료")


def example4_custom_output_directory():
    """예제 4: 커스텀 출력 디렉토리"""
    print("\n=== 예제 4: 커스텀 출력 디렉토리 ===")
    
    input_path = Path("private/07_타겟_converted.html")
    output_dir = Path("output/pptx")
    
    if not input_path.exists():
        print(f"❌ 입력 파일이 없습니다: {input_path}")
        return
    
    # 출력 디렉토리 생성
    output_dir.mkdir(parents=True, exist_ok=True)
    
    # 출력 파일 경로
    output_path = output_dir / "converted_presentation.pptx"
    
    # 변환
    convert_html_to_pptx(input_path, output_path)
    
    print(f"✅ 변환 완료: {output_path.absolute()}")


def main():
    """메인 실행 함수"""
    print("=" * 60)
    print("HTML to PPTX 변환 예제")
    print("=" * 60)
    
    # 예제 실행
    example1_simple_conversion()
    example2_converter_class()
    # example3_batch_conversion()  # 필요시 주석 해제
    example4_custom_output_directory()
    
    print("\n" + "=" * 60)
    print("모든 예제 완료!")
    print("=" * 60)


if __name__ == "__main__":
    main()
