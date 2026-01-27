#!/usr/bin/env python3
"""
HTML 파일을 PPTX로 변환하는 스크립트

Usage:
    python scripts/convert_html_to_pptx.py <input.html> <output.pptx>
    
Example:
    python scripts/convert_html_to_pptx.py private/07_타겟_converted.html output/result.pptx
"""
import sys
import logging
from pathlib import Path

# 프로젝트 루트를 path에 추가
project_root = Path(__file__).parent.parent
sys.path.insert(0, str(project_root / "src"))

from preforge.converters import convert_html_to_pptx

# 로깅 설정
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
)

logger = logging.getLogger(__name__)


def main():
    """메인 실행 함수"""
    if len(sys.argv) < 3:
        print("사용법: python scripts/convert_html_to_pptx.py <input.html> <output.pptx>")
        print("\n예시:")
        print("  python scripts/convert_html_to_pptx.py private/07_타겟_converted.html output/result.pptx")
        sys.exit(1)
    
    input_path = Path(sys.argv[1])
    output_path = Path(sys.argv[2])
    
    # 입력 파일 검증
    if not input_path.exists():
        logger.error(f"입력 파일을 찾을 수 없습니다: {input_path}")
        sys.exit(1)
    
    # 출력 디렉토리 생성
    output_path.parent.mkdir(parents=True, exist_ok=True)
    
    try:
        logger.info(f"변환 시작: {input_path} -> {output_path}")
        convert_html_to_pptx(input_path, output_path)
        logger.info(f"✅ 변환 완료: {output_path}")
        print(f"\n✅ 성공적으로 변환되었습니다!")
        print(f"   출력 파일: {output_path.absolute()}")
        
    except Exception as e:
        logger.error(f"변환 실패: {e}", exc_info=True)
        sys.exit(1)


if __name__ == "__main__":
    main()
