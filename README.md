# preforge
document processing

## 개요

각종 문서(docx, pptx, xlsx, html, md 등)를 읽거나 변환할 수 있는 파이썬 라이브러리입니다.

## 주요 기능

### 문서 파싱
- **DOCX**: Word 문서 파싱 및 분석
- **PPTX**: PowerPoint 프레젠테이션 파싱
- **HTML**: HTML 문서 파싱
- **PDF**: PDF 문서 파싱
- **XLSX**: Excel 파일 파싱 (계획)

### 문서 변환
- **PPTX → DOCX**: PowerPoint를 Word 문서로 변환
- **HTML → PPTX**: HTML 문서를 PowerPoint로 변환 ✨ NEW
  - 전체 문서 완전 변환 (34+ 슬라이드)
  - 테이블 가독성 대폭 향상
  - 자동 레이아웃 최적화

## 설치

```bash
# 개발 모드로 설치
pip install -e .

# 의존성만 설치
pip install -r requirements.txt
```

## 빠른 시작

### HTML to PPTX 변환

```python
from pathlib import Path
from preforge.converters import convert_html_to_pptx

# HTML 파일을 PPTX로 변환
input_path = Path("input.html")
output_path = Path("output.pptx")

convert_html_to_pptx(input_path, output_path)
```

### 명령줄에서 사용

```bash
# HTML to PPTX
python scripts/convert_html_to_pptx.py input.html output.pptx

# 예제 실행
python scripts/html_to_pptx_examples.py
```

## 문서

- [HTML to PPTX 변환 가이드](docs/html_to_pptx_guide.md)
- [아키텍처](docs/architecture.md)
- [요구사항](docs/requirements.md)
- [설정 가이드](docs/setup.md)

## 테스트

```bash
# 전체 테스트 실행
pytest tests/

# HTML to PPTX 변환 테스트만 실행
pytest tests/unit/test_html_to_pptx_converter.py -v
```

## 프로젝트 구조

```
preforge/
├── src/preforge/
│   ├── core/           # 핵심 문서 처리 모듈
│   ├── parsers/        # 문서 파서들
│   ├── converters/     # 문서 변환기들
│   └── extractors/     # 데이터 추출기들
├── tests/              # 테스트 코드
├── scripts/            # 유틸리티 스크립트
└── docs/               # 문서
```

## 의존성

- `python >= 3.13`
- `python-pptx >= 1.0.2`
- `python-docx >= 1.2.0`
- `beautifulsoup4 >= 4.14.3`
- `lxml >= 6.0.2`
- `pillow >= 12.1.0`

자세한 의존성은 [pyproject.toml](pyproject.toml)을 참조하세요.

## 라이센스

MIT License

## 기여

이슈나 풀 리퀘스트를 환영합니다!

