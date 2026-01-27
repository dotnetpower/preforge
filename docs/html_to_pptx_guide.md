# HTML to PPTX 변환 가이드

## 개요

`html_pptx` 패키지는 HTML 문서를 PowerPoint 프레젠테이션(.pptx)으로 변환하는 기능을 제공합니다. 모듈화된 구조로 유지보수성과 확장성을 높였습니다.

## 모듈 구조

```
src/preforge/converters/
├── html_to_pptx.py          # 래퍼 (하위 호환성)
└── html_pptx/
    ├── __init__.py
    ├── config.py             # 설정, 색상, 상수 (83줄)
    ├── style_utils.py        # 스타일 추출, 텍스트 유틸리티 (250줄)
    ├── table_builder.py      # 테이블 생성, 테두리 적용 (480줄)
    ├── slide_factory.py      # 슬라이드 빌더 클래스들 (648줄)
    └── converter.py          # 메인 변환 로직 (570줄)
```

## 주요 기능

- HTML 문서 구조를 슬라이드로 자동 변환
- 테이블 내 불릿(bullet) 및 줄바꿈 보존
- CSS 스타일 기반 색상 자동 적용
- 섹션별 슬라이드 자동 생성
- h3 하위 섹션 그룹화 (한 슬라이드에 여러 테이블)
- Reference 카드 및 Evidence 테이블 지원
- 독립적인 h3 섹션 처리 (gene-section 외부)

## 사용 방법

### 1. 명령줄에서 직접 실행

```bash
# 기본 사용법
python scripts/convert_html_to_pptx.py <input.html> <output.pptx>

# 예시
python scripts/convert_html_to_pptx.py private/07_타겟_converted.html output/result.pptx
```

### 2. Python 코드에서 사용

#### 편의 함수 사용 (권장)

```python
from pathlib import Path
from preforge.converters.html_pptx import convert_html_to_pptx

# HTML 파일을 PPTX로 변환
input_path = Path("input.html")
output_path = Path("output.pptx")

convert_html_to_pptx(input_path, output_path)
```

#### Converter 클래스 직접 사용

```python
from pathlib import Path
from preforge.converters.html_pptx import HtmlToPptxConverter

# Converter 인스턴스 생성
converter = HtmlToPptxConverter()

# 변환 실행
input_path = Path("input.html")
output_path = Path("output.pptx")

converter.convert(input_path, output_path)
```

#### 커스텀 설정 사용

```python
from pathlib import Path
from pptx.util import Inches, Pt
from preforge.converters.html_pptx import HtmlToPptxConverter
from preforge.converters.html_pptx.config import SlideConfig, TableConfig

# 슬라이드 설정 커스터마이징
slide_config = SlideConfig(
    width=Inches(13.333),   # 와이드스크린
    height=Inches(7.5),
    margin_left=Inches(0.5),
    margin_right=Inches(0.5)
)

# 테이블 설정 커스터마이징
table_config = TableConfig(
    header_font_size=Pt(11),
    body_font_size=Pt(10),
    max_rows_per_slide=12
)

converter = HtmlToPptxConverter(
    slide_config=slide_config,
    table_config=table_config
)
converter.convert(input_path, output_path)
```

### 3. 모듈로 직접 실행

```bash
python -m preforge.converters.html_pptx.converter input.html output.pptx
```

## 핵심 클래스 및 함수

### config.py
| 클래스 | 설명 |
|--------|------|
| `SlideConfig` | 슬라이드 크기, 여백 설정 |
| `TableConfig` | 테이블 폰트, 여백, 행 수 제한 설정 |
| `BorderConfig` | 테두리 두께 설정 |
| `ColorPalette` | 색상 팔레트 (TypedDict) |

### style_utils.py
| 클래스/함수 | 설명 |
|-------------|------|
| `StyleExtractor.extract_cell_styles()` | HTML 셀에서 bold, color, link 추출 |
| `StyleExtractor.parse_color()` | hex/rgb 색상 문자열을 RGBColor로 변환 |
| `TextUtils.clean_text()` | 공백 정리 |
| `TextUtils.extract_cell_text_with_formatting()` | bullet, 줄바꿈 유지하며 텍스트 추출 |

### table_builder.py
| 클래스 | 설명 |
|--------|------|
| `TableDataExtractor` | HTML 테이블에서 데이터, 머지 정보, 스타일 추출 |
| `TableBorderStyler` | 학술 논문 스타일 테두리 적용 |
| `TableColumnAdjuster` | 열 너비 자동/수동 조정 |
| `TableBuilder` | PowerPoint 테이블 생성 |

### slide_factory.py
| 클래스 | 설명 |
|--------|------|
| `TitleSlideBuilder` | 타이틀 슬라이드 생성 |
| `ContentSlideBuilder` | 일반 콘텐츠 슬라이드 생성 |
| `TableSlideBuilder` | 테이블 슬라이드 생성 (자동 분할) |
| `ImageSlideBuilder` | 이미지 슬라이드 생성 |
| `EvidenceSlideBuilder` | Evidence 테이블 슬라이드 생성 |

### converter.py
| 클래스/함수 | 설명 |
|-------------|------|
| `HtmlToPptxConverter` | 메인 변환 클래스 |
| `convert_html_to_pptx()` | 편의 함수 |

## 지원하는 HTML 구조

### 1. 헤더 섹션
```html
<div class="app-header">
    <div class="header-title">타이틀</div>
    <div class="header-subtitle">부제목</div>
</div>
```
→ 타이틀 슬라이드로 변환

### 2. 분석 요약 섹션
```html
<div class="analysis-summary">
    <div class="summary-section">
        <div class="section-header">요약</div>
        <table class="data-table">...</table>
    </div>
</div>
```
→ 요약 슬라이드로 변환

### 3. Gene 섹션
```html
<div class="gene-section">
    <h2 class="gene-title">Gene Name</h2>
    <div class="background-text">배경 설명...</div>
    <table class="major-institution-table">...</table>
    <div class="reference-card">...</div>
</div>
```
→ Gene별 여러 슬라이드로 변환

### 4. 테이블 (불릿/줄바꿈 지원)
```html
<table class="data-table">
    <thead>
        <tr>
            <th>헤더1</th>
            <th>헤더2</th>
        </tr>
    </thead>
    <tbody>
        <tr>
            <td>데이터1</td>
            <td>
                <ul>
                    <li>항목1</li>
                    <li>항목2</li>
                </ul>
            </td>
        </tr>
        <tr>
            <td>줄바꿈 예시</td>
            <td>첫 번째 줄<br>두 번째 줄</td>
        </tr>
    </tbody>
</table>
```
→ PowerPoint 테이블로 변환 (불릿: `•`, 줄바꿈: `\n` 유지)

### 5. h3 하위 섹션 (그룹화)
```html
<div class="gene-section">
    <h2 class="subsection-title">3) 제품 및 기관정보</h2>
    <h3>3.1) 자사 제품별 타겟 유전자 현황</h3>
    <table>...</table>
    <h3>3.2) 타사 제품별 타겟 유전자 현황</h3>
    <table>...</table>
</div>
```
→ 하나의 슬라이드에 모든 h3 제목과 테이블이 함께 표시됨

### 6. 독립적인 h3 섹션
```html
<div class="content-container">
    <!-- gene-section 외부의 h3 -->
    <h3>3.3) 주요기관 권고 현황</h3>
    <table class="major-institution-table">...</table>
    
    <h3>3.4) 주요 문헌 (Top 10 References)</h3>
    <div class="reference-card">...</div>
</div>
```
→ 각각 별도의 슬라이드로 변환

### 5. Reference 카드
```html
<div class="reference-card">
    <div class="reference-item">
        <div class="reference-number">Ref. 1</div>
        <div class="reference-title">논문 제목</div>
        <div class="reference-meta">
            <div class="reference-meta-item">저자: John Doe</div>
            <div class="reference-meta-item">연도: 2024</div>
        </div>
        <div class="reference-summary">요약 내용...</div>
    </div>
</div>
```
→ Reference 슬라이드로 변환

## 테이블 정렬 규칙

| 조건 | 정렬 |
|------|------|
| 헤더 행 | 가운데 정렬 |
| 불릿(`•`) 포함 | 왼쪽 정렬 |
| 줄바꿈(`\n`) 포함 | 왼쪽 정렬 |
| 그 외 | 가운데 정렬 |

## 슬라이드 레이아웃

### 슬라이드 크기
- 비율: 16:9
- 너비: 10 inches
- 높이: 7.5 inches

### 여백
- 좌측: 0.5 inches
- 우측: 0.5 inches
- 상단: 0.7 inches
- 하단: 0.5 inches

### 색상 테마
- Primary Red: #dc2626
- Gray Scale: #f9fafb ~ #111827
- White: #ffffff

## 출력 예시

변환된 PPTX 파일에는 다음과 같은 슬라이드가 포함됩니다:

1. **타이틀 슬라이드**: 문서 제목과 부제목
2. **요약 슬라이드**: 분석 요약 테이블
3. **Ranking 슬라이드**: Target Gene Ranking 테이블
4. **Gene 개요 슬라이드**: 각 Gene의 배경 정보
5. **상세 정보 슬라이드**: 기관 권고 현황, 상용화 키트 등
6. **h3 그룹 슬라이드**: 여러 h3 섹션과 테이블이 하나의 슬라이드에
7. **독립 h3 슬라이드**: gene-section 외부의 h3 섹션 (예: 3.3, 3.4)
8. **Reference 슬라이드**: 각 논문 참조 정보

## 변환 흐름

```
HTML 파일
    ↓
BeautifulSoup으로 파싱
    ↓
├── 타이틀 슬라이드 생성 (_create_title_slide)
├── 분석 요약 슬라이드 생성 (_create_analysis_summary_slides)
├── 메인 콘텐츠 처리 (_process_main_content)
│   ├── gene-section 순회
│   │   ├── 이미지 처리 (_process_images)
│   │   ├── 테이블 처리 (_process_tables) - 동적 그룹화
│   │   └── h3 하위 섹션 처리 (_process_subsections) - 그룹화
│   ├── 독립 h3 섹션 처리 (_process_standalone_h3_sections)
│   └── Evidence 테이블 처리 (_process_evidence_tables)
    ↓
PPTX 파일 저장
```

## 테스트

```bash
# 단위 테스트 실행
python -m pytest tests/unit/test_html_to_pptx_converter.py -v

# 전체 테스트 실행
python -m pytest tests/ -v
```

## 로깅

변환 과정에서 자세한 로그를 확인하려면:

```python
import logging

logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
)
```

## 제한사항

1. **이미지**: Base64 인코딩된 이미지 지원, Playwright 스크린샷 캡처 지원
2. **복잡한 CSS**: 일부 복잡한 CSS 스타일은 변환되지 않을 수 있습니다
3. **JavaScript**: JavaScript로 생성되는 동적 콘텐츠는 변환되지 않습니다
4. **깊은 중첩**: 3단계 이상 중첩된 리스트는 단순화됩니다

## 의존성

- `python-pptx >= 1.0.2`: PowerPoint 파일 생성
- `beautifulsoup4 >= 4.14.3`: HTML 파싱
- `lxml >= 6.0.2`: HTML 파서
- `pillow >= 12.1.0` (선택적): 이미지 처리
- `playwright` (선택적): 동적 요소 스크린샷 캡처

## 문제 해결

### 테이블이 너무 작게 표시됨
→ `TableConfig`의 `min_row_height` 파라미터를 조정하세요.

### 텍스트가 잘림
→ 텍스트 박스의 크기를 늘리거나 폰트 크기를 줄이세요.

### 색상이 다르게 표시됨
→ `DEFAULT_COLORS` 딕셔너리에서 색상 값을 조정하세요.

### 불릿이 표시되지 않음
→ HTML에서 `<ul><li>` 태그를 사용했는지 확인하세요.

### h3 섹션이 누락됨
→ h3가 `gene-section` 내부에 있는지, 또는 `content-container`의 직접 자식인지 확인하세요.

## 버전 히스토리

### v21 (2026-01-27)
- 테이블 셀 내 불릿(`<ul>`, `<li>`) 및 줄바꿈(`<br>`) 보존
- 정렬 로직 개선: 불릿/줄바꿈 포함 시에만 왼쪽 정렬
- h3 하위 섹션 그룹화 (한 슬라이드에 여러 h3+테이블)
- `content-container` 직접 자식 h3 섹션 처리 (3.3, 3.4 등)
- Reference 카드 슬라이드 생성 기능 추가

### v18 (2026-01-27)
- 모듈 리팩토링: 단일 파일(2,023줄) → 6개 모듈로 분리
- 단일 책임 원칙 적용
- dataclass 기반 설정 관리

## 향후 개선 계획

- [ ] 더 많은 HTML 구조 지원
- [ ] CSS 스타일 자동 변환 확대
- [ ] 슬라이드 테마 커스터마이징
- [ ] 차트 및 그래프 지원
- [ ] 테이블 자동 페이지 분할 개선

## 라이센스

MIT License

## 기여

이슈나 풀 리퀘스트를 환영합니다!
