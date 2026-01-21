# preforge 

## 개요
preforge는 각종 문서(docx, pptx, xlsx, html, md 등)를 읽거나 변환할 수 있는 파이썬 라이브러리이다. NER(Named Entity Recognition) 모델과 결합하여 문서 내의 특정 엔티티(예: 인물, 장소, 날짜 등)를 추출하는 기능을 제공한다. AI Agent 의 추론을 통해 문서 내용을 요약하거나 분석할 수도 있다.

## 기술적 요구사항

### 문서 파싱
**라이브러리**: python-docx, python-pptx, pypdf, pdfplumber, pymupdf, beautifulsoup4 등

**핵심 요구사항**:

#### 1. 기본 파싱 기능
- **지원 포맷**: DOCX, PPTX, PDF, HTML
- **추출 대상**: 텍스트, 테이블, 이미지, 메타데이터
- **구조 보존**: 제목 레벨, 단락 순서, 페이지 정보 유지

#### 2. 공통 고급 파싱 요구사항 (전체 포맷)

##### 2.1 중첩 구조 재귀 처리
- **중첩 객체 탐색**: 재귀적으로 모든 깊이의 중첩 요소 추출
  - PPTX: GROUP 안의 GROUP 안의 PICTURE/TEXT
  - DOCX: 중첩된 테이블, 텍스트 박스
  - PDF: 중첩된 콘텐츠 스트림
  - HTML: 중첩된 DOM 요소

##### 2.2 절대 좌표 계산
- **위치 추적**: 모든 요소의 문서 상 절대 위치 계산
  - 부모 컨테이너들의 좌표 누적 합산
  - top(y), left(x) 좌표 저장
  - 페이지/슬라이드 기준 절대 위치 유지

##### 2.3 다열 레이아웃 지원
- **자동 열 구분**: 문서 너비 기준 다열 레이아웃 감지
  - 2열: 50% 지점 기준 좌/우 구분
  - 3열: 33%, 66% 지점 기준 구분
  - N열: 동적 계산
- **열 단위 정렬**: 각 열 내에서 top 기준 정렬 후 좌→우 순서 병합
- **적용 대상**: PPTX, DOCX (다단 레이아웃), PDF (다단 논문), HTML (CSS 다단)

##### 2.4 텍스트 추출
- **구조 보존**: 중첩 컨테이너 내부 텍스트 재귀적 추출
- **위치 정렬**: 요소 위치(top, left) 기준 정렬
- **스타일 유지**: 제목 레벨, 폰트 스타일 정보 보존

##### 2.5 이미지 추출
- **재귀 탐색**: 중첩된 컨테이너 내 이미지 재귀적 추출
- **위치 정보**: 절대 좌표와 함께 저장
- **형식 지원**: JPG, PNG, GIF, BMP, TIFF 등
- **메타데이터**: 너비, 높이, 포맷, 페이지 번호

##### 2.6 테이블 처리
- **병합 셀 처리**: 
  - 병합된 셀의 빈 부분과 원본 셀 구분
  - 정확한 행/열 스팬 계산
- **셀 내 이미지**: 테이블 셀 내 이미지 위치 매핑 (row, col)
- **셀 서식**: 줄바꿈(\n) → `<br>` 변환, 텍스트 정렬 유지
- **중첩 테이블**: 테이블 내 테이블 재귀 처리

#### 3. 포맷별 특수 요구사항

##### 3.1 PowerPoint (PPTX)
- **Shape 타입**: MSO_SHAPE_TYPE.PICTURE, GROUP, TEXT_BOX 등 처리
- **슬라이드 좌표**: EMU 단위 (914만 = 슬라이드 너비)
- **마스터/레이아웃**: 슬라이드 마스터 상속 정보 추출
- **애니메이션**: 애니메이션 순서 정보 (선택적)
- **노트**: 발표자 노트 추출

##### 3.2 Word (DOCX)
- **섹션 구분**: 페이지 나누기, 섹션 구분 인식
- **머리글/바닥글**: 각 섹션별 헤더/푸터 추출
- **텍스트 박스**: 플로팅 텍스트 박스 위치 기반 추출
- **도형**: Drawing 객체 내 텍스트 및 이미지
- **수식**: MathML 수식 추출 (선택적)
- **주석/변경내역**: 문서 리뷰 정보 (선택적)

##### 3.3 PDF
- **텍스트 블록**: PDFMiner/PyMuPDF 기반 텍스트 블록 추출
- **좌표 시스템**: PDF 좌표계 (좌하단 원점) → 상단 기준 변환
- **폰트 정보**: 폰트 크기, 굵기로 제목 레벨 추정
- **벡터 그래픽**: 내장 벡터 이미지 래스터화
- **OCR**: 스캔 PDF의 경우 OCR 적용 (선택적)
- **암호화**: 암호화된 PDF 처리 (사용자 입력)

##### 3.4 HTML
- **DOM 파싱**: BeautifulSoup4 기반 DOM 트리 순회
- **CSS 레이아웃**: position, float, flexbox, grid 레이아웃 해석
- **시맨틱 태그**: header, article, section, aside 등 의미 구조 보존
- **이미지 소스**: 
  - 로컬 경로: 파일 시스템에서 로드
  - 원격 URL: 다운로드 (선택적)
  - Base64: 인라인 데이터 디코딩
- **스크립트 제거**: JavaScript 코드 제외
- **테이블**: colspan, rowspan 속성 해석

#### 4. 위치 기반 정렬 알고리즘
- **데이터 모델**: TextContent, ImageContent, TableContent에 position(top), left 필드 포함
- **다열 레이아웃 정렬**:
  ```python
  # 문서 너비 기반 열 구분 (예: PPTX)
  document_width = 9144000  # 포맷별 단위
  mid_point = document_width // 2
  
  # 좌/우 열 분류
  left_column = [e for e in elements if e.left < mid_point]
  right_column = [e for e in elements if e.left >= mid_point]
  
  # 각 열 내 정렬 (top 기준)
  left_column.sort(key=lambda x: x.position)
  right_column.sort(key=lambda x: x.position)
  
  # 열 순서 병합
  final_order = left_column + right_column
  ```
- **단일 열 정렬**: left 정보가 없거나 단일 열인 경우 position만으로 정렬
- **통합 출력**: 텍스트, 이미지, 테이블을 위치 기준으로 인라인 표시

#### 5. 출력 구조
```
parsing_results/
  {folder_name}/
    parsing_result.md      # 마크다운 통합 결과
    img/
      image_001.jpg        # 추출된 이미지들
      image_002.png
      table1_cell_0_1.jpg  # 테이블 셀 이미지
```

### NER 모델
- **라이브러리**: spacy
- **기능**: 개체명 인식 (인물, 장소, 날짜 등)

### AI Agent
- **라이브러리**: agent_framework
- **기능**: 추출된 문서의 제목, 부제목, 본문, 테이블, 이미지 분석

### 정교한 검색
- **라이브러리**: knowledge_graph
- **기능**: 지식 그래프 기반 검색 지원

## 개발 환경 설정
- python 3.13 이상
- uv
  - uv 로 .venv 생성 및 관리
  - uv 로 python 버전 관리
  - uv 로 초기화 및 패키지 설치(반드시 가상환경 내에서 실행)

## 필수 규칙
- 반드시 가상환경(venv) 내에서 개발 및 실행

---

## 레이아웃 시각화 기능

### 개요
PPT 슬라이드의 그리드 레이아웃을 이미지 위에 박스로 시각화하는 기능. 실제 렌더링된 컨텐츠를 기준으로 정확한 영역을 표시하며, 박스가 텍스트나 이미지와 겹치지 않도록 픽셀 단위로 조정한다.

### 핵심 기능

#### 1. 슬라이드 이미지 변환
**변환 파이프라인**: PPTX → PDF → PNG
- **LibreOffice**: `soffice --headless --convert-to pdf` 명령으로 PPTX를 PDF로 변환
- **poppler-utils**: `pdftoppm -png -r 150` 명령으로 PDF를 PNG 이미지로 변환
- **해상도**: 150 DPI (1500x1125px, 4:3 비율 기준)

#### 2. 마스터 요소 필터링
슬라이드 마스터의 반복 요소를 자동으로 제외하여 실제 컨텐츠 영역만 식별:

| 제외 요소 | 판별 기준 |
|-----------|-----------|
| **섹션 제목** (예: "1. 질병") | 상단 12% 영역 + `^\d+\.\s*[가-힣]+$` 정규식 패턴 |
| **헤더 구분선** | 상단 12% 영역 + 너비 80% 이상 + 높이 1% 미만 |
| **Confidential 워터마크** | 상단 12% + 우측 15% 영역 + 작은 크기 (너비 20% 미만) |
| **회사 로고** | 하단 8% + 우측 15% 영역 |
| **페이지 번호** | 하단 8% 영역 또는 `PP_PLACEHOLDER.SLIDE_NUMBER` 타입 |
| **푸터/날짜** | `PP_PLACEHOLDER.FOOTER`, `PP_PLACEHOLDER.DATE` 타입 |

#### 3. 픽셀 기반 컨텐츠 감지
shape 컨테이너 크기가 아닌 **실제 렌더링된 픽셀**을 분석하여 정확한 컨텐츠 영역 계산:

```python
# 배경색(흰색) 감지
bg_threshold = 250  # RGB 값이 250 이상이면 배경으로 간주
is_content = np.any(pixels < bg_threshold, axis=2)

# 컨텐츠가 있는 최소/최대 좌표 찾기
row_indices = np.where(np.any(is_content, axis=1))[0]
col_indices = np.where(np.any(is_content, axis=0))[0]

content_x1 = col_indices[0]
content_x2 = col_indices[-1] + 1
content_y1 = row_indices[0]
content_y2 = row_indices[-1] + 1
```

**장점**:
- 텍스트 박스의 빈 공백 영역 제외
- 실제 글자 크기에 맞는 정확한 바운딩 박스
- 이미지나 도형의 실제 렌더링 영역 감지

#### 4. 그리드 셀별 컨텐츠 분리
레이아웃 타입에 따라 슬라이드를 분할하고 각 영역에서 독립적으로 컨텐츠 감지:

**1x1 레이아웃** (단일 영역):
```python
region = (padding, content_top, img_width - padding, content_bottom)
bounds = find_content_bounds_pixel(image, region)
```

**1x2 레이아웃** (좌우 분할):
```python
mid_x = img_width // 2
gap = 10  # 좌우 영역 사이 여백

# 왼쪽 영역 (0 ~ 740px)
left_region = (padding, content_top, mid_x - gap, content_bottom)
left_bounds = find_content_bounds_pixel(image, left_region, margin=padding)

# 오른쪽 영역 (760 ~ 1500px)
right_region = (mid_x + gap, content_top, img_width - padding, content_bottom)
right_bounds = find_content_bounds_pixel(image, right_region, margin=padding)
```

**경계 제한**: `find_content_bounds_pixel()` 함수는 원래 검색 영역을 절대 벗어나지 않도록 결과를 제한:
```python
result = (
    max(orig_x1, x1 + content_x1 - margin),  # 검색 영역 왼쪽 경계 이상
    max(orig_y1, y1 + content_y1 - margin),  # 검색 영역 위쪽 경계 이상
    min(orig_x2, x1 + content_x2 + margin),  # 검색 영역 오른쪽 경계 이하
    min(orig_y2, y1 + content_y2 + margin)   # 검색 영역 아래쪽 경계 이하
)
```

#### 5. 박스 선과 컨텐츠 겹침 방지
박스 테두리가 텍스트나 이미지와 겹치지 않도록 **각 테두리를 개별적으로 조정**:

```python
def adjust_box_to_avoid_content(
    image: Image.Image,
    box: Tuple[int, int, int, int],
    line_width: int = 4,
    max_adjust: int = 20
) -> Tuple[int, int, int, int]:
    """박스 테두리가 컨텐츠와 겹치지 않도록 조정"""
    
    check_width = line_width + 2  # 선 두께 + 여유
    
    # 상단 테두리 확인 (y1 주변)
    for offset in range(max_adjust):
        y_check = max(0, y1 - offset)
        region = pixels[y_check - check_width//2 : y_check + check_width//2, x1:x2]
        if not np.any(is_content[region]):
            y1 = y_check  # 안전한 위치 발견
            break
    
    # 하단, 좌측, 우측 테두리도 동일하게 처리
    ...
```

**동작 원리**:
1. 각 테두리 주변 (선 두께 + 2px) 영역을 스캔
2. 해당 영역에 컨텐츠(배경이 아닌 픽셀)가 있는지 확인
3. 컨텐츠가 있으면 최대 20px까지 바깥쪽으로 이동
4. 컨텐츠가 없는 안전한 위치를 찾을 때까지 반복

### 실행 방법

**단일 페이지 시각화**:
```bash
python scripts/visualize_layout.py "path/to/file.pptx" -p 6 \
    -o "output/page_06.png" --no-show
```

**옵션**:
- `-p, --page`: 페이지 번호 (1부터 시작)
- `-o, --output`: 출력 이미지 경로
- `--no-show`: 이미지 창 표시 안 함 (저장만)
- `--box-color`: 박스 색상 (기본값: red)

### 시각적 예시

```
┌─────────────────────────────────────────────┐
│  1. 질병                      Confidential  │  ← 제외 (마스터 요소)
│─────────────────────────────────────────────│  ← 제외 (헤더 구분선)
│  ┌──────────────┐   ┌──────────────┐        │
│  │              │   │              │        │
│  │   좌측       │   │   우측       │        │  박스가 텍스트/이미지와
│  │   컨텐츠     │   │   컨텐츠     │        │  겹치지 않음
│  │              │   │              │        │
│  └──────────────┘   └──────────────┘        │
│                                    로고  6  │  ← 제외 (푸터 요소)
└─────────────────────────────────────────────┘
```

### 기술적 세부사항

**의존성**:
- `python-pptx`: PPTX 파일 파싱
- `Pillow (PIL)`: 이미지 처리 및 박스 그리기
- `numpy`: 픽셀 배열 분석
- `LibreOffice (soffice)`: PPTX→PDF 변환 (시스템 패키지)
- `poppler-utils (pdftoppm)`: PDF→PNG 변환 (시스템 패키지)

**좌표계**:
- **PPTX (EMU)**: 1인치 = 914,400 EMU (English Metric Unit)
- **이미지 (픽셀)**: 1500x1125px @ 150 DPI
- **변환 비율**: `scale_x = img_width / slide_width`

**성능**:
- 슬라이드당 변환 시간: ~2-3초 (LibreOffice + pdftoppm)
- 픽셀 분석 시간: ~0.1-0.2초 (numpy 배열 연산)
- 메모리 사용: ~30MB per 슬라이드 (1500x1125 RGB 이미지)

---

## 폴더 구조
```
preforge/
├── .github/
│   ├── workflows/
│   │   ├── ci.yml                    # CI/CD 파이프라인
│   │   └── release.yml               # 릴리스 자동화
│   └── copilot-instructions.md       # GitHub Copilot 지침
│
├── docs/
│   ├── requirements.md               # 기술적 요구사항 문서
│   ├── setup.md                      # 개발 환경 설정 가이드
│   ├── architecture.md               # 시스템 아키텍처 문서
│   ├── api-reference.md              # API 문서
│   └── examples.md                   # 사용 예제 모음
│
├── src/
│   ├── preforge/
│   │   ├── __init__.py               # 패키지 초기화
│   │   ├── core/
│   │   │   ├── __init__.py
│   │   │   ├── document.py           # 문서 기본 클래스 및 인터페이스
│   │   │   ├── parser.py             # 파서 인터페이스 정의
│   │   │   └── extractor.py          # 데이터 추출 기본 클래스
│   │   │
│   │   ├── parsers/
│   │   │   ├── __init__.py
│   │   │   ├── docx_parser.py        # Word 문서 파서
│   │   │   ├── pptx_parser.py        # PowerPoint 파서
│   │   │   ├── xlsx_parser.py        # Excel 파서
│   │   │   ├── pdf_parser.py         # PDF 파서
│   │   │   ├── html_parser.py        # HTML 파서
│   │   │   ├── markdown_parser.py    # Markdown 파서
│   │   │   └── txt_parser.py         # 텍스트 파일 파서
│   │   │
│   │   ├── extractors/
│   │   │   ├── __init__.py
│   │   │   ├── text_extractor.py     # 텍스트 추출기
│   │   │   ├── table_extractor.py    # 테이블 추출기
│   │   │   ├── image_extractor.py    # 이미지 추출기
│   │   │   ├── metadata_extractor.py # 메타데이터 추출기
│   │   │   └── structure_extractor.py # 구조 정보 추출기
│   │   │
│   │   ├── models/
│   │   │   ├── __init__.py
│   │   │   ├── ner/
│   │   │   │   ├── __init__.py
│   │   │   │   ├── ner_model.py      # NER 모델 기본 클래스
│   │   │   │   ├── spacy_ner.py      # spaCy 기반 NER
│   │   │   │   ├── transformers_ner.py # Transformers 기반 NER
│   │   │   │   └── custom_ner.py     # 커스텀 NER 모델
│   │   │   │
│   │   │   └── entity.py             # 엔티티 데이터 모델
│   │   │
│   │   ├── agents/
│   │   │   ├── __init__.py
│   │   │   ├── base_agent.py         # 에이전트 기본 클래스
│   │   │   ├── document_agent.py     # 문서 처리 에이전트
│   │   │   ├── summary_agent.py      # 요약 에이전트
│   │   │   ├── analysis_agent.py     # 분석 에이전트
│   │   │   └── qa_agent.py           # 질의응답 에이전트
│   │   │
│   │   ├── knowledge/
│   │   │   ├── __init__.py
│   │   │   ├── graph_builder.py      # 지식 그래프 구축
│   │   │   ├── graph_query.py        # 그래프 쿼리 엔진
│   │   │   ├── semantic_search.py    # 의미론적 검색
│   │   │   └── indexer.py            # 문서 인덱싱
│   │   │
│   │   ├── utils/
│   │   │   ├── __init__.py
│   │   │   ├── file_handler.py       # 파일 입출력 유틸
│   │   │   ├── text_processor.py     # 텍스트 전처리
│   │   │   ├── image_processor.py    # 이미지 처리
│   │   │   ├── validator.py          # 데이터 검증
│   │   │   └── logger.py             # 로깅 설정
│   │   │
│   │   └── config/
│   │       ├── __init__.py
│   │       ├── settings.py           # 전역 설정
│   │       └── constants.py          # 상수 정의
│   │
│   └── cli/
│       ├── __init__.py
│       └── main.py                   # CLI 진입점
│
├── tests/
│   ├── __init__.py
│   ├── unit/
│   │   ├── test_parsers.py           # 파서 단위 테스트
│   │   ├── test_extractors.py        # 추출기 테스트
│   │   ├── test_models.py            # 모델 테스트
│   │   └── test_agents.py            # 에이전트 테스트
│   │
│   ├── integration/
│   │   ├── test_pipeline.py          # 파이프라인 통합 테스트
│   │   └── test_end_to_end.py        # E2E 테스트
│   │
│   └── fixtures/
│       ├── sample.docx               # 테스트용 샘플 문서들
│       ├── sample.pptx
│       ├── sample.xlsx
│       ├── sample.pdf
│       └── sample.html
│
├── examples/
│   ├── basic_usage.py                # 기본 사용 예제
│   ├── ner_example.py                # NER 사용 예제
│   ├── agent_example.py              # AI Agent 예제
│   └── knowledge_graph_example.py    # 지식 그래프 예제
│
├── scripts/
│   ├── setup_env.sh                  # 환경 설정 스크립트
│   ├── download_models.py            # NER 모델 다운로드
│   └── benchmark.py                  # 성능 벤치마크
│
├── .env.example                      # 환경 변수 템플릿
├── .gitignore                        # Git 제외 파일 목록
├── .python-version                   # Python 버전 지정 (3.13+)
├── pyproject.toml                    # 프로젝트 메타데이터 및 의존성
├── uv.lock                           # uv 잠금 파일
├── README.md                         # 프로젝트 개요
└── LICENSE                           # 라이선스
```

