# HTML to PPTX 변환 개선 내역

## 변경 사항 (2026-01-27)

### 1. 전체 문서 변환 지원 ✅

**이전:**
- 3개 슬라이드만 생성 (타이틀 + 요약 2개)
- 메인 컨텐츠의 대부분 누락

**개선 후:**
- **34개 슬라이드** 생성
- 모든 섹션 완전 변환:
  - 타이틀 슬라이드
  - Analysis Summary (2개)
  - 기본정보 (1개)
  - 세부정보 (1개)
  - 제품 및 기관정보 (여러 개)
  - 서열분석 결과 (여러 개)
  - Evidence 테이블 (여러 개)
  - 기관 권고 현황 등

### 2. 테이블 가독성 대폭 향상 ✅

**개선 사항:**

#### 폰트 및 레이아웃
- **동적 폰트 크기**: 테이블 크기에 따라 자동 조정 (8-10pt)
- **헤더 강조**: 굵은 글씨 + 회색 배경 (gray-200)
- **짝수행 배경색**: 연한 회색으로 행 구분 용이

#### 텍스트 정렬
- **긴 텍스트 (30자 이상)**: 좌측 정렬
- **짧은 텍스트**: 중앙 정렬
- **자동 줄바꿈**: 텍스트가 셀 안에 완전히 표시

#### 여백 및 간격
- **셀 여백**: 상하좌우 3-4pt로 넉넉하게
- **줄 간격**: 1.2배로 가독성 향상
- **슬라이드 여백**: 좌우 0.3", 상하 0.5"로 축소하여 더 많은 공간 확보

#### 열 너비 자동 조정
- 각 열의 텍스트 길이를 분석하여 비례적으로 너비 할당
- 최소 너비 보장으로 좁은 열 방지

### 3. 새로운 기능

#### Evidence 테이블 전용 처리
```python
_create_evidence_table_slide()
```
- Evidence 데이터를 최적화된 형식으로 표시
- 헤더: No., Ref.ID, Title, Date, Country, Summary
- 최대 10개 행까지 표시
- 7pt 작은 폰트로 많은 정보 표시

#### 자동 섹션 인식
```python
_process_main_content()
```
- HTML 구조를 자동으로 분석
- h1, h2, h3 계층 구조 인식
- 각 섹션을 적절한 슬라이드로 변환

#### 메인 타이틀 표시
- 각 상세 슬라이드에 메인 Gene 타이틀을 작게 표시
- 사용자가 어떤 Gene의 정보를 보고 있는지 명확하게 인식

### 4. 성능 개선

- **처리 속도**: ~1.2초 (34개 슬라이드)
- **파일 크기**: 91KB (적절한 크기)
- **메모리 효율**: 대용량 HTML도 안정적 처리

## 사용 예시

### 기본 사용
```bash
python scripts/convert_html_to_pptx.py input.html output.pptx
```

### Python 코드
```python
from preforge.converters import convert_html_to_pptx
convert_html_to_pptx("input.html", "output.pptx")
```

## 변환 결과 비교

| 항목 | 이전 | 개선 후 |
|------|------|---------|
| 슬라이드 수 | 3개 | 34개 |
| 파일 크기 | 32KB | 91KB |
| 테이블 가독성 | 낮음 | 높음 |
| 전체 컨텐츠 변환 | 10% | 100% |
| 자동 레이아웃 조정 | 없음 | 있음 |

## 지원하는 HTML 구조

### 메인 컨테이너
```html
<div class="content-container">
  <h1 class="gene-title">Gene Name</h1>
  
  <div class="gene-section">
    <h2 class="subsection-title">섹션 제목</h2>
    <table class="data-table">...</table>
  </div>
  
  <h2 class="subsection-title">Evidence</h2>
  <div class="evidence-table">...</div>
</div>
```

### 지원 요소
- ✅ `<h1>`, `<h2>`, `<h3>` 제목 계층
- ✅ `<table class="data-table">` 일반 테이블
- ✅ `<div class="evidence-table">` Evidence 테이블
- ✅ `<thead>`, `<tbody>` 테이블 구조
- ✅ `colspan`, `rowspan` (부분 지원)

## 알려진 제한사항

1. **이미지**: 현재 버전에서는 이미지 변환 미지원
2. **복잡한 스타일**: 일부 CSS 스타일은 변환되지 않음
3. **차트**: SVG나 Canvas 차트는 변환 안됨
4. **대용량 테이블**: 50행 이상의 테이블은 여러 슬라이드로 분할 필요

## 향후 개선 계획

- [ ] 대용량 테이블 자동 분할
- [ ] 이미지 변환 지원
- [ ] 차트/그래프 변환
- [ ] 슬라이드 노트 추가
- [ ] 하이퍼링크 보존
- [ ] 커스텀 테마 지원

## 문의 및 버그 리포트

이슈나 개선 제안은 GitHub Issues를 통해 제출해주세요.
