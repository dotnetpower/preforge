# preforge
document processing

## Overview

A Python library for reading and converting various document formats (docx, pptx, xlsx, html, md, etc.).

## Key Features

### Document Parsing
- **DOCX**: Word document parsing and analysis
- **PPTX**: PowerPoint presentation parsing
- **HTML**: HTML document parsing
- **PDF**: PDF document parsing
- **XLSX**: Excel file parsing (planned)

### Document Conversion
- **PPTX to DOCX**: Convert PowerPoint to Word documents
- **HTML to PPTX**: Convert HTML documents to PowerPoint NEW
  - Complete document conversion (34+ slides)
  - Significantly improved table readability
  - Automatic layout optimization

## Installation

```bash
# Install in development mode
pip install -e .

# Install dependencies only
pip install -r requirements.txt
```

## Quick Start

### HTML to PPTX Conversion

```python
from pathlib import Path
from preforge.converters import convert_html_to_pptx

# Convert HTML file to PPTX
input_path = Path("input.html")
output_path = Path("output.pptx")

convert_html_to_pptx(input_path, output_path)
```

### Command Line Usage

```bash
# HTML to PPTX
python scripts/convert_html_to_pptx.py input.html output.pptx

# Run examples
python scripts/html_to_pptx_examples.py
```

## Documentation

- [HTML to PPTX Conversion Guide](docs/html_to_pptx_guide.md)
- [Architecture](docs/architecture.md)
- [Requirements](docs/requirements.md)
- [Setup Guide](docs/setup.md)

## Testing

```bash
# Run all tests
pytest tests/

# Run only HTML to PPTX conversion tests
pytest tests/unit/test_html_to_pptx_converter.py -v
```

## Project Structure

```
preforge/
├── src/preforge/
│   ├── core/           # Core document processing module
│   ├── parsers/        # Document parsers
│   ├── converters/     # Document converters
│   └── extractors/     # Data extractors
├── tests/              # Test code
├── scripts/            # Utility scripts
└── docs/               # Documentation
```

## Dependencies

- `python >= 3.13`
- `python-pptx >= 1.0.2`
- `python-docx >= 1.2.0`
- `beautifulsoup4 >= 4.14.3`
- `lxml >= 6.0.2`
- `pillow >= 12.1.0`

For detailed dependencies, refer to [pyproject.toml](pyproject.toml).

## 라이센스

MIT License

## 기여

이슈나 풀 리퀘스트를 환영합니다!

