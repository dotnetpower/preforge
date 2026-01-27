#!/usr/bin/env python3
"""
HTML to PPTX 변환 디버깅 스크립트
"""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent.parent / "src"))

from bs4 import BeautifulSoup

html_file = Path("private/07_타겟_converted.html")

with open(html_file, 'r', encoding='utf-8') as f:
    html_content = f.read()

soup = BeautifulSoup(html_content, 'lxml')

# 모든 테이블 찾기
tables = soup.find_all('table', class_='data-table')

print(f"전체 테이블 수: {len(tables)}")
print()

for idx, table in enumerate(tables, 1):
    thead = table.find('thead')
    tbody = table.find('tbody')
    
    header_rows = len(thead.find_all('tr')) if thead else 0
    body_rows = len(tbody.find_all('tr')) if tbody else len(table.find_all('tr'))
    
    total_rows = header_rows + body_rows
    
    print(f"테이블 {idx}:")
    print(f"  헤더 행: {header_rows}")
    print(f"  바디 행: {body_rows}")
    print(f"  전체 행: {total_rows}")
    
    if body_rows > 12:
        print(f"  ⚠️  분할 필요! ({body_rows}행 > 12행)")
    
    print()
