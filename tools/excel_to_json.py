#!/usr/bin/env python3
"""
엑셀(xlsx) → data/hydrants.json 변환 스크립트

사용법:
  pip install openpyxl
  python tools/excel_to_json.py 파일명.xlsx

컬럼 매핑 (0-indexed, 기존 Google Sheets 구조 기준):
  B(1) = 명칭(name)
  D(3) = 주소(addr)
  F(5) = 위도(lat)
  G(6) = 경도(lng)
  H(7) = 상세위치(desc)
  J(9) = 용수구역(area)

컬럼 위치가 다르면 아래 COL_* 상수를 수정하세요.
"""
import json
import sys
from pathlib import Path

try:
    import openpyxl
except ImportError:
    print("openpyxl 패키지가 필요합니다: pip install openpyxl")
    sys.exit(1)

COL_NAME = 1   # B열 (0-indexed)
COL_ADDR = 3   # D열
COL_LAT  = 5   # F열
COL_LNG  = 6   # G열
COL_DESC = 7   # H열
COL_AREA = 9   # J열

def convert(xlsx_path: str) -> None:
    wb = openpyxl.load_workbook(xlsx_path, read_only=True, data_only=True)
    ws = wb.active
    rows = list(ws.iter_rows(values_only=True))
    wb.close()

    data = []
    skipped = 0
    for i, row in enumerate(rows[1:], start=2):  # 1행은 헤더
        def cell(idx):
            try:
                return row[idx]
            except IndexError:
                return None

        try:
            lat = float(cell(COL_LAT))
            lng = float(cell(COL_LNG))
        except (TypeError, ValueError):
            skipped += 1
            continue

        data.append({
            "name": str(cell(COL_NAME) or "정보없음").strip(),
            "addr": str(cell(COL_ADDR) or "정보없음").strip(),
            "desc": str(cell(COL_DESC) or "").strip(),
            "area": str(cell(COL_AREA) or "").strip(),
            "lat":  lat,
            "lng":  lng,
        })

    out_path = Path(__file__).parent.parent / "data" / "hydrants.json"
    out_path.parent.mkdir(parents=True, exist_ok=True)
    with open(out_path, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)

    print(f"변환 완료: {len(data)}개 저장 → {out_path}")
    if skipped:
        print(f"  (위도/경도 없는 행 {skipped}개 제외)")

if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("사용법: python tools/excel_to_json.py 파일명.xlsx")
        sys.exit(1)
    convert(sys.argv[1])
