#!/usr/bin/env python3
"""
엑셀(xlsx) → data/hydrants.json 변환 스크립트

사용법:
  pip install openpyxl
  python tools/excel_to_json.py 소방용수현황.xlsx

컬럼 매핑 (소방용수시설 세부 현황 파일 기준):
  B(1) = 시설번호(name)
  D(3) = 소재지도로명주소(addr)
  F(5) = 위도(lat)
  G(6) = 경도(lng)
  H(7) = 상세위치(desc)
  I(8) = 안전센터명(center)
  J(9) = 응수구역(area)

헤더: 1행(제목), 2행(빈행), 3행(컬럼명) → 데이터는 4행부터
"""
import json
import sys
from pathlib import Path

try:
    import openpyxl
except ImportError:
    print("openpyxl 패키지가 필요합니다: pip install openpyxl")
    sys.exit(1)

COL_NAME   = 1   # B열 (0-indexed)
COL_ADDR   = 3   # D열
COL_LAT    = 5   # F열
COL_LNG    = 6   # G열
COL_DESC   = 7   # H열
COL_CENTER = 8   # I열
COL_AREA   = 9   # J열

HEADER_ROWS = 3  # 1행(제목) + 2행(빈행) + 3행(컬럼명)

def convert(xlsx_path: str) -> None:
    wb = openpyxl.load_workbook(xlsx_path, read_only=True, data_only=True)
    ws = wb.active
    rows = list(ws.iter_rows(values_only=True))
    wb.close()

    data = []
    skipped = 0
    for i, row in enumerate(rows[HEADER_ROWS:], start=HEADER_ROWS + 1):
        def cell(idx, r=row):
            try:
                return r[idx]
            except IndexError:
                return None

        try:
            lat = float(cell(COL_LAT))
            lng = float(cell(COL_LNG))
        except (TypeError, ValueError):
            skipped += 1
            continue

        data.append({
            "name":   str(cell(COL_NAME)   or "정보없음").strip(),
            "addr":   str(cell(COL_ADDR)   or "정보없음").strip(),
            "desc":   str(cell(COL_DESC)   or "").strip(),
            "center": str(cell(COL_CENTER) or "").strip(),
            "area":   str(cell(COL_AREA)   or "").strip(),
            "lat":    lat,
            "lng":    lng,
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
