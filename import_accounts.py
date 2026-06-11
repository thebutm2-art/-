"""
'배민 손익 계정 정보' 엑셀 → 점포마스터_v2.xlsx 변환

원본: 매장별 계정정보 시트 (C=매장명, D=배민아이디, E=배민비밀번호, 데이터 7행~)
생성: 점포마스터_v2.xlsx
  점포명 | 토더매장명 | 배민파트너명 | 파일암호 | 배민아이디 | 배민비밀번호 | 쿠팡아이디 | 쿠팡비밀번호 | 가게배달건수
  - 토더매장명: 기본값 = 점포명 (다르면 수정)
  - 배민파트너명/파일암호: 정산서 매칭·해제용 (계정파일에 없음 → 수기 입력 필요)
  - 가게배달건수: 월별로 배민셀프 가게통계에서 확인해 입력(선택)
"""
import sys, io
from pathlib import Path
import openpyxl

if sys.platform == "win32":
    try: sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")
    except Exception: pass

BASE = Path(__file__).parent
COLS = ["점포명","토더매장명","배민파트너명","파일암호",
        "배민아이디","배민비밀번호","쿠팡아이디","쿠팡비밀번호","가게배달건수"]


def import_accounts(src: Path, out: Path | None = None) -> Path:
    out = out or (BASE / "점포마스터_v2.xlsx")
    wb = openpyxl.load_workbook(src, data_only=True)
    ws = wb["매장별 계정정보"] if "매장별 계정정보" in wb.sheetnames else wb.active

    # 헤더(매장명/아이디/비밀번호)가 있는 행 탐색
    hdr = None
    for r in range(1, 15):
        row = [str(ws.cell(r, c).value or "") for c in range(1, ws.max_column + 1)]
        if any("매장명" in x for x in row) and any("아이디" in x for x in row):
            hdr = r; break
    if not hdr:
        raise ValueError("계정정보 헤더(매장명/아이디)를 찾지 못함")

    # 매장명/아이디/비번 열 위치
    def col_of(key):
        for c in range(1, ws.max_column + 1):
            if key in str(ws.cell(hdr, c).value or ""):
                return c
        return None
    c_name, c_id, c_pw = col_of("매장명"), col_of("아이디"), col_of("비밀번호")

    rows = []
    for r in range(hdr + 1, ws.max_row + 1):
        name = ws.cell(r, c_name).value
        if not name or not str(name).strip():
            continue
        rows.append({
            "점포명": str(name).strip(),
            "토더매장명": str(name).strip(),   # 기본값=점포명, 다르면 수정
            "배민파트너명": "",
            "파일암호": "",
            "배민아이디": str(ws.cell(r, c_id).value or "").strip() if c_id else "",
            "배민비밀번호": str(ws.cell(r, c_pw).value or "").strip() if c_pw else "",
            "쿠팡아이디": "",
            "쿠팡비밀번호": "",
            "가게배달건수": "",
        })

    nwb = openpyxl.Workbook(); nws = nwb.active; nws.title = "점포마스터"
    nws.append(COLS)
    for row in rows:
        nws.append([row[c] for c in COLS])
    nwb.save(out)
    print(f"[OK] 점포마스터 생성: {out}  ({len(rows)}개 매장)")
    print("  ※ 손익계산에 필요한 '파일암호'(생년월일6/법인번호뒤7), '배민파트너명'은 매장별로 채워주세요.")
    return out


if __name__ == "__main__":
    src = Path(sys.argv[1]) if len(sys.argv) > 1 else BASE / "계정정보_원본.xlsx"
    import_accounts(src)
