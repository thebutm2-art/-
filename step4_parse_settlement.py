"""
S4. 비밀번호 해제 / 데이터 파싱
- 배민 정산명세서 엑셀을 암호(생년월일 6자리 / 법인번호 뒷 7자리)로 해제
- '상세' 시트의 3단 헤더(그룹/중간/세부)를 해석해 손익 항목으로 집계
- 컬럼 위치가 아닌 '헤더명 키워드' 기준으로 매핑 → 배민 양식 변경 내성 확보

[실제 배민 정산명세서 구조 — 2026-06 검증]
  시트: '요약'(카테고리 순액), '상세'(거래단위 명세)
  상세 헤더 3행:
    R3 그룹: (A) 주문중개 / (B) 배달 / (C) 그외 / (D) 기타 / (E) 부가세 / (F) 우리가게클릭 / (G) 배민오더 / (H) 입금금액
    R4 중간: 주문금액 / 중개이용료 / 고객할인비용 / 가게배달팁 / 배민클럽 할인비용 / 배달비 / 결제정산수수료 / 만나서결제금액 ...
    R5 세부: 바로결제주문금액 / 배민1중개이용료 / 메뉴할인 / ...
  R6~ : 데이터
"""
import io
import warnings
from pathlib import Path
from rich.console import Console
import msoffcrypto
import openpyxl

warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")

console = Console(highlight=False)

# 손익 항목 키 (step5 입력과 동일)
NUMERIC_KEYS = {
    "total_sales", "brokerage_fee", "payment_fee",
    "delivery_fee", "discount_burden", "ad_cost",
    "vat", "expected_settlement",
}


def _decrypt(path: Path, password: str) -> io.BytesIO:
    """암호화 엑셀 → 복호화된 BytesIO. 암호화 안 됐으면 원본 그대로."""
    with open(path, "rb") as f:
        office = msoffcrypto.OfficeFile(f)
        if not office.is_encrypted():
            return io.BytesIO(path.read_bytes())
        office.load_key(password=password)
        out = io.BytesIO()
        office.decrypt(out)
        out.seek(0)
        return out


def _classify_column(group: str, mid: str) -> str | None:
    """그룹/중간 헤더명으로 컬럼을 손익 항목 키에 분류."""
    g = (group or "").replace(" ", "")
    m = (mid or "").replace(" ", "")

    if "주문중개" in g:
        if "주문금액" in m:        return "total_sales"
        if "부분환불" in m:        return "total_sales"   # 환불은 매출 차감(음수)
        if "중개이용료" in m:      return "brokerage_fee"
        if "할인" in m:            return "discount_burden"
    if "배달" in g and "주문중개" not in g:
        # (B) 배달 그룹 전체 = 점주 부담 배달비 (배달팁·배민클럽할인·배달비 모두 포함)
        return "delivery_fee"
    if "그외" in g:
        if "결제" in m and "수수료" in m:  return "payment_fee"
        if "만나서결제" in m:              return "manna_offset"  # 현금수령 상계 (손익 제외)
    if "기타" in g:                return "etc_income"
    if "부가세" in g:              return "vat"
    if "우리가게클릭" in g:        return "ad_cost"
    if "배민오더" in g:            return "baemin_order"
    if "입금금액" in g:            return "expected_settlement"
    return None


def _parse_baemin_detail(wb) -> dict:
    """배민 '상세' 시트를 손익 항목으로 집계."""
    ws = wb["상세"] if "상세" in wb.sheetnames else wb.active

    # 헤더 행 자동 탐지: '(A) 주문중개'가 들어있는 행을 그룹 헤더로
    group_row = None
    for r in range(1, 8):
        for c in range(1, ws.max_column + 1):
            v = ws.cell(r, c).value
            if v and "주문중개" in str(v):
                group_row = r
                break
        if group_row:
            break
    if not group_row:
        raise ValueError("상세 시트에서 헤더(주문중개)를 찾지 못함")

    mid_row  = group_row + 1
    data_row = group_row + 3  # 그룹/중간/세부 3행 뒤부터 데이터

    # 컬럼별 그룹/중간 헤더 forward-fill
    col_class = {}
    cur_group = ""
    for c in range(1, ws.max_column + 1):
        g = ws.cell(group_row, c).value
        # 실제 그룹은 '(A)~(H)' 마커 형식만 인정 (예: '입금 금액' 소계열 오인 방지)
        if g and "(" in str(g) and ")" in str(g):
            cur_group = str(g)
        mid = ws.cell(mid_row, c).value
        # 중간 헤더도 그룹 내에서 forward-fill
        if mid:
            col_class[c] = (cur_group, str(mid))
        else:
            # 중간 헤더 없으면 직전 중간 헤더 유지(같은 그룹 내)
            prev = col_class.get(c - 1)
            if prev and prev[0] == cur_group:
                col_class[c] = (cur_group, prev[1])
            else:
                col_class[c] = (cur_group, "")

    result = {k: 0.0 for k in NUMERIC_KEYS}
    result["manna_offset"] = 0.0
    result["etc_income"]   = 0.0
    result["baemin_order"] = 0.0

    for c, (g, m) in col_class.items():
        key = _classify_column(g, m)
        if not key or key not in result:
            continue
        total = 0.0
        for r in range(data_row, ws.max_row + 1):
            v = ws.cell(r, c).value
            if isinstance(v, (int, float)):
                total += v
        result[key] += total

    # 비용 항목은 정산파일에서 음수 → 손익계산용 양수(절대값)로 변환
    for k in ("brokerage_fee", "payment_fee", "delivery_fee",
              "discount_burden", "ad_cost", "vat"):
        result[k] = abs(result[k])

    return result


def parse_settlement_file(path: Path, store: dict) -> dict | None:
    """점포 정산 첨부파일 파싱 → 손익 항목 dict. 실패 시 None."""
    password = store.get("file_pw") or store.get("biz_no", "")
    try:
        dec = _decrypt(path, password)
        wb  = openpyxl.load_workbook(dec, data_only=True)

        if "상세" in wb.sheetnames or any("주문중개" in str(wb.active.cell(r, c).value)
                                          for r in range(1, 8)
                                          for c in range(1, min(wb.active.max_column + 1, 40))):
            data = _parse_baemin_detail(wb)
        else:
            console.print(f"[yellow]  [{store['name']}] 배민 표준 양식 아님 — 확인 필요[/yellow]")
            return None

        data["store_code"] = store["code"]
        data["store_name"] = store["name"]
        console.print(
            f"[green]  [{store['name']}] 파싱 완료 — "
            f"총매출 {data['total_sales']:,.0f}원, "
            f"입금예정 {data.get('expected_settlement',0):,.0f}원[/green]"
        )
        return data

    except Exception as e:
        console.print(f"[red]  [{store['name']}] 파싱 실패 ({path.name}): {e}[/red]")
        return None


def parse_all(file_map: dict[str, list[Path]], stores: list[dict]) -> list[dict]:
    """전 점포 첨부 파싱. file_map: {점포코드: [Path,...]}"""
    store_idx = {s["code"]: s for s in stores}
    results   = []
    for code, paths in file_map.items():
        if not paths:
            console.print(f"[yellow]  [{code}] 첨부파일 없음 — 스킵[/yellow]")
            continue
        store = store_idx.get(code)
        if not store:
            continue
        for path in paths:
            data = parse_settlement_file(path, store)
            if data:
                results.append(data)
                break
    console.print(f"[bold green]파싱 완료: {len(results)}/{len(stores)}개 점포[/bold green]")
    return results
