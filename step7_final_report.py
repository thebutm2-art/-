"""
S7. 최종 보고서 생성
- '최종보고양식.xlsx' 템플릿을 로드해 서식·수식을 보존한 채 금액만 채운다.
- 채널(배민/쿠팡)·월별 컬럼에 금액 기입 → 비율·공과금·순수익은 템플릿 수식이 자동 계산.

양식 구조:
  매장명 = C2 (병합 C2:D3)
  월 → 금액컬럼: {1:C,2:E,3:G,4:I,5:K,6:M}  (비율은 그 다음 열)
  배민 테이블 행:  매출8 식자재9 부자재10 중개11 우가클12 카드13 부가세14 배달비15 프로모션16 공과금17 순수익18
  쿠팡 테이블 행:  매출23 식자재24 부자재25 중개26 맞춤27 카드28 부가세29 배달비30 프로모션31 공과금32 순수익33
"""
from pathlib import Path
import shutil
import openpyxl
import config

TEMPLATE = config.BASE_DIR / "최종보고양식.xlsx"

# 월 → (금액 컬럼, 비율 컬럼)
MONTH_COLS = {1:("C","D"),2:("E","F"),3:("G","H"),4:("I","J"),5:("K","L"),6:("M","N")}

# 공과금 비율 (매출 대비 고정)
UTILITY_RATIO = 0.05

# 채널별 항목 행번호
ROWS_BAEMIN = {
    "sales":8, "sik":9, "bu":10, "ad_brokerage":11, "ad_click":12,
    "card_fee":13, "vat":14, "delivery":15, "promotion":16,
    "utility":17, "net":18,
}
ROWS_COUPANG = {
    "sales":23, "sik":24, "bu":25, "ad_brokerage":26, "ad_custom":27,
    "card_fee":28, "vat":29, "delivery":30, "promotion":31,
    "utility":32, "net":33,
}


def fill_report(store_name: str, month: int, channel: str,
                pl: dict, out_path: Path | None = None,
                template: Path | None = None,
                cost_only: bool = False) -> Path:
    """
    store_name : 매장명 (예: 방이점)
    month      : 1~6
    channel    : 'baemin' | 'coupang'
    pl         : {
        'sales','sik','bu','ad_brokerage','ad_click'(배민)/'ad_custom'(쿠팡),
        'card_fee','vat','delivery','promotion'
    }  (공과금·순수익은 수식 자동)
    out_path   : 저장 경로 (없으면 output/{매장}_손익보고서_{YYYY}-{MM}.xlsx 가 아닌 템플릿 복제)
    """
    template = template or TEMPLATE
    if out_path is None:
        out_path = config.OUTPUT_DIR / f"{store_name}_손익보고서.xlsx"

    # 기존 산출물이 있으면 누적, 없으면 템플릿 복제 후 시작
    src = out_path if out_path.exists() else template
    wb = openpyxl.load_workbook(src)
    ws = wb["최종 보고서"]

    # 매장명
    ws["C2"] = store_name

    amt_col, _ = MONTH_COLS[month]
    rows = ROWS_BAEMIN if channel == "baemin" else ROWS_COUPANG

    def put(key, value):
        ws[f"{amt_col}{rows[key]}"] = value

    # 매출·식자재·부자재 (배민·쿠팡 공통)
    put("sales", pl.get("sales", 0))
    put("sik",   pl.get("sik", 0))
    put("bu",    pl.get("bu", 0))

    if cost_only:
        # 쿠팡 등 정산서 없는 채널: 매출·식자재·부자재만 채움.
        # 수수료 데이터가 없어 순수익이 오산되므로 공과금·순수익 셀은 비운다.
        ws[f"{amt_col}{rows['utility']}"] = None
        ws[f"{amt_col}{rows['net']}"]     = None
        wb.save(out_path)
        return out_path

    put("ad_brokerage", pl.get("ad_brokerage", 0))
    if channel == "baemin":
        put("ad_click",  pl.get("ad_click", 0))
    else:
        put("ad_custom", pl.get("ad_custom", 0))
    put("card_fee",  pl.get("card_fee", 0))
    put("vat",       pl.get("vat", 0))
    put("delivery",  pl.get("delivery", 0))
    put("promotion", pl.get("promotion", 0))

    # 공과금: 매출 × 5% 고정 (비율셀 0.05로 설정), 금액 = 매출 × 5%
    util_ratio_col = MONTH_COLS[month][1]
    ws[f"{util_ratio_col}{rows['utility']}"] = UTILITY_RATIO
    ws[f"{amt_col}{rows['utility']}"] = f"=${amt_col}${rows['sales']}*{util_ratio_col}{rows['utility']}"
    # 순수익 = 매출 - SUM(식자재 ~ 공과금)  ← 공과금 포함 차감
    ws[f"{amt_col}{rows['net']}"]     = f"={amt_col}{rows['sales']}-SUM({amt_col}{rows['sik']}:{amt_col}{rows['utility']})"

    wb.save(out_path)
    return out_path
