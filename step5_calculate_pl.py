"""
S5. 손익계산서 산출
작업지시서 6장 산식 구현

[산식]
(+) 총매출액
(-) 중개수수료
(-) 결제정산수수료
(-) 배달비(점주부담)
(-) 할인·쿠폰 부담금
(-) 광고비
(=) 배달채널 기여이익
(-) 매출원가 = 총매출 × 원가율           ← 마스터 입력
(-) 고정비(인건비·임차료 등)              ← 마스터 입력 (월 기준 기간 안분)
(=) 영업이익
"""
from rich.console import Console

console = Console()


def calculate(raw: dict, store: dict, target_month: str) -> dict:
    """
    raw      : step4 파싱 결과 dict
    store    : 점포 마스터 dict (cost_rate, fixed_cost 포함)
    target_month: 'YYYY-MM'

    Returns  : 손익 항목 전체 dict
    """
    total_sales      = raw.get("total_sales",       0.0)
    brokerage_fee    = raw.get("brokerage_fee",     0.0)
    payment_fee      = raw.get("payment_fee",       0.0)
    delivery_fee     = raw.get("delivery_fee",      0.0)
    discount_burden  = raw.get("discount_burden",   0.0)
    ad_cost          = raw.get("ad_cost",           0.0)
    vat              = raw.get("vat",               0.0)

    # 기여이익 = 총매출 - 플랫폼 비용 합계
    platform_cost       = brokerage_fee + payment_fee + delivery_fee + discount_burden + ad_cost
    contribution_profit = total_sales - platform_cost

    # 매출원가 = 총매출 × 원가율
    cost_rate    = store.get("cost_rate") or 0.0
    cogs         = total_sales * cost_rate

    # 고정비: 월 고정비를 정산 대상 월 기준 그대로 사용
    # (기간이 월 단위가 아닌 경우 안분 로직 확장 가능)
    fixed_cost = store.get("fixed_cost") or 0.0

    # 영업이익
    operating_profit = contribution_profit - cogs - fixed_cost

    result = {
        # 식별
        "store_code":           store["code"],
        "store_name":           store["name"],
        "target_month":         target_month,
        # 수익
        "total_sales":          total_sales,
        # 비용 (플랫폼)
        "brokerage_fee":        brokerage_fee,
        "payment_fee":          payment_fee,
        "delivery_fee":         delivery_fee,
        "discount_burden":      discount_burden,
        "ad_cost":              ad_cost,
        "platform_cost_total":  platform_cost,
        # 중간이익
        "contribution_profit":  contribution_profit,
        # 비용 (원가·고정)
        "cogs":                 cogs,
        "fixed_cost":           fixed_cost,
        # 최종
        "operating_profit":     operating_profit,
        # 참고
        "vat":                  vat,
        "cost_rate":            cost_rate,
        "operating_margin_pct": (operating_profit / total_sales * 100) if total_sales else 0.0,
    }

    sign = "+" if operating_profit >= 0 else "-"
    console.print(
        f"  [{store['name']}] 매출 {total_sales:,.0f}원 | "
        f"기여이익 {contribution_profit:,.0f}원 | "
        f"영업이익 [{sign}]{abs(operating_profit):,.0f}원 "
        f"({result['operating_margin_pct']:.1f}%)"
    )
    return result


def calculate_all(parsed_list: list[dict], stores: list[dict], target_month: str) -> list[dict]:
    """전 점포 손익 계산"""
    store_idx = {s["code"]: s for s in stores}
    results   = []
    console.print("\n[bold]--- 손익 계산 ---[/bold]")
    for raw in parsed_list:
        code  = raw.get("store_code")
        store = store_idx.get(code)
        if not store:
            continue
        pl = calculate(raw, store, target_month)
        results.append(pl)

    console.print(f"[bold green]손익 계산 완료: {len(results)}개 점포[/bold green]")
    return results
