"""
최종 보고서 오케스트레이터 (전 매장 확장 + 완전 자동화)

per (매장 × 월):
  1) 배민 정산 엑셀 확보  → 파싱(step4)             [Gmail 자동수신 or 로컬]
  2) 토더 배민 채널 추출   → 식부자재 매칭(cost_matcher)  [Playwright or 로컬]
  3) 토더 쿠팡 채널 추출   → 식부자재 매칭              [Playwright or 로컬]
  4) 최종보고양식 채우기(step7) -> 바탕화면/{매장}/{YYMM 매장}.xlsx

사용법:
  python run_report.py --month 2026-05                 # 전 매장, 로컬 파일 사용
  python run_report.py --month 2026-05 --store 방이점   # 특정 매장
  python run_report.py --month 2026-05 --fetch          # Gmail+토더 자동수집까지

로컬 파일 규칙 (--fetch 미사용 시 downloads/ 에서 탐색):
  배민정산 : baemin_{토더매장명}_{YYYY-MM}.xlsx   (또는 {코드}_정산명세서.xlsx)
  토더배민 : toorder_{매장}_baemin_{YYYY-MM}.xlsx
  토더쿠팡 : toorder_{매장}_coupang_{YYYY-MM}.xlsx
  식부자재 : 식부자재_{월}월.xlsx  (예: 식부자재_5월.xlsx)
"""
import argparse, asyncio, sys
from pathlib import Path
import openpyxl
from rich.console import Console
from rich.rule import Rule

import config
from store_master import load_stores
from step4_parse_settlement import parse_settlement_file
from cost_matcher import match_costs
from step7_final_report import fill_report

console = Console(highlight=False)


def _load_sikbu(month: int) -> list[dict]:
    p = config.BASE_DIR / f"식부자재_{month}월.xlsx"
    if not p.exists():
        raise FileNotFoundError(f"식부자재 파일 없음: {p}")
    wb = openpyxl.load_workbook(p, data_only=True); ws = wb.active
    return [{"name": ws.cell(r,3).value, "sik": ws.cell(r,4).value, "bu": ws.cell(r,5).value}
            for r in range(3, ws.max_row+1) if ws.cell(r,3).value]


def _load_toorder(p: Path) -> list[dict]:
    wb = openpyxl.load_workbook(p, data_only=True); ws = wb.active
    return [{"category": ws.cell(r,1).value, "name": ws.cell(r,2).value,
             "qty": ws.cell(r,3).value, "sales": ws.cell(r,4).value}
            for r in range(5, ws.max_row+1) if ws.cell(r,2).value]


def _find_baemin_file(store: dict, ym: str) -> Path | None:
    cands = [
        config.DOWNLOAD_DIR / f"baemin_{store['toorder_name']}_{ym}.xlsx",
        config.DOWNLOAD_DIR / "001_정산명세서.xlsx",  # 데모 호환
    ]
    return next((c for c in cands if c.exists()), None)


def _find_coupang_file(store: dict, ym: str) -> Path | None:
    """쿠팡이츠 정산내역서 탐색 (downloads/ 또는 다운로드 폴더)."""
    import os
    pats = [f"coupang_{store['toorder_name']}_{ym}.xlsx",
            f"쿠팡이츠_{store['name']}_{ym}.xlsx"]
    for p in pats:
        f = config.DOWNLOAD_DIR / p
        if f.exists():
            return f
    # 사용자 다운로드 폴더에서 쿠팡/매출내역 키워드 탐색
    dl = Path(os.path.expanduser("~")) / "Downloads"
    if dl.exists():
        hits = sorted(
            [f for f in dl.glob("*.xlsx")
             if any(k in f.name for k in ("쿠팡", "쿠팡이츠", "매출내역", "coupang"))],
            key=lambda f: f.stat().st_mtime, reverse=True)
        if hits:
            return hits[0]
    return None


async def _fetch_toorder(store: dict, year: int, month: int, channel: str, dest: Path):
    import os
    from step_toorder import export_store_channel
    headless = os.getenv("TOORDER_HEADLESS", "true").lower() != "false"
    await export_store_channel(store["toorder_name"], year, month, channel, dest,
                               headless=headless)


def process_store(store: dict, year: int, month: int, sik_bu: list[dict],
                  fetch: bool) -> Path | None:
    ym = f"{year}-{month:02d}"
    console.print(Rule(f"[bold]{store['name']} {ym}[/bold]"))

    # ── 배민 정산 ──────────────────────────────
    bm_file = _find_baemin_file(store, ym)
    if not bm_file:
        console.print(f"[yellow]  배민 정산 파일 없음 — {store['name']} 스킵[/yellow]")
        return None
    bm = parse_settlement_file(bm_file, {"name": store["name"], "code": store["name"],
                                         "file_pw": store["file_pw"], "biz_no": ""},
                               gagae_count=store.get("gagae_count"))
    if not bm:
        return None

    # ── 토더 배민/쿠팡 ─────────────────────────
    t_bm = config.DOWNLOAD_DIR / f"toorder_{store['name']}_baemin_{ym}.xlsx"
    t_cp = config.DOWNLOAD_DIR / f"toorder_{store['name']}_coupang_{ym}.xlsx"
    if fetch:
        asyncio.run(_fetch_toorder(store, year, month, "baemin", t_bm))
        asyncio.run(_fetch_toorder(store, year, month, "coupang", t_cp))
    # 데모 호환: 기존 방이점 파일명 fallback
    if not t_bm.exists():
        alt = config.BASE_DIR / "토더_방이점_배민_5월.xlsx"
        if alt.exists() and store["name"] == "방이점" and ym == "2026-05": t_bm = alt
    if not t_cp.exists():
        alt = config.BASE_DIR / "토더_방이점_쿠팡_5월.xlsx"
        if alt.exists() and store["name"] == "방이점" and ym == "2026-05": t_cp = alt

    out = Path.home() / "Desktop" / store["name"] / f"{year%100:02d}{month:02d} {store['name']}.xlsx"

    # 배민 섹션
    bm_cost = match_costs(_load_toorder(t_bm), sik_bu) if t_bm.exists() else {"total_sik":0,"total_bu":0}
    pl_bm = {
        "sales": bm["total_sales"], "sik": round(bm_cost["total_sik"]), "bu": round(bm_cost["total_bu"]),
        "ad_brokerage": bm["brokerage_fee"], "ad_click": bm["ad_cost"], "card_fee": bm["payment_fee"],
        "vat": bm["vat"], "delivery": bm["delivery_fee"], "promotion": bm["discount_burden"],
    }
    if out.exists(): out.unlink()
    fill_report(store["name"], month, "baemin", pl_bm, out_path=out)

    # 쿠팡 섹션: 식자재/부자재=토더, 매출·수수료=쿠팡이츠 정산내역서(있으면 정식 섹션)
    if t_cp.exists():
        cp = _load_toorder(t_cp); cp_cost = match_costs(cp, sik_bu)
        cp_pl = {"sik": round(cp_cost["total_sik"]), "bu": round(cp_cost["total_bu"])}
        cp_settle = _find_coupang_file(store, ym)
        if cp_settle:
            from step_coupang_parse import parse_coupang_settlement
            cs = parse_coupang_settlement(cp_settle)
            cp_pl.update({
                "sales": round(cs["sales"]), "ad_brokerage": round(cs["ad_brokerage"]),
                "ad_custom": round(cs["ad_custom"]), "card_fee": round(cs["card_fee"]),
                "vat": round(cs["vat"]), "delivery": round(cs["delivery"]),
            })
            fill_report(store["name"], month, "coupang", cp_pl, out_path=out)
        else:
            # 정산내역서 없으면 토더 매출 + 식자재/부자재만
            cp_sales = sum(x["sales"] for x in cp if isinstance(x["sales"], (int, float)))
            cp_pl["sales"] = round(cp_sales)
            console.print(f"[yellow]  쿠팡이츠 정산내역서 없음 — 매출·식자재·부자재만 기재[/yellow]")
            fill_report(store["name"], month, "coupang", cp_pl, out_path=out, cost_only=True)
    else:
        console.print(f"[yellow]  쿠팡 토더 파일 없음 — 쿠팡 섹션 생략[/yellow]")

    console.print(f"[bold green]  ✓ 저장: {out}[/bold green]")
    return out


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--month", required=True, help="YYYY-MM")
    ap.add_argument("--store", default=None, help="특정 매장명만")
    ap.add_argument("--fetch", action="store_true", help="Gmail+토더 자동수집")
    ap.add_argument("--master", default=None)
    ap.add_argument("--gagae", type=int, default=None,
                    help="가게배달 건수(배민셀프서비스 주문내역 기준). 단일매장 실행 시 적용")
    args = ap.parse_args()

    year, month = map(int, args.month.split("-"))
    stores = load_stores(Path(args.master) if args.master else None)
    if args.store:
        stores = [s for s in stores if s["name"] == args.store]
    if args.gagae is not None:
        for s in stores:
            s["gagae_count"] = args.gagae
    sik_bu = _load_sikbu(month)

    done, fail = [], []
    for s in stores:
        try:
            r = process_store(s, year, month, sik_bu, args.fetch)
            (done if r else fail).append(s["name"])
        except Exception as e:
            console.print(f"[red]  {s['name']} 오류: {e}[/red]")
            fail.append(s["name"])

    console.print(Rule("[bold green]완료[/bold green]"))
    console.print(f"  성공 {len(done)} / 실패 {len(fail)}")
    if fail: console.print(f"[yellow]  실패: {fail}[/yellow]")


if __name__ == "__main__":
    main()
