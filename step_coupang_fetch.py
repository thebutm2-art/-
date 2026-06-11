"""
쿠팡이츠 매출내역서 자동 다운로드 (봇차단 우회 / 비헤드리스)

흐름: 로그인 → 매출 관리 → 매출내역서 다운로드 → 매장 선택 → 월별 선택(년/월) → 다운로드

store.coupangeats.com 은 자동화 탐지가 있어 browser_util.stealth_browser(실제 Chrome,
AutomationControlled 해제, webdriver 스푸핑) 로 우회. 로그인 검증 완료.
"""
import asyncio
from pathlib import Path
from rich.console import Console
from playwright.async_api import TimeoutError as PWTimeout
from browser_util import stealth_browser, type_human
import config

console = Console(highlight=False)
LOGIN_URL = "https://store.coupangeats.com/merchant/login"


async def _login(page, uid, pw):
    await page.goto(LOGIN_URL, wait_until="domcontentloaded", timeout=40000)
    await page.wait_for_timeout(3500)
    ins = page.locator("input")
    if await ins.count() < 2:
        raise RuntimeError("쿠팡 로그인 폼을 찾지 못함(봇차단/속도제한 가능)")
    await type_human(page, ins.nth(0), uid)
    await type_human(page, ins.nth(1), pw)
    await page.get_by_role("button", name="로그인").click()
    # 네비게이션 느릴 수 있어 URL 폴링 (최대 30초)
    for _ in range(30):
        await page.wait_for_timeout(1000)
        if "/management/" in page.url:
            break
    if "/management/" not in page.url:
        raise RuntimeError("로그인 후 대시보드 진입 실패(속도제한 의심 — 잠시 후 재시도)")
    await page.wait_for_timeout(2500)


async def _dismiss_popups(page):
    """스마트모드/광고 등 오버레이 팝업 닫기."""
    for _ in range(3):
        closed = False
        for label in ["일주일간 보지 않기", "오늘 하루 보지 않기", "다음에 하기", "닫기"]:
            try:
                el = page.get_by_text(label, exact=False).first
                if await el.count() and await el.is_visible():
                    await el.click(timeout=2000); await page.wait_for_timeout(400); closed = True
            except Exception:
                pass
        # 다이얼로그 내 X(아이콘) 버튼
        try:
            dlg = page.locator("div[role='dialog'], .MuiDialog-paper").last
            if await dlg.count() and await dlg.is_visible():
                xbtn = dlg.locator("button").first
                if await xbtn.count():
                    await xbtn.click(timeout=2000); await page.wait_for_timeout(400); closed = True
        except Exception:
            pass
        await page.keyboard.press("Escape"); await page.wait_for_timeout(300)
        if not closed:
            break


async def _open_download_modal(page):
    # 매출 관리로 이동 (홈 URL의 shopid 사용)
    url = page.url
    if "/management/" in url:
        base = url.split("/management/")[0]
        shopid = url.rstrip("/").split("/")[-1]
        await page.goto(f"{base}/management/orders/{shopid}",
                        wait_until="domcontentloaded", timeout=30000)
        await page.wait_for_timeout(1500)
    await _dismiss_popups(page)
    await page.get_by_role("button", name="매출내역서 다운로드").click()
    await page.get_by_text("매장 선택", exact=False).first.wait_for(timeout=10000)
    await page.wait_for_timeout(800)


async def _select_store(page, store_keyword: str):
    # 매장 선택 드롭다운 열기 (모달 내 '선택')
    dlg = page.locator(".MuiDialog-paper, div[role='dialog']").last
    await dlg.get_by_text("선택", exact=False).first.click()
    await page.wait_for_timeout(700)
    # 79대포 포함 옵션 우선, 없으면 매장 키워드, 없으면 첫 옵션
    opts = page.locator("li, [role='option'], [role='menuitem']")
    n = await opts.count()
    target = None
    for i in range(n):
        t = (await opts.nth(i).inner_text()).strip()
        if not t:
            continue
        if "79대포" in t or (store_keyword and store_keyword in t):
            target = opts.nth(i); break
    if target is None and n:
        # 주먹밥(2브랜드) 제외 첫 옵션
        for i in range(n):
            t = (await opts.nth(i).inner_text()).strip()
            if t and "주먹밥" not in t:
                target = opts.nth(i); break
    if target is None:
        raise RuntimeError("쿠팡 매장 옵션을 찾지 못함")
    await target.click()
    await page.wait_for_timeout(600)


async def _select_month(page, year: int, month: int):
    dlg = page.locator(".MuiDialog-paper, div[role='dialog']").last
    # 년 드롭다운
    await dlg.get_by_text("년", exact=True).first.click()
    await page.wait_for_timeout(500)
    await page.get_by_text(f"{year}년", exact=True).first.click()
    await page.wait_for_timeout(500)
    # 월 드롭다운
    await dlg.get_by_text("월", exact=True).first.click()
    await page.wait_for_timeout(500)
    await page.get_by_text(f"{month}월", exact=True).first.click()
    await page.wait_for_timeout(500)


async def fetch_coupang(store: dict, year: int, month: int,
                        dest: Path, headless: bool = False) -> Path | None:
    """쿠팡이츠 매출내역서 다운로드 → dest 저장. 실패 시 None."""
    uid = store.get("coupang_id"); pw = store.get("coupang_pw")
    if not uid or not pw:
        console.print(f"[yellow]  [{store['name']}] 쿠팡 계정 없음 — 생략[/yellow]")
        return None
    try:
        async with stealth_browser(headless=headless) as page:
            await _login(page, uid, pw)
            await _open_download_modal(page)
            await _select_store(page, store.get("name", ""))
            await _select_month(page, year, month)
            async with page.expect_download(timeout=30000) as dl:
                await page.get_by_role("button", name="다운로드").click()
            download = await dl.value
            dest.parent.mkdir(parents=True, exist_ok=True)
            await download.save_as(str(dest))
            console.print(f"[green]  [{store['name']}] 쿠팡 매출내역서 저장: {dest.name}[/green]")
            return dest
    except Exception as e:
        console.print(f"[red]  [{store['name']}] 쿠팡 다운로드 실패: {repr(e)[:120]}[/red]")
        return None


if __name__ == "__main__":
    import sys, io
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")
    store = {"name": "방이점", "coupang_id": "79dp0019", "coupang_pw": "thebut007+"}
    dest = config.DOWNLOAD_DIR / "coupang_방이점_2026-05.xlsx"
    asyncio.run(fetch_coupang(store, 2026, 5, dest))
