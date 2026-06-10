"""
토더(ToOrder) 메뉴별 판매량 자동 추출 (Playwright)

흐름: 로그인 → 메뉴분석>메뉴별 판매량 → 매장/날짜/채널 설정 → Excel 내보내기

[중요 UI 특이점 — toorder-ui-quirks]
  - 검색(돋보기) 버튼은 채널 필터를 리셋한다. 채널 '적용하기' 후 검색 누르지 말고
    바로 '내보내기' 한다.
  - 매장/채널 모달은 헤더 체크박스로 전체 토글 후 원하는 항목만 체크.
  - 내보내기 Excel 헤더의 기간 표기는 부정확할 수 있으나 데이터는 필터대로 정확.

자격증명은 .env 에서 읽는다 (TOORDER_ID, TOORDER_PW). 평문 하드코딩 금지.
"""
import os
import asyncio
from pathlib import Path
from datetime import date
from dateutil.relativedelta import relativedelta
from rich.console import Console
from playwright.async_api import async_playwright, Page, TimeoutError as PWTimeout
import config

console = Console(highlight=False)

TOORDER_URL = "https://ceo.toorder.co.kr"
MENU_SALES_PATH = "/dashboard/product-analysis/product-sales-date"

# 채널 그룹 정의 (채널 상세명 기준)
CHANNELS = {
    "baemin":  ["배달의민족", "배민 포장", "배민1"],
    "coupang": ["쿠팡이츠", "쿠팡 포장"],
    "yogiyo":  ["요기요", "요기요 포장", "요기배달"],
}


async def login(page: Page):
    uid = os.getenv("TOORDER_ID", "")
    pw  = os.getenv("TOORDER_PW", "")
    await page.goto(TOORDER_URL, wait_until="networkidle", timeout=30000)

    # 이미 로그인되어 있으면 대시보드로 진입됨
    if "/dashboard" in page.url:
        console.print("[green]토더 이미 로그인됨[/green]")
        return

    # 로그인 폼: id / password / isCompany(본사 로그인 체크박스)
    await page.fill("input[name='id']", uid)
    await page.fill("input[name='password']", pw)
    # '본사 로그인' = isCompany 체크박스
    company = page.locator("input[name='isCompany']")
    if await company.count() and not await company.is_checked():
        await company.check()
    await page.get_by_role("button", name="로그인").click()
    await page.wait_for_url("**/dashboard/**", timeout=20000)
    console.print("[green]토더 로그인 성공[/green]")


async def goto_menu_sales(page: Page):
    # SPA 세션은 localStorage 기반 → 하드 네비게이션(goto) 시 로그인으로 튕김.
    # 반드시 앱 내 사이드바 클릭(클라이언트 라우팅)으로 이동한다.
    await page.wait_for_load_state("networkidle", timeout=20000)
    await page.get_by_text("메뉴 분석", exact=False).first.click()  # 부모 메뉴 펼치기
    await page.wait_for_timeout(500)
    await page.get_by_text("메뉴별 판매량", exact=False).first.click()
    await page.get_by_text("매장 선택", exact=False).first.wait_for(timeout=15000)
    await page.wait_for_timeout(1500)


def _dialog(page: Page):
    """현재 열린 MUI 다이얼로그(모달 페이퍼) 스코프."""
    return page.locator(".MuiDialog-paper").last


async def _open_modal(page: Page, trigger_text: str):
    await page.get_by_text(trigger_text, exact=False).first.click()
    dlg = _dialog(page)
    await dlg.get_by_role("button", name="적용하기").wait_for(timeout=5000)
    await page.wait_for_timeout(800)
    return dlg


async def _deselect_all(page: Page):
    """모달 헤더 '전체선택' 체크박스로 전체 해제 (체크박스 상태 기반, 표시문구 무관).

    MUI 전체선택: checked=전체 / indeterminate=일부 / unchecked=없음.
    1클릭하면 어떤 상태든 '전체선택(checked)'으로 가거나, 전체였다면 '없음'으로 간다.
    클릭 후 checked면 한 번 더 눌러 '없음'으로 만든다.
    """
    dlg = _dialog(page)
    header_cb = dlg.locator("input[type='checkbox']").first
    await header_cb.click()
    await page.wait_for_timeout(400)
    if await header_cb.is_checked():      # 전체선택 상태가 됨 → 다시 눌러 전체 해제
        await header_cb.click()
        await page.wait_for_timeout(400)


async def _check_rows_by_name(page: Page, names: list[str]):
    """검색창에 이름을 넣어 해당 행만 남긴 뒤 그 행 체크박스 클릭 (다이얼로그 스코프)."""
    dlg = _dialog(page)
    search = dlg.locator("input[type='search']").last
    for name in names:
        await search.fill(name)
        await page.wait_for_timeout(700)
        boxes = dlg.locator("input[type='checkbox']")
        cnt = await boxes.count()
        clicked = False
        for i in range(cnt):
            box = boxes.nth(i)
            row = box.locator("xpath=ancestor::*[self::tr or @role='row'][1]")
            try:
                txt = await row.inner_text(timeout=800)
            except Exception:
                continue
            if name in txt and not await box.is_checked():
                await box.click()
                clicked = True
                break
        if not clicked:
            console.print(f"[yellow]  '{name}' 행 못찾음 — 스킵[/yellow]")
    await search.fill("")


async def set_store(page: Page, store_name: str):
    dlg = await _open_modal(page, "매장 선택")
    await _deselect_all(page)
    await _check_rows_by_name(page, [store_name])
    await dlg.get_by_role("button", name="적용하기").click()
    await page.wait_for_timeout(1800)


async def set_channels(page: Page, channel_keys: list[str]):
    targets = [c for k in channel_keys for c in CHANNELS[k]]
    dlg = await _open_modal(page, "채널 선택")
    await _deselect_all(page)
    await _check_rows_by_name(page, targets)
    await dlg.get_by_role("button", name="적용하기").click()
    await page.wait_for_timeout(1800)  # 적용 후 자동 새로고침 (검색 X)


async def click_search(page: Page) -> bool:
    """우상단 검색(돋보기) 버튼 클릭 → 필터를 서버 쿼리에 커밋.
    주의: 검색은 채널 필터를 전체로 리셋하므로, 검색 후 채널을 마지막에 설정한다.
    aria-label이 없어 위치(우상단 SVG 버튼)로 식별한다.
    """
    vw = (page.viewport_size or {"width": 1568})["width"]
    btns = page.locator("button")
    for i in range(await btns.count()):
        try:
            bb = await btns.nth(i).bounding_box(timeout=300)
        except Exception:
            continue
        if bb and bb["x"] > vw - 120 and 40 < bb["y"] < 110:
            if not await btns.nth(i).get_attribute("aria-label"):
                await btns.nth(i).click()
                await page.wait_for_timeout(1500)
                return True
    console.print("[yellow]  검색 버튼을 찾지 못함[/yellow]")
    return False


async def set_month(page: Page, year: int, month: int):
    """대상 월 1일~말일. 날짜 입력칸(YY-MM-DD) 2개를 직접 채운다."""
    start = date(year, month, 1)
    end   = (start + relativedelta(months=1)) - relativedelta(days=1)
    fields = page.locator("input[placeholder='YY-MM-DD']")
    if await fields.count() >= 2:
        await fields.nth(0).fill(start.strftime("%y-%m-%d"))
        await fields.nth(1).fill(end.strftime("%y-%m-%d"))
        await page.keyboard.press("Enter")
        await page.wait_for_timeout(800)
        # 날짜 팝업(MuiPickersPopper)이 열려있으면 닫아 다음 클릭 차단 방지
        await page.keyboard.press("Escape")
        await page.mouse.click(5, 5)  # 빈 영역 클릭으로 팝퍼 닫기
        await page.wait_for_timeout(800)
    else:
        console.print("[yellow]  날짜 입력칸을 찾지 못함 — 기본 기간 사용[/yellow]")


async def export_excel(page: Page, dest: Path) -> Path:
    """테이블 '내보내기' > 'Excel로 내보내기' → 다운로드 저장."""
    await page.get_by_role("button", name="내보내기").first.click()
    await page.wait_for_timeout(500)
    async with page.expect_download(timeout=30000) as dl_info:
        await page.get_by_text("Excel로 내보내기").click()
    download = await dl_info.value
    dest.parent.mkdir(parents=True, exist_ok=True)
    await download.save_as(str(dest))
    console.print(f"[green]토더 내보내기 저장: {dest.name}[/green]")
    return dest


async def export_store_channel(store_name: str, year: int, month: int,
                               channel_key: str, dest: Path,
                               headless: bool = False) -> Path:
    """단일 매장·채널·월 토더 데이터 추출 (독립 실행용)."""
    async with async_playwright() as pw:
        browser = await pw.chromium.launch(headless=headless)
        ctx = await browser.new_context(accept_downloads=True)
        page = await ctx.new_page()
        try:
            await login(page)
            await goto_menu_sales(page)
            # 순서 중요: 매장·날짜 설정 → 검색(서버 커밋, 채널 리셋) → 채널 마지막 → 내보내기
            await set_store(page, store_name)
            await set_month(page, year, month)
            await click_search(page)            # 매장·날짜 커밋
            await set_channels(page, [channel_key])  # 채널은 검색 후 마지막에
            return await export_excel(page, dest)
        finally:
            await browser.close()


if __name__ == "__main__":
    # 예시: 방이점 2026-05 배민 채널
    dest = config.DOWNLOAD_DIR / "toorder_방이점_baemin_2026-05.xlsx"
    asyncio.run(export_store_channel("방이점", 2026, 5, "baemin", dest))
