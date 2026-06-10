"""
S2. 배민 셀프서비스 자동 로그인 / 정산 요청
- Playwright 기반 웹 자동화
- 주의: 캡차·2단계 인증 발생 시 수동 개입 필요 (section 11.1 참조)
"""
import asyncio
import time
from datetime import datetime, date
from dateutil.relativedelta import relativedelta
from pathlib import Path
from rich.console import Console
from playwright.async_api import async_playwright, Page, TimeoutError as PWTimeout
import config

console = Console()

SELFSERVICE_URL = config.BAEMIN_SELFSERVICE_URL


async def _login(page: Page, store: dict) -> bool:
    """로그인 시도. 성공 시 True 반환."""
    try:
        await page.goto(SELFSERVICE_URL, wait_until="networkidle", timeout=30000)
        await page.fill("input[name='loginId'], input[type='email']", store["uid"])
        await page.fill("input[name='password'], input[type='password']", store["pw"])
        await page.click("button[type='submit']")
        await page.wait_for_load_state("networkidle", timeout=15000)

        # 캡차·2단계 인증 감지
        if await page.locator("text=캡차").count() or await page.locator("text=인증번호").count():
            console.print(f"[yellow]  [{store['name']}] 캡차/2FA 감지 — 수동 처리 대기 (최대 60초)[/yellow]")
            await page.wait_for_url(f"{SELFSERVICE_URL}/**", timeout=60000)

        # 로그인 실패 메시지 확인
        if await page.locator("text=아이디 또는 비밀번호").count():
            console.print(f"[red]  [{store['name']}] 로그인 실패: ID/PW 오류[/red]")
            return False

        console.print(f"[green]  [{store['name']}] 로그인 성공[/green]")
        return True
    except PWTimeout:
        console.print(f"[red]  [{store['name']}] 로그인 타임아웃[/red]")
        return False


async def _request_settlement(page: Page, store: dict, target_month: str) -> bool:
    """
    정산 메일 발송 요청.
    target_month: 'YYYY-MM' 형식
    """
    try:
        year, month = target_month.split("-")
        start_date = date(int(year), int(month), 1)
        end_date   = (start_date + relativedelta(months=1)) - relativedelta(days=1)

        # 정산 내역 페이지 이동 (실제 URL은 배민 정책에 따라 변경될 수 있음)
        await page.goto(f"{SELFSERVICE_URL}/settlement/history", wait_until="networkidle", timeout=20000)

        # 날짜 선택 (배민 셀프서비스 UI 기준)
        await page.fill("input[placeholder='시작일'], input[name='startDate']",
                        start_date.strftime("%Y.%m.%d"))
        await page.fill("input[placeholder='종료일'], input[name='endDate']",
                        end_date.strftime("%Y.%m.%d"))
        await page.click("button:has-text('조회'), button:has-text('검색')")
        await page.wait_for_load_state("networkidle", timeout=15000)

        # 메일 발송(다운로드 요청) 버튼 클릭
        btn = page.locator("button:has-text('메일'), button:has-text('이메일'), button:has-text('다운로드')")
        if not await btn.count():
            console.print(f"[yellow]  [{store['name']}] 정산 요청 버튼을 찾을 수 없음 — 수동 확인 필요[/yellow]")
            return False

        await btn.first.click()
        await page.wait_for_selector("text=발송, text=완료", timeout=10000)
        console.print(f"[green]  [{store['name']}] 정산 메일 발송 요청 완료 ({target_month})[/green]")
        return True

    except PWTimeout:
        console.print(f"[red]  [{store['name']}] 정산 요청 타임아웃[/red]")
        return False
    except Exception as e:
        console.print(f"[red]  [{store['name']}] 정산 요청 오류: {e}[/red]")
        return False


async def run_settlement_requests(stores: list[dict], target_month: str) -> list[str]:
    """
    전 점포 대상 정산 메일 발송 요청.
    성공한 점포코드 리스트 반환.
    """
    succeeded = []
    async with async_playwright() as pw:
        browser = await pw.chromium.launch(headless=False)  # 캡차 대응용 headless=False
        for store in stores:
            console.print(f"\n[bold]▶ [{store['name']}] 처리 중...[/bold]")
            context = await browser.new_context()
            page    = await context.new_page()
            try:
                ok = await _login(page, store)
                if ok:
                    ok2 = await _request_settlement(page, store, target_month)
                    if ok2:
                        succeeded.append(store["code"])
            finally:
                await context.close()
            await asyncio.sleep(2)  # 서버 부하 방지
        await browser.close()

    console.print(f"\n[bold green]정산 요청 완료: {len(succeeded)}/{len(stores)}개 점포[/bold green]")
    return succeeded


if __name__ == "__main__":
    from step1_load_master import load_master
    stores = load_master()
    month  = config.TARGET_MONTH or input("정산 대상 월 입력 (YYYY-MM): ").strip()
    asyncio.run(run_settlement_requests(stores, month))
