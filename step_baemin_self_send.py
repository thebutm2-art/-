"""
배민셀프서비스 정산명세서 자동 발송 (봇차단 우회 / 비헤드리스)

흐름: 배민셀프 로그인 → 정산내역(/orders/billing) → 정산명세서 모달
      → 이메일을 수신계정으로 변경 → 메일 보내기

배민셀프 정산명세서는 '다운로드'가 아니라 '입력 이메일로 발송' 방식.
목적지 이메일을 thebut_m2@79daepo.com 으로 지정해 발송 → Gmail에서 수신.
기간은 기본 범위(데이터 존재)로 발송하고, 수신 단계에서 대상월 파일을 골라낸다.

biz-member.baemin.com 도 자동화 탐지가 있어 browser_util.stealth_browser 사용.
검증: 강릉교동점(79dp0225) 로그인→발송 성공.
"""
import asyncio
from rich.console import Console
from playwright.async_api import TimeoutError as PWTimeout
from browser_util import stealth_browser, type_human
import config

console = Console(highlight=False)
LOGIN_URL = "https://biz-member.baemin.com/login"
BILLING_URL = "https://self.baemin.com/orders/billing"


async def _dismiss_popups(page):
    for label in ["오늘 하루 보지 않기", "일주일간 보지 않기", "닫기", "다음에 하기"]:
        try:
            el = page.get_by_text(label, exact=False).first
            if await el.count() and await el.is_visible():
                await el.click(timeout=1500)
                await page.wait_for_timeout(300)
        except Exception:
            pass
    await page.keyboard.press("Escape")
    await page.wait_for_timeout(300)


async def _login(page, uid, pw):
    # 이미 로그인된 세션이면 스킵
    await page.goto("https://self.baemin.com/", wait_until="domcontentloaded", timeout=40000)
    await page.wait_for_timeout(2500)
    if "self.baemin.com" in page.url and "login" not in page.url:
        return

    await page.goto(LOGIN_URL, wait_until="domcontentloaded", timeout=40000)
    await page.wait_for_timeout(4500)  # 폼 렌더 대기(봇차단 우회 후)

    # 이름 기반 셀렉터 우선, 없으면 인덱스
    id_box = page.locator("input[name='id'], input[type='text']").first
    pw_box = page.locator("input[name='password'], input[type='password']").first
    if not await id_box.count() or not await pw_box.count():
        ins = page.locator("input")
        if await ins.count() < 2:
            raise RuntimeError("배민셀프 로그인 폼을 찾지 못함(봇차단/속도제한 가능)")
        id_box, pw_box = ins.nth(0), ins.nth(1)

    await type_human(page, id_box, uid)
    await type_human(page, pw_box, pw)

    # 로그인 클릭 (최대 2회 시도)
    for click_try in range(2):
        try:
            await page.get_by_role("button", name="로그인").first.click(timeout=5000)
        except Exception:
            await pw_box.press("Enter")
        for _ in range(25):
            await page.wait_for_timeout(1000)
            if "self.baemin.com" in page.url:
                await page.wait_for_timeout(2000)
                return
        # 아직이면 한 번 더 시도
    raise RuntimeError("배민셀프 로그인 후 진입 실패(2단계 인증/캡차 가능 — 창에서 직접 로그인 필요)")


async def send_settlement_mail(store: dict, year: int, month: int,
                               dest_email: str, headless: bool = False) -> bool:
    """배민셀프에서 정산명세서를 dest_email로 발송. 성공 시 True."""
    uid = store.get("baemin_id"); pw = store.get("baemin_pw")
    if not uid or not pw:
        console.print(f"[yellow]  [{store['name']}] 배민 계정 없음 — 발송 생략[/yellow]")
        return False
    try:
        async with stealth_browser(headless=headless) as page:
            await _login(page, uid, pw)
            await page.goto(BILLING_URL, wait_until="domcontentloaded", timeout=30000)
            await page.wait_for_timeout(2000)
            await _dismiss_popups(page)

            # 정산명세서 모달 열기
            await page.get_by_role("button", name="정산명세서").first.click()
            await page.get_by_text("이메일로 보내드려요", exact=False).wait_for(timeout=10000)
            await page.wait_for_timeout(800)

            # 이메일 입력칸 → 수신계정으로 변경 (기간은 기본범위=데이터 존재)
            email_box = page.get_by_placeholder("baemin@baemin.com").last
            await email_box.click()
            await page.keyboard.press("Control+a")
            await email_box.type(dest_email, delay=40)

            # 메일 보내기
            await page.get_by_role("button", name="메일 보내기").click()
            await page.wait_for_timeout(2500)
            console.print(f"[green]  [{store['name']}] 정산명세서 발송 → {dest_email}[/green]")
            return True
    except Exception as e:
        console.print(f"[red]  [{store['name']}] 배민셀프 발송 실패: {repr(e)[:120]}[/red]")
        return False


if __name__ == "__main__":
    import sys, io
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")
    store = {"name": "강릉교동점", "baemin_id": "79dp0225", "baemin_pw": "sh4fkdgo@@"}
    asyncio.run(send_settlement_mail(store, 2026, 5, "thebut_m2@79daepo.com"))
