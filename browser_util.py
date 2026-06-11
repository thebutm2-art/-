"""
봇차단 우회 브라우저 런처 (쿠팡이츠/배민셀프 등 자동화탐지 사이트용)

기법:
  - 실제 Chrome 채널 사용 (channel='chrome')
  - --disable-blink-features=AutomationControlled
  - navigator.webdriver 등 스푸핑 (add_init_script)
  - 사람처럼 타이핑(type_human)
검증: 쿠팡이츠 store.coupangeats.com 로그인 성공.
"""
from contextlib import asynccontextmanager
from playwright.async_api import async_playwright, Page

STEALTH_JS = """
Object.defineProperty(navigator,'webdriver',{get:()=>undefined});
Object.defineProperty(navigator,'languages',{get:()=>['ko-KR','ko']});
Object.defineProperty(navigator,'plugins',{get:()=>[1,2,3,4,5]});
window.chrome = window.chrome || {runtime:{}};
"""
UA = ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
      "(KHTML, like Gecko) Chrome/131.0.0.0 Safari/537.36")


@asynccontextmanager
async def stealth_browser(headless: bool = False):
    """안티탐지 브라우저 컨텍스트. headless=False 권장(탐지 회피)."""
    async with async_playwright() as pw:
        args = ["--disable-blink-features=AutomationControlled"]
        try:
            browser = await pw.chromium.launch(channel="chrome", headless=headless, args=args)
        except Exception:
            browser = await pw.chromium.launch(headless=headless, args=args)
        ctx = await browser.new_context(user_agent=UA, locale="ko-KR",
                                        viewport={"width": 1366, "height": 768},
                                        accept_downloads=True)
        await ctx.add_init_script(STEALTH_JS)
        page = await ctx.new_page()
        try:
            yield page
        finally:
            await browser.close()


async def type_human(page: Page, selector_or_locator, text: str, delay: int = 60):
    """사람처럼 천천히 입력."""
    loc = page.locator(selector_or_locator) if isinstance(selector_or_locator, str) else selector_or_locator
    await loc.click()
    await loc.fill("")
    await loc.type(text, delay=delay)
