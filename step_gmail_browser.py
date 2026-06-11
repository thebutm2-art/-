"""
Gmail 정산서 수신 (브라우저 방식 — 앱 비밀번호 불필요)

워크스페이스 계정이 앱 비밀번호(IMAP)를 막은 경우, 로그인 세션이 유지되는
전용 브라우저 프로필(browser_util.stealth_persistent)로 Gmail 웹에 접속해
정산명세서 첨부(대상월)를 직접 다운로드한다.

첫 실행 때 Gmail 로그인 1회 필요(창에서) → 이후 세션 유지로 자동.
"""
import re
import asyncio
from pathlib import Path
from rich.console import Console
from browser_util import stealth_persistent
import config

console = Console(highlight=False)


async def _is_logged_in(page) -> bool:
    return "mail.google.com" in page.url and "accounts.google.com" not in page.url


async def fetch_settlement(store: dict, year: int, month: int, dest: Path,
                           ctx=None, page=None) -> Path | None:
    """대상월 정산명세서 첨부를 Gmail에서 다운로드 → dest. 실패/미로그인 시 None."""
    partner = store.get("partner")
    if not partner:
        console.print(f"[yellow]  [{store['name']}] 배민파트너명 없음 — Gmail 수신 생략[/yellow]")
        return None

    own = ctx is None
    if own:
        cm = stealth_persistent(headless=False)
        ctx, page = await cm.__aenter__()

    try:
        q = f"{partner}+정산명세서"
        await page.goto(f"https://mail.google.com/mail/u/0/#search/{q}",
                        wait_until="domcontentloaded", timeout=40000)
        await page.wait_for_timeout(3500)
        if not await _is_logged_in(page):
            console.print("[red]  Gmail 로그인이 필요합니다. 열린 창에서 한 번 로그인해 주세요.[/red]")
            return None

        # 검색 결과 첫 메일 열기
        try:
            await page.locator("tr.zA").first.click(timeout=8000)
        except Exception:
            console.print(f"[yellow]  [{store['name']}] '{partner}' 정산명세서 메일 없음[/yellow]")
            return None
        await page.wait_for_timeout(2500)

        # 대상월 첨부 다운로드 버튼 (aria-label에 'YYYY년 M월 ... 정산명세서')
        pat = re.compile(rf"{year}년\s*{month}월.*정산명세서")
        btn = page.get_by_role("button", name=pat).first
        try:
            await btn.wait_for(timeout=8000)
        except Exception:
            console.print(f"[yellow]  [{store['name']}] {year}년 {month}월 첨부를 못 찾음[/yellow]")
            return None

        async with page.expect_download(timeout=30000) as dl:
            await btn.click()
        download = await dl.value
        dest.parent.mkdir(parents=True, exist_ok=True)
        await download.save_as(str(dest))
        console.print(f"[green]  [{store['name']}] Gmail 정산서 수신: {dest.name}[/green]")
        return dest
    finally:
        if own:
            await cm.__aexit__(None, None, None)


async def fetch_many(stores: list[dict], year: int, month: int,
                     save_dir: Path | None = None, max_wait_sec: int = 480) -> dict:
    """여러 매장 정산서를 한 브라우저 세션에서 폴링 수신. {점포명: 경로}.
    메일 도착까지 최대 max_wait_sec 동안 30초 간격 재시도.
    """
    save_dir = save_dir or config.DOWNLOAD_DIR
    ym = f"{year}-{month:02d}"
    targets = [s for s in stores if s.get("partner")]
    saved: dict[str, Path] = {}
    if not targets:
        return saved

    async with stealth_persistent(headless=False) as (ctx, page):
        # 로그인 확인 — 안 돼 있으면 열린 창에서 로그인할 때까지 대기(최대 5분)
        await page.goto("https://mail.google.com/mail/u/0/", wait_until="domcontentloaded", timeout=40000)
        await page.wait_for_timeout(3000)
        if not await _is_logged_in(page):
            console.print("[yellow]  Gmail 로그인이 필요합니다. 열린 창에서 로그인해 주세요 (대기 중, 최대 5분)...[/yellow]")
            for _ in range(60):  # 5초 × 60 = 5분
                await page.wait_for_timeout(5000)
                if await _is_logged_in(page):
                    console.print("[green]  Gmail 로그인 확인됨 — 계속 진행[/green]")
                    break
            if not await _is_logged_in(page):
                console.print("[red]  Gmail 로그인 시간 초과 — 다음에 다시 시도하세요.[/red]")
                return saved

        waited = 0
        while len(saved) < len(targets):
            for s in targets:
                if s["name"] in saved:
                    continue
                dest = save_dir / f"baemin_{s['toorder_name']}_{ym}.xlsx"
                r = await fetch_settlement(s, year, month, dest, ctx=ctx, page=page)
                if r:
                    saved[s["name"]] = r
            if len(saved) >= len(targets) or waited >= max_wait_sec:
                break
            await page.wait_for_timeout(30000)
            waited += 30
    return saved


if __name__ == "__main__":
    import sys, io
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")
    store = {"name": "강릉교동점", "partner": "장승환"}
    dest = config.DOWNLOAD_DIR / "baemin_강릉교동점_2026-05.xlsx"
    asyncio.run(fetch_settlement(store, 2026, 5, dest))
