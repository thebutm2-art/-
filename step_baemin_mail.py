"""
배민 정산서 Gmail 자동 수신 (매장별 파트너명 + 대상월 매칭)

배민 정산 메일:
  발신자 : noreply@woowahan.com
  제목   : [배달의민족] {파트너명} {기간} 정산명세서
  첨부   : [배달의민족] {파트너명} {YYYY}년 {M}월 정산명세서.xlsx (암호화)

매장별로 '배민파트너명'과 '{YYYY}년 {M}월'이 모두 들어맞는 메일의 첨부를
  downloads/baemin_{토더매장명}_{YYYY-MM}.xlsx 로 저장한다 (run_report 탐색 규칙과 일치).

.env: MAIL_HOST, MAIL_PORT, MAIL_USER, MAIL_PASS(앱 비밀번호), BAEMIN_MAIL_SENDER
"""
import imaplib, email, email.header
from pathlib import Path
from rich.console import Console
import config

console = Console(highlight=False)


def _decode(h: str) -> str:
    out = ""
    for part, enc in email.header.decode_header(h or ""):
        out += part.decode(enc or "utf-8", "replace") if isinstance(part, bytes) else part
    return out


def _attach_xlsx(msg):
    """메일에서 첫 xlsx 첨부의 (파일명, 바이트) 반환."""
    for part in msg.walk():
        if "attachment" not in (part.get("Content-Disposition") or ""):
            continue
        fn = _decode(part.get_filename() or "")
        if fn.lower().endswith((".xlsx", ".xls")):
            return fn, part.get_payload(decode=True)
    return None, None


def fetch_all(stores: list[dict], year: int, month: int,
              save_dir: Path | None = None,
              mail_user: str | None = None,
              mail_pass: str | None = None) -> dict[str, Path]:
    """대상월 배민 정산서를 매장별로 수신·저장. Returns {점포명: 저장경로}.

    mail_pass: 실행 시 입력받은 Gmail 앱 비밀번호 (없으면 .env의 MAIL_PASS).
    """
    save_dir = save_dir or config.DOWNLOAD_DIR
    user = mail_user or config.MAIL_USER
    pw   = mail_pass or config.MAIL_PASS
    ym = f"{year}-{month:02d}"
    month_tag = f"{year}년 {month}월"      # 첨부 파일명 매칭 키
    sender = config.BAEMIN_SENDER_KEYWORD   # woowahan.com

    # 파트너명 있는 매장만 대상
    targets = {s["partner"]: s for s in stores if s.get("partner")}
    if not targets:
        console.print("[yellow]  배민파트너명이 입력된 매장이 없습니다 — 메일 수신 생략[/yellow]")
        return {}

    saved: dict[str, Path] = {}
    imap = imaplib.IMAP4_SSL(config.MAIL_HOST, config.MAIL_PORT)
    imap.login(user, pw)
    try:
        imap.select("INBOX")
        # 발신자 + 정산명세서 제목으로 1차 검색
        typ, data = imap.search(None, 'FROM', f'"{sender}"')
        uids = data[0].split()
        console.print(f"[cyan]  '{sender}' 메일 {len(uids)}건 스캔 ({month_tag})[/cyan]")

        for uid in reversed(uids):  # 최신부터
            _, raw = imap.fetch(uid, "(RFC822)")
            msg = email.message_from_bytes(raw[0][1])
            subj = _decode(msg.get("Subject", ""))
            if "정산명세서" not in subj:
                continue
            # 어떤 매장(파트너명)인지 + 대상월인지
            fn, payload = _attach_xlsx(msg)
            if not payload:
                continue
            for partner, store in targets.items():
                if store["name"] in saved:
                    continue
                if partner in subj and (month_tag in fn or month_tag in subj):
                    dest = save_dir / f"baemin_{store['toorder_name']}_{ym}.xlsx"
                    dest.write_bytes(payload)
                    saved[store["name"]] = dest
                    console.print(f"[green]  [{store['name']}] 정산서 수신: {dest.name}[/green]")
                    break
    finally:
        imap.logout()

    missing = [s["name"] for s in stores if s.get("partner") and s["name"] not in saved]
    if missing:
        console.print(f"[yellow]  정산서 미수신: {missing}[/yellow]")
    console.print(f"[bold green]배민 정산서 수신: {len(saved)}건[/bold green]")
    return saved


if __name__ == "__main__":
    import sys, io
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")
    from store_master import load_stores
    y, m = (2026, 5)
    if len(sys.argv) > 1:
        y, m = map(int, sys.argv[1].split("-"))
    fetch_all(load_stores(), y, m)
