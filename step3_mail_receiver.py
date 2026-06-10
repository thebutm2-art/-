"""
S3. 메일 자동 수신 / 첨부 추출
- IMAP 연결 → 배민 정산 메일 식별 → 첨부파일 저장
"""
import imaplib
import email
import email.header
import time
from pathlib import Path
from rich.console import Console
import config

console = Console()


def _decode_header(h: str) -> str:
    parts = email.header.decode_header(h)
    result = ""
    for part, enc in parts:
        if isinstance(part, bytes):
            result += part.decode(enc or "utf-8", errors="replace")
        else:
            result += part
    return result


def _is_baemin_settlement(msg, store: dict) -> bool:
    """배민 정산 메일 여부 확인 (발신자·제목 기준)"""
    sender  = _decode_header(msg.get("From", ""))
    subject = _decode_header(msg.get("Subject", ""))
    return (config.BAEMIN_SENDER_KEYWORD in sender and
            config.BAEMIN_SUBJECT_KEYWORD in subject)


def _save_attachments(msg, store: dict, save_dir: Path) -> list[Path]:
    """첨부파일을 점포코드_파일명으로 저장, 경로 리스트 반환"""
    saved = []
    for part in msg.walk():
        disposition = part.get("Content-Disposition", "")
        if "attachment" not in disposition:
            continue
        filename = part.get_filename()
        if not filename:
            continue
        filename = _decode_header(filename)
        # 점포코드 prefix로 저장 (겹침 방지)
        dest = save_dir / f"{store['code']}_{filename}"
        dest.write_bytes(part.get_payload(decode=True))
        saved.append(dest)
        console.print(f"    저장: {dest.name}")
    return saved


def fetch_settlement_mails(
    stores: list[dict],
    save_dir: Path | None = None,
    poll_interval: int | None = None,
    max_wait: int | None = None,
) -> dict[str, list[Path]]:
    """
    모든 점포 대상 정산 첨부파일 수집.
    Returns: {점포코드: [첨부파일 경로, ...]}
    """
    save_dir     = save_dir     or config.DOWNLOAD_DIR
    poll_interval = poll_interval or config.MAIL_POLL_INTERVAL_SEC
    max_wait     = max_wait     or config.MAIL_POLL_MAX_WAIT_SEC

    pending   = {s["code"]: s for s in stores}  # 아직 수신 안 된 점포
    collected: dict[str, list[Path]] = {s["code"]: [] for s in stores}

    deadline = time.time() + max_wait

    console.print(f"\n[bold]메일 수신 대기 (최대 {max_wait}초, {poll_interval}초 간격)...[/bold]")

    imap = imaplib.IMAP4_SSL(config.MAIL_HOST, config.MAIL_PORT)
    imap.login(config.MAIL_USER, config.MAIL_PASS)

    try:
        while pending and time.time() < deadline:
            imap.select("INBOX")
            _, data = imap.search(None, "UNSEEN")
            uids = data[0].split()

            for uid in uids:
                _, raw = imap.fetch(uid, "(RFC822)")
                msg = email.message_from_bytes(raw[0][1])

                for code, store in list(pending.items()):
                    if _is_baemin_settlement(msg, store):
                        console.print(f"[green]  [{store['name']}] 정산 메일 수신[/green]")
                        paths = _save_attachments(msg, store, save_dir)
                        if paths:
                            collected[code].extend(paths)
                            del pending[code]
                            break  # 다음 메일로

            if pending:
                remaining = int(deadline - time.time())
                console.print(f"  대기 중: {list(pending.keys())} (남은 시간 {remaining}초)")
                time.sleep(poll_interval)

    finally:
        imap.logout()

    if pending:
        console.print(f"[yellow]  타임아웃 — 미수신 점포: {list(pending.keys())}[/yellow]")

    ok_count = sum(1 for v in collected.values() if v)
    console.print(f"[bold green]메일 수신 완료: {ok_count}/{len(stores)}개 점포[/bold green]")
    return collected


if __name__ == "__main__":
    from step1_load_master import load_master
    stores = load_master()
    result = fetch_settlement_mails(stores)
    for code, paths in result.items():
        console.print(f"  {code}: {[p.name for p in paths]}")
