"""공통 실행 로그 관리"""
from datetime import datetime
from pathlib import Path
import config

_entries: list[dict] = []


def log(code: str, name: str, status: str, reason: str = ""):
    entry = {
        "code":   code,
        "name":   name,
        "status": status,
        "reason": reason,
        "ts":     datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
    }
    _entries.append(entry)
    _write_to_file(entry)


def _write_to_file(entry: dict):
    log_path = config.LOG_DIR / f"run_{datetime.now().strftime('%Y%m%d')}.log"
    with open(log_path, "a", encoding="utf-8") as f:
        f.write(f"[{entry['ts']}] [{entry['status']}] {entry['code']} {entry['name']} — {entry['reason']}\n")


def get_all() -> list[dict]:
    return list(_entries)
