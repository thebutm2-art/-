"""
점포 마스터 (최종 보고서 파이프라인용 v2)

새 손익 로직에서는:
  - 식자재/부자재 = 토더 판매량 × 식부자재 단가  (원가율 불필요)
  - 공과금 = 매출 × 5% 고정               (월고정비 불필요)
따라서 마스터는 매장 식별·매칭 정보만 관리한다.

컬럼:
  점포명        : 보고서 매장명 + 저장 폴더명 (예: 방이점)
  토더매장명     : 토더 '매장 선택'에서의 이름 (보통 점포명과 동일)
  배민파트너명   : 배민 정산 메일 제목/식별용 (예: 김효성)
  파일암호      : 배민 정산 엑셀 잠금 해제 (일반/간이=생년월일6, 법인=법인번호 뒷7)
  셀프서비스ID   : (선택) 배민 셀프서비스 자동 로그인용 — 수동 다운로드 시 불필요
  셀프서비스PW   : (선택)
"""
from pathlib import Path
import pandas as pd
from rich.console import Console
import config

console = Console(highlight=False)

MASTER_V2 = config.BASE_DIR / "점포마스터_v2.xlsx"
REQUIRED = ["점포명"]   # 점포명만 필수, 나머지는 있으면 사용


def _g(row, key):
    return str(row.get(key, "") or "").strip()


def load_stores(path: Path | None = None) -> list[dict]:
    path = path or MASTER_V2
    if not path.exists():
        raise FileNotFoundError(f"마스터 없음: {path} (import_accounts 실행)")
    df = pd.read_excel(path, dtype=str).fillna("")
    df.columns = [c.strip() for c in df.columns]
    if "점포명" not in df.columns:
        raise ValueError("마스터에 '점포명' 컬럼이 없습니다")

    stores, skipped = [], 0
    for i, row in df.iterrows():
        name = _g(row, "점포명")
        if not name:
            skipped += 1
            continue
        gagae = _g(row, "가게배달건수")
        stores.append({
            "name": name,
            "toorder_name": _g(row, "토더매장명") or name,
            "partner": _g(row, "배민파트너명"),
            "file_pw": _g(row, "파일암호").replace("-", ""),
            "baemin_id": _g(row, "배민아이디"),
            "baemin_pw": _g(row, "배민비밀번호"),
            "coupang_id": _g(row, "쿠팡아이디"),
            "coupang_pw": _g(row, "쿠팡비밀번호"),
            "gagae_count": int(gagae) if gagae.isdigit() else None,
        })
    console.print(f"[green][OK] 마스터 로딩: {len(stores)}개 점포 (스킵 {skipped})[/green]")
    return stores


def create_template(path: Path | None = None):
    path = path or MASTER_V2
    sample = {
        "점포명":      ["방이점"],
        "토더매장명":   ["방이점"],
        "배민파트너명": ["김효성"],
        "파일암호":    ["0218839"],
        "셀프서비스ID": [""],
        "셀프서비스PW": [""],
    }
    pd.DataFrame(sample).to_excel(path, index=False)
    console.print(f"[green]마스터 v2 템플릿 생성: {path}[/green]")


if __name__ == "__main__":
    if not MASTER_V2.exists():
        create_template()
    for s in load_stores():
        console.print(f"  - {s['name']} (토더:{s['toorder_name']}, 파트너:{s['partner']})")
