"""S1. 점포 마스터 로딩"""
import pandas as pd
from pathlib import Path
from rich.console import Console
from rich.table import Table
import config

console = Console(highlight=False)

# 파일암호 = 배민 정산 엑셀 잠금 해제용 (일반/간이=생년월일6자리, 법인=법인번호 뒷7자리)
REQUIRED_COLS = ["점포명", "점포코드", "셀프서비스ID", "셀프서비스PW", "사업자번호", "파일암호", "원가율", "월고정비"]

def load_master(path: Path | None = None) -> list[dict]:
    """
    점포 마스터 Excel을 읽어 유효한 점포 리스트를 반환한다.
    필수값 누락 행은 경고 후 스킵.
    """
    path = path or config.MASTER_FILE
    if not path.exists():
        console.print(f"[red][오류] 마스터 파일 없음: {path}[/red]")
        console.print("[yellow]  → create_master_template() 실행 후 점포 정보를 입력하세요.[/yellow]")
        raise FileNotFoundError(path)

    df = pd.read_excel(path, dtype=str)
    df.columns = [c.strip() for c in df.columns]

    missing_cols = [c for c in REQUIRED_COLS if c not in df.columns]
    if missing_cols:
        raise ValueError(f"마스터 파일에 필수 컬럼 없음: {missing_cols}")

    stores = []
    skipped = 0
    for i, row in df.iterrows():
        row_num = i + 2  # Excel 행번호
        errors = []

        name  = str(row.get("점포명", "")).strip()
        code  = str(row.get("점포코드", "")).strip()
        uid   = str(row.get("셀프서비스ID", "")).strip()
        pw    = str(row.get("셀프서비스PW", "")).strip()
        bizno = str(row.get("사업자번호", "")).strip().replace("-", "")
        file_pw = str(row.get("파일암호", "")).strip().replace("-", "")
        try:
            cost_rate = float(str(row.get("원가율", "0")).replace("%", "")) / 100
        except ValueError:
            cost_rate = None
            errors.append("원가율 숫자 아님")
        try:
            fixed_cost = float(str(row.get("월고정비", "0")).replace(",", ""))
        except ValueError:
            fixed_cost = 0.0

        for field, val in [("점포명", name), ("점포코드", code),
                           ("셀프서비스ID", uid), ("셀프서비스PW", pw),
                           ("사업자번호", bizno), ("파일암호", file_pw)]:
            if not val or val in ("nan", "None"):
                errors.append(f"{field} 누락")

        if bizno and len(bizno) not in (10, 13):  # 10자리(숫자만) 또는 13자리(구분자 포함)
            errors.append(f"사업자번호 형식 오류({bizno!r})")
        # 파일암호: 생년월일 6자리 또는 법인번호 뒷 7자리
        if file_pw and len(file_pw) not in (6, 7):
            errors.append(f"파일암호 자릿수 오류({file_pw!r}, 6 또는 7자리)")

        if errors:
            console.print(f"[yellow][경고] {row_num}행 스킵 - {name or '미입력'}: {', '.join(errors)}[/yellow]")
            skipped += 1
            continue

        stores.append({
            "name":       name,
            "code":       code,
            "uid":        uid,
            "pw":         pw,
            "biz_no":     bizno,
            "file_pw":    file_pw,
            "cost_rate":  cost_rate,
            "fixed_cost": fixed_cost,
        })

    console.print(f"[green][OK] 마스터 로딩 완료: {len(stores)}개 점포 (스킵 {skipped}개)[/green]")
    return stores


def create_master_template(path: Path | None = None):
    """샘플 점포 마스터 Excel 템플릿 생성"""
    path = path or config.MASTER_FILE
    sample = {
        "점포명":      ["방이점", "강남점"],
        "점포코드":    ["001",    "002"],
        "셀프서비스ID": ["id001@example.com", "id002@example.com"],
        "셀프서비스PW": ["pw001",  "pw002"],
        "사업자번호":  ["1234567890", "9876543210"],
        # 파일암호: 일반/간이사업자=생년월일 6자리(YYMMDD), 법인사업자=법인번호 뒷 7자리
        "파일암호":    ["900101",  "1234567"],
        "원가율":      ["35%",    "33%"],
        "월고정비":    ["5000000", "6000000"],
    }
    df = pd.DataFrame(sample)
    df.to_excel(path, index=False)
    console.print(f"[green]템플릿 생성: {path}[/green]")


def print_master_table(stores: list[dict]):
    table = Table(title="점포 마스터 목록", show_lines=True)
    for col in ["#", "점포명", "코드", "ID", "사업자번호", "원가율", "월고정비"]:
        table.add_column(col)
    for i, s in enumerate(stores, 1):
        table.add_row(
            str(i), s["name"], s["code"], s["uid"],
            f"{s['biz_no'][:3]}-{s['biz_no'][3:5]}-{s['biz_no'][5:]}",
            f"{s['cost_rate']*100:.1f}%" if s["cost_rate"] else "-",
            f"{s['fixed_cost']:,.0f}원",
        )
    console.print(table)


if __name__ == "__main__":
    if not config.MASTER_FILE.exists():
        create_master_template()
    stores = load_master()
    print_master_table(stores)
