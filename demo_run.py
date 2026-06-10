"""
데모 실행: 샘플 암호화 정산 파일을 생성한 뒤
S4(암호해제·파싱) → S5(손익계산) → S6(리포트)까지 실제 코드로 완주.
(S2 로그인·S3 메일수신은 외부 자격증명이 필요하므로 샘플 파일로 대체)
"""
import sys, io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")

import io as _io
from pathlib import Path
import openpyxl
import msoffcrypto

import config
import logger
from step1_load_master import load_master, print_master_table
from step4_parse_settlement import parse_all
from step5_calculate_pl import calculate_all
from step6_report import generate_report
from rich.console import Console
from rich.rule import Rule

console = Console(highlight=False)

# 점포별 샘플 정산 수치 (배민 정산 파일을 모사)
SAMPLE = {
    "001": {"주문금액": 15_000_000, "중개수수료": 1_050_000, "결제수수료": 150_000,
            "배달비": 800_000, "할인부담금": 200_000, "광고비": 300_000, "부가세": 0},
    "002": {"주문금액": 22_000_000, "중개수수료": 1_540_000, "결제수수료": 220_000,
            "배달비": 1_200_000, "할인부담금": 350_000, "광고비": 500_000, "부가세": 0},
}


def make_settlement_file(store: dict, values: dict) -> Path:
    """항목명-금액 2열 구조의 Excel 생성 (배민 정산 파일 모사).
    실제 운영에서는 암호화 파일이 오며, S4가 사업자번호로 해제 후
    동일한 파싱 로직을 탄다. 여기서는 평문으로 S4 파싱 로직을 검증."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["항목", "금액"])
    for k, v in values.items():
        ws.append([k, v])
    path = config.DOWNLOAD_DIR / f"{store['code']}_정산내역.xlsx"
    wb.save(path)
    return path


def main():
    console.print(Rule("[bold blue]배민 정산 자동화 — 데모 완주[/bold blue]"))

    # S1
    console.print(Rule("S1. 점포 마스터 로딩"))
    stores = load_master(Path("점포마스터.xlsx"))
    print_master_table(stores)

    # S2/S3 대체: 샘플 정산 파일 생성
    console.print(Rule("S2/S3 대체. 샘플 정산 파일 생성 (외부연결 불필요)"))
    file_map = {}
    for s in stores:
        vals = SAMPLE.get(s["code"])
        if not vals:
            file_map[s["code"]] = []
            logger.log(s["code"], s["name"], "실패", "샘플데이터없음")
            continue
        p = make_settlement_file(s, vals)
        file_map[s["code"]] = [p]
        logger.log(s["code"], s["name"], "성공", "정산파일생성")
        console.print(f"  생성: {p.name}")

    # S4
    console.print(Rule("S4. 비밀번호 해제 / 데이터 파싱"))
    parsed = parse_all(file_map, stores)

    # S5
    console.print(Rule("S5. 손익계산서 산출"))
    pl_list = calculate_all(parsed, stores, "2026-05")

    # S6
    console.print(Rule("S6. 리포트 출력"))
    out = generate_report(pl_list, logger.get_all(), "2026-05")

    console.print(Rule("[bold green]데모 완료[/bold green]"))
    console.print(f"  처리: {len(pl_list)}개 점포")
    console.print(f"  리포트: {out}")


if __name__ == "__main__":
    main()
