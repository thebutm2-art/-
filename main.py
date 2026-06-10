"""
배민 정산 자동화 — 메인 실행 파일
사용법:
  python main.py                        # 대화형 (월 입력)
  python main.py --month 2026-05        # 특정 월
  python main.py --skip-login           # S2(로그인/요청) 스킵 — 메일이 이미 수신된 경우
  python main.py --retry CODES          # 특정 점포코드만 재처리 (쉼표 구분)
  python main.py --template             # 점포 마스터 템플릿 생성
"""
import argparse
import asyncio
import sys
from datetime import datetime
from pathlib import Path
from rich.console import Console
from rich.rule import Rule

import config
import logger
from step1_load_master  import load_master, create_master_template, print_master_table
from step2_baemin_login import run_settlement_requests
from step3_mail_receiver import fetch_settlement_mails
from step4_parse_settlement import parse_all
from step5_calculate_pl import calculate_all
from step6_report import generate_report

console = Console()


def parse_args():
    p = argparse.ArgumentParser(description="배민 정산 자동화")
    p.add_argument("--month",       default=None,  help="정산 대상 월 YYYY-MM")
    p.add_argument("--skip-login",  action="store_true", help="S2 로그인/요청 단계 스킵")
    p.add_argument("--retry",       default=None,  help="재처리할 점포코드 (쉼표 구분)")
    p.add_argument("--template",    action="store_true", help="점포 마스터 템플릿 생성 후 종료")
    p.add_argument("--master",      default=None,  help="점포 마스터 파일 경로")
    return p.parse_args()


def get_target_month(args) -> str:
    month = args.month or config.TARGET_MONTH
    if not month:
        default = datetime.now().strftime("%Y-%m")
        month   = input(f"정산 대상 월 입력 (YYYY-MM, 엔터=전월): ").strip()
        if not month:
            # 전월 기본값
            from dateutil.relativedelta import relativedelta
            month = (datetime.now() - relativedelta(months=1)).strftime("%Y-%m")
    return month


def main():
    args = parse_args()

    console.print(Rule("[bold blue]배달의민족 정산 자동화 시스템[/bold blue]"))

    # 템플릿 생성 모드
    if args.template:
        create_master_template(Path(args.master) if args.master else None)
        return

    # ── S1. 점포 마스터 로딩 ─────────────────────────────
    console.print(Rule("S1. 점포 마스터 로딩"))
    stores = load_master(Path(args.master) if args.master else None)

    if args.retry:
        retry_codes = {c.strip() for c in args.retry.split(",")}
        stores      = [s for s in stores if s["code"] in retry_codes]
        console.print(f"[yellow]재처리 대상: {[s['name'] for s in stores]}[/yellow]")

    if not stores:
        console.print("[red]처리 대상 점포가 없습니다.[/red]")
        sys.exit(1)

    print_master_table(stores)
    target_month = get_target_month(args)
    console.print(f"\n[bold]대상 기간: {target_month}[/bold]")

    # ── S2. 셀프서비스 로그인 / 정산 요청 ─────────────────
    if not args.skip_login:
        console.print(Rule("S2. 셀프서비스 로그인 / 정산 요청"))
        requested = asyncio.run(run_settlement_requests(stores, target_month))
        for s in stores:
            if s["code"] in requested:
                logger.log(s["code"], s["name"], "성공", "정산요청완료")
            else:
                logger.log(s["code"], s["name"], "실패", "정산요청실패")
    else:
        console.print("[yellow]S2 스킵 — 메일 수신 대기부터 처리합니다.[/yellow]")

    # ── S3. 메일 수신 / 첨부 추출 ──────────────────────────
    console.print(Rule("S3. 메일 자동 수신 / 첨부 추출"))
    file_map = fetch_settlement_mails(stores)
    for s in stores:
        if file_map.get(s["code"]):
            logger.log(s["code"], s["name"], "성공", "메일수신완료")
        else:
            logger.log(s["code"], s["name"], "실패", "메일미수신")

    # ── S4. 비밀번호 해제 / 파싱 ──────────────────────────
    console.print(Rule("S4. 비밀번호 해제 / 데이터 파싱"))
    parsed = parse_all(file_map, stores)
    parsed_codes = {p["store_code"] for p in parsed}
    for s in stores:
        if s["code"] not in parsed_codes:
            logger.log(s["code"], s["name"], "실패", "파싱실패")

    if not parsed:
        console.print("[red]파싱된 데이터가 없습니다. 종료.[/red]")
        sys.exit(1)

    # ── S5. 손익 계산 ──────────────────────────────────────
    console.print(Rule("S5. 손익계산서 산출"))
    pl_list = calculate_all(parsed, stores, target_month)

    # ── S6. 리포트 출력 ────────────────────────────────────
    console.print(Rule("S6. 리포트 출력"))
    log_entries = logger.get_all()
    out_path    = generate_report(pl_list, log_entries, target_month)

    # 최종 요약
    console.print(Rule("[bold green]완료[/bold green]"))
    ok_cnt   = sum(1 for e in log_entries if e["status"] == "성공" and e["reason"] == "파싱실패" is False)
    fail_cnt = sum(1 for s in stores if s["code"] not in {p["store_code"] for p in pl_list})
    console.print(f"  처리 완료: {len(pl_list)}개 | 실패: {fail_cnt}개")
    console.print(f"  리포트: [link={out_path}]{out_path}[/link]")

    if fail_cnt:
        failed = [s["name"] for s in stores if s["code"] not in {p["store_code"] for p in pl_list}]
        console.print(f"[yellow]  실패 점포: {failed}[/yellow]")
        console.print(f"  → 재처리: python main.py --skip-login --retry {','.join(s['code'] for s in stores if s['code'] not in {p['store_code'] for p in pl_list})}")


if __name__ == "__main__":
    main()
