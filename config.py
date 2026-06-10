"""전역 설정 및 환경변수 로드"""
import os
from pathlib import Path
from dotenv import load_dotenv

load_dotenv()

BASE_DIR = Path(__file__).parent

# 폴더 경로
DOWNLOAD_DIR = BASE_DIR / os.getenv("DOWNLOAD_DIR", "downloads")
OUTPUT_DIR   = BASE_DIR / os.getenv("OUTPUT_DIR", "output")
LOG_DIR      = BASE_DIR / os.getenv("LOG_DIR", "logs")

for d in (DOWNLOAD_DIR, OUTPUT_DIR, LOG_DIR):
    d.mkdir(exist_ok=True)

# 마스터 파일
MASTER_FILE = BASE_DIR / os.getenv("MASTER_FILE", "점포마스터.xlsx")

# 메일 설정
MAIL_HOST   = os.getenv("MAIL_HOST", "imap.gmail.com")
MAIL_PORT   = int(os.getenv("MAIL_PORT", "993"))
MAIL_USER   = os.getenv("MAIL_USER", "")
MAIL_PASS   = os.getenv("MAIL_PASS", "")

# 배민 메일 발신자 조건 (부분 일치)
# 실제 배민 정산 메일 확인 결과 (2026-06, 더벗F&C 메일함 기준)
#   발신자: 배민셀프서비스 <noreply@woowahan.com>
#   제목  : [배달의민족] {파트너명} {기간} 정산명세서
#   첨부  : [배달의민족] {파트너명} {연월} 정산명세서.xlsx (암호화)
BAEMIN_SENDER_KEYWORD = os.getenv("BAEMIN_MAIL_SENDER", "woowahan.com")
BAEMIN_SUBJECT_KEYWORD = "정산명세서"

# 메일 폴링 설정
MAIL_POLL_INTERVAL_SEC = 30   # 폴링 간격
MAIL_POLL_MAX_WAIT_SEC = 600  # 최대 대기 시간 (10분)

# 배민 셀프서비스 URL
BAEMIN_SELFSERVICE_URL = "https://self.baemin.com"

# 대상 월 (None이면 실행 시 입력받음)
TARGET_MONTH = os.getenv("TARGET_MONTH", "") or None
