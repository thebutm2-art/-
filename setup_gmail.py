"""
Gmail 연결 설정 도우미 (사장님 PC에서 1회 실행)
- 앱 비밀번호를 입력받아 .env 파일에 안전하게 저장
- IMAP 접속을 즉시 테스트해 배민 정산 메일이 보이는지 확인
"""
import sys, io, getpass, imaplib, email, email.header
from pathlib import Path

if sys.platform == "win32":
    try:
        sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")
    except Exception:
        pass

BASE = Path(__file__).parent
ENV  = BASE / ".env"


def decode(h):
    out = ""
    for part, enc in email.header.decode_header(h or ""):
        out += part.decode(enc or "utf-8", "replace") if isinstance(part, bytes) else part
    return out


def write_env(user, app_pw, sender, subject):
    lines = [
        "MAIL_HOST=imap.gmail.com",
        "MAIL_PORT=993",
        f"MAIL_USER={user}",
        f"MAIL_PASS={app_pw}",
        f"BAEMIN_MAIL_SENDER={sender}",
        "DOWNLOAD_DIR=downloads",
        "OUTPUT_DIR=output",
        "LOG_DIR=logs",
        "MASTER_FILE=점포마스터.xlsx",
        "TARGET_MONTH=",
    ]
    ENV.write_text("\n".join(lines) + "\n", encoding="utf-8")
    print(f"\n[OK] .env 저장 완료: {ENV}")


def main():
    print("=" * 56)
    print("  Gmail 연결 설정 도우미")
    print("=" * 56)
    print("※ 앱 비밀번호는 Google 계정 > 보안 > 2단계 인증 > 앱 비밀번호 에서 발급(16자리)")
    print()

    user   = input("Gmail 주소: ").strip()
    app_pw = getpass.getpass("앱 비밀번호 (16자리, 입력 시 화면에 안 보임): ").strip().replace(" ", "")
    sender = input("배민 정산 메일 발신자 [기본 woowahan.com]: ").strip() or "woowahan.com"
    subject= input("정산 메일 제목 키워드 [기본 정산명세서]: ").strip() or "정산명세서"

    print("\nIMAP 접속 테스트 중...")
    try:
        imap = imaplib.IMAP4_SSL("imap.gmail.com", 993)
        imap.login(user, app_pw)
        imap.select("INBOX")
        # 발신자 기준 최근 메일 검색
        typ, data = imap.search(None, 'FROM', f'"{sender}"')
        ids = data[0].split()
        print(f"[OK] 로그인 성공. '{sender}' 발신 메일 {len(ids)}건 발견.")

        # 최근 3건 제목 미리보기
        for uid in ids[-3:]:
            _, raw = imap.fetch(uid, "(RFC822.HEADER)")
            msg = email.message_from_bytes(raw[0][1])
            print(f"   - {decode(msg.get('Subject',''))[:50]}")
        imap.logout()
    except imaplib.IMAP4.error as e:
        print(f"[실패] 로그인 거부: {e}")
        print("   → 앱 비밀번호가 맞는지, IMAP이 켜져 있는지 확인하세요.")
        print("   (Gmail 설정 > 전달 및 POP/IMAP > IMAP 사용)")
        return
    except Exception as e:
        print(f"[실패] 접속 오류: {e}")
        return

    write_env(user, app_pw, sender, subject)
    print("\n다음 단계: python main.py --skip-login --month 2026-05")


if __name__ == "__main__":
    main()
