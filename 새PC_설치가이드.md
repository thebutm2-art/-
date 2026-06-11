# 새 PC 설치 가이드 (같은 계정/같은 저장소)

깃에는 **코드만** 있고, 데이터·설정·로그인은 PC마다 새로 준비합니다.

---

## 1. 코드 받기 (git clone)
```bat
git clone https://github.com/thebutm2-art/-.git 배달손익
cd 배달손익
```
> 이미 git이 없으면 https://git-scm.com 에서 Git 설치.

## 2. 파이썬 + 패키지 설치
```bat
:: Python 3.11+ 설치 (https://www.python.org, 설치 시 "Add to PATH" 체크)
pip install -r requirements.txt
python -m playwright install chromium
```
> 봇차단 우회는 **실제 Chrome**을 씁니다 → PC에 Google Chrome도 설치되어 있어야 함.

## 3. 설정 파일 만들기 (.env)
`.env.example` 를 복사해 `.env` 로 만들고 값 입력:
```
TOORDER_ID=79daepo
TOORDER_PW=thebut007+
TOORDER_HEADLESS=true
MAIL_USER=thebut_m2@79daepo.com
BAEMIN_MAIL_SENDER=woowahan.com
```
> 메일 앱 비밀번호는 불필요(브라우저 방식). MAIL_PASS는 비워둠.

## 4. 점포 마스터 준비 (점포마스터_v2.xlsx)
방법 A — 계정정보 엑셀로 생성:
```bat
:: 배민+쿠팡 계정정보 엑셀을 이 폴더에 두고
python import_accounts.py "계정정보파일.xlsx"
```
방법 B — 기존 PC의 `점포마스터_v2.xlsx` 를 USB/메일로 복사해 이 폴더에 넣기 (배민파트너명·파일암호까지 보존되어 편함)

필수 컬럼: 점포명 | 토더매장명 | 배민파트너명 | 파일암호 | 배민아이디 | 배민비밀번호 | 쿠팡아이디 | 쿠팡비밀번호 | 가게배달건수

## 5. 식부자재 단가표 넣기
해당 월 파일을 폴더에 `식부자재_5월.xlsx` 형식으로 둠 (공통, 월 1개).

## 6. 바탕화면 바로가기 만들기
```bat
python make_shortcut.py
```
→ 바탕화면에 **"배달 손익 생성기"** 아이콘 생성.

## 7. 첫 실행 (로그인 1회)
바로가기 더블클릭 → 매장 선택 → 생성.
처음엔 브라우저 창이 열리며:
- **Gmail**(thebut_m2@79daepo.com) 로그인 1회 → 이후 세션 유지
- **배민셀프/쿠팡**은 매장 계정으로 자동 로그인 (2단계 인증/캡차 뜨면 그 창에서 1회 직접 로그인)

이후엔 매장 선택만으로 손익 보고서가 `바탕화면\{매장}\{YYMM 매장}.xlsx` 에 저장됩니다.

---

## 코드 수정 후 동기화 (두 PC 공통)
```bat
git pull            # 작업 전: 최신 코드 받기
:: ...수정...
git add -A && git commit -m "수정내용" && git push   # 작업 후: 올리기
```
> 데이터/설정(.env, 마스터, 정산서, .browser_profile)은 깃에 안 올라가니 PC마다 따로 관리.
