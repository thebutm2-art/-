@echo off
echo [배민 정산 자동화] 패키지 설치 중...
pip install -r requirements.txt
playwright install chromium
echo.
echo [완료] 설치가 끝났습니다.
echo.
echo 다음 단계:
echo   1. .env.example 을 .env 로 복사 후 메일 계정 정보 입력
echo   2. python create_template.py  → 점포마스터.xlsx 생성 후 점포 정보 입력
echo   3. python main.py             → 실행
pause
