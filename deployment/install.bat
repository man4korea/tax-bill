@echo off
echo ============================================
echo HomeTax 거래처 등록 자동화 프로그램 설치
echo ============================================

echo.
echo 필요한 패키지를 설치합니다...
pip install -r requirements.txt

echo.
echo Playwright 브라우저를 설치합니다...
playwright install chromium

echo.
echo ============================================
echo 설치가 완료되었습니다!
echo ============================================
echo.
echo 실행 방법:
echo python hometax_excel_integration.py
echo.
pause