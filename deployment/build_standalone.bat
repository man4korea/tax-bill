@echo off
echo ====================================
echo HomeTax 시스템 스탠드얼론 앱 빌드
echo ====================================
echo.

REM 가상환경 활성화 (선택사항)
REM call venv\Scripts\activate

echo 1. Playwright 브라우저 설치...
playwright install chromium

echo.
echo 2. PyInstaller로 실행파일 빌드...
pyinstaller --clean build_app.spec

echo.
echo 3. 빌드 완료 확인...
if exist "dist\HomeTax_System\HomeTax_System.exe" (
    echo ✅ 빌드 성공! 
    echo 📁 실행파일 위치: dist\HomeTax_System\HomeTax_System.exe
    echo.
    echo 배포용 폴더: dist\HomeTax_System\
    echo 이 폴더 전체를 복사하여 배포하세요.
) else (
    echo ❌ 빌드 실패!
    echo 오류를 확인하고 다시 시도하세요.
)

echo.
pause