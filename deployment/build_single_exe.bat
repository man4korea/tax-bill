@echo off
echo ====================================
echo HomeTax 시스템 단일 EXE 파일 빌드
echo ====================================
echo.

echo 1. Playwright 브라우저 설치...
playwright install chromium

echo.
echo 2. PyInstaller로 단일 실행파일 빌드...
pyinstaller --onefile --windowed ^
    --name "HomeTax_System" ^
    --icon "hometax_icon.ico" ^
    --add-data ".env;." ^
    --add-data "hometax_quick.py;." ^
    --add-data "hometax_excel_integration.py;." ^
    --add-data "tax_invoice_generator.py;." ^
    --hidden-import "playwright" ^
    --hidden-import "playwright.async_api" ^
    --hidden-import "pandas" ^
    --hidden-import "tkinter" ^
    --hidden-import "python-dotenv" ^
    --clean ^
    hometax_main.py

echo.
echo 3. 빌드 완료 확인...
if exist "dist\HomeTax_System.exe" (
    echo ✅ 단일 EXE 빌드 성공! 
    echo 📁 실행파일 위치: dist\HomeTax_System.exe
    echo.
    echo 이 단일 파일만 배포하면 됩니다!
    echo ⚠️  주의: 첫 실행시 압축 해제로 인해 시작이 느릴 수 있습니다.
) else (
    echo ❌ 빌드 실패!
    echo 오류를 확인하고 다시 시도하세요.
)

echo.
pause