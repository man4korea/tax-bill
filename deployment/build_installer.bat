@echo off
echo ====================================
echo HomeTax 시스템 설치 프로그램 빌드
echo ====================================
echo.

REM NSIS가 설치되어 있는지 확인
where makensis >nul 2>&1
if %ERRORLEVEL% NEQ 0 (
    echo ❌ NSIS가 설치되지 않았습니다.
    echo.
    echo NSIS 다운로드: https://nsis.sourceforge.io/Download
    echo 또는 Chocolatey로 설치: choco install nsis
    echo.
    pause
    exit /b 1
)

echo ✅ NSIS 확인됨

echo.
echo 1단계: Playwright 브라우저 설치...
playwright install chromium

echo.
echo 2단계: Python 앱 빌드...
call build_standalone.bat

echo.
echo 3단계: 빌드 결과 확인...
if not exist "dist\HomeTax_System\HomeTax_System.exe" (
    echo ❌ Python 앱 빌드가 실패했습니다.
    echo build_standalone.bat를 먼저 성공시켜주세요.
    pause
    exit /b 1
)

echo ✅ Python 앱 빌드 완료

echo.
echo 4단계: 설치 프로그램 생성...
makensis installer.nsi

echo.
echo 5단계: 최종 결과 확인...
if exist "HomeTax_System_Setup.exe" (
    echo.
    echo 🎉 설치 프로그램 생성 완료!
    echo ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
    echo 📦 파일명: HomeTax_System_Setup.exe
    echo 📂 위치: %CD%\HomeTax_System_Setup.exe
    echo.
    echo 📋 설치 내용:
    echo   • HomeTax 전자세금계산서 시스템
    echo   • 세금계산서.xlsx 템플릿
    echo   • 바탕화면 바로가기
    echo   • 시작 메뉴 등록
    echo   • Playwright 브라우저
    echo.
    echo ✨ 이 파일을 사용자에게 배포하세요!
    echo ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
) else (
    echo ❌ 설치 프로그램 생성 실패!
    echo installer.nsi 스크립트를 확인하세요.
)

echo.
pause