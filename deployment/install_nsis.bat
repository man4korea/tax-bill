@echo off
echo ====================================
echo NSIS 설치 도구
echo ====================================
echo.

echo NSIS (Nullsoft Scriptable Install System)는
echo 설치 프로그램을 만들기 위한 도구입니다.
echo.

REM NSIS가 이미 설치되어 있는지 확인
where makensis >nul 2>&1
if %ERRORLEVEL% EQU 0 (
    echo ✅ NSIS가 이미 설치되어 있습니다!
    makensis /VERSION
    echo.
    echo build_installer.bat를 실행하여 설치 프로그램을 만드세요.
    pause
    exit /b 0
)

echo ❌ NSIS가 설치되지 않았습니다.
echo.
echo 설치 방법을 선택하세요:
echo.
echo 1. 웹사이트에서 직접 다운로드
echo 2. Chocolatey로 자동 설치 (권장)
echo 3. 종료
echo.

choice /C 123 /M "선택하세요 (1-3)"

if %ERRORLEVEL%==1 goto manual_download
if %ERRORLEVEL%==2 goto chocolatey_install
if %ERRORLEVEL%==3 goto end

:manual_download
echo.
echo 웹 브라우저에서 NSIS 다운로드 페이지를 엽니다...
start https://nsis.sourceforge.io/Download
echo.
echo 다운로드 후 설치를 완료하고 다시 build_installer.bat를 실행하세요.
pause
goto end

:chocolatey_install
echo.
echo Chocolatey가 설치되어 있는지 확인 중...
where choco >nul 2>&1
if %ERRORLEVEL% NEQ 0 (
    echo ❌ Chocolatey가 설치되지 않았습니다.
    echo.
    echo Chocolatey 설치 방법:
    echo 1. 관리자 권한으로 PowerShell 실행
    echo 2. 다음 명령어 실행:
    echo    Set-ExecutionPolicy Bypass -Scope Process -Force; [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.ServicePointManager]::SecurityProtocol -bor 3072; iex ((New-Object System.Net.WebClient).DownloadString('https://community.chocolatey.org/install.ps1'))
    echo.
    pause
    goto end
)

echo ✅ Chocolatey가 설치되어 있습니다.
echo.
echo NSIS를 설치하는 중...
choco install nsis -y

echo.
echo 설치 확인 중...
where makensis >nul 2>&1
if %ERRORLEVEL% EQU 0 (
    echo ✅ NSIS 설치 완료!
    echo.
    echo 이제 build_installer.bat를 실행할 수 있습니다.
) else (
    echo ❌ NSIS 설치에 실패했습니다.
    echo 수동으로 설치를 시도해주세요.
)

pause
goto end

:end
echo.
echo 프로그램을 종료합니다.