; HomeTax 전자세금계산서 시스템 설치 스크립트
; NSIS (Nullsoft Scriptable Install System) 사용

!define PRODUCT_NAME "HomeTax 전자세금계산서 시스템"
!define PRODUCT_VERSION "1.0.0"
!define PRODUCT_PUBLISHER "HomeTax Automation Team"
!define PRODUCT_WEB_SITE "https://hometax.go.kr"
!define PRODUCT_DIR_REGKEY "Software\Microsoft\Windows\CurrentVersion\App Paths\HomeTax_System.exe"
!define PRODUCT_UNINST_KEY "Software\Microsoft\Windows\CurrentVersion\Uninstall\${PRODUCT_NAME}"
!define PRODUCT_UNINST_ROOT_KEY "HKLM"

; 설치 파일 정보
Name "${PRODUCT_NAME} ${PRODUCT_VERSION}"
OutFile "HomeTax_System_Setup.exe"
InstallDir "$PROGRAMFILES\HomeTax System"
InstallDirRegKey HKLM "${PRODUCT_DIR_REGKEY}" ""
ShowInstDetails show
ShowUnInstDetails show

; 관리자 권한 요청
RequestExecutionLevel admin

; 아이콘 설정
Icon "hometax_icon.ico"
UninstallIcon "hometax_icon.ico"

; 모던 UI 사용
!include "MUI2.nsh"

; MUI 설정
!define MUI_ABORTWARNING
!define MUI_ICON "hometax_icon.ico"
!define MUI_UNICON "hometax_icon.ico"
!define MUI_WELCOMEFINISHPAGE_BITMAP "installer_banner.bmp"
!define MUI_UNWELCOMEFINISHPAGE_BITMAP "installer_banner.bmp"

; 설치 페이지들
!insertmacro MUI_PAGE_WELCOME
!insertmacro MUI_PAGE_LICENSE "LICENSE.txt"
!insertmacro MUI_PAGE_COMPONENTS
!insertmacro MUI_PAGE_DIRECTORY
!insertmacro MUI_PAGE_INSTFILES

; 완료 페이지 (바탕화면 바로가기 옵션)
!define MUI_FINISHPAGE_SHOWREADME ""
!define MUI_FINISHPAGE_SHOWREADME_TEXT "바탕화면에 바로가기 만들기"
!define MUI_FINISHPAGE_SHOWREADME_FUNCTION CreateDesktopShortcut
!define MUI_FINISHPAGE_RUN "$INSTDIR\HomeTax_System.exe"
!define MUI_FINISHPAGE_RUN_TEXT "HomeTax 시스템 실행"
!insertmacro MUI_PAGE_FINISH

; 제거 페이지들
!insertmacro MUI_UNPAGE_WELCOME
!insertmacro MUI_UNPAGE_CONFIRM
!insertmacro MUI_UNPAGE_INSTFILES
!insertmacro MUI_UNPAGE_FINISH

; 언어 설정
!insertmacro MUI_LANGUAGE "Korean"

; 버전 정보
VIProductVersion "1.0.0.0"
VIAddVersionKey "ProductName" "${PRODUCT_NAME}"
VIAddVersionKey "Comments" "HomeTax 전자세금계산서 자동화 시스템"
VIAddVersionKey "CompanyName" "${PRODUCT_PUBLISHER}"
VIAddVersionKey "LegalTrademarks" ""
VIAddVersionKey "LegalCopyright" "© 2025 HomeTax Automation Team"
VIAddVersionKey "FileDescription" "${PRODUCT_NAME}"
VIAddVersionKey "FileVersion" "${PRODUCT_VERSION}"

; 설치 섹션들
Section "메인 프로그램" SEC01
    SectionIn RO  ; 필수 설치
    
    ; 설치 디렉터리 설정
    SetOutPath "$INSTDIR"
    
    ; 메인 실행파일들 복사
    File "dist\HomeTax_System\HomeTax_System.exe"
    File /r "dist\HomeTax_System\_internal"
    
    ; Python 스크립트들 복사
    File "hometax_main.py"
    File "hometax_quick.py"
    File "hometax_excel_integration.py"
    File "hometax_partner_input.py"
    File "tax_invoice_generator.py"
    
    ; 설정 파일
    File ".env"
    
    ; 아이콘 파일
    File "hometax_icon.ico"
    File "hometax_logo.png"
    
    ; 문서 파일들
    SetOutPath "$INSTDIR\docs"
    File "DEPLOYMENT_GUIDE.md"
    File "CLAUDE.md"
    File "field_mapping.md"
    
    ; 시작 메뉴 바로가기
    CreateDirectory "$SMPROGRAMS\HomeTax System"
    CreateShortCut "$SMPROGRAMS\HomeTax System\HomeTax 전자세금계산서 시스템.lnk" "$INSTDIR\HomeTax_System.exe" "" "$INSTDIR\hometax_icon.ico" 0
    CreateShortCut "$SMPROGRAMS\HomeTax System\제거.lnk" "$INSTDIR\uninst.exe" "" "" 0
    CreateShortCut "$SMPROGRAMS\HomeTax System\사용자 가이드.lnk" "$INSTDIR\docs\DEPLOYMENT_GUIDE.md" "" "" 0
    
    ; 레지스트리 등록
    WriteRegStr HKLM "${PRODUCT_DIR_REGKEY}" "" "$INSTDIR\HomeTax_System.exe"
    WriteRegStr ${PRODUCT_UNINST_ROOT_KEY} "${PRODUCT_UNINST_KEY}" "DisplayName" "$(^Name)"
    WriteRegStr ${PRODUCT_UNINST_ROOT_KEY} "${PRODUCT_UNINST_KEY}" "UninstallString" "$INSTDIR\uninst.exe"
    WriteRegStr ${PRODUCT_UNINST_ROOT_KEY} "${PRODUCT_UNINST_KEY}" "DisplayIcon" "$INSTDIR\hometax_icon.ico"
    WriteRegStr ${PRODUCT_UNINST_ROOT_KEY} "${PRODUCT_UNINST_KEY}" "DisplayVersion" "${PRODUCT_VERSION}"
    WriteRegStr ${PRODUCT_UNINST_ROOT_KEY} "${PRODUCT_UNINST_KEY}" "Publisher" "${PRODUCT_PUBLISHER}"
    
    ; 제거 프로그램 생성
    WriteUninstaller "$INSTDIR\uninst.exe"
SectionEnd

Section "엑셀 템플릿 파일" SEC02
    ; 사용자 문서 폴더에 엑셀 템플릿 설치
    SetOutPath "$DOCUMENTS"
    File "세금계산서.xlsx"
    
    ; 엑셀 파일 경로를 레지스트리에 저장
    WriteRegStr HKCU "Software\HomeTax System" "ExcelTemplate" "$DOCUMENTS\세금계산서.xlsx"
    
    ; 사용자에게 알림
    DetailPrint "엑셀 템플릿을 문서 폴더에 설치했습니다: $DOCUMENTS\세금계산서.xlsx"
SectionEnd

Section "바탕화면 바로가기" SEC03
    ; 바탕화면 바로가기 생성
    CreateShortCut "$DESKTOP\HomeTax 전자세금계산서 시스템.lnk" "$INSTDIR\HomeTax_System.exe" "" "$INSTDIR\hometax_icon.ico" 0
SectionEnd

Section "Playwright 브라우저" SEC04
    ; Playwright Chromium 브라우저 설치
    DetailPrint "Playwright Chromium 브라우저를 설치하는 중..."
    
    ; Python과 Playwright 설치 확인
    ExecWait '"$INSTDIR\HomeTax_System.exe" --install-browser' $0
    
    ; 실패 시 수동 안내
    ${If} $0 != 0
        DetailPrint "브라우저 자동 설치에 실패했습니다."
        DetailPrint "프로그램 첫 실행 시 브라우저가 자동으로 다운로드됩니다."
    ${EndIf}
SectionEnd

; 섹션 설명
LangString DESC_SecMain ${LANG_KOREAN} "HomeTax 전자세금계산서 시스템 메인 프로그램 (필수)"
LangString DESC_SecExcel ${LANG_KOREAN} "세금계산서.xlsx 템플릿 파일을 문서 폴더에 설치"
LangString DESC_SecDesktop ${LANG_KOREAN} "바탕화면에 프로그램 바로가기 생성"
LangString DESC_SecBrowser ${LANG_KOREAN} "자동화에 필요한 Playwright 브라우저 설치"

!insertmacro MUI_FUNCTION_DESCRIPTION_BEGIN
  !insertmacro MUI_DESCRIPTION_TEXT ${SEC01} $(DESC_SecMain)
  !insertmacro MUI_DESCRIPTION_TEXT ${SEC02} $(DESC_SecExcel)
  !insertmacro MUI_DESCRIPTION_TEXT ${SEC03} $(DESC_SecDesktop)
  !insertmacro MUI_DESCRIPTION_TEXT ${SEC04} $(DESC_SecBrowser)
!insertmacro MUI_FUNCTION_DESCRIPTION_END

; 바탕화면 바로가기 생성 함수
Function CreateDesktopShortcut
    CreateShortCut "$DESKTOP\HomeTax 전자세금계산서 시스템.lnk" "$INSTDIR\HomeTax_System.exe" "" "$INSTDIR\hometax_icon.ico" 0
FunctionEnd

; 설치 시작 시 체크
Function .onInit
    ; 이미 설치되어 있는지 확인
    ReadRegStr $R0 ${PRODUCT_UNINST_ROOT_KEY} "${PRODUCT_UNINST_KEY}" "UninstallString"
    StrCmp $R0 "" done
    
    MessageBox MB_OKCANCEL|MB_ICONEXCLAMATION \
        "${PRODUCT_NAME}이(가) 이미 설치되어 있습니다.$\r$\n$\r$\n기존 버전을 제거하고 계속하시겠습니까?" \
        /SD IDOK IDOK uninst
    Abort
    
    uninst:
        ClearErrors
        ExecWait '$R0 /S _?=$INSTDIR'
        
        IfErrors no_remove_uninstaller done
        IfFileExists "$INSTDIR\uninst.exe" 0 no_remove_uninstaller
            Delete "$R0"
            RMDir "$INSTDIR"
        
    no_remove_uninstaller:
    done:
FunctionEnd

; 제거 섹션
Section Uninstall
    ; 시작 메뉴 제거
    Delete "$SMPROGRAMS\HomeTax System\HomeTax 전자세금계산서 시스템.lnk"
    Delete "$SMPROGRAMS\HomeTax System\제거.lnk"
    Delete "$SMPROGRAMS\HomeTax System\사용자 가이드.lnk"
    RMDir "$SMPROGRAMS\HomeTax System"
    
    ; 바탕화면 바로가기 제거
    Delete "$DESKTOP\HomeTax 전자세금계산서 시스템.lnk"
    
    ; 프로그램 파일들 제거
    Delete "$INSTDIR\HomeTax_System.exe"
    Delete "$INSTDIR\*.py"
    Delete "$INSTDIR\*.env"
    Delete "$INSTDIR\*.ico"
    Delete "$INSTDIR\*.png"
    Delete "$INSTDIR\uninst.exe"
    
    ; 폴더들 제거
    RMDir /r "$INSTDIR\_internal"
    RMDir /r "$INSTDIR\docs"
    RMDir "$INSTDIR"
    
    ; 레지스트리 제거
    DeleteRegKey ${PRODUCT_UNINST_ROOT_KEY} "${PRODUCT_UNINST_KEY}"
    DeleteRegKey HKLM "${PRODUCT_DIR_REGKEY}"
    DeleteRegKey HKCU "Software\HomeTax System"
    
    ; 엑셀 파일 제거 여부 묻기
    MessageBox MB_YESNO|MB_ICONQUESTION \
        "문서 폴더의 세금계산서.xlsx 템플릿 파일도 함께 제거하시겠습니까?" \
        /SD IDNO IDYES remove_excel IDNO keep_excel
    
    remove_excel:
        Delete "$DOCUMENTS\세금계산서.xlsx"
        DetailPrint "엑셀 템플릿 파일을 제거했습니다."
        Goto done_excel
    
    keep_excel:
        DetailPrint "엑셀 템플릿 파일을 보존했습니다: $DOCUMENTS\세금계산서.xlsx"
    
    done_excel:
        SetAutoClose true
SectionEnd

; 제거 시작 시 체크
Function un.onInit
    MessageBox MB_ICONQUESTION|MB_YESNO|MB_DEFBUTTON2 \
        "$(^Name)을(를) 완전히 제거하시겠습니까?" \
        /SD IDYES IDYES +2
    Abort
FunctionEnd

Function un.onUninstSuccess
    HideWindow
    MessageBox MB_ICONINFORMATION|MB_OK \
        "$(^Name)이(가) 성공적으로 제거되었습니다." \
        /SD IDOK
FunctionEnd