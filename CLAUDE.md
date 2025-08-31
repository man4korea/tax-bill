# 🏛️ 홈택스 세금계산서 자동화 시스템 v2.3

Korean HomeTax Tax Invoice Automation System - Claude Code AI 개발 가이드 (Playwright 독립 로그인 시스템)

---

## 📋 프로젝트 개요

한국 홈택스(hometax.go.kr) 시스템의 세금계산서 작성을 자동화하는 Python 기반 통합 솔루션입니다.

**핵심 기능**:
- 🔐 홈택스 인증서 자동 로그인
- 📊 엑셀 데이터 기반 거래처 자동 등록  
- 🧾 세금계산서 자동 작성 및 발급
- 💼 배포 가능한 독립실행형 애플리케이션

**개발 환경**:
- **언어**: Python 3.8+
- **웹 자동화**: Playwright 
- **엑셀 처리**: openpyxl, xlwings, pandas
- **UI 프레임워크**: tkinter, ttkbootstrap
- **배포**: PyInstaller, NSIS Installer

---

## 📁 프로젝트 구조 (v2.3 - Playwright 통합 로그인 시스템)

```
C:\APP\tax-bill\
├── 📂 core/                    # 핵심 시스템 모듈 (Playwright 통합 완료)
│   ├── hometax_main.py         # 메인 UI 애플리케이션
│   ├── hometax_login_module.py # 🔄 공통 로그인 모듈 (독립 모듈 subprocess 방식)
│   ├── auto-login.py           # 🆕 자동 로그인 모듈 (Playwright, 독립 실행)
│   ├── manual-login.py         # 🆕 수동 로그인 모듈 (Playwright, 독립 실행)
│   ├── excel_unified_processor.py # 🆕 통합 엑셀 처리 모듈 (600+줄)
│   ├── hometax_partner_registration.py # 거래처 등록 (Playwright 통합)
│   ├── hometax_security_manager.py # 보안 관리 모듈
│   ├── hometax_cert_manager.py # 인증서 관리 모듈
│   ├── field_mapping.md        # 필드 매핑 문서
│   ├── requirements.txt        # Python 의존성 패키지 목록
│   │
│   ├── 📂 tax-invoice/         # 세금계산서 전용 모듈
│   │   ├── hometax_tax_invoice.py # 세금계산서 자동화 (2200+줄 → 통합 모듈 사용)
│   │   ├── excel_data_manager.py # 엑셀 데이터 관리자
│   │   ├── excel_reader.py     # 엑셀 파일 리더
│   │   ├── hometax_transaction_processor.py # 거래내역 처리 모듈
│   │   ├── hometax_utils.py    # 세금계산서 전용 유틸리티
│   │   └── README.md          # 세금계산서 모듈 설명서
│   │
│   └── 📂 utils/               # 유틸리티 및 리소스
│       ├── create_hometax_icon.py # 아이콘 생성 도구
│       ├── extract_logo.py     # 로고 추출 도구
│       ├── hometax_icon.ico    # 애플리케이션 아이콘
│       ├── hometax_logo.png    # 홈택스 로고
│       └── auto_login_error.png # 오류 스크린샷
│
├── 📂 tests/                   # 테스트 파일
│   ├── test_functions.py       # 함수 검증 테스트
│   ├── check_data_rows.py      # 데이터 검증 테스트
│   ├── test_magicline_detection.py # 매직라인 감지 테스트
│   └── 📂 archive/             # 레거시 테스트 파일
│       ├── hometax_customer.py # 거래처 관련 레거시
│       ├── hometax_partner_input.py # 파트너 입력 레거시
│       ├── hometax_tax_invoice_automation.py # 세금계산서 레거시
│       └── README.md          # 아카이브 설명서
│
├── 📂 deployment/              # 배포 관련 파일
│   ├── build_single_exe.bat    # 단일 실행파일 빌드
│   ├── build_standalone.bat    # 독립실행형 빌드
│   ├── build_installer.bat     # 인스톨러 빌드
│   ├── install.bat             # Windows 설치 스크립트
│   ├── install_nsis.bat        # NSIS 설치 스크립트
│   ├── license_system.py       # 라이선스 시스템
│   └── LICENSE.txt             # 라이선스 파일
│
├── 📂 docs/                    # 문서 및 가이드
│   ├── DEPLOYMENT_GUIDE.md     # 배포 가이드
│   ├── ONLINE_LICENSE_SERVER_GUIDE.md # 온라인 라이선스 서버 가이드
│   ├── CERT_MANAGER_GUIDE.md   # 인증서 관리 가이드
│   └── GEMINI.md              # Gemini AI 가이드 (템플릿)
│
├── README.md                   # 프로젝트 메인 README
└── CLAUDE.md                   # 본 파일 - AI 개발 가이드
```
## 표준 헤더 적용 지침

### 파일 헤더 형식
```
# 📁 C:\APP\tax-bill\[경로]\[파일명]
# Create at YYMMDDhhmm Ver1.00
```

### 파일별 주석 형식
- **Python**: `# 헤더`
- **HTML/XML**: `<!-- 헤더 -->`
- **JavaScript/CSS**: `// 헤더`
- **SQL**: `-- 헤더`
- **Markdown**: `<!-- 헤더 -->`

### 헤더 업데이트 규칙
파일 변경 시 헤더에 Update 라인 추가 (JSON 파일 제외):
```
# 📁 C:\APP\tax-bill\[경로]\[파일명]
# Create at YYMMDDhhmm Ver1.00
# Update at YYMMDDhhmm Ver1.01
```

### 버전 관리
- **초기 생성**: Ver1.00
- **마이너 수정**: Ver1.01, Ver1.02...
- **메이저 변경**: Ver2.00, Ver3.00...

### 시간 형식
- **형식**: `YYMMDDhhmm` (연월일시분, 한국시간 기준)
- **확인 원칙**: 반드시 `date "+%Y년 %m월 %d일 %H시 %M분"` 명령어로 실시간 확인
- **중요**: 추측이나 대략적인 시간 사용 금지, 매번 실제 명령어 실행
- **AI 주의사항**: Claude는 실시간을 알 수 없으므로 반드시 Bash 도구로 `date` 명령어 실행 후 시간 확인
- **위반 사례**: 추측으로 작성하면 잘못된 시간이 기록됨 (예: 실제 21:42인데 15:50으로 잘못 기록)

### 제외 파일
- JSON, 바이너리, 자동 생성 파일

## 🚀 실행 방법

### 1. 메인 애플리케이션 실행
```bash
cd C:\APP\tax-bill
python core/hometax_main.py
```

### 2. 세금계산서 자동화 (통합 모듈 사용)
```bash
# 세금계산서 자동화 시스템 (새로운 통합 모듈)
python core/tax-invoice/hometax_tax_invoice.py
```

### 3. 거래처 등록 자동화 (통합 모듈 사용)
```bash
# 거래처 등록 자동화 (새로운 통합 모듈)
python core/hometax_partner_registration.py
```

### 4. 로그인 노트북 실행 (개별 테스트)
```bash

# 로그인 모듈 테스트
python core/hometax_login_module.py
```

### 5. 배포용 실행파일 생성
```bash
# 단일 실행파일 생성
deployment/build_single_exe.bat

# 독립실행형 배포판 생성
deployment/build_standalone.bat

# 인스톨러 생성 (NSIS 필요)
deployment/build_installer.bat
```
## 🔧 개발 환경 설정

### 필수 의존성 설치
```bash
# 기본 의존성 (중복 파일 정리 완료)
pip install -r core/requirements.txt

# Playwright 설치 (권장)
pip install playwright python-dotenv openpyxl xlwings pandas ttkbootstrap
playwright install chromium

---

## 🏗️ 시스템 아키텍처

### 통합 모듈 아키텍처 (v2.3 - Playwright 독립 로그인 시스템)
```
[hometax_main.py] - 메인 UI 애플리케이션
         │
         ├─► [hometax_partner_registration.py] - 거래처 등록 (Playwright 통합)
         │            │
         │            ├─► [hometax_login_module.py] - 🔄 공통 로그인 디스패처
         │            │    ├── hometax_login_dispatcher() (분기 처리)
         │            │    ├── subprocess.run("auto-login.py") - 독립 프로세스 실행
         │            │    ├── subprocess.run("manual-login.py") - 독립 프로세스 실행
         │            │    └── _fallback_manual_login() - 기존 방식 fallback
         │            │
         │            ├─► [auto-login.py] - 🆕 독립 자동 로그인 모듈 (Playwright)
         │            ├─► [manual-login.py] - 🆕 독립 수동 로그인 모듈 (Playwright)
         │            │
         │            └─► [excel_unified_processor.py] (파트너 시트)
         │                         │
         │                         ├── ExcelUnifiedProcessor (600+줄 통합 모듈)
         │                         ├── ExcelFileManager (3단계 파일 열기)
         │                         ├── RowSelector (GUI 기반 행 선택)
         │                         ├── DataProcessor (데이터 가공)
         │                         └── StatusRecorder (상태 기록)
         │
         └─► [tax-invoice/hometax_tax_invoice.py] - 세금계산서 (Playwright 통합)
                      │
                      ├─► [excel_unified_processor.py] (거래명세표 시트)
                      ├─► [hometax_login_module.py] - 🔄 공통 로그인 디스패처
                      │    ├── 독립 모듈 subprocess 실행
                      │    └── 로그인 상태 확인 및 콜백 처리
                      │
                      ├─► [auto-login.py] - 독립 자동 로그인 (Playwright)
                      ├─► [manual-login.py] - 독립 수동 로그인 (Playwright)
                      ├─► [partner-menu-navigation.ipynb] - 알림창 처리 (Selenium)
                      │
                      ├─► [hometax_utils.py] - 공통 유틸리티
                      │    ├── FieldCollector, SelectorManager, MenuNavigator
                      │    ├── DialogHandler, format functions
                      │    └── 242줄 유틸리티 함수들
                      │
                      ├─► [hometax_transaction_processor.py] - 거래내역 처리
                      ├─► [hometax_security_manager.py] - 보안 관리
                      └─► [hometax_cert_manager.py] - 인증서 관리
```

### Playwright 독립 로그인 플로우 (v2.3 - 신규)
```
hometax_login_dispatcher(callback_function)
├── 환경변수 확인 (.env HOMETAX_LOGIN_MODE)
├── 홈택스 페이지 열기 (Playwright)
├── 로그인 모드 분기
│   ├── auto: hometax_auto_login()
│   │   ├── subprocess.run("auto-login.py") - 독립 프로세스 실행
│   │   ├── auto_login_with_playwright() 함수 실행
│   │   ├── Playwright 브라우저 독립 실행
│   │   ├── 자동 인증서 비밀번호 입력 및 로그인
│   │   └── 기존 브라우저에서 로그인 상태 확인
│   └── manual: hometax_manual_login()
│       ├── subprocess.run("manual-login.py") - 독립 프로세스 실행
│       ├── manual_login_with_playwright() 함수 실행
│       ├── 1단계: 홈택스 페이지 열기 (Playwright)
│       ├── 1.5단계: 확인/취소 버튼 메시지박스 (tkinter)
│       ├── 2단계: 공동·금융인증서 버튼 클릭
│       ├── 3단계: 공인인증서 비밀번호 자동 입력 (선택적)
│       ├── 4단계: 확인 버튼 클릭 또는 수동 로그인 대기
│       ├── 5단계: URL 변화 감지로 로그인 완료 확인 (5분)
│       └── Fallback: _fallback_manual_login() 기존 방식
└── 로그인 완료 후 → callback_function(page, browser) 실행
```

### 통합 자동화 워크플로우 (v2.3 - Playwright 독립 모듈)
1. **독립 로그인**: `hometax_login_module.py` → `.py` 독립 프로세스로 로그인 처리
2. **통합 엑셀 처리**: `excel_unified_processor.py`로 데이터 로드
3. **3단계 파일 열기**: 열린 파일 확인 → 문서 폴더 → 파일 다이얼로그
4. **GUI 행 선택**: tkinter 기반 사용자 친화적 선택
5. **데이터 처리**: 
   - 거래처 등록: 거래처 시트 처리 → 홈택스 등록
   - 세금계산서: 거래명세표 시트 처리 → 세금계산서 작성
6. **결과 기록**: 통합 상태 기록 시스템 (xlwings + openpyxl 듀얼 지원)

### 독립 모듈 시스템 특징 (v2.3)
- **독립 프로세스 실행**: subprocess.run()으로 완전 분리된 모듈 실행
- **모듈 간 의존성 제거**: auto-login.py, manual-login.py 완전 독립
- **Playwright 통일**: 모든 웹 자동화 Playwright로 통합 (Selenium은 특수 용도)
- **에러 처리**: 독립 모듈 실패 시 fallback 시스템
- **세션 연속성**: 로그인 완료 후 기존 브라우저에서 상태 확인
- **테스트 도구**: `test_python_execution()` 함수로 의존성 및 파일 확인

---


## 📊 TODO 작업 계획 및 완료 현황

---

## 📞 지원 및 문의

### 개발 관련 문의
- **개발자**: Claude AI Assistant
- **개발 기간**: 2024.08.25 - 2024.08.30
- **버전**: v2.1 (통합 모듈화 완료)

### 기술 지원
- **GitHub Issues**: [프로젝트 저장소 이슈 탭]
- **문서**: `docs/` 폴더 내 상세 가이드
- **테스트**: `tests/` 폴더 내 검증 도구

---

## 📄 라이선스 및 저작권

본 프로젝트는 교육 및 개발 목적으로 제작되었습니다.
상업적 사용 시에는 별도의 라이선스 협의가 필요합니다.

**© 2024 HomeTax Automation System. All rights reserved.**

---

*마지막 업데이트: 2024년 8월 31일*
*문서 버전: v2.3 (Playwright 독립 로그인 시스템)*