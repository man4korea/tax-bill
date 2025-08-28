# 🏛️ 홈택스 세금계산서 자동화 시스템 v2.0

Korean HomeTax Tax Invoice Automation System - Claude Code AI 개발 가이드

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
- **웹 자동화**: Playwright (권장), Selenium (레거시)
- **엑셀 처리**: openpyxl, xlwings, pandas
- **UI 프레임워크**: tkinter, ttkbootstrap
- **배포**: PyInstaller, NSIS Installer

---

## 📁 프로젝트 구조

```
C:\APP\tax-bill\
├── 📂 core/                    # 핵심 시스템 모듈
│   ├── hometax_main.py         # 메인 UI 애플리케이션
│   ├── hometax_quick.py        # 세금계산서 자동화 엔진
│   ├── hometax_transaction_processor.py  # 거래내역 처리 모듈 
│   ├── excel_data_manager.py   # 엑셀 데이터 관리자
│   └── excel_reader.py         # 엑셀 파일 리더
│
├── 📂 automation/              # 자동화 모듈
│   ├── hometax_customer.py     # 거래처 등록 자동화
│   ├── hometax_partner_input.py # 거래처 입력 처리
│   ├── hometax_excel_integration.py # 엑셀 연동
│   └── hometax_tax_invoice_automation.py # 세금계산서 자동화
│
├── 📂 utils/                   # 유틸리티 및 리소스
│   ├── tax_invoice_generator.py # 세금계산서 생성기
│   ├── create_hometax_icon.py  # 아이콘 생성 도구
│   ├── extract_logo.py         # 로고 추출 도구
│   ├── hometax_icon.ico        # 애플리케이션 아이콘
│   ├── hometax_logo.png        # 홈택스 로고
│   └── auto_login_error.png    # 오류 스크린샷
│
├── 📂 tests/                   # 테스트 파일
│   ├── test_functions.py       # 함수 검증 테스트
│   └── check_data_rows.py      # 데이터 검증 테스트
│
├── 📂 deployment/              # 배포 관련 파일
│   ├── build_single_exe.bat    # 단일 실행파일 빌드
│   ├── build_standalone.bat    # 독립실행형 빌드
│   ├── build_installer.bat     # 인스톨러 빌드
│   ├── build_app.spec          # PyInstaller 설정
│   ├── installer.nsi           # NSIS 인스톨러 스크립트
│   ├── install.bat             # Windows 설치 스크립트
│   ├── install.sh              # Linux 설치 스크립트
│   ├── license_system.py       # 라이선스 시스템
│   └── LICENSE.txt             # 라이선스 파일
│
├── 📂 docs/                    # 문서 및 가이드
│   ├── DEPLOYMENT_GUIDE.md     # 배포 가이드
│   ├── ONLINE_LICENSE_SERVER_GUIDE.md # 온라인 라이선스 서버 가이드
│   ├── field_mapping.md        # 필드 매핑 문서
│   └── GEMINI.md              # Gemini AI 가이드 (템플릿)
│
├── requirements.txt           # Python 의존성
├── .env                      # 인증서 비밀번호 (PW=your_password)
└── CLAUDE.md                 # 본 파일 - AI 개발 가이드
```

---

## 🚀 실행 방법

### 1. 메인 애플리케이션 실행
```bash
cd C:\APP\tax-bill
python core/hometax_main.py
```

### 2. 빠른 세금계산서 자동화
```bash
# 추천 - 가장 안정적인 버전
python core/hometax_quick.py
```

### 3. 거래처 등록 자동화
```bash
python automation/hometax_customer.py
```

### 4. 테스트 실행
```bash
# 함수 검증 테스트
python tests/test_functions.py

# 데이터 검증 테스트  
python tests/check_data_rows.py
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

---

## 🔧 개발 환경 설정

### 필수 의존성 설치
```bash
# 기본 의존성
pip install -r requirements.txt

# Playwright 설치 (권장)
pip install playwright python-dotenv openpyxl xlwings pandas ttkbootstrap
playwright install chromium

# Selenium 설치 (레거시 지원)  
pip install selenium python-dotenv
# Chrome WebDriver 별도 설치 필요
```

### 환경 설정
```bash
# .env 파일 생성
echo PW=your_certificate_password > .env
```

---

## 🏗️ 시스템 아키텍처

### 핵심 모듈 관계도
```
[hometax_main.py] - 메인 UI 애플리케이션
         │
         └─► [hometax_quick.py] - 세금계산서 자동화 엔진
                      │
                      ├─► [hometax_transaction_processor.py] - 거래내역 처리
                      └─► [excel_data_manager.py] - 엑셀 데이터 관리
                               │
                               └─► [excel_reader.py] - 엑셀 파일 읽기
```

### 자동화 워크플로우
1. **로그인 단계**: 홈택스 인증서 자동 로그인
2. **데이터 로드**: 엑셀 파일에서 거래 데이터 읽기
3. **거래처 처리**: 거래처별 데이터 그룹핑 및 검증
4. **세금계산서 작성**: 홈택스 폼에 자동 입력
5. **발급 처리**: 세금계산서 발급보류 또는 즉시발급
6. **결과 기록**: 엑셀 시트에 처리 결과 기록

---

## 📊 TODO 작업 계획 및 완료 현황

### ✅ 완료된 작업 (2024.08.25-28)

#### 🎯 Phase 1: 핵심 시스템 구축
- ✅ **홈택스 로그인 자동화** (Playwright 기반)
- ✅ **엑셀 데이터 처리 엔진** 구현
- ✅ **세금계산서 자동 작성 기능** 완성
- ✅ **거래내역 상세 입력 프로세스** 구현
  - 공급일자 년월 비교 및 자동 변경 (5회 beep 알림)
  - 거래 수량별 처리 (1-4건: 기본, 5-16건: 확장, 16+건: 분할)
  - 합계 검증 및 외상미수금 자동 계산
  - 금액 불일치 시 연속 beep 알림 및 사용자 수정 대기

#### 🔧 Phase 2: 시스템 최적화
- ✅ **토큰 효율성 개선** - 파일 분할 모듈화
  - `hometax_transaction_processor.py` 분리 (12개 함수)
  - 메인 파일 크기 대폭 감소 (2344줄 → 최적화)
- ✅ **프로젝트 구조 정리** - 서브폴더 체계 구성
- ✅ **함수 검증 시스템** 구축
- ✅ **에러 처리 및 복구 로직** 강화

#### 🎨 Phase 3: UI/UX 개선
- ✅ **메인 UI 애플리케이션** (tkinter + ttkbootstrap)
- ✅ **진행 상황 표시** 및 로깅 시스템
- ✅ **사용자 알림 시스템** (beep 음성 알림)
- ✅ **폼 필드 초기화** 자동 처리

#### 🚀 Phase 4: 배포 시스템
- ✅ **PyInstaller 빌드 설정** 구성
- ✅ **NSIS 인스톨러** 스크립트 작성
- ✅ **라이선스 시스템** 기본 구조
- ✅ **배포 자동화** 스크립트

### 🔄 진행 중인 작업

#### 📋 현재 작업: 문서화 및 정리
- 🔄 **프로젝트 구조 정리** - 서브폴더 분류 완료
- 🔄 **CLAUDE.md 한글 재작성** - 진행 중

### 🎯 예정된 작업 (Priority Order)

#### 🔥 High Priority
- 📝 **사용자 매뉴얼 작성** (한글)
- 🧪 **통합 테스트 스위트** 구축
- 🛡️ **오류 처리 강화** (네트워크 오류, 페이지 변경 등)
- ⚡ **성능 최적화** (대용량 엑셀 처리)

#### 🔸 Medium Priority  
- 🎛️ **설정 파일 시스템** 구축
- 📊 **상세 로그 및 리포트** 기능
- 🔄 **자동 업데이트** 시스템
- 💾 **데이터 백업 및 복구** 기능

#### 🔹 Low Priority
- 🌐 **웹 기반 UI** 개발 검토
- 📱 **모바일 지원** 검토
- 🔗 **다른 세무 시스템 연동** 검토
- 🤖 **AI 기반 데이터 검증** 기능

---

## 🐛 알려진 이슈 및 해결 방법

### 🚨 중요 이슈
1. **인증서 비밀번호 보안**
   - `.env` 파일 보안 관리 필요
   - 배포 시 민감정보 제외 처리

2. **홈택스 페이지 구조 변경**
   - iframe 구조 변경 시 selector 업데이트 필요
   - 필드명 변경 대응 필요

### 🔧 해결 완료된 이슈
- ✅ **작성일자 컬럼 없음 오류** - 다중 컬럼명 fallback 로직 구현
- ✅ **합계금액 input_value 오류** - 다중 방법 시도 로직 구현  
- ✅ **거래내역 입력 불가** - 필드 매핑 개선
- ✅ **폼 필드 미초기화** - 자동 초기화 루틴 추가

---

## 🔐 보안 고려사항

### 인증 정보 관리
- **인증서 비밀번호**: `.env` 파일에 암호화하여 저장
- **세션 관리**: 브라우저 세션 자동 종료
- **로그 보안**: 민감정보 로그 기록 방지

### 데이터 보호
- **엑셀 파일**: 읽기 전용 모드로 접근
- **임시 파일**: 작업 완료 후 자동 삭제
- **오류 로그**: 개인정보 마스킹 처리

---

## 📞 지원 및 문의

### 개발 관련 문의
- **개발자**: Claude AI Assistant
- **개발 기간**: 2024.08.25 - 2024.08.28
- **버전**: v2.0 (모듈화 완료)

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

*마지막 업데이트: 2024년 8월 28일*
*문서 버전: v2.0*