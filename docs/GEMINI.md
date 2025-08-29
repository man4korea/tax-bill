# 🏛️ HomeTax 세금계산서 자동화 시스템 - Gemini 가이드

이 파일은 Gemini가 HomeTax 자동화 프로젝트를 이해하고 더 나은 지원을 제공하기 위한 컨텍스트를 제공합니다.

---

# Gemini Workspace

This file provides context for Gemini to understand the HomeTax automation project and assist you better.

## Project Overview

한국 홈택스(hometax.go.kr) 시스템의 세금계산서 작성을 자동화하는 Python 기반 통합 솔루션

*   **Purpose:** 홈택스 거래처 등록 및 세금계산서 발급 자동화
*   **Technologies:** Python 3.8+, Playwright, tkinter, openpyxl, pandas, cryptography
*   **Architecture:** 모듈형 시스템 (거래처 등록 + 세금계산서 발급 분리)

## Project Structure

```
C:\APP\tax-bill\
├── 📂 core\                          # 핵심 시스템 모듈
│   ├── hometax_main.py              # 메인 UI 애플리케이션
│   ├── hometax_partner_registration.py # 거래처 등록 자동화 시스템
│   ├── hometax_security_manager.py   # 보안 관리 (AES 암호화)
│   ├── hometax_cert_manager.py      # 인증서 관리
│   ├── field_mapping.md             # 필드 매핑 정보
│   ├── 📂 tax-invoice\              # 세금계산서 자동화 시스템
│       ├── hometax_tax_invoice.py   # 메인 세금계산서 자동화
│       ├── excel_data_manager.py    # 엑셀 데이터 관리
│       ├── hometax_utils.py         # 홈택스 유틸리티 (242줄)
│       ├── hometax_transaction_processor.py # 거래내역 처리 (1,346줄)
│       ├── hometax_security_manager.py # 보안 관리
│       ├── excel_reader.py          # 엑셀 파일 읽기
│       ├── README.md                # 시스템 구성 및 실행 가이드
│       ├── requirements.txt         # Python 의존성
│       └── .env                     # 환경 변수
│   │
│   └── 📂 utils\                    # 개발 및 배포 지원 도구
│       ├── create_hometax_icon.py   # HomeTax 아이콘 생성 도구
│       ├── extract_logo.py          # HomeTax 로고 추출 도구
│       ├── tax_invoice_generator.py # 독립 세금계산서 생성 유틸리티
│       ├── hometax_icon.ico         # 애플리케이션 아이콘
│       ├── hometax_logo.png         # HomeTax 로고 이미지
│       └── auto_login_error.png     # 디버깅용 오류 스크린샷
│
├── 📂 tests\                        # 테스트 및 아카이브
│   ├── test_functions.py            # 함수 검증 테스트
│   ├── check_data_rows.py           # 데이터 검증
│   └── 📂 archive\                   # 레거시 파일들
│
├── 📂 deployment\                    # 배포 관련
│   ├── build_single_exe.bat         # 단일 실행파일 빌드
│   ├── build_standalone.bat         # 독립실행형 빌드
│   └── installer.nsi                # NSIS 인스톨러
│
└── 📂 docs\                         # 문서
    ├── DEPLOYMENT_GUIDE.md          # 배포 가이드
    ├── field_mapping.md             # 필드 매핑 (레거시, core/field_mapping.md 사용 권장)
    └── GEMINI.md                    # 본 파일 - Gemini AI 가이드
```

## Building and Running

프로젝트를 빌드, 실행 및 테스트하기 위한 주요 명령어

### 실행 명령어
*   **메인 UI:** `cd core && python hometax_main.py`
*   **거래처 등록:** `cd core && python hometax_partner_registration.py`
*   **세금계산서 발급:** `cd core/tax-invoice && python hometax_tax_invoice.py`

### 빌드 명령어
*   **단일 실행파일:** `deployment/build_single_exe.bat`
*   **독립실행형:** `deployment/build_standalone.bat`
*   **인스톨러:** `deployment/build_installer.bat`

### 테스트 명령어
*   **함수 검증:** `cd tests && python test_functions.py`
*   **데이터 검증:** `cd tests && python check_data_rows.py`

## Development Conventions

개발 규칙 및 지침

*   **Coding Style:** Python PEP 8, 한글 주석 사용, 모듈화된 구조
*   **Testing:** 수동 테스트 위주, 실제 홈택스 사이트와 연동 테스트
*   **Security:** AES 암호화된 인증서 비밀번호, 평문 저장 금지
*   **Architecture:** 거래처 등록과 세금계산서 발급을 독립된 시스템으로 분리
*   **Gemini-CLI:** 명확한 한국어 가이드 제공, 실행 경로 구분 명시

## Key Files

가장 중요한 파일들과 포함된 내용

*   **`core/hometax_main.py`**: 통합 UI 시스템, 두 자동화 시스템 연동
*   **`core/hometax_partner_registration.py`**: 거래처 등록 자동화 (1,651줄)
*   **`core/tax-invoice/hometax_tax_invoice.py`**: 세금계산서 발급 자동화 (2,214줄)
*   **`core/tax-invoice/README.md`**: 세금계산서 시스템 구성 및 실행 가이드
*   **`core/hometax_security_manager.py`**: AES 암호화 보안 관리
*   **`core/field_mapping.md`**: 홈택스 폼 필드 매핑 테이블
*   **`.env`**: 암호화된 인증서 비밀번호 (`PW_ENCRYPTED=...`)
*   **`requirements.txt`**: Python 패키지 의존성 목록

## Workflow

자동화 워크플로우 순서

1. **1단계: 거래처 등록** → `hometax_partner_registration.py`
   - 엑셀에서 거래처 정보 읽기
   - 홈택스에 거래처 자동 등록

2. **2단계: 세금계산서 발급** → `tax-invoice/hometax_tax_invoice.py`
   - 등록된 거래처 정보 활용
   - 세금계산서 자동 작성 및 발급

## Security Notes

보안 관련 주의사항

*   **인증서 비밀번호**: `.env` 파일에 AES 암호화하여 저장
*   **평문 비밀번호**: 완전히 제거됨, `PW=` 형식 지원 중단
*   **Git 제외**: `.env`, `__pycache__`, `tests/archive/` 폴더
*   **브라우저 세션**: 작업 완료 후 자동 종료 및 정리