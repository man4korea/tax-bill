# 🚀 HomeTax 시스템 배포 가이드

## 📋 개요
HomeTax 전자세금계산서 시스템을 데스크탑 스탠드얼론 앱(.exe)으로 빌드하고 배포하는 방법을 안내합니다.

## 🔧 배포 준비사항

### 1. 필수 라이브러리 설치
```bash
pip install -r requirements.txt
```

### 2. Playwright 브라우저 설치
```bash
playwright install chromium
```

## 🏗️ 빌드 방법

### **방법 1: 설치 프로그램 배포 (권장)**
```bash
build_installer.bat
```

**장점:**
- ✅ 전문적인 설치/제거 프로그램
- ✅ 자동 바탕화면 바로가기 생성
- ✅ 시작 메뉴 등록
- ✅ 엑셀 템플릿 자동 설치
- ✅ Windows 프로그램 목록 등록
- ✅ 완전한 제거 지원

**결과:**
- `HomeTax_System_Setup.exe` 설치 프로그램
- 사용자가 실행하면 자동 설치

### **방법 2: 폴더 형태 배포**
```bash
build_standalone.bat
```

**장점:**
- ✅ 빠른 실행 속도
- ✅ 안정적인 동작
- ✅ 각 모듈별 파일 접근 가능

**결과:**
- `dist/HomeTax_System/` 폴더 생성
- `HomeTax_System.exe` + 필요한 라이브러리 파일들
- 전체 폴더를 복사하여 배포

### **방법 3: 단일 EXE 파일 배포**
```bash
build_single_exe.bat
```

**장점:**
- ✅ 단일 파일로 배포 간편
- ✅ 설치 필요 없음

**단점:**
- ⚠️ 첫 실행시 압축 해제로 느림
- ⚠️ 파일 크기 큼 (100MB+)

**결과:**
- `dist/HomeTax_System.exe` 단일 파일

## 📦 배포 패키지 구성

### 폴더 형태 배포 시
```
HomeTax_System/
├── HomeTax_System.exe          # 메인 실행파일
├── hometax_quick.py            # 세금계산서 자동발행
├── hometax_excel_integration.py # 거래처 등록관리  
├── tax_invoice_generator.py    # 세금계산서 생성기
├── .env                        # 환경설정 (인증서 비밀번호)
├── _internal/                  # PyInstaller 라이브러리
└── playwright/                 # Playwright 브라우저 파일
```

## 🎯 사용자 설치 방법

### 설치 프로그램 (권장)
1. `HomeTax_System_Setup.exe` 다운로드
2. 관리자 권한으로 실행
3. 설치 마법사 따라 진행
4. 설치 완료 후 바탕화면 바로가기 자동 생성
5. 시작 메뉴에서도 접근 가능

**설치 내용:**
- 프로그램 파일: `C:\Program Files\HomeTax System\`
- 엑셀 템플릿: `문서\세금계산서.xlsx`
- 바탕화면 바로가기: `HomeTax 전자세금계산서 시스템.lnk`
- 시작 메뉴: `HomeTax System` 폴더

### 폴더 형태
1. `HomeTax_System` 폴더 전체를 원하는 위치에 복사
2. `HomeTax_System.exe` 실행
3. 바탕화면 바로가기 생성 (선택사항)

### 단일 EXE
1. `HomeTax_System.exe` 파일을 원하는 위치에 복사
2. 파일 실행

## ⚙️ 환경설정

### .env 파일 설정
```
PW=인증서_비밀번호
```

사용자는 본인의 공인인증서 비밀번호를 `.env` 파일에 입력해야 합니다.

## 🔍 문제 해결

### 빌드 실패 시
1. **의존성 오류**: `pip install -r requirements.txt` 재실행
2. **Playwright 오류**: `playwright install chromium` 재실행
3. **경로 오류**: 절대경로 사용, 한글 경로 피하기

### 실행 시 오류
1. **인증서 오류**: `.env` 파일의 비밀번호 확인
2. **브라우저 오류**: Playwright chromium 설치 확인
3. **Excel 오류**: Microsoft Excel 설치 확인

## 📊 빌드 시간 및 크기

| 빌드 방법 | 빌드 시간 | 파일 크기 | 실행 속도 |
|----------|-----------|-----------|-----------|
| 폴더 형태 | 2-3분     | ~200MB    | 빠름      |
| 단일 EXE  | 3-5분     | ~150MB    | 느림      |

## 🚀 배포 권장사항

1. **폴더 형태 배포** 추천
2. **인증서 관리 가이드** 함께 제공
3. **Windows 10/11** 호환성 테스트
4. **바이러스 검사 통과** 확인
5. **디지털 서명** (선택사항, 보안 향상)

## 📝 사용자 가이드

### 시스템 요구사항
- Windows 10 이상
- .NET Framework 4.8 이상
- 인터넷 연결 (HomeTax 접속용)

### 기능
- 📄 전자세금계산서 자동발행
- 🏢 거래처 등록관리
- 📊 거래명세서 조회 (개발 예정)
- 🔍 세금계산서 조회 (개발 예정)
- 🔐 공인인증서 비밀번호 관리 (개발 예정)

## 🔐 보안 고려사항

- `.env` 파일의 비밀번호 보안
- 공인인증서 파일 백업 권장
- 정기적인 프로그램 업데이트
- 신뢰할 수 있는 경로에서만 다운로드