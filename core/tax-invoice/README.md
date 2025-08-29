# 세금계산서 자동화 시스템

## 시스템 구성

### 핵심 모듈
- **hometax_tax_invoice.py** - 메인 실행 및 UI 제어
- **excel_data_manager.py** - 엑셀 데이터 관리
- **hometax_utils.py** - 공통 유틸리티 (242줄)
- **hometax_transaction_processor.py** - 홈택스 웹 처리 (1,346줄)
- **hometax_security_manager.py** - 보안 관리

### 주요 클래스
- `FieldCollector` - 홈택스 필드 수집
- `SelectorManager` - CSS 셀렉터 관리  
- `MenuNavigator` - 메뉴 네비게이션
- `DialogHandler` - 팝업 처리

## 실행 흐름

1. **엑셀 데이터 로드** → 행 선택 GUI
2. **홈택스 로그인** → 인증서 자동 로그인
3. **사업자번호별 그룹화** → 16건 단위 분할
4. **세금계산서 작성** → 거래내역 입력
5. **발급보류 처리** → 결과 기록

## 주요 기능

### 데이터 처리
- 사업자번호별 자동 그룹화
- 공급일자 비교 및 자동 수정 (5회 beep)
- 합계 검증 및 외상미수금 계산

### 웹 자동화
- iframe 동적 감지 및 로그인
- 거래 건수별 입력 방식 분기 (4건↓/5-16건)
- 발급보류 후 연속 알림창 처리

## 실행 방법

```bash
cd C:\APP\tax-bill\core\tax-invoice
python hometax_tax_invoice.py
```

## 환경 설정

- **Python**: 3.8+
- **필수 패키지**: requirements.txt 참조
- **인증서**: .env 파일에 암호화된 비밀번호