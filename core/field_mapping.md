<!-- 📁 C:\APP\tax-bill\core\field_mapping.md -->
<!-- Create at 2508312118 Ver1.00 -->
# 거래처 등록 필드 매핑표

| 입력화면 라벨명 | 변수명 (UI) | Excel 열명 | HomeTax 셀렉션명 |
|:---------------|:------------|:----------|:----------------|
|  |c_date  |등록일  |  |
| 거래처 사업자등록번호 | business_num | 사업자등록번호 |#mf_txppWframe_txtBsno1|
| 종사업장번호 | sub_business_num | 종사업장번호 |  |
| 상호(법인명) | company_name | 거래처명 |#mf_txppWframe_txtTnmNm  |
| 대표자 | ceo_name | 대표자 |#mf_txppWframe_txtRprs|
| 사업장 주소 | address | 사업장주소 |#mf_txppWframe_edtSplrPfbAdrTop  |
| 업태 | business_type | 업태 |#mf_txppWframe_edtSplrBcNmTop  |
| 종목 | business_item | 종목 |#mf_txppWframe_edtSplrItmNmTop  |
| 주담당부서명 | main_dept_name | 주담당부서명 |#mf_txppWframe_txtChrgDprtNm  |
| 주담당자명 | main_manager_name | 주담당자명 |#mf_txppWframe_txtChrgNm |
| 주담당자전화번호 | main_phone | 주담당자전화번호 |#mf_txppWframe_txtChrgTelNo1  |
| 주담당자휴대전화번호 | main_mobile | 주담당자휴대전화번호 |#mf_txppWframe_txtChrgMpNo2  |
| 주담당자팩스번호 | main_fax | 주담당자팩스번호 |#mf_txppWframe_txtChrgFaxNo3  |
| 주담당자이메일 앞 | main_email_1 | 주담당자이메일주소_앞 | #mf_txppWframe_txtChrgEmlAdr1 |
| 주담당자이메일 뒤 | main_email_2 | 주담당자이메일주소_뒤 | #mf_txppWframe_txtChrgEmlAdr2 |
| 주담당자 비고 | main_memo | 주담당자비고 |#mf_txppWframe_txtChrgRmrk  |
| 부담당부서명 | sub_dept_name | 부담당부서명 |#mf_txppWframe_txtSchrgDprtNm  |
| 부담당자명 | sub_manager_name | 부담당자명 |#mf_txppWframe_txtSchrgNm  |
| 부담당자전화번호 | sub_phone | 부담당자전화번호 |#mf_txppWframe_txtSchrgTelNo1  |
| 부담당자휴대전화번호 | sub_mobile | 부담당자휴대전화번호 |#mf_txppWframe_txtSchrgMpNo2  |
| 부담당자팩스번호 | sub_fax | 부담당자팩스번호 | #mf_txppWframe_txtSchrgFaxNo3 |
| 부담당자이메일 앞 | sub_email_1 | 부담당자이메일주소_앞 |#mf_txppWframe_txtSchrgEmlAdr1  |
| 부담당자이메일 뒤 | sub_email_2 | 부담당자이메일주소_뒤 |#mf_txppWframe_txtSchrgEmlAdr2  |
| 부담당자 비고 | sub_memo | 부담당자비고 |#mf_txppWframe_txtSchrgRmrk  |

## 특별 처리 필드

| 필드 | 특별 기능 |
|:-----|:----------|
| business_num | 확인 버튼 포함 |
| business_type | 조회 버튼 포함 |
| sub_business_num | 주소조회 라벨 포함 |
| main_email, sub_email | @ 분할 입력 + 직접입력 버튼 |

## 사용법

1. HomeTax 페이지에서 각 필드의 셀렉션명(CSS 선택자, ID, Name 등)을 파악
2. 위 표의 "HomeTax 셀렉션명" 열에 해당 값 작성
3. 완성된 매핑표로 자동 등록 프로그램 개발 진행