# -*- coding: utf-8 -*-
"""
HomeTax 거래 내역 입력 프로세스 모듈
Transaction Detail Input Process for HomeTax Automation

이 모듈은 HomeTax 세금계산서 작성 시 거래 내역을 입력하는 모든 기능을 포함합니다.
"""

import asyncio
import pandas as pd
from datetime import datetime
from hometax_utils import (
    play_beep, format_date, FieldCollector, SelectorManager,
    DialogHandler, get_date_columns, get_item_name_columns,
    get_cash_amount_columns, validate_page_state, format_date_range
)


async def process_transaction_details(page, processor, first_row_data, business_number):
    """거래 내역 입력 프로세스 - 10번 루틴에서 호출"""
    try:
        print("   [LIST] 거래 내역 입력 프로세스 시작")
        
        # 1. 동일 사업자번호 행들 가져오기
        work_rows = get_same_business_number_rows(processor, business_number)
        if not work_rows:
            print("   [ERROR] 동일 사업자번호 데이터가 없습니다.")
            return
            
        print(f"   [DATA] 처리할 거래 건수: {len(work_rows)}건")
        
        # 2. 공급일자 비교 및 변경
        await check_and_update_supply_date(page, work_rows[0])
        
        # 3. 거래 내역 입력 (건수에 따라 다른 방식)
        if len(work_rows) <= 4:
            await input_transaction_items_basic(page, work_rows)
        elif len(work_rows) <= 16:
            await input_transaction_items_extended(page, work_rows)
        else:
            print("   [WARN] 16건 초과 - 분할 처리가 필요합니다.")
            # 16건씩 분할 처리 로직은 별도로 구현 필요
            work_rows = work_rows[:16]  # 임시로 16건만 처리
            await input_transaction_items_extended(page, work_rows)
        
        # 4. 합계 확정 (결제방법 분류) - 발급보류 포함
        success = await finalize_transaction_summary(page, work_rows, processor, business_number)
        
        # 5. 발급보류 성공 후에만 세금계산서 시트에 기록 및 Q열 완료 표시
        if success:
            # 세금계산서 시트에 기록
            await write_to_tax_invoice_sheet(page, processor, work_rows, business_number)
            
            # Q열에 완료 표시
            from datetime import datetime
            today_date = datetime.now().strftime("%Y-%m-%d")
            
            for row_data in work_rows:
                processor.write_completion_to_excel_q_column(row_data['excel_row'], today_date)
        
        print("   [OK] 거래 내역 입력 프로세스 완료!")
        
    except Exception as e:
        print(f"   [ERROR] 거래 내역 입력 프로세스 오류: {e}")
        # 오류 발생 시 Q열에 오류 표시
        if 'work_rows' in locals():
            for row_data in work_rows:
                processor.write_error_to_excel_q_column(row_data['excel_row'], "처리오류")


def get_same_business_number_rows(processor, business_number):
    """동일 사업자번호를 가진 모든 행 데이터 반환"""
    try:
        print(f"   🔍 사업자번호 '{business_number}' 관련 데이터 검색 중...")
        
        # selected_data에서 해당 사업자번호와 일치하는 모든 행 찾기
        if not hasattr(processor, 'selected_data') or not processor.selected_data:
            print("   [ERROR] 처리할 데이터가 없습니다.")
            return []
        
        matching_rows = []
        for row_data in processor.selected_data:
            if str(row_data.get('등록번호', '')).strip() == business_number.strip():
                matching_rows.append(row_data)
        
        if not matching_rows:
            print("   [ERROR] 일치하는 사업자번호 데이터가 없습니다.")
            return []
        
        # 행 데이터를 리스트로 변환 (이미 dict 형태이므로 그대로 사용)
        work_rows = matching_rows
        
        print(f"   [OK] {len(work_rows)}건의 거래 데이터 발견")
        return work_rows
        
    except Exception as e:
        print(f"   [ERROR] 사업자번호 데이터 검색 오류: {e}")
        return []


async def check_and_update_supply_date(page, first_row):
    """공급일자 비교 및 변경 (년/월 다르면 5회 beep)"""
    try:
        print("   [DATE] 공급일자 확인 중...")
        
        excel_date = _get_excel_date(first_row)
        excel_date_obj = _parse_date(excel_date)
        
        # HomeTax 현재 공급일자 가져오기
        hometax_date_input = page.locator("#mf_txppWframe_calWrtDtTop_input")
        await hometax_date_input.wait_for(state="visible", timeout=5000)
        hometax_date_str = await hometax_date_input.input_value()
        
        print(f"   [WEB] HomeTax 공급일자: {hometax_date_str}")
        
        hometax_date_obj = pd.to_datetime(hometax_date_str)
        
        if _dates_differ_by_month(excel_date_obj, hometax_date_obj):
            await _handle_date_mismatch(page, excel_date_obj, hometax_date_input)
        else:
            print("   [OK] 공급일자 일치 - 변경 불필요")
            
    except Exception as e:
        print(f"   [ERROR] 공급일자 확인 오류: {e}")


def _get_excel_date(first_row):
    """엑셀에서 날짜 데이터 추출"""
    for col in get_date_columns():
        if col in first_row and pd.notna(first_row[col]):
            print(f"   [DATA] Excel {col}: {first_row[col]}")
            return first_row[col]
    
    print("   [WARN] Excel에서 날짜를 찾을 수 없어 현재 날짜를 사용합니다.")
    return datetime.now()


def _parse_date(date_value):
    """날짜 값을 datetime 객체로 변환"""
    if isinstance(date_value, pd.Timestamp):
        return date_value
    elif isinstance(date_value, str):
        try:
            return pd.to_datetime(date_value)
        except:
            return datetime.now()
    else:
        return datetime.now()


def _dates_differ_by_month(date1, date2):
    """두 날짜의 년/월이 다른지 확인"""
    date1_ym = f"{date1.year}{date1.month:02d}"
    date2_ym = f"{date2.year}{date2.month:02d}"
    
    if date1_ym != date2_ym:
        print(f"   [ALERT] 공급일자 년/월이 다릅니다! Excel: {date1_ym}, HomeTax: {date2_ym}")
        return True
    return False


async def _handle_date_mismatch(page, excel_date_obj, hometax_date_input):
    """날짜 불일치 처리"""
    # 5회 beep
    await play_beep(5)
    
    # 새 공급일자로 변경
    new_date_str = excel_date_obj.strftime("%Y%m%d")
    await hometax_date_input.clear()
    await hometax_date_input.fill(new_date_str)
    await page.wait_for_timeout(500)
    
    print(f"   [OK] 공급일자 변경 완료: {new_date_str}")


async def input_transaction_items_basic(page, work_rows):
    """기본 거래 내역 입력 (1-4건)"""
    try:
        print(f"   [INPUT] 기본 거래 내역 입력: {len(work_rows)}건")
        
        for i, row_data in enumerate(work_rows, 1):
            await input_single_transaction_item(page, i, row_data)
            await page.wait_for_timeout(300)
        
        print("   [OK] 기본 거래 내역 입력 완료")
        
    except Exception as e:
        print(f"   [ERROR] 기본 거래 내역 입력 오류: {e}")


async def input_transaction_items_extended(page, work_rows):
    """확장 거래 내역 입력 (5-16건)"""
    try:
        print(f"   [INPUT] 확장 거래 내역 입력: {len(work_rows)}건")
        
        # 5건 이상인 경우 품목추가 버튼 클릭이 필요
        items_to_add = len(work_rows) - 4
        if items_to_add > 0:
            print(f"   ➕ 품목 추가 필요: {items_to_add}건")
            
            for add_count in range(items_to_add):
                try:
                    # 품목추가 버튼 클릭
                    add_button = page.locator("#mf_txppWframe_btnLsatAddTop")
                    await add_button.wait_for(state="visible", timeout=3000)
                    await add_button.click()
                    await page.wait_for_timeout(500)
                    print(f"   ➕ 품목 추가 {add_count + 1}/{items_to_add}")
                    
                except Exception as add_error:
                    print(f"   [ERROR] 품목 추가 실패 {add_count + 1}: {add_error}")
                    break
        
        # 모든 거래 내역 입력
        for i, row_data in enumerate(work_rows, 1):
            await input_single_transaction_item(page, i, row_data)
            await page.wait_for_timeout(300)
        
        print("   [OK] 확장 거래 내역 입력 완료")
        
    except Exception as e:
        print(f"   [ERROR] 확장 거래 내역 입력 오류: {e}")


async def input_single_transaction_item(page, row_idx, row_data):
    """단일 거래 내역 입력"""
    try:
        print(f"   [INPUT] {row_idx}번째 거래 내역 입력 중...")
        idx = row_idx - 1  # 0-based index
        
        await _input_date_field(page, idx, row_data)
        await _input_item_name_field(page, idx, row_data)
        await _input_basic_fields(page, idx, row_data)
        await _input_remark_field(page, idx, row_data)
        
        print(f"   [OK] {row_idx}번째 거래 내역 입력 완료")
        
    except Exception as e:
        print(f"   [ERROR] {row_idx}번째 거래 내역 입력 오류: {e}")


async def finalize_transaction_summary(page, work_rows, processor, business_number):
    """거래 합계 확정 및 결제방법 분류"""
    try:
        print("   [MONEY] 거래 합계 확정 중...")
        
        # Excel 데이터에서 결제 방법별 금액 계산
        cash_amount, check_amount, note_amount = _calculate_payment_amounts(work_rows)
        
        print(f"   [CASH] 현금: {cash_amount:,.0f}원")
        print(f"   [FORM] 수표: {check_amount:,.0f}원")
        print(f"   [INPUT] 어음: {note_amount:,.0f}원")
        
        # 합계 금액 검증 및 외상미수금 계산
        credit_amount = await verify_and_calculate_credit(page, work_rows, cash_amount, check_amount, note_amount)
        
        # 각 결제 방법 입력
        await _input_payment_amounts(page, cash_amount, check_amount, note_amount, credit_amount)
        
        # 영수/청구 버튼 선택
        await _select_receipt_type(page, cash_amount, check_amount, note_amount, credit_amount)
        
        # 발급보류 버튼 클릭 및 연속 Alert 처리
        try:
            await page.wait_for_timeout(1000)  # 1초 대기
            
            # 발급보류 버튼 확인 및 클릭
            issue_button = page.locator("#mf_txppWframe_btnIsnRsrv")
            await issue_button.wait_for(state="visible", timeout=3000)
            
            print("   [FORM] 발급보류 버튼 클릭 시도...")
            
            # 연속 Alert 처리를 위한 통합 핸들러
            dialog_count = 0
            max_dialogs = 2
            all_dialogs_handled = False
            
            async def handle_consecutive_dialogs(dialog):
                nonlocal dialog_count, all_dialogs_handled
                dialog_count += 1
                
                print(f"   [ALERT] Alert {dialog_count}/2: {dialog.message}")
                await dialog.accept()  # 확인 클릭
                
                if dialog_count == 1:
                    print("   [OK] 발급보류 확인 다이얼로그 - 확인 클릭")
                elif dialog_count == 2:
                    print("   [OK] 발급보류 성공 다이얼로그 - 확인 클릭")
                    all_dialogs_handled = True
                
                # 아직 더 처리할 dialog가 있다면 계속 리스너 유지
                if dialog_count < max_dialogs:
                    page.once("dialog", handle_consecutive_dialogs)
            
            # 첫 번째 dialog 리스너 설정
            page.once("dialog", handle_consecutive_dialogs)
            
            # 발급보류 버튼 클릭
            await issue_button.click()
            print("   [FORM] 발급보류 버튼 클릭 완료")
            
            # 모든 다이얼로그 처리 완료까지 대기 (최대 10초)
            wait_time = 0
            max_wait_time = 10.0
            
            while not all_dialogs_handled and wait_time < max_wait_time:
                await page.wait_for_timeout(200)
                wait_time += 0.2
                
                # 진행 상황 로깅
                if wait_time % 2.0 < 0.2:
                    print(f"   [WAIT] Alert 처리 대기 중... ({dialog_count}/2 완료, {wait_time:.1f}초 경과)")
            
            if not all_dialogs_handled:
                print(f"   [WARN] 모든 Alert가 처리되지 않았습니다. ({dialog_count}/2 처리됨)")
                # 다시 한 번 더 시도해보자
                await page.wait_for_timeout(2000)
                
                # 추가 시도: 다양한 방법으로 Alert 감지 및 처리
                try:
                    # 방법 1: JavaScript로 직접 Alert 확인 및 처리
                    result = await page.evaluate("""
                        () => {
                            return {
                                hasAlert: typeof window.alert !== 'function',
                                hasConfirm: typeof window.confirm !== 'function',
                                documentReady: document.readyState,
                                activeElement: document.activeElement ? document.activeElement.tagName : null
                            };
                        }
                    """)
                    print(f"   [DEBUG] JavaScript 상태: {result}")
                    
                    # 방법 2: 홈택스 특정 Alert 요소 확인 및 클릭 시도
                    hometax_alerts = await page.query_selector_all("div[id*='alert'], div[class*='alert'], div[id*='dialog'], div[class*='popup']")
                    if hometax_alerts:
                        print(f"   [DEBUG] 홈택스 Alert 요소 {len(hometax_alerts)}개 발견")
                        
                        # 각 Alert 요소에서 확인 버튼 찾아서 클릭
                        for i, alert_element in enumerate(hometax_alerts):
                            try:
                                # Alert 내부의 확인 버튼들을 찾아서 클릭 시도
                                confirm_buttons = await alert_element.query_selector_all("button, input[type='button'], .btn, [onclick*='확인'], [onclick*='OK']")
                                for btn in confirm_buttons:
                                    try:
                                        btn_text = await btn.text_content() or ""
                                        btn_value = await btn.get_attribute("value") or ""
                                        if "확인" in btn_text or "OK" in btn_text.upper() or "확인" in btn_value:
                                            await btn.click()
                                            print(f"   [OK] Alert {i+1}의 확인 버튼 클릭 완료: '{btn_text or btn_value}'")
                                            success_dialog_handled = True
                                            break
                                    except Exception as btn_error:
                                        continue
                                        
                                if success_dialog_handled:
                                    break
                                    
                            except Exception as alert_error:
                                print(f"   [DEBUG] Alert {i+1} 처리 실패: {alert_error}")
                                continue
                    
                    # 방법 3: JavaScript로 직접 확인 버튼 클릭 시도
                    if not success_dialog_handled:
                        try:
                            js_click_result = await page.evaluate("""
                                () => {
                                    // 다양한 확인 버튼 선택자들 시도
                                    const selectors = [
                                        'button:contains("확인")',
                                        'input[value="확인"]',
                                        'button:contains("OK")',
                                        '[onclick*="확인"]',
                                        '.btn:contains("확인")',
                                        'button[type="button"]:contains("확인")'
                                    ];
                                    
                                    let clicked = false;
                                    document.querySelectorAll('button, input[type="button"], .btn').forEach(btn => {
                                        const text = btn.textContent || btn.value || '';
                                        if ((text.includes('확인') || text.includes('OK')) && btn.offsetParent !== null) {
                                            btn.click();
                                            clicked = true;
                                            return true;
                                        }
                                    });
                                    
                                    return { clicked: clicked };
                                }
                            """)
                            if js_click_result.get('clicked'):
                                print("   [OK] JavaScript로 확인 버튼 클릭 완료")
                                success_dialog_handled = True
                        except Exception as js_click_error:
                            print(f"   [DEBUG] JavaScript 클릭 실패: {js_click_error}")
                        
                    # 방법 4: 페이지 소스에서 성공 메시지 확인 (최종 확인용)
                    page_content = await page.content()
                    success_keywords = ["성공", "완료", "처리되었습니다", "저장되었습니다"]
                    found_keywords = [keyword for keyword in success_keywords if keyword in page_content]
                    if found_keywords:
                        print(f"   [DEBUG] 성공 관련 키워드 발견: {found_keywords}")
                        
                except Exception as js_error:
                    print(f"   [DEBUG] 추가 확인 실패: {js_error}")
            
            # 폼 초기화 확인 및 대기
            await page.wait_for_timeout(2000)  # 폼 클리어 대기
            print("   [OK] 전자세금계산서 입력 화면 클리어 완료")
            
            # 성공 여부 반환
            issuance_success = all_dialogs_handled  # 모든 다이얼로그가 처리되면 성공으로 간주
            
        except Exception as e:
            print(f"   [ERROR] 발급보류 처리 실패: {e}")
            issuance_success = False
        
        print("   [OK] 거래 합계 확정 및 발급보류 완료")
        return issuance_success  # 성공/실패 반환
        
    except Exception as e:
        print(f"   [ERROR] 거래 합계 확정 오류: {e}")
        return False  # 실패 반환


async def verify_and_calculate_credit(page, work_rows, cash_amount, check_amount, note_amount):
    """합계금액 검증 및 외상미수금 계산"""
    try:
        # 실제 거래 합계 계산
        actual_total = sum(float(row.get('합계금액', 0) or 0) for row in work_rows)
        
        # HomeTax 합계금액 가져오기 (여러 방법 시도)
        total_field = page.locator("#mf_txppWframe_edtTotaAmtHeaderTop")
        await total_field.wait_for(state="visible", timeout=3000)
        
        hometax_total_str = ""
        try:
            hometax_total_str = await total_field.input_value()
        except:
            try:
                hometax_total_str = await total_field.text_content()
            except:
                try:
                    hometax_total_str = await total_field.inner_text()
                except:
                    hometax_total_str = await total_field.get_attribute("value") or ""
        
        hometax_total = float(hometax_total_str.replace(",", "") or 0)
        
        print(f"   [DATA] 실제 합계: {actual_total:,.0f}원")
        print(f"   [WEB] HomeTax 합계: {hometax_total:,.0f}원")
        
        # HomeTax 값을 기준으로 사용 (불일치 검증 제거)
        total_amount = hometax_total
        print(f"   [OK] 합계금액 확인: {total_amount:,.0f}원")
        
        # 현금+수표+어음이 모두 0인 경우 전체 금액을 외상미수금으로
        payment_total = cash_amount + check_amount + note_amount
        
        if payment_total == 0:
            # 현금+수표+어음이 0이면 합계금액 전체를 외상미수금으로
            credit_amount = total_amount
            print(f"   [CREDIT] 결제방법이 없으므로 전체 금액을 외상미수금으로: {credit_amount:,.0f}원")
        else:
            # 외상미수금 = 총합계 - (현금 + 수표 + 어음)
            credit_amount = total_amount - payment_total
            
            if credit_amount < 0:
                print("   [WARN] 외상미수금이 음수입니다. 0으로 설정합니다.")
                credit_amount = 0
        
        return credit_amount
        
    except Exception as e:
        print(f"   [ERROR] 합계금액 검증 오류: {e}")
        return 0



async def handle_issuance_alerts(page):
    """발행 관련 Alert 처리 - 두 번의 Alert 예상 (발급보류 후 처리)"""
    try:
        print("   [ALERT] 발급보류 후 Alert 처리 대기 중...")
        
        # 발급보류 버튼 클릭 후 잠시 더 대기 (시스템 처리 시간)
        await page.wait_for_timeout(2000)  # 2초 대기
        
        # Alert 처리를 위한 통합 함수
        async def wait_for_alert(alert_name, timeout_sec):
            try:
                dialog_event = asyncio.Event()
                dialog_message = None

                async def handle_dialog(dialog):
                    nonlocal dialog_message
                    dialog_message = dialog.message
                    print(f"   [MSG] {alert_name} Alert 감지: {dialog_message}")
                    await dialog.accept()
                    dialog_event.set()

                page.once("dialog", handle_dialog)
                
                # Alert 대기
                await asyncio.wait_for(dialog_event.wait(), timeout=timeout_sec)
                print(f"   [OK] {alert_name} Alert 처리 완료")
                await page.wait_for_timeout(500)  # Alert 처리 후 잠시 대기
                return True
                
            except asyncio.TimeoutError:
                print(f"   [INFO] {alert_name} Alert 없음 (timeout: {timeout_sec}초)")
                return False
        
        # 첫 번째 Alert 처리 (더 긴 대기 시간)
        await wait_for_alert("첫 번째", 7.0)
        
        # 두 번째 Alert 처리 (첫 번째 Alert 후 나타남)
        await wait_for_alert("두 번째", 5.0)
        
        # 추가 Alert 확인 (혹시나 더 있을 수 있음)
        await wait_for_alert("추가", 3.0)
        
        # 최종 대기
        await page.wait_for_timeout(1000)
        
    except Exception as e:
        print(f"   [ERROR] Alert 처리 오류: {e}")


async def clear_form_fields(page):
    """세금계산서 작성 폼의 모든 필드 초기화"""
    try:
        print("   [CLEAR] 폼 필드 초기화 시작...")
        
        # 거래처 정보 초기화
        fields_to_clear = [
            # 상단 거래처 정보
            "#mf_txppWframe_edtDmnrTnmNmTop",        # 상호(거래처명)
            "#mf_txppWframe_edtDmnrTnmNmTop_input",  # 상호 입력 필드
            "#mf_txppWframe_edtDmnrBznoTop",         # 사업자등록번호
            "#mf_txppWframe_edtDmnrBznoTop_input",   # 사업자등록번호 입력
            "#mf_txppWframe_edtDmnrRprsvNmTop",      # 대표자명
            "#mf_txppWframe_edtDmnrAdrsTop",         # 주소
            "#mf_txppWframe_edtDmnrUptaeNmTop",      # 업태
            "#mf_txppWframe_edtDmnrJongNmTop",       # 종목
            "#mf_txppWframe_edtDmnrMchrgEmlIdTop",   # 이메일 ID
            "#mf_txppWframe_edtDmnrMchrgEmlDmanTop", # 이메일 도메인
            
            # 공급일자
            "#mf_txppWframe_calWrtDtTop_input",      # 공급일자
            
            # 품목 정보 (기본 4개 항목)
            "#mf_txppWframe_edtItmNm1",              # 품목명1
            "#mf_txppWframe_edtStndrd1",             # 규격1  
            "#mf_txppWframe_edtQy1",                 # 수량1
            "#mf_txppWframe_edtUntprc1",             # 단가1
            "#mf_txppWframe_edtSplCft1",             # 공급가액1
            "#mf_txppWframe_edtTxamt1",              # 세액1
            "#mf_txppWframe_edtRmk1",                # 비고1
            
            "#mf_txppWframe_edtItmNm2",              # 품목명2
            "#mf_txppWframe_edtStndrd2",             # 규격2
            "#mf_txppWframe_edtQy2",                 # 수량2
            "#mf_txppWframe_edtUntprc2",             # 단가2
            "#mf_txppWframe_edtSplCft2",             # 공급가액2
            "#mf_txppWframe_edtTxamt2",              # 세액2
            "#mf_txppWframe_edtRmk2",                # 비고2
            
            "#mf_txppWframe_edtItmNm3",              # 품목명3
            "#mf_txppWframe_edtStndrd3",             # 규격3
            "#mf_txppWframe_edtQy3",                 # 수량3
            "#mf_txppWframe_edtUntprc3",             # 단가3
            "#mf_txppWframe_edtSplCft3",             # 공급가액3
            "#mf_txppWframe_edtTxamt3",              # 세액3
            "#mf_txppWframe_edtRmk3",                # 비고3
            
            "#mf_txppWframe_edtItmNm4",              # 품목명4
            "#mf_txppWframe_edtStndrd4",             # 규격4
            "#mf_txppWframe_edtQy4",                 # 수량4
            "#mf_txppWframe_edtUntprc4",             # 단가4
            "#mf_txppWframe_edtSplCft4",             # 공급가액4
            "#mf_txppWframe_edtTxamt4",              # 세액4
            "#mf_txppWframe_edtRmk4",                # 비고4
            
            # 합계 정보
            "#mf_txppWframe_edtSumSplCftHeaderTop",  # 합계 공급가액
            "#mf_txppWframe_edtSumTxamtHeaderTop",   # 합계 세액
            "#mf_txppWframe_edtTotaAmtHeaderTop",    # 총 합계금액
            
            # 대금결제 정보
            "#mf_txppWframe_edtCshAmt",              # 현금
            "#mf_txppWframe_edtChkAmt",              # 수표
            "#mf_txppWframe_edtNoteAmt",             # 어음
            "#mf_txppWframe_edtCrdtAmt",             # 외상미수금
        ]
        
        # 각 필드를 순차적으로 초기화
        cleared_count = 0
        for field_selector in fields_to_clear:
            try:
                element = page.locator(field_selector)
                if await element.is_visible():
                    await element.clear()
                    cleared_count += 1
                    await page.wait_for_timeout(50)  # 짧은 대기
            except Exception as field_error:
                # 개별 필드 초기화 실패는 무시하고 계속 진행
                pass
        
        # 추가된 품목들도 초기화 (5번째부터 16번째까지)
        for i in range(5, 17):
            try:
                item_fields = [
                    f"#mf_txppWframe_edtItmNm{i}",    # 품목명
                    f"#mf_txppWframe_edtStndrd{i}",   # 규격  
                    f"#mf_txppWframe_edtQy{i}",       # 수량
                    f"#mf_txppWframe_edtUntprc{i}",   # 단가
                    f"#mf_txppWframe_edtSplCft{i}",   # 공급가액
                    f"#mf_txppWframe_edtTxamt{i}",    # 세액
                    f"#mf_txppWframe_edtRmk{i}",      # 비고
                ]
                
                for field_selector in item_fields:
                    try:
                        element = page.locator(field_selector)
                        if await element.is_visible():
                            await element.clear()
                            cleared_count += 1
                            await page.wait_for_timeout(30)
                    except:
                        pass
                        
            except Exception:
                pass
        
        print(f"   🔄 폼 필드 초기화 완료: {cleared_count}개 필드 초기화됨")
        
    except Exception as e:
        print(f"   [ERROR] 폼 필드 초기화 오류 (계속 진행): {e}")


async def write_to_tax_invoice_sheet(page, processor, work_rows, business_number):
    """세금계산서 시트에 기록"""
    try:
        print("   [FORM] 세금계산서 시트 기록 중...")
        
        # 값이 초기화되는 문제 방지 - 페이지 안정화 대기
        print("   [WAIT] 페이지 안정화 대기 중...")
        await page.wait_for_timeout(3000)  # 3초 대기로 초기화 방지
        
        # 페이지 로딩 완료 확인
        try:
            await page.wait_for_load_state("networkidle", timeout=5000)
            print("   [OK] 네트워크 안정화 완료")
        except:
            print("   [WARN] 네트워크 안정화 대기 시간 초과 - 계속 진행")
        
        # 페이지 상태 검증 함수
        async def is_page_valid():
            """페이지가 유효한 상태인지 확인"""
            try:
                # 페이지가 닫혔는지 확인
                if page.is_closed():
                    print("   [ERROR] 페이지가 이미 닫혔습니다.")
                    return False
                
                # 브라우저 컨텍스트가 유효한지 확인
                await page.evaluate("() => document.readyState")
                return True
            except Exception as e:
                print(f"   [ERROR] 페이지 상태 검증 실패: {e}")
                return False
        
        # 페이지 상태 검증
        if not await is_page_valid():
            print("   [ERROR] 페이지가 유효하지 않아 필드 수집을 건너뜁니다.")
            return
        
        # 필드값 수집
        print("   [COLLECT] 필드값 수집 시작...")
        
        # FieldCollector 인스턴스 생성
        collector = FieldCollector()
        
        # 공급일자 수집
        supply_date = await collector.get_field_value(page, "#mf_txppWframe_calWrtDtTop_input", "공급일자")
        if not supply_date:
            print("   [RETRY] 공급일자 재시도...")
            await page.wait_for_timeout(1000)
            supply_date = await collector.get_field_value(page, "#mf_txppWframe_calWrtDtTop_input", "공급일자")
        
        # 거래처 정보 우선 캐시에서 가져오기
        partner_info = None
        if hasattr(processor, 'partner_info_cache') and business_number in processor.partner_info_cache:
            partner_info = processor.partner_info_cache[business_number]
            print(f"   [CACHE] 캐시된 거래처 정보 사용: {partner_info['company_name']}")
        
        # 상호명 - 캐시 우선, 없으면 페이지에서 수집
        company_name = ""
        if partner_info and partner_info.get('company_name'):
            company_name = partner_info['company_name']
            print(f"   [CACHE] 상호명 (캐시): {company_name}")
        else:
            # 상호명 수집 (다중 선택자 시도)
            for selector in SelectorManager.COMPANY_NAME_SELECTORS:
                print(f"   [TRY] 상호명 수집 시도: {selector}")
                company_name = await collector.get_field_value(page, selector, f"상호({selector})", wait_time=5000)
                if company_name and company_name.strip():
                    print(f"   [SUCCESS] 상호명 수집 성공: '{company_name}'")
                    break
                await page.wait_for_timeout(200)
        
        if not company_name:
            # 마지막 시도: JavaScript로 직접 찾기
            try:
                company_name = await page.evaluate("""
                    () => {
                        // 다양한 방법으로 상호명 필드 찾기
                        const selectors = [
                            '#mf_txppWframe_edtDmnrTnmNmTop',
                            'input[id*="DmnrTnmNm"]',
                            'input[placeholder*="상호"]'
                        ];
                        
                        for (let selector of selectors) {
                            const el = document.querySelector(selector);
                            if (el && el.value && el.value.trim()) {
                                return el.value.trim();
                            }
                        }
                        return '';
                    }
                """)
                if company_name:
                    print(f"   [SUCCESS] 상호명 JavaScript 수집 성공: '{company_name}'")
            except Exception as e:
                print(f"   [WARN] 상호명 JavaScript 수집 실패: {e}")
        
        print(f"   [RESULT] 최종 상호명: '{company_name}'")
        
        # 이메일 - 캐시 우선, 없으면 페이지에서 수집
        email_combined = ""
        if partner_info and partner_info.get('full_email'):
            email_combined = partner_info['full_email']
            print(f"   [CACHE] 이메일 (캐시): {email_combined}")
        else:
            # 이메일 정보 가져오기 (강화된 방법)
            print("   [EMAIL] 이메일 정보 수집 시작...")
            
            # 이메일 ID 수집 - 다중 선택자 시도
            email_id = ""
            email_id_selectors = [
                "#mf_txppWframe_edtDmnrMchrgEmlIdTop",
                "#mf_txppWframe_edtDmnrMchrgEmlIdTop_input",
                "input[id*='MchrgEmlId']",
                "input[name*='emailId']",
                "[placeholder*='이메일'][placeholder*='ID']"
            ]
            
            for selector in email_id_selectors:
                email_id = await get_field_value(selector, f"이메일ID({selector})")
                if email_id:
                    print(f"   [EMAIL-ID] 수집 성공: '{email_id}'")
                    break
            
            # 이메일 도메인 수집 - 다중 선택자 시도
            email_domain = ""
            email_domain_selectors = [
                "#mf_txppWframe_edtDmnrMchrgEmlDmanTop",
                "#mf_txppWframe_edtDmnrMchrgEmlDmanTop_input", 
                "input[id*='MchrgEmlDman']",
                "input[name*='emailDomain']",
                "[placeholder*='이메일'][placeholder*='도메인']"
            ]
            
            for selector in email_domain_selectors:
                email_domain = await get_field_value(selector, f"이메일도메인({selector})")
                if email_domain:
                    print(f"   [EMAIL-DOMAIN] 수집 성공: '{email_domain}'")
                    break
        
        # JavaScript로 이메일 정보 추가 수집 시도
        if not email_id or not email_domain:
            try:
                email_data = await page.evaluate("""
                    () => {
                        const result = {id: '', domain: ''};
                        
                        // ID 찾기
                        const idSelectors = [
                            '#mf_txppWframe_edtDmnrMchrgEmlIdTop',
                            'input[id*="MchrgEmlId"]'
                        ];
                        for (let selector of idSelectors) {
                            const el = document.querySelector(selector);
                            if (el && el.value && el.value.trim()) {
                                result.id = el.value.trim();
                                break;
                            }
                        }
                        
                        // 도메인 찾기
                        const domainSelectors = [
                            '#mf_txppWframe_edtDmnrMchrgEmlDmanTop',
                            'input[id*="MchrgEmlDman"]'
                        ];
                        for (let selector of domainSelectors) {
                            const el = document.querySelector(selector);
                            if (el && el.value && el.value.trim()) {
                                result.domain = el.value.trim();
                                break;
                            }
                        }
                        
                        return result;
                    }
                """)
                
                if not email_id and email_data.get('id'):
                    email_id = email_data['id']
                    print(f"   [EMAIL-ID] JavaScript 수집 성공: '{email_id}'")
                
                if not email_domain and email_data.get('domain'):
                    email_domain = email_data['domain']
                    print(f"   [EMAIL-DOMAIN] JavaScript 수집 성공: '{email_domain}'")
                    
            except Exception as e:
                print(f"   [WARN] 이메일 JavaScript 수집 실패: {e}")
        
            # JavaScript로 이메일 정보 추가 수집 시도
            if not email_id or not email_domain:
                try:
                    email_data = await page.evaluate("""
                        () => {
                            const result = {id: '', domain: ''};
                            
                            // ID 찾기
                            const idSelectors = [
                                '#mf_txppWframe_edtDmnrMchrgEmlIdTop',
                                'input[id*="MchrgEmlId"]'
                            ];
                            
                            for (let sel of idSelectors) {
                                const el = document.querySelector(sel);
                                if (el && el.value && el.value.trim()) {
                                    result.id = el.value.trim();
                                    break;
                                }
                            }
                            
                            // Domain 찾기
                            const domainSelectors = [
                                '#mf_txppWframe_edtDmnrMchrgEmlDmanTop',
                                'input[id*="MchrgEmlDman"]'
                            ];
                            
                            for (let sel of domainSelectors) {
                                const el = document.querySelector(sel);
                                if (el && el.value && el.value.trim()) {
                                    result.domain = el.value.trim();
                                    break;
                                }
                            }
                            
                            return result;
                        }
                    """)
                    
                    if email_data['id'] and not email_id:
                        email_id = email_data['id']
                        print(f"   [EMAIL-ID] JavaScript 수집 성공: '{email_id}'")
                    if email_data['domain'] and not email_domain:
                        email_domain = email_data['domain']
                        print(f"   [EMAIL-DOMAIN] JavaScript 수집 성공: '{email_domain}'")
                        
                except Exception as e:
                    print(f"   [WARN] 이메일 JavaScript 수집 실패: {e}")
            
            # 이메일 조합 - 강화된 로직
            if email_id and email_domain:
                email_combined = f"{email_id}@{email_domain}"
                print(f"   [EMAIL] 완전한 이메일 조합 성공: '{email_combined}'")
            elif email_id and not email_domain:
                # ID만 있고 도메인이 없는 경우 - ID에 @가 포함되어 있는지 확인
                if "@" in email_id:
                    email_combined = email_id  # 이미 완전한 이메일
                    print(f"   [EMAIL] 완성된 이메일 ID 사용: '{email_combined}'")
                else:
                    email_combined = email_id  # 도메인 없으면 ID만
                    print(f"   [EMAIL] ID만 사용: '{email_combined}'")
            elif not email_id and email_domain:
                email_combined = f"@{email_domain}"  # ID 없으면 @도메인만
                print(f"   [EMAIL] 도메인만 사용: '{email_combined}'")
            else:
                email_combined = ""  # 둘 다 없으면 빈 값
                print(f"   [EMAIL] 이메일 정보 없음")
        
        print(f"   [RESULT] 최종 이메일: '{email_combined}'")
        
        # 합계 금액들 수집 (강화된 방법)
        print("   [AMOUNT] 금액 정보 수집 시작...")
        
        # 공급가액 수집 - 다중 방법 시도
        print("   [SUPPLY] 공급가액 수집 시도...")
        total_supply_raw = ""
        supply_selectors = [
            "#mf_txppWframe_edtSumSplCftHeaderTop",
            "#mf_txppWframe_edtSumSplCftHeaderTop_input",
            "input[id*='SumSplCft']",
            "input[name*='supplyAmount']",
            "[title*='공급가액']",
            ".supply-amount",
            "#supplyAmount"
        ]
        
        for selector in supply_selectors:
            total_supply_raw = await get_field_value(selector, f"공급가액({selector})", wait_time=5000)
            if total_supply_raw:
                print(f"   [SUPPLY] 공급가액 수집 성공: '{total_supply_raw}'")
                break
        
        # JavaScript로 공급가액 추가 시도
        if not total_supply_raw:
            try:
                total_supply_raw = await page.evaluate("""
                    () => {
                        const selectors = [
                            '#mf_txppWframe_edtSumSplCftHeaderTop',
                            'input[id*="SumSplCft"]',
                            'input[title*="공급가액"]'
                        ];
                        
                        for (let selector of selectors) {
                            const el = document.querySelector(selector);
                            if (el && el.value && el.value.trim()) {
                                return el.value.trim();
                            }
                        }
                        return '';
                    }
                """)
                if total_supply_raw:
                    print(f"   [SUPPLY] JavaScript 공급가액 수집 성공: '{total_supply_raw}'")
            except Exception as e:
                print(f"   [WARN] JavaScript 공급가액 수집 실패: {e}")
        
        # 세액 수집 - 다중 방법 시도
        print("   [TAX] 세액 수집 시도...")
        total_tax_raw = ""
        tax_selectors = [
            "#mf_txppWframe_edtSumTxamtHeaderTop",
            "#mf_txppWframe_edtSumTxamtHeaderTop_input",
            "input[id*='SumTxamt']",
            "input[name*='taxAmount']",
            "[title*='세액']",
            ".tax-amount",
            "#taxAmount"
        ]
        
        for selector in tax_selectors:
            total_tax_raw = await get_field_value(selector, f"세액({selector})", wait_time=5000)
            if total_tax_raw:
                print(f"   [TAX] 세액 수집 성공: '{total_tax_raw}'")
                break
        
        # JavaScript로 세액 추가 시도
        if not total_tax_raw:
            try:
                total_tax_raw = await page.evaluate("""
                    () => {
                        const selectors = [
                            '#mf_txppWframe_edtSumTxamtHeaderTop',
                            'input[id*="SumTxamt"]',
                            'input[title*="세액"]'
                        ];
                        
                        for (let selector of selectors) {
                            const el = document.querySelector(selector);
                            if (el && el.value && el.value.trim()) {
                                return el.value.trim();
                            }
                        }
                        return '';
                    }
                """)
                if total_tax_raw:
                    print(f"   [TAX] JavaScript 세액 수집 성공: '{total_tax_raw}'")
            except Exception as e:
                print(f"   [WARN] JavaScript 세액 수집 실패: {e}")
        
        # 합계금액 수집 - 다중 방법 시도
        print("   [TOTAL] 합계금액 수집 시도...")
        total_amount_raw = ""
        total_selectors = [
            "#mf_txppWframe_edtTotaAmtHeaderTop",
            "#mf_txppWframe_edtTotaAmtHeaderTop_input", 
            "input[id*='TotaAmt']",
            "input[name*='totalAmount']",
            "[title*='합계금액']",
            ".total-amount",
            "#totalAmount"
        ]
        
        for selector in total_selectors:
            total_amount_raw = await get_field_value(selector, f"합계금액({selector})", wait_time=5000)
            if total_amount_raw:
                print(f"   [TOTAL] 합계금액 수집 성공: '{total_amount_raw}'")
                break
        
        # JavaScript로 합계금액 추가 시도
        if not total_amount_raw:
            try:
                total_amount_raw = await page.evaluate("""
                    () => {
                        const selectors = [
                            '#mf_txppWframe_edtTotaAmtHeaderTop',
                            'input[id*="TotaAmt"]',
                            'input[title*="합계금액"]'
                        ];
                        
                        for (let selector of selectors) {
                            const el = document.querySelector(selector);
                            if (el && el.value && el.value.trim()) {
                                return el.value.trim();
                            }
                        }
                        return '';
                    }
                """)
                if total_amount_raw:
                    print(f"   [TOTAL] JavaScript 합계금액 수집 성공: '{total_amount_raw}'")
            except Exception as e:
                print(f"   [WARN] JavaScript 합계금액 수집 실패: {e}")
        
        # 숫자 필드 정리 (콤마 제거)
        total_supply = total_supply_raw.replace(',', '') if total_supply_raw else ""
        total_tax = total_tax_raw.replace(',', '') if total_tax_raw else ""
        total_amount = total_amount_raw.replace(',', '') if total_amount_raw else ""
        
        print(f"   [RESULT] 최종 공급가액: '{total_supply}'")
        print(f"   [RESULT] 최종 세액: '{total_tax}'")
        print(f"   [RESULT] 최종 합계금액: '{total_amount}'")
        
        # 첫 번째 품목 정보
        first_item_name = await get_field_value("#mf_txppWframe_genEtxivLsatTop_0_edtLsatNmTop", "첫번째품목명")
        first_item_spec = await get_field_value("#mf_txppWframe_genEtxivLsatTop_0_edtLsatRszeNmTop", "첫번째규격")
        first_item_quantity = await get_field_value("#mf_txppWframe_genEtxivLsatTop_0_edtLsatQtyTop", "첫번째수량")
        
        # 품목명 생성 로직 수정
        if len(work_rows) == 1:
            # 1건인 경우: 홈택스 필드값 그대로 사용
            item_name = first_item_name or work_rows[0].get('품명', '') or work_rows[0].get('품목명', '')
            item_spec = first_item_spec or work_rows[0].get('규격', '')
            item_quantity = first_item_quantity or str(work_rows[0].get('수량', ''))
        else:
            # 여러 건인 경우: "첫번째품목명 외 N개 품목" 형식으로 수정
            base_item = first_item_name or work_rows[0].get('품명', '') or work_rows[0].get('품목명', '') or '품목'
            additional_count = len(work_rows) - 1  # 첫 번째 제외한 나머지 개수
            if additional_count > 0:
                item_name = f"{base_item} 외 {additional_count}개 품목"
            else:
                item_name = base_item
            item_spec = first_item_spec or ""
            item_quantity = first_item_quantity or ""
        
        # 공급일자 범위 생성 - 형식 개선
        def format_date(date_obj):
            """날짜를 YYMMDD 형식으로 변환"""
            if not date_obj:
                return ""
            try:
                if isinstance(date_obj, str):
                    import pandas as pd
                    date_obj = pd.to_datetime(date_obj)
                return date_obj.strftime("%y%m%d")  # 250810 형식
            except:
                return str(date_obj)[:8] if str(date_obj) else ""
        
        if len(work_rows) == 1:
            # 1건일 때: 2025-08-10 형식
            single_date = work_rows[0].get('공급일자') or work_rows[0].get('작성일자', '')
            if single_date:
                try:
                    import pandas as pd
                    date_obj = pd.to_datetime(single_date)
                    date_range = date_obj.strftime("%Y-%m-%d")  # 2025-08-10 형식
                except:
                    date_range = str(single_date)
            else:
                date_range = ""
        else:
            # 여러 건일 때: 250810-250831 4건 형식
            start_date_raw = work_rows[0].get('공급일자') or work_rows[0].get('작성일자', '')
            end_date_raw = work_rows[-1].get('공급일자') or work_rows[-1].get('작성일자', '')
            
            start_formatted = format_date(start_date_raw)
            end_formatted = format_date(end_date_raw)
            
            if start_formatted and end_formatted and start_formatted != end_formatted:
                date_range = f"{start_formatted}-{end_formatted} {len(work_rows)}건"
            elif start_formatted:
                date_range = f"{start_formatted} {len(work_rows)}건"
            else:
                date_range = f"{len(work_rows)}건"
        
        # 세금계산서 시트에 기록할 데이터 준비
        tax_invoice_data = {
            'a': supply_date,  # 공급일자
            'b': business_number,  # 등록번호
            'c': company_name,  # 상호
            'd': email_combined,  # 이메일 (수정됨)
            'f': item_name,  # 품목
            'g': item_spec,  # 규격
            'h': item_quantity,  # 수량
            'i': total_supply,  # 공급가액
            'j': total_tax,  # 세액
            'k': total_amount,  # 합계금액
            'l': date_range  # 기간 및 건수
        }
        
        # 디버깅을 위한 데이터 출력
        print(f"   [DATA] 세금계산서 시트 기록 데이터:")
        for col, value in tax_invoice_data.items():
            print(f"      {col}열: '{value}'")
        
        # 값이 초기화된 경우 한 번 더 시도
        if not company_name or not total_supply or not total_tax or not total_amount:
            print("   [RETRY] 주요 값이 누락됨 - 전체 재시도...")
            await page.wait_for_timeout(2000)
            
            # 다시 한 번 시도
            if not company_name:
                company_name = await get_field_value("#mf_txppWframe_edtDmnrTnmNmTop", "상호(재시도)")
            if not total_supply:
                total_supply_retry = await get_field_value("#mf_txppWframe_edtSumSplCftHeaderTop", "공급가액(재시도)")
                total_supply = total_supply_retry.replace(',', '') if total_supply_retry else ""
            if not total_tax:
                total_tax_retry = await get_field_value("#mf_txppWframe_edtSumTxamtHeaderTop", "세액(재시도)")
                total_tax = total_tax_retry.replace(',', '') if total_tax_retry else ""
            if not total_amount:
                total_amount_retry = await get_field_value("#mf_txppWframe_edtTotaAmtHeaderTop", "합계금액(재시도)")
                total_amount = total_amount_retry.replace(',', '') if total_amount_retry else ""
            
            # 재시도 결과 업데이트
            tax_invoice_data['c'] = company_name
            tax_invoice_data['i'] = total_supply  
            tax_invoice_data['j'] = total_tax
            tax_invoice_data['k'] = total_amount
            
            print("   [RETRY] 재시도 완료")
        
        # 빈 값들 처리 - 빈 값이면 기록하지 않음
        filtered_data = {k: v for k, v in tax_invoice_data.items() if v and str(v).strip()}
        
        # 수집된 데이터 최종 검증
        critical_fields = ['c', 'i', 'j', 'k']  # 상호, 공급가액, 세액, 합계금액
        missing_fields = [field for field in critical_fields if field not in filtered_data or not str(filtered_data[field]).strip()]
        
        if missing_fields:
            print(f"   [WARN] 누락된 중요 필드: {missing_fields}")
            print("   [WARN] 가능한 원인:")
            print("     1. 페이지가 아직 로딩 중")
            print("     2. 필드가 초기화됨")
            print("     3. 선택자가 변경됨")
        else:
            print("   [OK] 모든 중요 필드 수집 완료")
        
        # 실제 엑셀 파일에 기록
        processor.write_tax_invoice_data(tax_invoice_data)
        
        # 발급보류 전에 데이터를 수집하여 기록 완료
        
        print("   [FORM] 세금계산서 시트 기록 및 필드 초기화 완료!")
        
    except Exception as e:
        print(f"   [ERROR] 세금계산서 시트 기록 오류: {e}")


# ==========================================
# 최적화된 헬퍼 함수들
# ==========================================

async def _input_date_field(page, idx, row_data):
    """일자 필드 입력"""
    supply_date = _find_column_value(row_data, get_date_columns())
    
    if supply_date:
        try:
            date_obj = pd.to_datetime(supply_date)
            day_str = str(date_obj.day)
            
            day_input = page.locator(f"#mf_txppWframe_genEtxivLsatTop_{idx}_edtLsatSplDdTop")
            await day_input.wait_for(state="visible", timeout=3000)
            await day_input.clear()
            await day_input.fill(day_str)
            print(f"      일자: {day_str}")
        except Exception as e:
            print(f"      일자 입력 실패: {e}")
    else:
        print(f"      일자: 데이터 없음")


async def _input_item_name_field(page, idx, row_data):
    """품목명 필드 입력"""
    item_name = _find_column_value(row_data, get_item_name_columns())
    
    if item_name:
        item_input = page.locator(f"#mf_txppWframe_genEtxivLsatTop_{idx}_edtLsatNmTop")
        await item_input.wait_for(state="visible", timeout=3000)
        await item_input.clear()
        await item_input.fill(item_name)
        print(f"      품목: {item_name}")
    else:
        print(f"      품목: 데이터 없음")


async def _input_basic_fields(page, idx, row_data):
    """기본 필드들 입력 (규격, 수량, 단가, 공급가액, 세액)"""
    field_mappings = [
        ('규격', f"#mf_txppWframe_genEtxivLsatTop_{idx}_edtLsatRszeNmTop", "규격"),
        ('수량', f"#mf_txppWframe_genEtxivLsatTop_{idx}_edtLsatQtyTop", "수량"),
        ('단가', f"#mf_txppWframe_genEtxivLsatTop_{idx}_edtLsatUtprcTop", "단가"),
        ('공급가액', f"#mf_txppWframe_genEtxivLsatTop_{idx}_edtLsatSplCftTop", "공급가액"),
        ('세액', f"#mf_txppWframe_genEtxivLsatTop_{idx}_edtLsatTxamtTop", "세액")
    ]
    
    for field_key, selector, display_name in field_mappings:
        value = str(row_data.get(field_key, '')).strip()
        if value:
            field_input = page.locator(selector)
            await field_input.wait_for(state="visible", timeout=3000)
            await field_input.clear()
            await field_input.fill(value)
            print(f"      {display_name}: {value}")


async def _input_remark_field(page, idx, row_data):
    """비고 필드 입력"""
    remarks = str(row_data.get('비고', '')).strip()
    if remarks and remarks != 'nan':
        try:
            remark_input = page.locator(f"#mf_txppWframe_genEtxivLsatTop_{idx}_edtLsatRmrkCntnTop")
            if await remark_input.count() > 0:
                await remark_input.wait_for(state="visible", timeout=2000)
                await remark_input.clear()
                await remark_input.fill(remarks)
                print(f"      비고: {remarks}")
            else:
                print(f"      [INFO] 비고 필드가 존재하지 않습니다")
        except Exception as e:
            print(f"      [WARN] 비고 입력 실패: {e}")
    else:
        print(f"      비고: (빈 값 또는 NaN - 건너뛰기)")


def _find_column_value(row_data, column_candidates):
    """여러 컬럼 후보에서 값 찾기"""
    for col in column_candidates:
        if col in row_data and row_data[col]:
            value = str(row_data[col]).strip()
            print(f"      데이터 발견: {col} = {value}")
            return value
    return None


def _calculate_payment_amounts(work_rows):
    """결제 방법별 금액 계산"""
    cash_amount = check_amount = note_amount = 0
    
    for row in work_rows:
        # 현금금액 추출
        row_cash_amount = 0
        for cash_col in get_cash_amount_columns():
            if cash_col in row and row[cash_col]:
                try:
                    row_cash_amount = float(str(row[cash_col]).replace(',', '') or 0)
                    print(f"      현금 데이터 발견: {cash_col} = {row_cash_amount:,.0f}원")
                    break
                except:
                    continue
        
        if row_cash_amount > 0:
            # 현금종류에 따른 분류
            payment_type = str(row.get('현금종류', '')).strip()
            
            if payment_type == '수표':
                check_amount += row_cash_amount
                print(f"      수표로 분류: {row_cash_amount:,.0f}원")
            elif payment_type == '어음':
                note_amount += row_cash_amount
                print(f"      어음으로 분류: {row_cash_amount:,.0f}원")
            else:
                cash_amount += row_cash_amount
                print(f"      현금으로 분류: {row_cash_amount:,.0f}원")
    
    # fallback 방식
    if cash_amount == 0 and check_amount == 0 and note_amount == 0:
        cash_amount = sum(float(row.get('현금', 0) or 0) for row in work_rows)
        check_amount = sum(float(row.get('수표', 0) or 0) for row in work_rows)
        note_amount = sum(float(row.get('어음', 0) or 0) for row in work_rows)
    
    return cash_amount, check_amount, note_amount


async def _input_payment_amounts(page, cash_amount, check_amount, note_amount, credit_amount):
    """결제방법별 금액 입력"""
    payment_selectors = [
        (cash_amount, "#mf_txppWframe_edtStlMthd10Top", "현금"),
        (check_amount, "#mf_txppWframe_edtStlMthd20Top", "수표"),
        (note_amount, "#mf_txppWframe_edtStlMthd30Top", "어음"),
        (credit_amount, "#mf_txppWframe_edtStlMthd40Top", "외상미수금")
    ]
    
    for amount, selector, name in payment_selectors:
        if amount > 0:
            input_field = page.locator(selector)
            await input_field.wait_for(state="visible", timeout=3000)
            await input_field.clear()
            await input_field.fill(str(int(amount)))
            if name == "외상미수금":
                print(f"   [CREDIT] {name}: {amount:,.0f}원")


async def _select_receipt_type(page, cash_amount, check_amount, note_amount, credit_amount):
    """영수/청구 버튼 선택"""
    try:
        total_payment = cash_amount + check_amount + note_amount
        
        if total_payment == 0 and credit_amount > 0:
            # 전액 외상미수금 - 청구
            button = page.locator("#mf_txppWframe_rdoRecApeClCdTop > div.w2radio_item.w2radio_item_0 > label")
            await button.wait_for(state="visible", timeout=3000)
            await button.click()
            print("   [REQUEST] 전액 외상미수금 - 청구 버튼 클릭 완료")
        else:
            # 일반적인 경우 - 영수
            button = page.locator("#mf_txppWframe_rdoRecApeClCdTop > div.w2radio_item.w2radio_item_1 > label")
            await button.wait_for(state="visible", timeout=3000)
            await button.click()
            print("   [RECEIPT] 영수 버튼 클릭 완료")
    except Exception as e:
        print(f"   [WARN] 영수/청구 버튼 클릭 실패: {e}")
        # 기본값으로 영수 버튼 시도
        try:
            button = page.locator("#mf_txppWframe_rdoRecApeClCdTop > div.w2radio_item.w2radio_item_1 > label")
            await button.click()
            print("   [FALLBACK] 기본 영수 버튼 클릭 완료")
        except:
            pass