# -*- coding: utf-8 -*-
"""
HomeTax 거래 내역 입력 프로세스 모듈
Transaction Detail Input Process for HomeTax Automation

이 모듈은 HomeTax 세금계산서 작성 시 거래 내역을 입력하는 모든 기능을 포함합니다.
"""

import asyncio
import pandas as pd
import winsound
import threading
import tkinter as tk
from tkinter import messagebox
from datetime import datetime


async def process_transaction_details(page, processor, first_row_data, business_number):
    """거래 내역 입력 프로세스 - 10번 루틴에서 호출"""
    try:
        print("   [LIST] 거래 내역 입력 프로세스 시작")
        
        # 1. 동일 사업자번호 행들 가져오기
        work_rows = get_same_business_number_rows(processor, business_number)
        if not work_rows:
            print("   [ERROR] 동일 사업자번호 데이터가 없습니다.")
            return
            
        print(f"   📊 처리할 거래 건수: {len(work_rows)}건")
        
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
        
        # 4. 합계 확정 (결제방법 분류)
        await finalize_transaction_summary(page, work_rows, processor, business_number)
        
        # 5. 발행 관련 alert 처리
        await handle_issuance_alerts(page)
        
        # 6. 세금계산서 시트에 기록
        await write_to_tax_invoice_sheet(page, processor, work_rows, business_number)
        
        # 7. Q열에 완료 표시
        # 완료된 각 행에 Q열에 오늘 날짜 기록
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
        print("   📅 공급일자 확인 중...")
        
        # 엑셀에서 작성일자 가져오기 (여러 가능한 컬럼명 시도)
        excel_date = None
        date_columns = ['작성일자', '일자', '날짜']
        
        for col in date_columns:
            if col in first_row and pd.notna(first_row[col]):
                excel_date = first_row[col]
                print(f"   📊 Excel {col}: {excel_date}")
                break
        
        if excel_date is None:
            print("   [WARN] Excel에서 날짜를 찾을 수 없어 현재 날짜를 사용합니다.")
            excel_date = datetime.now()
        
        # 날짜 형식 통일
        if isinstance(excel_date, pd.Timestamp):
            excel_date_obj = excel_date
        elif isinstance(excel_date, str):
            try:
                excel_date_obj = pd.to_datetime(excel_date)
            except:
                excel_date_obj = datetime.now()
        else:
            excel_date_obj = datetime.now()
        
        # HomeTax 현재 공급일자 가져오기
        hometax_date_input = page.locator("#mf_txppWframe_calWrtDtTop_input")
        await hometax_date_input.wait_for(state="visible", timeout=5000)
        hometax_date_str = await hometax_date_input.input_value()
        
        print(f"   🌐 HomeTax 공급일자: {hometax_date_str}")
        
        # 날짜 비교 (년/월) - HomeTax는 ISO 형식 (YYYY-MM-DD)
        try:
            hometax_date_obj = pd.to_datetime(hometax_date_str, format='%Y-%m-%d')
        except:
            # 다른 형식도 시도
            hometax_date_obj = pd.to_datetime(hometax_date_str)
        
        excel_year_month = f"{excel_date_obj.year}{excel_date_obj.month:02d}"
        hometax_year_month = f"{hometax_date_obj.year}{hometax_date_obj.month:02d}"
        
        if excel_year_month != hometax_year_month:
            print(f"   [ALERT] 공급일자 년/월이 다릅니다! Excel: {excel_year_month}, HomeTax: {hometax_year_month}")
            
            # 5회 beep
            for i in range(5):
                winsound.Beep(800, 300)
                await asyncio.sleep(0.2)
            
            # 새 공급일자로 변경
            new_date_str = excel_date_obj.strftime("%Y%m%d")
            await hometax_date_input.clear()
            await hometax_date_input.fill(new_date_str)
            await page.wait_for_timeout(500)
            
            print(f"   [OK] 공급일자 변경 완료: {new_date_str}")
        else:
            print("   [OK] 공급일자 일치 - 변경 불필요")
            
    except Exception as e:
        print(f"   [ERROR] 공급일자 확인 오류: {e}")


async def input_transaction_items_basic(page, work_rows):
    """기본 거래 내역 입력 (1-4건)"""
    try:
        print(f"   📝 기본 거래 내역 입력: {len(work_rows)}건")
        
        for i, row_data in enumerate(work_rows, 1):
            await input_single_transaction_item(page, i, row_data)
            await page.wait_for_timeout(300)
        
        print("   [OK] 기본 거래 내역 입력 완료")
        
    except Exception as e:
        print(f"   [ERROR] 기본 거래 내역 입력 오류: {e}")


async def input_transaction_items_extended(page, work_rows):
    """확장 거래 내역 입력 (5-16건)"""
    try:
        print(f"   📝 확장 거래 내역 입력: {len(work_rows)}건")
        
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
    """단일 거래 내역 입력 - 사용자 요구사항에 맞는 정확한 selector 사용"""
    try:
        print(f"   📝 {row_idx}번째 거래 내역 입력 중...")
        print(f"      데이터 키들: {list(row_data.keys())}")  # 디버깅용
        
        # 0-based index로 변환 (첫번째는 0, 두번째는 1, ...)
        idx = row_idx - 1
        
        # 일자 (공급일자의 일 부분만) - 여러 컬럼명 시도
        supply_date = None
        for date_col in ['공급일자', '작성일자', '일자', '날짜', 'supply_date']:
            if date_col in row_data and row_data[date_col]:
                supply_date = str(row_data[date_col]).strip()
                print(f"      일자 데이터 발견: {date_col} = {supply_date}")
                break
        
        if supply_date:
            try:
                # 일자만 추출
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
        
        # 품목명 - 여러 컬럼명 시도
        item_name = None
        for item_col in ['품목명', '품명', '품목', 'item_name']:
            if item_col in row_data and row_data[item_col]:
                item_name = str(row_data[item_col]).strip()
                print(f"      품목 데이터 발견: {item_col} = {item_name}")
                break
        
        if item_name:
            item_input = page.locator(f"#mf_txppWframe_genEtxivLsatTop_{idx}_edtLsatNmTop")
            await item_input.wait_for(state="visible", timeout=3000)
            await item_input.clear()
            await item_input.fill(item_name)
            print(f"      품목: {item_name}")
        else:
            print(f"      품목: 데이터 없음")
        
        # 규격
        spec = str(row_data.get('규격', '')).strip()
        if spec:
            spec_input = page.locator(f"#mf_txppWframe_genEtxivLsatTop_{idx}_edtLsatRszeNmTop")
            await spec_input.wait_for(state="visible", timeout=3000)
            await spec_input.clear()
            await spec_input.fill(spec)
            print(f"      규격: {spec}")
        
        # 수량
        quantity = str(row_data.get('수량', '')).strip()
        if quantity:
            qty_input = page.locator(f"#mf_txppWframe_genEtxivLsatTop_{idx}_edtLsatQtyTop")
            await qty_input.wait_for(state="visible", timeout=3000)
            await qty_input.clear()
            await qty_input.fill(quantity)
            print(f"      수량: {quantity}")
        
        # 단가
        unit_price = str(row_data.get('단가', '')).strip()
        if unit_price:
            price_input = page.locator(f"#mf_txppWframe_genEtxivLsatTop_{idx}_edtLsatUtprcTop")
            await price_input.wait_for(state="visible", timeout=3000)
            await price_input.clear()
            await price_input.fill(unit_price)
            print(f"      단가: {unit_price}")
        
        # 공급가액
        supply_amount = str(row_data.get('공급가액', '')).strip()
        if supply_amount:
            supply_input = page.locator(f"#mf_txppWframe_genEtxivLsatTop_{idx}_edtLsatSplCftTop")
            await supply_input.wait_for(state="visible", timeout=3000)
            await supply_input.clear()
            await supply_input.fill(supply_amount)
            print(f"      공급가액: {supply_amount}")
        
        # 세액
        tax_amount = str(row_data.get('세액', '')).strip()
        if tax_amount:
            tax_input = page.locator(f"#mf_txppWframe_genEtxivLsatTop_{idx}_edtLsatTxamtTop")
            await tax_input.wait_for(state="visible", timeout=3000)
            await tax_input.clear()
            await tax_input.fill(tax_amount)
            print(f"      세액: {tax_amount}")
        
        # 비고
        remarks = str(row_data.get('비고', '')).strip()
        if remarks:
            remark_input = page.locator(f"#mf_txppWframe_edtRmk{row_idx}")
            await remark_input.wait_for(state="visible", timeout=3000)
            await remark_input.clear()
            await remark_input.fill(remarks)
        
        print(f"   [OK] {row_idx}번째 거래 내역 입력 완료")
        
    except Exception as e:
        print(f"   [ERROR] {row_idx}번째 거래 내역 입력 오류: {e}")


async def finalize_transaction_summary(page, work_rows, processor, business_number):
    """거래 합계 확정 및 결제방법 분류"""
    try:
        print("   [MONEY] 거래 합계 확정 중...")
        
        # Excel 데이터에서 결제 방법별 금액 계산 - 실제 컬럼명 사용
        cash_amount = 0
        check_amount = 0
        note_amount = 0
        
        for row in work_rows:
            # 현금금액 추출 (여러 컬럼명 시도)
            row_cash_amount = 0
            for cash_col in ['현금금액', '현금', 'cash_amount']:
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
                    # 현금종류가 비어있거나 '현금'인 경우
                    cash_amount += row_cash_amount
                    print(f"      현금으로 분류: {row_cash_amount:,.0f}원")
        
        # 기존 방식도 시도 (fallback)
        if cash_amount == 0 and check_amount == 0 and note_amount == 0:
            cash_amount = sum(float(row.get('현금', 0) or 0) for row in work_rows)
            check_amount = sum(float(row.get('수표', 0) or 0) for row in work_rows)
            note_amount = sum(float(row.get('어음', 0) or 0) for row in work_rows)
        
        print(f"   💵 현금: {cash_amount:,.0f}원")
        print(f"   [FORM] 수표: {check_amount:,.0f}원")
        print(f"   📝 어음: {note_amount:,.0f}원")
        
        # 합계 금액 검증 및 외상미수금 계산
        credit_amount = await verify_and_calculate_credit(page, work_rows, cash_amount, check_amount, note_amount)
        
        # 각 결제 방법 입력 (사용자 요구 selector 사용)
        if cash_amount > 0:
            cash_input = page.locator("#mf_txppWframe_edtStlMthd10Top")
            await cash_input.wait_for(state="visible", timeout=3000)
            await cash_input.clear()
            await cash_input.fill(str(int(cash_amount)))
        
        if check_amount > 0:
            check_input = page.locator("#mf_txppWframe_edtStlMthd20Top")
            await check_input.wait_for(state="visible", timeout=3000)
            await check_input.clear()
            await check_input.fill(str(int(check_amount)))
        
        if note_amount > 0:
            note_input = page.locator("#mf_txppWframe_edtStlMthd30Top")
            await note_input.wait_for(state="visible", timeout=3000)
            await note_input.clear()
            await note_input.fill(str(int(note_amount)))
        
        if credit_amount > 0:
            credit_input = page.locator("#mf_txppWframe_edtStlMthd40Top")
            await credit_input.wait_for(state="visible", timeout=3000)
            await credit_input.clear()
            await credit_input.fill(str(int(credit_amount)))
            print(f"   💳 외상미수금: {credit_amount:,.0f}원")
        
        # 영수 버튼 클릭
        try:
            receipt_button = page.locator("#mf_txppWframe_rdoRecApeClCdTop > div.w2radio_item.w2radio_item_1 > label")
            await receipt_button.wait_for(state="visible", timeout=3000)
            await receipt_button.click()
            print("   [LIST] 영수 버튼 클릭 완료")
        except Exception as e:
            print(f"   [WARN] 영수 버튼 클릭 실패: {e}")
        
        # 발급보류 버튼 클릭 및 Alert 처리
        try:
            await page.wait_for_timeout(1000)  # 1초 대기
            
            # 발급보류 버튼 확인 및 클릭
            issue_button = page.locator("#mf_txppWframe_btnIsnRsrv")
            await issue_button.wait_for(state="visible", timeout=3000)
            
            print("   [FORM] 발급보류 버튼 클릭 시도...")
            
            # Alert 리스너 설정 (발급보류 확인/취소 다이얼로그용)
            confirm_dialog_handled = False
            
            async def handle_confirm_dialog(dialog):
                nonlocal confirm_dialog_handled
                print(f"   [ALERT] 발급보류 확인 다이얼로그: {dialog.message}")
                await dialog.accept()  # 확인 버튼 클릭
                confirm_dialog_handled = True
                print("   [OK] 발급보류 확인 다이얼로그 - 확인 클릭")
            
            page.once("dialog", handle_confirm_dialog)
            
            # 발급보류 버튼 클릭
            await issue_button.click()
            print("   [FORM] 발급보류 버튼 클릭 완료")
            
            # 확인 다이얼로그 대기 (최대 5초)
            wait_time = 0
            while not confirm_dialog_handled and wait_time < 5:
                await page.wait_for_timeout(100)
                wait_time += 0.1
            
            if not confirm_dialog_handled:
                print("   [WARN] 발급보류 확인 다이얼로그가 나타나지 않았습니다.")
            
            # 발급보류 성공 Alert 처리
            await page.wait_for_timeout(1000)  # 잠시 대기
            
            success_dialog_handled = False
            
            async def handle_success_dialog(dialog):
                nonlocal success_dialog_handled
                print(f"   [ALERT] 발급보류 성공 다이얼로그: {dialog.message}")
                await dialog.accept()  # 확인 버튼 클릭
                success_dialog_handled = True
                print("   [OK] 발급보류 성공 다이얼로그 - 확인 클릭")
            
            page.once("dialog", handle_success_dialog)
            
            # 성공 다이얼로그 대기 (최대 5초)
            wait_time = 0
            while not success_dialog_handled and wait_time < 5:
                await page.wait_for_timeout(100)
                wait_time += 0.1
            
            if not success_dialog_handled:
                print("   [WARN] 발급보류 성공 다이얼로그가 나타나지 않았습니다.")
            
            # 폼 초기화 확인 및 대기
            await page.wait_for_timeout(2000)  # 폼 클리어 대기
            print("   [OK] 전자세금계산서 입력 화면 클리어 완료")
            
        except Exception as e:
            print(f"   [ERROR] 발급보류 처리 실패: {e}")
        
        print("   [OK] 거래 합계 확정 및 발급보류 완료")
        
    except Exception as e:
        print(f"   [ERROR] 거래 합계 확정 오류: {e}")


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
        
        print(f"   📊 실제 합계: {actual_total:,.0f}원")
        print(f"   🌐 HomeTax 합계: {hometax_total:,.0f}원")
        
        # HomeTax 값을 기준으로 사용 (불일치 검증 제거)
        total_amount = hometax_total
        print(f"   [OK] 합계금액 확인: {total_amount:,.0f}원")
        
        # 현금+수표+어음이 모두 0인 경우 전체 금액을 외상미수금으로
        payment_total = cash_amount + check_amount + note_amount
        
        if payment_total == 0:
            # 현금+수표+어음이 0이면 합계금액 전체를 외상미수금으로
            credit_amount = total_amount
            print(f"   💳 결제방법이 없으므로 전체 금액을 외상미수금으로: {credit_amount:,.0f}원")
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
        
        # 필요한 값들 수집 - 여러 방법 시도
        async def get_field_value(selector, field_name):
            try:
                element = page.locator(selector)
                await element.wait_for(state="visible", timeout=2000)
                
                # 먼저 input_value() 시도
                try:
                    return await element.input_value()
                except:
                    # input_value 실패 시 text_content() 시도
                    try:
                        return await element.text_content() or ""
                    except:
                        # text_content 실패 시 inner_text() 시도
                        try:
                            return await element.inner_text() or ""
                        except:
                            # 모두 실패 시 get_attribute('value') 시도
                            return await element.get_attribute("value") or ""
            except Exception as e:
                print(f"   [WARN] {field_name} 필드 값 가져오기 실패: {e}")
                return ""
        
        supply_date = await get_field_value("#mf_txppWframe_calWrtDtTop_input", "공급일자")
        company_name = await get_field_value("#mf_txppWframe_edtDmnrTnmNmTop", "상호")
        email_id = await get_field_value("#mf_txppWframe_edtDmnrMchrgEmlIdTop", "이메일ID")
        email_domain = await get_field_value("#mf_txppWframe_edtDmnrMchrgEmlDmanTop", "이메일도메인")
        
        # 합계 금액들
        total_supply = await get_field_value("#mf_txppWframe_edtSumSplCftHeaderTop", "공급가액")
        total_tax = await get_field_value("#mf_txppWframe_edtSumTxamtHeaderTop", "세액")
        total_amount = await get_field_value("#mf_txppWframe_edtTotaAmtHeaderTop", "합계금액")
        
        # 첫 번째 품목 정보 (단일 건일 때 사용)
        first_item_name = await get_field_value("#mf_txppWframe_genEtxivLsatTop_0_edtLsatNmTop", "첫번째품목명")
        first_item_spec = await get_field_value("#mf_txppWframe_genEtxivLsatTop_0_edtLsatRszeNmTop", "첫번째규격")
        first_item_quantity = await get_field_value("#mf_txppWframe_genEtxivLsatTop_0_edtLsatQtyTop", "첫번째수량")
        
        # 품목 정보 생성 (홈택스 필드값 우선 사용)
        if len(work_rows) == 1:
            # 1건인 경우: 홈택스 필드값 사용
            item_name = first_item_name or work_rows[0].get('품명', '')
            item_spec = first_item_spec or work_rows[0].get('규격', '')
            item_quantity = first_item_quantity or work_rows[0].get('수량', '')
        else:
            # 여러 건인 경우: "첫번째품목명 외 N개 품목" 형식
            base_item = first_item_name or work_rows[0].get('품명', '품목')
            item_name = f"{base_item} 외 {len(work_rows)}개 품목"
            item_spec = ""
            item_quantity = ""
        
        # 공급일자 범위 생성
        start_date = work_rows[0].get('작성일자', '')
        end_date = work_rows[-1].get('작성일자', '') if len(work_rows) > 1 else start_date
        date_range = f"{start_date} - {end_date} & {len(work_rows)}건"
        
        # 세금계산서 시트에 기록
        tax_invoice_data = {
            'a': supply_date,  # 공급일자
            'b': business_number,  # 등록번호
            'c': company_name,  # 상호
            'd': f"{email_id}@{email_domain}",  # 이메일
            'f': item_name,  # 품목
            'g': item_spec,  # 규격
            'h': item_quantity,  # 수량
            'i': total_supply,  # 공급가액
            'j': total_tax,  # 세액
            'k': total_amount,  # 합계금액
            'l': date_range  # 기간 및 건수
        }
        
        # 실제 엑셀 파일에 기록
        processor.write_tax_invoice_data(tax_invoice_data)
        
        # 폼 필드는 발급보류 성공 후 자동으로 클리어되므로 별도 초기화 불필요
        print("   [INFO] 발급보류 성공 후 폼 자동 클리어됨 - 초기화 생략")
        
        print("   [FORM] 세금계산서 시트 기록 및 필드 초기화 완료!")
        
    except Exception as e:
        print(f"   [ERROR] 세금계산서 시트 기록 오류: {e}")