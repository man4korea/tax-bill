# -*- coding: utf-8 -*-
"""
HomeTax 유틸리티 함수 모듈
Common utility functions for HomeTax automation
"""

import asyncio
import pandas as pd
import winsound
from typing import List, Dict, Any, Optional


async def play_beep(count: int = 1, frequency: int = 800, duration: int = 300):
    """지정된 횟수만큼 Beep음을 재생"""
    try:
        print(f"      [BEEP] 알림 {count}회...")
        for i in range(count):
            winsound.Beep(frequency, duration)
            if i < count - 1:
                await asyncio.sleep(0.2)
        print("      [BEEP] 알림 완료")
    except Exception as beep_error:
        print(f"      Beep 처리 오류: {beep_error}")


def format_date(value) -> str:
    """날짜 형식 변환 (YYYY-MM-DD → YYYYMMDD)"""
    if pd.isna(value) or not value:
        return ""
    date_str = str(value).replace('-', '').replace('/', '').replace('.', '')
    return date_str[:8] if len(date_str) >= 8 else date_str


def format_business_number(value) -> str:
    """사업자번호 형식 변환"""
    if pd.isna(value) or not value:
        return ""
    number_str = str(value).replace('-', '').replace(' ', '')
    return number_str[:10] if len(number_str) >= 10 else number_str


def format_number(value) -> str:
    """숫자 형식 변환 (콤마 제거)"""
    if pd.isna(value) or not value:
        return ""
    return str(value).replace(',', '').strip()


def clean_string_value(value) -> str:
    """문자열 값 정리"""
    if pd.isna(value) or not value:
        return ""
    return str(value).strip()


class FieldCollector:
    """필드 값 수집을 위한 유틸리티 클래스"""
    
    @staticmethod
    async def get_field_value(page, selector: str, field_name: str, wait_time: int = 3000) -> str:
        """필드 값을 다양한 방법으로 시도하여 수집"""
        try:
            element = page.locator(selector)
            
            try:
                await element.wait_for(state="visible", timeout=wait_time)
            except:
                await page.wait_for_timeout(500)
            
            # 여러 방법 시도
            methods = [
                ("input_value", lambda el: el.input_value()),
                ("attribute", lambda el: el.get_attribute("value")),
                ("text_content", lambda el: el.text_content()),
                ("inner_text", lambda el: el.inner_text()),
                ("evaluate", lambda el: el.evaluate("el => el.value || el.textContent || el.innerText"))
            ]
            
            for method_name, method_func in methods:
                try:
                    value = await method_func(element)
                    if value and str(value).strip():
                        print(f"   [OK] {field_name} 수집 성공 ({method_name}): '{value}'")
                        return str(value).strip()
                except Exception as e:
                    print(f"   [WARN] {field_name} {method_name} 실패: {e}")
                    continue
            
            print(f"   [ERROR] {field_name} 값 수집 실패 - 모든 방법 실패")
            return ""
            
        except Exception as e:
            print(f"   [ERROR] {field_name} 필드 접근 실패: {e}")
            return ""


class SelectorManager:
    """선택자 관리 클래스"""
    
    # 공통 선택자 그룹
    COMPANY_NAME_SELECTORS = [
        "#mf_txppWframe_edtDmnrTnmNmTop",
        "#mf_txppWframe_edtDmnrTnmNmTop_input",
        "input[id*='DmnrTnmNm']",
        "input[placeholder*='상호']"
    ]
    
    EMAIL_ID_SELECTORS = [
        "#mf_txppWframe_edtDmnrMchrgEmlIdTop",
        "#mf_txppWframe_edtDmnrMchrgEmlIdTop_input",
        "input[id*='MchrgEmlId']"
    ]
    
    EMAIL_DOMAIN_SELECTORS = [
        "#mf_txppWframe_edtDmnrMchrgEmlDmanTop",
        "#mf_txppWframe_edtDmnrMchrgEmlDmanTop_input",
        "input[id*='MchrgEmlDman']"
    ]
    
    SUPPLY_AMOUNT_SELECTORS = [
        "#mf_txppWframe_edtSumSplCftHeaderTop",
        "#mf_txppWframe_edtSumSplCftHeaderTop_input",
        "input[id*='SumSplCft']"
    ]
    
    TAX_AMOUNT_SELECTORS = [
        "#mf_txppWframe_edtSumTxamtHeaderTop",
        "#mf_txppWframe_edtSumTxamtHeaderTop_input",
        "input[id*='SumTxamt']"
    ]
    
    TOTAL_AMOUNT_SELECTORS = [
        "#mf_txppWframe_edtTotaAmtHeaderTop",
        "#mf_txppWframe_edtTotaAmtHeaderTop_input",
        "input[id*='TotaAmt']"
    ]


class MenuNavigator:
    """메뉴 네비게이션 유틸리티"""
    
    @staticmethod
    async def click_menu_with_fallback(page, selectors: List[str], menu_name: str, wait_time: int = 10000) -> bool:
        """여러 선택자로 메뉴 클릭 시도"""
        print(f"   {menu_name} 선택 시도...")
        
        for selector in selectors:
            try:
                print(f"   시도: {selector}")
                element = page.locator(selector).first
                await element.wait_for(state="visible", timeout=3000)
                await element.click()
                print(f"   {menu_name} 클릭 성공: {selector}")
                return True
            except:
                continue
        
        print(f"   {menu_name}를 찾을 수 없습니다 - 수동으로 선택하세요")
        await page.wait_for_timeout(wait_time)
        return False


class DialogHandler:
    """다이얼로그 처리 유틸리티"""
    
    @staticmethod
    async def handle_consecutive_dialogs(page, max_dialogs: int = 2) -> bool:
        """연속된 다이얼로그 처리"""
        dialog_count = 0
        all_handled = False
        
        async def handle_dialog(dialog):
            nonlocal dialog_count, all_handled
            dialog_count += 1
            print(f"   [ALERT] Alert {dialog_count}/{max_dialogs}: {dialog.message}")
            await dialog.accept()
            
            if dialog_count >= max_dialogs:
                all_handled = True
            else:
                page.once("dialog", handle_dialog)
        
        page.once("dialog", handle_dialog)
        
        # 최대 10초 대기
        wait_time = 0
        while not all_handled and wait_time < 10.0:
            await page.wait_for_timeout(200)
            wait_time += 0.2
        
        return all_handled


def get_date_columns() -> List[str]:
    """날짜 컬럼 후보 목록"""
    return ['공급일자', '작성일자', '일자', '날짜', 'supply_date', 'date']


def get_item_name_columns() -> List[str]:
    """품목명 컬럼 후보 목록"""
    return ['품목명', '품명', '품목', 'item_name', 'item', 'product_name', '상품명']


def get_cash_amount_columns() -> List[str]:
    """현금금액 컬럼 후보 목록"""
    return ['현금금액', '현금', 'cash_amount']


def validate_page_state(page) -> bool:
    """페이지 상태 검증"""
    try:
        return not page.is_closed()
    except:
        return False


def format_date_range(work_rows: List[Dict], single_format: str = "%Y-%m-%d", multi_format: str = "%y%m%d") -> str:
    """날짜 범위 형식화"""
    if len(work_rows) == 1:
        single_date = work_rows[0].get('공급일자') or work_rows[0].get('작성일자', '')
        if single_date:
            try:
                date_obj = pd.to_datetime(single_date)
                return date_obj.strftime(single_format)
            except:
                return str(single_date)
        return ""
    else:
        start_date = work_rows[0].get('공급일자') or work_rows[0].get('작성일자', '')
        end_date = work_rows[-1].get('공급일자') or work_rows[-1].get('작성일자', '')
        
        try:
            start_formatted = pd.to_datetime(start_date).strftime(multi_format) if start_date else ""
            end_formatted = pd.to_datetime(end_date).strftime(multi_format) if end_date else ""
            
            if start_formatted and end_formatted and start_formatted != end_formatted:
                return f"{start_formatted}-{end_formatted} {len(work_rows)}건"
            elif start_formatted:
                return f"{start_formatted} {len(work_rows)}건"
        except:
            pass
        
        return f"{len(work_rows)}건"