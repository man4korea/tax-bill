# 📁 C:\APP\tax-bill\core\hometax_partner_registration.py
# Create at 2508312118 Ver1.00
# -*- coding: utf-8 -*-
"""
HomeTax 거래처 등록 자동화 프로그램 (엑셀 통합 버전)
1. 엑셀 파일 열기/확인
2. 행 선택 GUI
3. HomeTax 자동 로그인 및 수동 로그인 여부 파악 
   자동 혹은 수동 로그인 완료 후 거래처 등록 화면 이동
4. 엑셀에서 가져온 거래처 등록번호로 오류체크
5. 홈텍스에 거래처 등록
6. 결과 엑셀에 기록 (성공: 오늘 날짜, 실패: error)
"""

# Windows 콘솔 유니코드 출력 설정
import sys
import io
if sys.platform.startswith('win'):
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

import asyncio
import os
import subprocess
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from dotenv import load_dotenv
from playwright.async_api import async_playwright
import pandas as pd
from pathlib import Path
import re

# 보안 관리자 import
sys.path.append(str(Path(__file__).parent.parent / "core"))
from hometax_security_manager import HomeTaxSecurityManager

# 통합 엑셀 처리 모듈 import
from excel_unified_processor import create_partner_processor

def check_and_install_dependencies():
    """필수 의존성 패키지 확인 및 자동 설치"""
    required_packages = {
        'xlwings': 'xlwings>=0.30.0',
        'openpyxl': 'openpyxl>=3.1.0'
    }
    
    missing_packages = []
    
    for package_name, package_spec in required_packages.items():
        try:
            __import__(package_name)
            print(f"[OK] {package_name} 설치됨")
        except ImportError:
            missing_packages.append(package_spec)
            print(f"❌ {package_name} 미설치")
    
    if missing_packages:
        print(f"\n📦 {len(missing_packages)}개의 패키지를 설치합니다...")
        for package in missing_packages:
            try:
                print(f"설치 중: {package}")
                subprocess.check_call([sys.executable, "-m", "pip", "install", package])
                print(f"✅ {package} 설치 완료")
            except subprocess.CalledProcessError as e:
                print(f"❌ {package} 설치 실패: {e}")
                print(f"수동으로 설치하세요: pip install {package}")
        print("📦 패키지 설치가 완료되었습니다.\n")
    else:
        print("✅ 모든 필수 패키지가 설치되어 있습니다.\n")

class ExcelRowSelector:
    """ExcelUnifiedProcessor 어댑터 클래스 - 기존 인터페이스 호환성 유지"""
    
    def __init__(self):
        # 통합 프로세서 생성 - 거래처 시트용
        self.processor = create_partner_processor()
        
        # 기존 인터페이스 호환을 위한 속성들
        self.selected_rows = None
        self.selected_data = None
        self.excel_file_path = None
        self.headers = None
        self.processed_data = []
        self.field_mapping = {}
    
    def initialize(self):
        """초기화 - 파일 열기 및 컴포넌트 생성"""
        if not self.processor.initialize():
            return False
        
        # 호환성을 위한 속성 동기화
        self.excel_file_path = self.processor.file_manager.excel_file_path
        return True
    
    def check_and_open_excel(self):
        """엑셀 파일 확인 및 열기"""
        return self.initialize()
    
    def show_row_selection_gui(self):
        """행 선택 GUI 표시"""
        if not self.processor.select_rows():
            return False
        
        # 호환성을 위한 속성 동기화
        self.selected_rows = self.processor.get_selected_rows()
        
        # 첫 번째 행의 첫 번째 열 값 저장 (기존 로직 유지)
        if self.selected_rows and self.excel_file_path:
            try:
                import pandas as pd
                from openpyxl import load_workbook
                
                wb = load_workbook(self.excel_file_path)
                ws = wb.active
                max_row = ws.max_row
                
                df = pd.read_excel(self.excel_file_path, sheet_name="거래처", header=None, dtype=str, 
                                 keep_default_na=False, engine='openpyxl', na_filter=False, nrows=max_row)
                
                first_row = self.selected_rows[0]
                if first_row <= len(df) and len(df.columns) > 0:
                    self.selected_data = df.iloc[first_row-1, 0]
                    print(f"✅ 첫 번째 행({first_row})의 첫 번째 열 값 저장: {self.selected_data}")
                else:
                    self.selected_data = None
                    
            except Exception as e:
                print(f"❌ 엑셀 데이터 읽기 실패: {e}")
                self.selected_data = None
        
        return True
    
    def process_excel_data(self):
        """엑셀 데이터 처리"""
        if not self.processor.process_data():
            return False
        
        # 호환성을 위한 데이터 동기화
        processed_data = self.processor.get_processed_data()
        
        # 기존 형식으로 변환
        self.processed_data = processed_data
        
        if processed_data:
            self.headers = list(processed_data[0]['data'].keys())
            print(f"헤더: {self.headers}")
        
        return True
    
    def load_field_mapping(self):
        """field_mapping.md 파일을 읽어서 매핑 정보 추출"""
        mapping_file = Path(__file__).parent / "field_mapping.md"
        print(f"[DEBUG] 매핑 파일 경로: {mapping_file}")
        print(f"[DEBUG] 파일 존재 여부: {mapping_file.exists()}")
        
        if not mapping_file.exists():
            print(f"❌ {mapping_file} 파일을 찾을 수 없습니다.")
            return False
        
        try:
            with open(mapping_file, 'r', encoding='utf-8') as f:
                content = f.read()
            
            # 테이블 파싱 (| 구분자 사용)
            lines = content.split('\n')
            mapping_found = False
            
            for line in lines:
                line = line.strip()
                if not line or not line.startswith('|'):
                    continue
                
                # 테이블 헤더나 구분선 스킵
                if '입력화면 라벨명' in line or ':--' in line:
                    mapping_found = True
                    continue
                
                if mapping_found and line.startswith('|'):
                    # 테이블 행 파싱
                    parts = [p.strip() for p in line.split('|')]
                    if len(parts) >= 5:  # |빈칸|라벨|변수명|Excel열명|셀렉터|
                        excel_column = parts[3]  # Excel 열명
                        selector = parts[4]      # HomeTax 셀렉션명
                        
                        if excel_column and selector and excel_column != "Excel 열명":
                            self.field_mapping[excel_column] = {
                                'selector': selector,
                                'label': parts[1] if len(parts) > 1 else '',
                                'variable': parts[2] if len(parts) > 2 else ''
                            }
            
            print(f"✅ 필드 매핑 로드 완료: {len(self.field_mapping)}개 필드")
            
            # 매핑 정보 일부 출력
            print("매핑 예시:")
            count = 0
            for excel_col, info in self.field_mapping.items():
                if count < 3 and info['selector']:
                    print(f"  - {excel_col} → {info['selector']}")
                    count += 1
            
            return True
            
        except Exception as e:
            print(f"❌ 필드 매핑 파일 로드 실패: {e}")
            return False
    
    def write_error_to_excel(self, row_number, error_message="error"):
        """에러 상태 기록"""
        return self.processor.record_error(row_number, error_message)
    
    def write_today_to_excel(self, row_number):
        """성공 상태 기록 (오늘 날짜)"""
        return self.processor.record_success(row_number)
    
    def split_email(self, email_str):
        """이메일 주소를 @ 기준으로 분리"""
        if pd.isna(email_str) or not str(email_str).strip():
            return "", ""
        
        email_str = str(email_str).strip()
        if '@' in email_str:
            parts = email_str.split('@', 1)  # 첫 번째 @에서만 분리
            return parts[0].strip(), parts[1].strip()
        else:
            # @ 없는 경우 전체를 앞부분으로 처리
            return email_str, ""

async def prepare_next_registration(page):
    """다음 거래처 등록을 위한 페이지 준비"""
    try:
        print("다음 거래처 등록을 위한 페이지 준비...")
        
        # 1. 사업자번호 필드 찾기 및 클리어
        business_number_selectors = [
            "#mf_txppWframe_txtBsno1",      # 기본 사업자번호 필드
            "input[name*='txtBsno']",       # 사업자번호 관련 필드
            "input[id*='Bsno']",           # Bsno가 포함된 ID
            "input[placeholder*='사업자']",  # placeholder에 사업자가 포함된 필드
            "input[title*='사업자']",       # title에 사업자가 포함된 필드
        ]
        
        business_field = None
        for selector in business_number_selectors:
            try:
                business_field = page.locator(selector).first
                await business_field.wait_for(state="visible", timeout=1000)
                print(f"  ✅ 사업자번호 입력 필드 찾음: {selector}")
                break
            except:
                continue
        
        if not business_field:
            print("  ⚠️ 사업자번호 입력 필드를 찾을 수 없습니다.")
            return False
        
        # 2. 필드 클리어 및 포커스 설정
        try:
            # 필드 클리어
            await business_field.clear()
            
            # 포커스 설정
            await business_field.focus()
            
            # 잠시 대기
            await page.wait_for_timeout(1000)
            
            print("  ✅ 사업자번호 필드 클리어 및 포커스 설정 완료")
            return True
            
        except Exception as e:
            print(f"  ❌ 필드 클리어/포커스 설정 실패: {e}")
            return False
            
    except Exception as e:
        print(f"  ❌ 페이지 준비 실패: {e}")
        return False

async def fill_hometax_form(page, row_data, field_mapping, excel_selector, current_row_number, is_first_record=False):
    """HomeTax 폼에 데이터 자동 입력"""
    print("\n=== 거래처 데이터 입력 시작 ===")
    
    # 첫 번째 거래처가 아닌 경우 페이지 준비
    if not is_first_record:
        if not await prepare_next_registration(page):
            raise Exception("다음 거래처 등록을 위한 페이지 준비에 실패했습니다.")
    
    success_count = 0
    failed_fields = []

    # 2. 사업자번호 입력 및 확인 (전달받은 page_context 사용)
    business_number = row_data.get('사업자번호') or row_data.get('사업자등록번호') or row_data.get('거래처등록번호')
    if business_number and ('사업자번호' in field_mapping or '사업자등록번호' in field_mapping or '거래처등록번호' in field_mapping):
        try:
            # 사업자번호 필드 찾기
            if '사업자번호' in field_mapping:
                business_field = '사업자번호'
            elif '사업자등록번호' in field_mapping:
                business_field = '사업자등록번호'
            else:
                business_field = '거래처등록번호'
            selector = field_mapping[business_field]['selector'].strip()
            
            print(f"사업자번호 입력: {business_number} → {selector}")
            
            # 사업자번호 입력
            element = page.locator(selector).first
            await element.wait_for(state="visible", timeout=1000)
            await element.clear()
            await element.fill(str(business_number))
            await page.wait_for_timeout(1000)
            
            print(f"  ✅ 사업자번호 입력 완료")
            success_count += 1
            
            # 사업자번호 확인 버튼 클릭 및 검증
            await handle_business_number_validation(page, business_number, excel_selector, current_row_number)
            
        except Exception as e:
            print(f"  ❌ 사업자번호 처리 실패: {e}")
            if "BUSINESS_NUMBER_ERROR" in str(e):
                error_message = str(e).split("|")[1] if "|" in str(e) else "사업자번호 오류"
                excel_selector.write_error_to_excel(current_row_number, "error")
                import tkinter as tk
                from tkinter import messagebox
                root = tk.Tk()
                root.withdraw()
                messagebox.showerror(
                    "사업자번호 오류", 
                    f"사업자번호가 올바르지 않습니다.\n\n행 번호: {current_row_number}\n사업자번호: {business_number}\n메시지: {error_message}\n\n엑셀 파일에 'error'가 기록되었습니다.\n프로그램을 종료합니다."
                )
                root.destroy()
            raise e
    
    # 3. 나머지 필드들 입력 (새 팝업 페이지에서 수행)
    for excel_column, value in row_data.items():
        if excel_column in ['사업자번호', '사업자등록번호', '거래처등록번호']:
            continue
        
        if not value:
            continue

        if excel_column not in field_mapping:
            print(f"  ⚠️ 매핑되지 않은 필드: '{excel_column}' (값: '{value}') - field_mapping.md에 해당 항목이 없거나 Excel 헤더가 다릅니다.")
            continue
        
        mapping_info = field_mapping[excel_column]
        selector = mapping_info['selector'].strip()
        
        if not selector:
            continue
        
        try:
            print(f"입력 중: {excel_column} = '{value}' → {selector}")
            element = page.locator(selector).first
            
            # 먼저 일반적인 방법으로 시도
            try:
                await element.clear(timeout=1000)
                await element.fill(str(value), timeout=1000)
                await page.wait_for_timeout(200)
                success_count += 1
                print(f"  ✅ 입력 완료")
                continue  # 성공하면 다음 필드로
            except Exception as normal_error:
                print(f"  ⚠️ 일반 입력 실패 ({normal_error}), JavaScript 방법 시도...")
                
                # JavaScript로 강제 입력 시도
                try:
                    # disabled 속성을 제거하고 값을 설정
                    await page.evaluate(f"""
                        const element = document.querySelector('{selector}');
                        if (element) {{
                            element.removeAttribute('disabled');
                            element.value = '{value}';
                            element.dispatchEvent(new Event('input', {{ bubbles: true }})));
                            element.dispatchEvent(new Event('change', {{ bubbles: true }})));
                        }}
                    """)
                    success_count += 1
                    print(f"  ✅ JavaScript로 입력 완료")
                except Exception as js_error:
                    failed_fields.append({'field': excel_column, 'selector': selector, 'error': f"일반: {normal_error}, JS: {js_error}"})
                    print(f"  ❌ JavaScript 입력도 실패: {js_error}")
                    
        except Exception as e:
            failed_fields.append({'field': excel_column, 'selector': selector, 'error': str(e)})
            print(f"  ❌ 입력 실패: {e}")

    
    # 4. 기타 특별 처리 필드들 (새 팝업 페이지에서 수행)
    await handle_other_special_fields(page, row_data, field_mapping)
    
    print(f"\n=== 입력 완료 ===")
    print(f"성공: {success_count}개 필드")
    if failed_fields:
        print(f"실패: {len(failed_fields)}개 필드")
        for failed in failed_fields:
            print(f"  - {failed['field']}: {failed['error']}")

    # 5. 최종 등록 버튼 클릭 및 Alert 처리
    try:
        print(f"최종 등록 버튼 클릭: #mf_txppWframe_btnRgt")
        
        # Alert 리스너 설정
        alert_handled = False
        alert_message = ""
        
        async def handle_final_alert(dialog):
            nonlocal alert_handled, alert_message
            alert_message = dialog.message
            print(f"등록 확인 Alert 메시지: {alert_message}")
            
            # 품목 등록 또는 담당자 추가 Alert인 경우 취소 클릭
            if "품목 등록" in alert_message:
                await dialog.dismiss()  # 취소 버튼 클릭
                print("  ✅ Alert의 '취소' 버튼을 클릭했습니다. (품목 등록 거부)")
            elif "담당자를 추가 등록" in alert_message:
                await dialog.dismiss()  # 취소 버튼 클릭
                print("  ✅ Alert의 '취소' 버튼을 클릭했습니다. (담당자 추가 등록 거부)")
            else:
                await dialog.accept()  # 확인 버튼 클릭
                print("  ✅ Alert의 '확인' 버튼을 클릭했습니다.")
            
            alert_handled = True

        page.on("dialog", handle_final_alert)

        # 등록 버튼 클릭 (여러 방법 시도)
        register_btn = page.locator("#mf_txppWframe_btnRgt").first
        
        try:
            # 방법 1: 일반 클릭
            await register_btn.click(timeout=1000)
            print("  ✅ 등록 버튼 클릭 성공 (일반 클릭)")
        except Exception as e1:
            print(f"  일반 클릭 실패: {e1}")
            try:
                # 방법 2: 강제 클릭 (다른 요소가 가리고 있어도 클릭)
                await register_btn.click(force=True, timeout=1000)
                print("  ✅ 등록 버튼 클릭 성공 (강제 클릭)")
            except Exception as e2:
                print(f"  강제 클릭 실패: {e2}")
                try:
                    # 방법 3: JavaScript를 통한 클릭
                    await page.evaluate("document.getElementById('mf_txppWframe_btnRgt').click()")
                    print("  ✅ 등록 버튼 클릭 성공 (JavaScript 클릭)")
                except Exception as e3:
                    print(f"  JavaScript 클릭도 실패: {e3}")
                    raise Exception("모든 등록 버튼 클릭 방법이 실패했습니다")

        # Alert 대기 (더 긴 시간)
        for i in range(100): # 10초 대기
            if alert_handled:
                break
            if i % 10 == 0:  # 1초마다 상태 출력
                print(f"  Alert 대기 중... {i//10 + 1}/10초")
            await page.wait_for_timeout(100)

        page.remove_listener("dialog", handle_final_alert)

        if alert_handled:
            print("  ✅ 등록 확인 Alert 처리 완료.")
            # 등록 성공 시 엑셀 파일에 오늘 날짜 기록
            excel_selector.write_today_to_excel(current_row_number)
        else:
            print("  ⚠️ 등록 확인 Alert가 나타나지 않았습니다. (정상일 수 있음)")
            # Alert가 없어도 등록이 성공한 것으로 간주하고 날짜 기록
            excel_selector.write_today_to_excel(current_row_number)

    except Exception as e:
        print(f"  ❌ 최종 등록 처리 실패: {e}")
        raise e

    return success_count, failed_fields


async def handle_business_number_validation(page, business_number, excel_selector, current_row_number):
    """사업자번호 확인 버튼 클릭 및 검증 (종사업장 선택 창 처리 포함)"""
    try:
        print(f"사업자번호 확인 버튼 클릭: {business_number}")
        
        # 확인 버튼 클릭 (여러 방법 시도)
        confirm_btn = page.locator("#mf_txppWframe_btnValidCheck").first
        
        try:
            # 방법 1: 일반 클릭
            await confirm_btn.click(timeout=1000)
            print("  ✅ 사업자번호 확인 버튼 클릭 성공 (일반 클릭)")
        except Exception as e1:
            print(f"  일반 클릭 실패: {e1}")
            try:
                # 방법 2: 강제 클릭
                await confirm_btn.click(force=True, timeout=1000)
                print("  ✅ 사업자번호 확인 버튼 클릭 성공 (강제 클릭)")
            except Exception as e2:
                print(f"  강제 클릭 실패: {e2}")
                try:
                    # 방법 3: JavaScript 클릭
                    await page.evaluate("document.getElementById('mf_txppWframe_btnValidCheck').click()")
                    print("  ✅ 사업자번호 확인 버튼 클릭 성공 (JavaScript 클릭)")
                except Exception as e3:
                    print(f"  JavaScript 클릭도 실패: {e3}")
                    print("  ⚠️ 사업자번호 확인 버튼을 클릭할 수 없습니다. 계속 진행합니다...")
                    return  # 버튼 클릭 실패해도 계속 진행

        # 잠시 대기 후 종사업장 선택 창 확인
        await page.wait_for_timeout(1000)
        
        # 1단계: 종사업장 선택 창이 나타났는지 확인
        workplace_popup_selectors = [
            "#mf_txppWframe_ABTIBsnoUnitPopup2",  # 종사업장 팝업
            ".popup:has-text('종사업장')",         # 종사업장 텍스트가 있는 팝업
            "[id*='BsnoUnit']",                   # BsnoUnit이 포함된 ID
        ]
        
        workplace_popup_found = False
        for selector in workplace_popup_selectors:
            try:
                element = page.locator(selector).first
                if await element.is_visible():
                    workplace_popup_found = True
                    print("  🏢 종사업장 선택 창이 나타났습니다!")
                    break
            except:
                continue
        
        if workplace_popup_found:
            print("  🔊 BEEP! 종사업장을 선택하고 확인 버튼을 클릭하세요!")
            print("  ⏳ 사용자 종사업장 선택 대기 중...")
            
            # 종사업장 확인 버튼 대기
            workplace_confirm_btn = page.locator("#mf_txppWframe_ABTIBsnoUnitPopup2_wframe_trigger66").first
            
            # 종사업장 확인 버튼이 클릭될 때까지 대기 (최대 60초) + 반복 beep
            for i in range(600):  # 60초 동안 0.1초마다 확인
                try:
                    # 버튼이 존재하는지 확인
                    if await workplace_confirm_btn.is_visible():
                        # 5초마다 beep 소리 울리기 (i % 50 == 0이면 5초마다)
                        if i % 50 == 0:  # 5초마다 (50 * 100ms = 5000ms)
                            try:
                                import winsound
                                winsound.Beep(1000, 300)  # 1000Hz 주파수로 300ms 동안 beep
                            except:
                                print("\a")  # 시스템 beep 소리
                        await page.wait_for_timeout(100)
                    else:
                        # 버튼이 사라졌으면 클릭된 것으로 간주
                        break
                except:
                    # 에러가 발생하면 팝업이 사라진 것으로 간주
                    break
            
            print("  ✅ 종사업장 확인 버튼이 클릭된 것으로 보입니다.")
            
            # 종사업장 확인 후 Alert 처리
            alert_handled = False
            alert_message = ""
            
            async def handle_workplace_alert(dialog):
                nonlocal alert_handled, alert_message
                alert_message = dialog.message
                print(f"  종사업장 확인 후 Alert 메시지: {alert_message}")
                await dialog.accept()
                alert_handled = True
            
            page.on("dialog", handle_workplace_alert)
            
            # Alert 대기 (최대 5초)
            for i in range(50):
                if alert_handled:
                    break
                await page.wait_for_timeout(100)
            
            page.remove_listener("dialog", handle_workplace_alert)
            
            if alert_handled:
                print(f"  ✅ 종사업장 선택 완료: {alert_message}")
            
        else:
            # 종사업장 선택 창이 없는 경우 일반 Alert 처리
            print("  📋 일반 사업자번호 확인 처리...")
            
            # Alert 이벤트 리스너 설정
            alert_handled = False
            alert_message = ""
            
            async def handle_alert(dialog):
                nonlocal alert_handled, alert_message
                alert_message = dialog.message
                print(f"  Alert 메시지: {alert_message}")
                await dialog.accept()
                alert_handled = True
            
            page.on("dialog", handle_alert)
            
            # Alert 대기 (최대 5초)
            for i in range(50):
                if alert_handled:
                    break
                await page.wait_for_timeout(100)
            
            page.remove_listener("dialog", handle_alert)
            
            if alert_handled:
                if "비정상적인 등록번호" in alert_message or "이미 등록된 사업자등록번호" in alert_message:
                    print(f"⚠️ 사업자번호 처리 불가: {alert_message}")
                    print("  ➡️ 다음 행으로 스킵합니다.")
                    raise Exception(f"SKIP_TO_NEXT_ROW|{alert_message}")
                elif "정상적인 사업자번호" in alert_message:
                    print("  ✅ 사업자번호 확인 완료 - 정상적인 사업자번호입니다.")
                else:
                    print(f"  ⚠️ 예상하지 못한 메시지: {alert_message}")
            else:
                print("  ⚠️ Alert 메시지를 받지 못했습니다.")
        
        # 필드들이 활성화되는지 대기
        print("  ⏳ 입력 필드 활성화 대기 중...")
        await page.wait_for_timeout(2000)
        
        # 거래처명 필드가 활성화되었는지 확인
        try:
            await page.wait_for_selector("#mf_txppWframe_txtTnmNm:not([disabled])", timeout=1000)
            print("  ✅ 입력 필드가 활성화되었습니다.")
        except:
            print("  ⚠️ 입력 필드 활성화를 확인할 수 없습니다. 계속 진행합니다...")
            
    except Exception as e:
        print(f"  ❌ 사업자번호 확인 처리 실패: {e}")
        # SKIP_TO_NEXT_ROW 예외는 다시 발생시켜서 상위에서 처리
        if "SKIP_TO_NEXT_ROW" in str(e):
            raise e
        # 기타 오류는 계속 진행
        print("  ⚠️ 사업자번호 확인에 실패했지만 계속 진행합니다...")
        return

async def handle_other_special_fields(page, row_data, field_mapping):
    """기타 특별 처리가 필요한 필드들 처리 (사업자번호 제외)"""
    
    # 이메일 직접입력 버튼들
    email_fields = [
        ('주이메일앞', '주이메일뒤'),
        ('부이메일앞', '부이메일뒤')
    ]
    
    for email_front, email_back in email_fields:
        if email_front in row_data and row_data[email_front]:
            try:
                # 직접입력 버튼 찾기
                if '주이메일' in email_front:
                    direct_btn = page.locator("#mf_txppWframe_btnMainEmailDirect").first
                else:
                    direct_btn = page.locator("#mf_txppWframe_btnSubEmailDirect").first
                
                await direct_btn.click(timeout=1000)
                await page.wait_for_timeout(300)
                print(f"  ✅ {email_front.replace('앞', '')} 직접입력 버튼 클릭 완료")
            except Exception as e:
                print(f"  ❌ {email_front.replace('앞', '')} 직접입력 버튼 클릭 실패: {e}")

# 공통 로그인 모듈에서 로그인 함수들을 import
from hometax_login_module import hometax_login_dispatcher


async def main():
    """메인 프로그램 실행"""
    print("=== HomeTax 거래처 등록 자동화 프로그램 ===")
    print("(엑셀 통합 버전)\n")
    
    # 필수 패키지 확인 및 설치
    check_and_install_dependencies()
    
    # 1. 엑셀 파일 확인 및 열기
    print("1단계: 엑셀 파일 확인 및 열기")
    excel_selector = ExcelRowSelector()
    
    if not excel_selector.check_and_open_excel():
        print("❌ 엑셀 파일 열기에 실패했습니다.")
        return
    
    # 2. 행 선택 GUI
    print("\n2단계: 행 선택")
    if not excel_selector.show_row_selection_gui():
        print("❌ 행 선택이 취소되었습니다.")
        return
    
    print(f"✅ 선택된 행: {excel_selector.selected_rows}")
    if excel_selector.selected_data is not None:
        print(f"✅ 첫 번째 행의 첫 번째 열 값: {excel_selector.selected_data}")
    
    # 2.5. 필드 매핑 로드
    print("\n2.5단계: 필드 매핑 로드")
    if not excel_selector.load_field_mapping():
        print("❌ 필드 매핑 로드에 실패했습니다.")
        return
    
    # 2.6. 엑셀 데이터 처리
    print("\n2.6단계: 엑셀 데이터 처리")
    if not excel_selector.process_excel_data():
        print("❌ 엑셀 데이터 처리에 실패했습니다.")
        return
    
    # 3. HomeTax 로그인 모듈 실행
    print("\n3단계: HomeTax 로그인 실행")
    
    # 로그인 완료 후 콜백으로 거래처 등록 메뉴로 이동
    async def login_callback(page=None, browser=None):
        print("\n3.5단계: 거래처 등록 메뉴 이동")
        if page:
            try:
                # 거래처등록 메뉴로 이동
                print("거래처등록 메뉴 클릭 시도...")
                await page.click("text=거래처등록")
                print("[OK] 거래처등록 메뉴 클릭 성공")
                return page, browser
            except Exception as e:
                print(f"[ERROR] 거래처등록 메뉴 이동 실패: {e}")
                return page, browser
        return page, browser
    
    # 로그인 및 메뉴 이동 실행
    result = await hometax_login_dispatcher(login_callback)
    page, browser = result if result else (None, None)
    
    # 4. 실제 거래처 등록 자동화 실행
    print("\n4단계: 거래처 등록 자동화 실행")
    # page, browser = await hometax_auto_partner()  # 이미 메뉴 이동 완료
    
    if page:
        print("\n✅ 거래처 등록 자동화 시작!")
        print("- 엑셀 파일이 열렸습니다.")
        print(f"- 선택된 행: {excel_selector.selected_rows}")
        print(f"- 처리된 데이터: {len(excel_selector.processed_data)}개 행")
        print("- HomeTax 화면에 접속했습니다.")
        
        try:
            # 선택된 데이터에 대해 거래처 등록 수행
            success_count = 0
            failed_count = 0
            
            for idx, row_info in enumerate(excel_selector.processed_data):
                current_row_number = row_info['row_number']
                row_data = row_info['data']
                
                print(f"\n--- 거래처 등록 {idx+1}/{len(excel_selector.processed_data)} (행번호: {current_row_number}) ---")
                
                try:
                    # 각 거래처에 대해 폼 입력 실행
                    is_first_record = (idx == 0)
                    success_count_fields, failed_fields = await fill_hometax_form(
                        page, row_data, excel_selector.field_mapping, 
                        excel_selector, current_row_number, is_first_record
                    )
                    
                    if success_count_fields > 0:
                        success_count += 1
                        print(f"✅ 거래처 등록 완료: 행 {current_row_number}")
                    else:
                        failed_count += 1
                        print(f"❌ 거래처 등록 실패: 행 {current_row_number}")
                        
                except Exception as e:
                    failed_count += 1
                    if "SKIP_TO_NEXT_ROW" in str(e):
                        print(f"⏭️ 행 {current_row_number} 스킵됨: {e}")
                        excel_selector.write_error_to_excel(current_row_number, "error")
                    else:
                        print(f"❌ 거래처 등록 중 오류 (행 {current_row_number}): {e}")
                        excel_selector.write_error_to_excel(current_row_number, "error")
                
                # 다음 거래처 등록을 위한 대기
                if idx < len(excel_selector.processed_data) - 1:
                    print("⏳ 다음 거래처 등록을 위해 3초 대기...")
                    await page.wait_for_timeout(3000)
            
            # 최종 결과 출력
            print(f"\n{'='*50}")
            print(f"🎉 거래처 등록 자동화 완료!")
            print(f"✅ 성공: {success_count}개")
            print(f"❌ 실패: {failed_count}개")
            print(f"📊 전체: {len(excel_selector.processed_data)}개")
            print(f"{'='*50}")
            
        except Exception as e:
            print(f"❌ 거래처 등록 자동화 중 전체 오류: {e}")
        
        # 브라우저 정리
        print("\n브라우저를 5초 후 종료합니다...")
        await page.wait_for_timeout(5000)
        
        if browser:
            try:
                await browser.close()
                print("✅ 브라우저가 정상적으로 종료되었습니다.")
            except Exception as e:
                print(f"❌ 브라우저 종료 중 오류: {e}")
    else:
        print("\n❌ HomeTax 자동화에 실패했습니다.")
        if browser:
            try:
                # 모든 페이지 닫기
                pages = browser.contexts[0].pages if browser.contexts else []
                for page in pages:
                    try:
                        await page.close()
                    except:
                        pass
                        
                await browser.close()
                print("브라우저가 정상적으로 종료되었습니다.")
            except Exception as e:
                print(f"브라우저 종료 중 오류: {e}")


if __name__ == "__main__":
    asyncio.run(main())
