# -*- coding: utf-8 -*-
"""
HomeTax 거래처 등록 자동화 프로그램 (엑셀 통합 버전)
1. 엑셀 파일 열기/확인
2. 행 선택 GUI
3. HomeTax 자동 로그인 및 거래처 등록 화면 이동
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
    def __init__(self):
        self.selected_rows = None
        self.selected_data = None
        self.excel_file_path = None
        self.headers = None  # 첫 번째 행(헤더) 데이터
        self.processed_data = []  # 처리된 데이터 리스트
        self.field_mapping = {}  # 필드 매핑 정보
    
    def write_error_to_excel(self, row_number, error_message="error"):
        """엑셀 파일의 지정된 행 첫 번째 열에 에러 메시지 작성"""
        if not self.excel_file_path:
            print("❌ 엑셀 파일 경로가 없습니다.")
            return False
        
        try:
            # 엑셀 파일이 열려있는 경우를 대비해 openpyxl 사용
            from openpyxl import load_workbook
            
            print(f"엑셀 파일에 에러 기록 중: 행 {row_number}, 메시지: {error_message}")
            
            # 엑셀 파일 로드
            workbook = load_workbook(self.excel_file_path)
            worksheet = workbook.active
            
            # 첫 번째 열(A열)에 에러 메시지 작성
            worksheet.cell(row=row_number, column=1, value=error_message)
            
            # 저장
            workbook.save(self.excel_file_path)
            workbook.close()
            
            print(f"✅ 엑셀 파일에 에러 기록 완료: 행 {row_number}")
            return True
            
        except Exception as e:
            print(f"❌ 엑셀 파일 에러 기록 실패: {e}")
            print("엑셀 파일이 열려있는 경우 파일을 닫고 다시 시도하세요.")
            return False
    
    def write_today_to_excel(self, row_number):
        """엑셀 파일의 지정된 행 첫 번째 열(A열)에 오늘 날짜 기록"""
        if not self.excel_file_path:
            print("❌ 엑셀 파일 경로가 없습니다.")
            return False
        
        try:
            from datetime import datetime
            
            # 오늘 날짜 생성
            today = datetime.now().strftime("%Y-%m-%d")
            
            print(f"엑셀 파일에 날짜 기록 중: 행 {row_number}, 날짜: {today}")
            
            # 방법 1: xlwings를 사용해서 열린 엑셀 파일에 직접 쓰기 시도
            try:
                import xlwings as xw
                
                # 현재 열려있는 엑셀 앱에 연결 (새로운 API 사용)
                try:
                    app = xw.apps.active
                except:
                    app = xw.App(visible=True, add_book=False)
                
                # 열린 워크북 찾기
                workbook_name = self.excel_file_path.split("\\")[-1]  # 파일명만 추출
                wb = None
                
                for book in app.books:
                    if book.name == workbook_name:
                        wb = book
                        break
                
                if wb:
                    # "거래처" 시트 선택
                    ws = None
                    for sheet in wb.sheets:
                        if sheet.name == "거래처":
                            ws = sheet
                            break
                    
                    if not ws:
                        ws = wb.sheets[0]  # 첫 번째 시트 사용
                    
                    # A열에 날짜 기록
                    ws.range(f'A{row_number}').value = today
                    
                    # 저장
                    wb.save()
                    
                    print(f"✅ 엑셀 파일에 날짜 기록 완료 (xlwings): 행 {row_number}, 날짜: {today}")
                    return True
                    
            except ImportError:
                print("  xlwings가 설치되지 않았습니다. openpyxl 방법을 시도합니다...")
                print("  💡 xlwings 설치 방법: pip install xlwings")
                print("  (xlwings를 설치하면 엑셀 파일이 열린 상태에서도 날짜를 직접 기록할 수 있습니다)")
            except Exception as e:
                print(f"  xlwings 방법 실패: {e}")
            
            # 방법 2: openpyxl 사용 (파일이 닫혀있을 때)
            from openpyxl import load_workbook
            
            workbook = load_workbook(self.excel_file_path)
            
            # "거래처" 시트 선택
            if "거래처" in workbook.sheetnames:
                worksheet = workbook["거래처"]
            else:
                worksheet = workbook.active
            
            # 첫 번째 열(A열)에 오늘 날짜 작성
            worksheet.cell(row=row_number, column=1, value=today)
            
            # 저장
            workbook.save(self.excel_file_path)
            workbook.close()
            
            print(f"✅ 엑셀 파일에 날짜 기록 완료 (openpyxl): 행 {row_number}, 날짜: {today}")
            return True
            
        except PermissionError:
            # 방법 3: 임시 파일로 백업 후 나중에 수동 적용하도록 안내
            try:
                import tempfile
                import os
                
                temp_file = os.path.join(tempfile.gettempdir(), f"hometax_update_{row_number}_{today}.txt")
                with open(temp_file, 'w', encoding='utf-8') as f:
                    f.write(f"행 {row_number}에 {today} 날짜를 기록하세요.\n")
                    f.write(f"파일: {self.excel_file_path}\n")
                    f.write(f"시트: 거래처\n")
                    f.write(f"위치: A{row_number} 셀\n")
                
                print(f"⚠️ 엑셀 파일이 열려있어 직접 기록할 수 없습니다.")
                print(f"임시 파일에 기록 정보를 저장했습니다: {temp_file}")
                print(f"수동으로 A{row_number} 셀에 {today}를 입력하세요.")
                return False
                
            except Exception as temp_error:
                print(f"❌ 임시 파일 생성도 실패: {temp_error}")
                return False
            
        except Exception as e:
            print(f"❌ 엑셀 파일 날짜 기록 실패: {e}")
            print("엑셀 파일이 열려있는 경우 파일을 닫고 다시 시도하세요.")
            return False
    
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
    
    def load_field_mapping(self):
        """field_mapping.md 파일을 읽어서 매핑 정보 추출"""
        # 절대 경로로 docs 폴더에서 파일 찾기
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
    
    def process_excel_data(self):
        """엑셀 데이터를 딕셔너리 형태로 처리"""
        if not self.excel_file_path or not self.selected_rows:
            print("❌ 엑셀 파일 경로나 선택된 행이 없습니다.")
            return False
        
        try:
            # 먼저 openpyxl로 실제 사용된 범위 확인 (거래처 시트)
            from openpyxl import load_workbook
            wb = load_workbook(self.excel_file_path)
            
            # "거래처" 시트 선택
            if "거래처" in wb.sheetnames:
                ws = wb["거래처"]
            else:
                ws = wb.active
                print(f"경고: '거래처' 시트를 찾을 수 없어 기본 시트({ws.title}) 사용")
            max_row = ws.max_row
            max_col = ws.max_column
            print(f"디버그: process_excel_data - openpyxl 최대 행 = {max_row}, 최대 열 = {max_col}")
            
            # 엑셀 파일 읽기 (헤더 없이, 모든 데이터 타입을 문자열로, 빈 값 유지)
            df = pd.read_excel(self.excel_file_path, sheet_name="거래처", header=None, dtype=str, keep_default_na=False,
                             engine='openpyxl', na_filter=False, nrows=max_row)
            print(f"엑셀 파일 읽기 성공: {len(df)}행 × {len(df.columns)}열")
            
            # B열(사업자번호) 기준으로 실제 데이터 행 수 확인
            if len(df.columns) > 1:
                b_column = df.iloc[:, 1]  # B열 (사업자번호)
                data_rows_with_business_num = b_column[b_column.str.strip() != '']
                print(f"B열(사업자번호) 기준 데이터 행 수: {len(data_rows_with_business_num)}개")
            
            # 첫 번째 행을 헤더로 사용
            if len(df) < 1:
                print("❌ 엑셀 파일에 데이터가 없습니다.")
                return False
            

            self.headers = [str(h).strip() for h in df.iloc[0].fillna("").tolist()]  # 헤더의 양쪽 공백 제거
            print(f"헤더: {self.headers}")
            
            # 선택된 행들을 처리
            self.processed_data = []
            
            for row_num in self.selected_rows:
                if row_num > len(df):
                    print(f"⚠️ 행 {row_num}은 데이터 범위({len(df)})를 초과합니다.")
                    continue
                
                # 행 데이터 추출 (0-based index이므로 -1)
                row_data = df.iloc[row_num - 1].fillna("").tolist()
                
                # 헤더와 데이터를 매핑하여 딕셔너리 생성
                row_dict = {}
                for i, header in enumerate(self.headers):
                    if i < len(row_data):
                        value = str(row_data[i]).strip()
                        
                        # 사업자번호 필드 처리 (하이픈 제거)
                        if '사업자번호' in header or '사업자등록번호' in header or '거래처등록번호' in header:
                            # 하이픈 제거하고 숫자만 추출
                            value = ''.join(filter(str.isdigit, value))
                            row_dict[header] = value
                        # 이메일 필드 처리
                        elif '이메일' in header:
                            if '앞' in header or '뒤' in header:
                                # 이미 분리된 이메일 필드는 그대로 사용
                                row_dict[header] = value
                            else:
                                # 통합 이메일 필드인 경우 분리
                                email_front, email_back = self.split_email(value)
                                row_dict[f"{header}_앞"] = email_front
                                row_dict[f"{header}_뒤"] = email_back
                        else:
                            row_dict[header] = value
                    else:
                        row_dict[header] = ""
                
                self.processed_data.append({
                    'row_number': row_num,
                    'data': row_dict
                })
                
                print(f"✅ 행 {row_num} 처리 완료")
                
                # 첫 번째 행의 몇 가지 데이터 샘플 출력
                if row_num == self.selected_rows[0]:
                    print("   샘플 데이터:")
                    sample_count = 0
                    for key, value in row_dict.items():
                        if value and sample_count < 3:  # 값이 있는 첫 3개만 출력
                            print(f"   - {key}: {value}")
                            sample_count += 1
            
            print(f"✅ 총 {len(self.processed_data)}개 행 처리 완료")
            return True
            
        except Exception as e:
            print(f"❌ 엑셀 데이터 처리 실패: {e}")
            return False
    
    def check_and_open_excel(self):
        """엑셀 파일 확인 및 열기 (명확한 3단계 로직)"""
        target_file = r"C:\Users\man4k\OneDrive\문서\세금계산서.xlsx"
        target_filename = "세금계산서.xlsx"
        
        print("=== 엑셀 파일 확인 (3단계 체크) ===")
        
        # === 1단계: 세금계산서.xlsx가 이미 열려있는가? ===
        print(f"1단계: '{target_filename}'가 이미 열려있는지 확인...")
        
        try:
            result = subprocess.run(['tasklist', '/fi', 'imagename eq excel.exe'], 
                                  capture_output=True, text=True)
            if 'excel.exe' in result.stdout.lower():
                print("   ✅ Excel 프로세스 실행 중")
                
                # xlwings로 정확한 파일 확인
                try:
                    import xlwings as xw
                    
                    # Excel 앱이 있는지 먼저 확인 (새로운 API 사용)
                    try:
                        app = xw.apps.active
                        if not app:
                            raise Exception("No active app")
                    except Exception:
                        print("   ⚠️ 활성 Excel 앱을 찾을 수 없습니다.")
                        raise ImportError("No active Excel app")
                    
                    # 모든 열린 워크북 확인 (더 정확한 검사)
                    found_file = False
                    opened_books = []
                    
                    print("   🔍 현재 열린 Excel 파일들을 확인합니다...")
                    
                    for book in app.books:
                        book_name = book.name.lower()
                        target_name = target_filename.lower()
                        opened_books.append(book.name)
                        
                        print(f"   📋 열린 파일: {book.name}")
                        
                        # 정확한 파일명 매치 (더 엄격한 검사)
                        if book_name == target_name:
                            print(f"   ✅ 정확히 일치: '{book.name}' 파일이 이미 열려있습니다!")
                            self.excel_file_path = book.fullname  # 전체 경로 사용
                            found_file = True
                            break
                        elif target_name in book_name and len(book_name) - len(target_name) <= 5:
                            # 비슷한 이름이지만 약간의 차이 (읽기전용 표시 등)
                            print(f"   ✅ 유사한 파일명 발견: '{book.name}' (읽기 전용일 수 있음)")
                            self.excel_file_path = book.fullname
                            found_file = True
                            break
                    
                    if found_file:
                        print("   → 파일을 다시 열지 않고 자동화를 계속 진행합니다.")
                        return True
                    else:
                        print(f"   ⚠️ Excel은 실행 중이지만 '{target_filename}' 파일이 열려있지 않습니다.")
                        print(f"   📋 현재 열린 파일 목록: {opened_books}")
                        print("   → 2단계로 진행합니다.")
                    
                except ImportError:
                    print("   ⚠️ xlwings가 설치되지 않았습니다.")
                    print("   💡 xlwings 설치하면 자동 감지 가능: pip install xlwings")
                    print("   → 2단계로 진행합니다.")
                        
                except Exception as e:
                    print(f"   ❌ xlwings 확인 중 오류: {e}")
                    print("   → 2단계로 진행합니다.")
            else:
                print("   ⚠️ Excel 프로세스가 실행되지 않았습니다.")
        except Exception as e:
            print(f"   ❌ 프로세스 확인 중 오류: {e}")
        
        # === 2단계: 문서 폴더에 세금계산서.xlsx가 있는가? ===
        print(f"2단계: 문서 폴더에 '{target_filename}' 파일이 있는지 확인...")
        
        if os.path.exists(target_file):
            print(f"   ✅ 파일 발견: {target_file}")
            print(f"   📂 '{target_filename}' 파일을 자동으로 엽니다...")
            
            try:
                os.startfile(target_file)
                self.excel_file_path = target_file
                
                # Excel 로딩 대기
                import time
                time.sleep(3)
                
                # 포커스 복원
                try:
                    import win32gui
                    console_hwnd = win32gui.GetConsoleWindow()
                    if console_hwnd:
                        win32gui.SetForegroundWindow(console_hwnd)
                        print("   ✅ 포커스를 콘솔로 복원")
                except:
                    pass
                
                print(f"   ✅ '{target_filename}' 파일이 열렸습니다!")
                print("   → 자동화를 계속 진행합니다.")
                return True
                
            except Exception as e:
                print(f"   ❌ 파일 열기 실패: {e}")
                print("   → 3단계로 진행합니다.")
        else:
            print(f"   ❌ 문서 폴더에 '{target_filename}' 파일이 없습니다.")
            print("   → 3단계로 진행합니다.")
        
        # === 3단계: 파일 열기 창으로 세금계산서.xlsx 선택 ===
        print(f"3단계: 파일 선택 창에서 '{target_filename}' 파일을 선택해주세요...")
        
        root = tk.Tk()
        root.withdraw()
        root.attributes('-topmost', True)
        
        messagebox.showinfo(
            "파일 선택 - 3단계", 
            f"다음 창에서 '{target_filename}' 파일을 선택해주세요.\n\n파일이 선택되면 자동으로 열고 자동화를 계속 진행합니다."
        )
        
        file_path = filedialog.askopenfilename(
            title=f"'{target_filename}' 파일을 선택하세요",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
            initialdir=os.path.dirname(target_file) if os.path.dirname(target_file) else os.path.expanduser("~/Documents")
        )
        
        if file_path:
            print(f"   ✅ 선택된 파일: {file_path}")
            
            # 선택된 파일명 확인
            selected_filename = os.path.basename(file_path)
            if target_filename.lower() in selected_filename.lower():
                print(f"   ✅ 올바른 파일이 선택되었습니다: {selected_filename}")
            else:
                print(f"   ⚠️ 다른 파일이 선택되었지만 계속 진행합니다: {selected_filename}")
            
            try:
                os.startfile(file_path)
                self.excel_file_path = file_path
                print(f"   📂 선택된 파일을 엽니다: {selected_filename}")
                
                # 포커스 복원
                import time
                time.sleep(3)
                try:
                    import win32gui
                    console_hwnd = win32gui.GetConsoleWindow()
                    if console_hwnd:
                        win32gui.SetForegroundWindow(console_hwnd)
                        print("   ✅ 포커스를 콘솔로 복원")
                except:
                    pass
                
                print(f"   ✅ '{selected_filename}' 파일이 열렸습니다!")
                print("   → 자동화를 계속 진행합니다.")
                
                root.destroy()
                return True
                
            except Exception as e:
                print(f"   ❌ 선택된 파일 열기 실패: {e}")
                messagebox.showerror("오류", f"파일을 열 수 없습니다: {e}")
                root.destroy()
                return False
        else:
            print("   ❌ 파일이 선택되지 않았습니다.")
            messagebox.showerror("오류", "파일을 선택하지 않으면 프로그램을 계속할 수 없습니다.")
            root.destroy()
            return False
    
    def parse_row_selection(self, selection_str, silent=False):
        """행 선택 문자열 파싱"""
        if not selection_str.strip():
            return []
        
        rows = []
        parts = selection_str.split(',')
        
        for part in parts:
            part = part.strip()
            if '-' in part:
                # 범위 처리 (예: 2-8)
                try:
                    parts_split = part.split('-', 1)  # 첫 번째 -만으로 분리
                    if len(parts_split) == 2:
                        start_str = parts_split[0].strip()
                        end_str = parts_split[1].strip()
                        
                        if start_str and end_str:  # 둘 다 비어있지 않은 경우
                            start_num = int(start_str)
                            end_num = int(end_str)
                            rows.extend(range(start_num, end_num + 1))
                        else:
                            if not silent:
                                print(f"❌ 잘못된 범위 형식: {part}")
                    else:
                        if not silent:
                            print(f"❌ 잘못된 범위 형식: {part}")
                except (ValueError, IndexError):
                    if not silent:
                        print(f"❌ 잘못된 범위 형식: {part}")
            else:
                # 단일 행 (예: 2)
                try:
                    row_num = int(part.strip())
                    rows.append(row_num)
                except ValueError:
                    if not silent:
                        print(f"❌ 잘못된 행 번호: {part}")
        
        return sorted(set(rows))  # 중복 제거 및 정렬
    
    def show_row_selection_gui(self):
        """행 선택 GUI 표시"""
        print("\n=== 행 선택 GUI ===")
        
        root = tk.Tk()
        root.title("행 선택")
        root.resizable(False, False)
        
        # 화면 상단 중앙에 위치
        screen_width = root.winfo_screenwidth()
        screen_height = root.winfo_screenheight()
        window_width = 500
        window_height = 550
        x = (screen_width - window_width) // 2
        y = (screen_height - window_height) // 4  # 화면 상단 1/4 지점
        root.geometry(f"{window_width}x{window_height}+{x}+{y}")
        
        # 메인 프레임
        main_frame = ttk.Frame(root, padding="20")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # 제목
        title_label = ttk.Label(main_frame, text="처리할 행을 선택하세요", 
                               font=('맑은 고딕', 14, 'bold'))
        title_label.pack(pady=(0, 20))
        
        # 안내 메시지
        guide_frame = ttk.LabelFrame(main_frame, text="행 선택 방법", padding="10")
        guide_frame.pack(fill=tk.X, pady=(0, 20))
        
        guide_text = """• 단일 행: 2
• 복수 행: 2,4,8
• 범위: 2-8
• 혼합: 2,5-7,10

예시: 2행, 5~7행, 10행을 처리하려면 → 2,5-7,10"""
        
        guide_label = ttk.Label(guide_frame, text=guide_text, justify=tk.LEFT)
        guide_label.pack(anchor=tk.W)
        
        # 입력 프레임
        input_frame = ttk.Frame(main_frame)
        input_frame.pack(fill=tk.X, pady=(0, 20))
        
        ttk.Label(input_frame, text="행 선택:").pack(anchor=tk.W)
        
        entry_var = tk.StringVar()
        entry = ttk.Entry(input_frame, textvariable=entry_var, font=('맑은 고딕', 11))
        entry.pack(fill=tk.X, pady=(5, 0))
        entry.focus()
        
        # 결과 표시 프레임
        result_frame = ttk.LabelFrame(main_frame, text="선택 결과", padding="10")
        result_frame.pack(fill=tk.X, pady=(0, 20))
        
        # 스크롤바 추가를 위한 프레임
        text_frame = ttk.Frame(result_frame)
        text_frame.pack(fill=tk.BOTH, expand=True)
        
        result_text = tk.Text(text_frame, height=8, width=50, wrap=tk.WORD, font=('맑은 고딕', 9))
        scrollbar = ttk.Scrollbar(text_frame, orient="vertical", command=result_text.yview)
        result_text.configure(yscrollcommand=scrollbar.set)
        
        result_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # 버튼 프레임
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill=tk.X)
        
        def preview_selection():
            """선택 미리보기"""
            selection = entry_var.get()
            if not selection.strip():
                result_text.delete(1.0, tk.END)
                result_text.insert(1.0, "행을 입력하세요.")
                return
            
            try:
                rows = self.parse_row_selection(selection, silent=True)  # 실시간 미리보기에서는 오류 메시지 숨김
                if rows:
                    result_text.delete(1.0, tk.END)
                    result_text.insert(1.0, f"선택된 행: {rows}\n")
                    result_text.insert(tk.END, f"총 {len(rows)}개 행이 선택됩니다.\n\n")
                    
                    # 첫 번째 행 정보 미리보기
                    if self.excel_file_path and os.path.exists(self.excel_file_path):
                        try:
                            # openpyxl로 직접 데이터 읽기 (거래처 시트)
                            from openpyxl import load_workbook
                            wb = load_workbook(self.excel_file_path)
                            
                            # "거래처" 시트 선택
                            if "거래처" in wb.sheetnames:
                                ws = wb["거래처"]
                            else:
                                ws = wb.active
                                print(f"경고: '거래처' 시트를 찾을 수 없어 기본 시트({ws.title}) 사용")
                            
                            # 실제 사용된 범위
                            max_row = ws.max_row
                            max_col = ws.max_column
                            print(f"디버그: openpyxl 최대 행 = {max_row}, 최대 열 = {max_col}")
                            
                            # 실제 데이터가 있는 행 수 계산 (B열 기준)
                            actual_data_rows = 0
                            for row in range(1, max_row + 1):
                                b_value = ws.cell(row=row, column=2).value or ""
                                if str(b_value).strip() and str(b_value).strip() not in ['사업자등록번호', '등록번호']:
                                    actual_data_rows = row
                            
                            print(f"디버그: openpyxl 최대 행 = {max_row}, 실제 데이터 마지막 행 = {actual_data_rows}")
                            
                            result_text.insert(tk.END, f"총 데이터 {actual_data_rows-1}개중 {len(rows)}개행이 선택되었습니다.\n")
                            
                            # 선택된 행들의 거래처명 바로 표시
                            for row_num in rows:
                                if row_num <= max_row:
                                    d_value = ws.cell(row=row_num, column=4).value or "데이터 없음"
                                    result_text.insert(tk.END, f"행{row_num} : {d_value}\n")
                                else:
                                    result_text.insert(tk.END, f"행{row_num} : 범위 초과\n")
                            
                            # 유효성 검사
                            invalid_rows = [r for r in rows if r > max_row]
                            if invalid_rows:
                                result_text.insert(tk.END, f"\n❌ 범위 초과 행: {invalid_rows}")
                                
                        except Exception as e:
                            result_text.insert(tk.END, f"엑셀 데이터 미리보기 실패: {e}")
                else:
                    result_text.delete(1.0, tk.END)
                    result_text.insert(1.0, "올바른 행 번호를 입력하세요.")
            except Exception as e:
                result_text.delete(1.0, tk.END)
                result_text.insert(1.0, f"오류: {e}")
        
        def confirm_selection():
            """선택 확정"""
            selection = entry_var.get()
            rows = self.parse_row_selection(selection)
            
            if not rows:
                messagebox.showerror("오류", "올바른 행을 선택하세요.")
                return
            
            # 바로 진행
            self.selected_rows = rows
            
            # 첫 번째 행의 첫 번째 열 값 저장 (기존 로직 유지)
            if self.excel_file_path and os.path.exists(self.excel_file_path):
                try:
                    from openpyxl import load_workbook
                    wb = load_workbook(self.excel_file_path)
                    ws = wb.active
                    max_row = ws.max_row
                    
                    df = pd.read_excel(self.excel_file_path, sheet_name="거래처", header=None, dtype=str, keep_default_na=False,
                                     engine='openpyxl', na_filter=False, nrows=max_row)
                    
                    first_row = rows[0]
                    if first_row <= len(df) and len(df.columns) > 0:
                        self.selected_data = df.iloc[first_row-1, 0]
                        print(f"✅ 첫 번째 행({first_row})의 첫 번째 열 값 저장: {self.selected_data}")
                    else:
                        self.selected_data = None
                        print(f"❌ 행 {first_row}의 데이터를 찾을 수 없습니다.")
                        
                except Exception as e:
                    print(f"❌ 엑셀 데이터 읽기 실패: {e}")
                    self.selected_data = None
            
            root.destroy()
        
        def cancel_selection():
            """선택 취소 및 프로그램 종료"""
            self.selected_rows = None
            self.selected_data = None
            root.destroy()
            print("사용자가 프로그램을 취소했습니다.")
            sys.exit(0)
        
        # 실시간 미리보기를 위한 이벤트 바인딩
        entry_var.trace('w', lambda *args: preview_selection())
        
        ttk.Button(button_frame, text="확인", command=confirm_selection).pack(side=tk.LEFT, padx=(0, 10))
        ttk.Button(button_frame, text="취소", command=cancel_selection).pack(side=tk.LEFT)
        
        # Enter 키로 확인
        def on_enter(event):
            confirm_selection()
        
        root.bind('<Return>', on_enter)
        
        root.mainloop()
        
        return self.selected_rows is not None

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

async def hometax_auto_login():
    """HomeTax 자동 로그인 및 거래처 등록 화면 이동 (암호화된 비밀번호 사용)"""
    load_dotenv()
    
    # 보안 관리자를 통해 암호화된 비밀번호 로드
    print("[SECURITY] 암호화된 비밀번호 로드 중...")
    security_manager = HomeTaxSecurityManager()
    cert_password = security_manager.load_password_from_env()
    
    if not cert_password:
        print("[ERROR] 암호화된 비밀번호가 설정되지 않았습니다.")
        print("[HELP] hometax_cert_manager.py를 실행하여 비밀번호를 저장하세요.")
        return None, None
    else:
        print("[OK] 암호화된 비밀번호 로드 성공")

    playwright = await async_playwright().start()
    browser = await playwright.chromium.launch(
        headless=False, 
        slow_mo=500,
        args=[
            '--disable-web-security',
            '--disable-features=VizDisplayCompositor'
        ]
    )
    
    try:
        page = await browser.new_page()
        page.set_default_timeout(30000)  # 30초로 증가
        
        print("홈택스 페이지 이동...")
        await page.goto("https://hometax.go.kr/websquare/websquare.html?w2xPath=/ui/pp/index_pp.xml&menuCd=index3", timeout=60000)  # 60초 타임아웃
        await page.wait_for_load_state('domcontentloaded')
        
        await page.wait_for_timeout(1500)

        # MagicLine 로딩 대기 (최대 15초)
        print("MagicLine 보안 모듈 로딩 대기...")
        try:
            await page.wait_for_function(
                "typeof magicline !== 'undefined' && typeof magicline.AGENT_VER !== 'undefined'",
                timeout=1000)
            print("✅ MagicLine 로딩 완료!")
        except Exception as e:
            print("⚠️ MagicLine 로딩 실패 또는 타임아웃. 계속 진행합니다...")
            print(f"   오류: {e}")
        
        # 공동·금융인증서 버튼 클릭
        print("공동·금융인증서 버튼 검색...")
        
        button_selectors = [
            "#mf_txppWframe_loginboxFrame_anchor22",
            "#anchor22",
            "a:has-text('공동인증서')",
            "a:has-text('공동·금융인증서')",
            "a:has-text('금융인증서')"
        ]
        
        login_clicked = False
        for selector in button_selectors:
            try:
                print(f"시도: {selector}")
                await page.locator(selector).first.click(timeout=1000)
                print(f"클릭 성공: {selector}")
                login_clicked = True
                break
            except:
                continue
        
        # iframe 내부에서도 시도
        if not login_clicked:
            try:
                iframe = page.frame_locator("#txppIframe")
                await iframe.locator("a:has-text('공동')").first.click(timeout=1000)
                login_clicked = True
                print("iframe 내부 클릭 성공")
            except:
                pass
        
        if not login_clicked:
            print("자동 클릭 실패 - 수동으로 '공동·금융인증서' 버튼을 클릭하세요")
            await page.wait_for_timeout(3000)
        
        # 인증서 창 대기 (더 긴 시간과 더 많은 시도)
        print("인증서 창 대기...")
        dscert_found = False
        
        # 먼저 페이지가 완전히 로드되길 기다림
        await page.wait_for_timeout(1500)
        
        for i in range(20):  # 20초 동안 시도
            try:
                print(f"인증서 창 찾는 중... 시도 {i+1}/20")
                
                # 다양한 방법으로 인증서 창 찾기
                selectors_to_try = ["#dscert", "iframe[id='dscert']", "iframe[name='dscert']"]
                
                for selector in selectors_to_try:
                    try:
                        await page.wait_for_selector(selector, timeout=1000)
                        dscert_iframe = page.frame_locator(selector)
                        await dscert_iframe.locator("body").wait_for(timeout=1000)
                        print(f"인증서 창 발견! (선택자: {selector})")
                        dscert_found = True
                        break
                    except:
                        continue
                
                if dscert_found:
                    break
                    
            except Exception as e:
                print(f"시도 {i+1} 실패: {e}")
                await page.wait_for_timeout(1000)
        
        if not dscert_found:
            print("❌ 인증서 창을 찾을 수 없습니다.")
            print("💡 수동으로 인증서 로그인을 진행하세요.")
            
            # 수동 로그인 대기 옵션 제공
            print("인증서 로그인을 수동으로 완료한 후 계속 진행하시겠습니까? (y/n)")
            
            # 사용자 입력 대기를 위한 간단한 방법
            import asyncio
            import sys
            
            # 15초 동안 로그인 완료 확인
            for i in range(15):
                await page.wait_for_timeout(1000)
                
                # URL 변경이나 특정 요소로 로그인 완료 확인
                current_url = page.url.lower()
                if "main" in current_url or "home" in current_url:
                    print("✅ 수동 로그인 완료 감지!")
                    return page, browser
                
                # 인증서 창이 사라졌는지 확인
                try:
                    dscert_visible = await page.locator("#dscert").is_visible()
                    if not dscert_visible:
                        print("✅ 인증서 창 사라짐 - 로그인 완료로 간주")
                        await page.wait_for_timeout(3000)  # 추가 대기
                        return page, browser
                except:
                    pass
            
            print("❌ 수동 로그인도 완료되지 않았습니다.")
            return None, browser
        
        # 인증서 선택 (Firefox용 최적화)
        print("인증서 선택...")
        try:
            # Firefox에서 더 안정적인 방법으로 인증서 선택
            await page.wait_for_timeout(2000)  # 페이지 안정화 대기
            
            # JavaScript로 강제 클릭 (blockUI 무시)
            try:
                await page.evaluate("""
                    (function() {
                        const iframe = document.getElementById('dscert');
                        if (iframe && iframe.contentDocument) {
                            const firstCert = iframe.contentDocument.querySelector('#row0dataTable > td:nth-child(1) > a');
                            if (firstCert) {
                                firstCert.click();
                                console.log('인증서 선택 완료 (JavaScript 강제 클릭)');
                            }
                        }
                    })()
                """)
                print("인증서 선택 완료 (JavaScript 강제 클릭)")
                await page.wait_for_timeout(2000)  # 더 긴 대기 시간
                
            except Exception as js_error:
                print(f"JavaScript 방법 실패: {js_error}")
                
                # 대체 방법: 테이블의 첫 번째 행 클릭
                try:
                    await page.evaluate("""
                        (function() {
                            const iframe = document.getElementById('dscert');
                            if (iframe && iframe.contentDocument) {
                                const rows = iframe.contentDocument.querySelectorAll('#row0dataTable tr');
                                if (rows.length > 0) {
                                    rows[0].click();
                                    console.log('대체 방법으로 인증서 선택 완료');
                                }
                            }
                        })()
                    """)
                    print("대체 방법으로 인증서 선택 완료 (행 클릭)")
                    await page.wait_for_timeout(2000)
                except Exception as alt_error:
                    print(f"대체 방법도 실패: {alt_error}")
                    print("인증서 선택 실패 - 수동으로 선택하세요")
                    await page.wait_for_timeout(3000)  # 수동 선택 대기
                
        except Exception as e:
            print(f"인증서 선택 처리 실패: {e}")
            print("수동으로 인증서를 선택하세요")
            await page.wait_for_timeout(3000)  # 3초 수동 대기
        
        # 비밀번호 입력 (개선된 방법)
        print("비밀번호 입력...")
        try:
            # 먼저 비밀번호 입력란이 나타날 때까지 기다림
            await page.wait_for_timeout(1500)
            
            # JavaScript로 비밀번호 입력 시도
            password_filled = await page.evaluate(f"""
                (function() {{
                    const iframe = document.getElementById('dscert');
                    if (iframe && iframe.contentDocument) {{
                        const passwordInput = iframe.contentDocument.querySelector('#input_cert_pw');
                        if (passwordInput) {{
                            passwordInput.value = '{cert_password}';
                            passwordInput.dispatchEvent(new Event('input', {{ bubbles: true }}));
                            passwordInput.dispatchEvent(new Event('change', {{ bubbles: true }}));
                            return true;
                        }}
                    }}
                    return false;
                }})()""")
            
            if password_filled:
                print("비밀번호 입력 완료 (JavaScript)")
            else:
                # 대체 방법: Playwright locator 사용
                password_input = dscert_iframe.locator("#input_cert_pw").first
                await password_input.wait_for(state="visible", timeout=1000)
                await password_input.fill(cert_password)
                print("비밀번호 입력 완료 (Playwright)")
                
        except Exception as e:
            print(f"비밀번호 입력 실패: {e}")
            print("수동으로 비밀번호를 입력하세요")
            await page.wait_for_timeout(3000)  # 수동 입력 대기
        
        # 확인 버튼 클릭 (개선된 방법)
        print("확인 버튼 클릭...")
        await page.wait_for_timeout(1000)
        
        try:
            # JavaScript로 확인 버튼 클릭 시도
            confirm_clicked = await page.evaluate("""
                (function() {{
                    const iframe = document.getElementById('dscert');
                    if (iframe && iframe.contentDocument) {{
                        // 여러 가능한 확인 버튼 셀렉터 시도
                        const selectors = [
                            '#btn_confirm_iframe',
                            '#btn_confirm_iframe > span',
                            'input[value*="확인"]',
                            'button:contains("확인")',
                            '[id*="confirm"]'
                        ];
                        
                        for (const selector of selectors) {{
                            const btn = iframe.contentDocument.querySelector(selector);
                            if (btn) {{
                                btn.click();
                                console.log('확인 버튼 클릭 완료 (JavaScript): ' + selector);
                                return true;
                            }}
                        }}
                    }}
                    return false;
                }})()""")
            
            if confirm_clicked:
                print("확인 버튼 클릭 완료 (JavaScript)")
            else:
                # 대체 방법: Playwright locator 사용
                try:
                    confirm_btn = dscert_iframe.locator("#btn_confirm_iframe > span").first
                    await confirm_btn.wait_for(state="visible", timeout=1000)
                    await confirm_btn.click()
                    print("확인 버튼 클릭 완료 (Playwright 정확한 셀렉터)")
                except:
                    confirm_btn = dscert_iframe.locator("#btn_confirm_iframe").first
                    await confirm_btn.click(timeout=1000)
                    print("확인 버튼 클릭 완료 (Playwright 대체 방법)")
                    
        except Exception as e:
            print(f"모든 확인 버튼 클릭 방법 실패: {e}")
            print("수동으로 확인 버튼을 클릭하세요")
            await page.wait_for_timeout(3000)  # 수동 클릭 대기
        
        # 로그인 완료 대기
        print("로그인 처리 중...")
        login_confirmed = False
        for i in range(15):
            await page.wait_for_timeout(1000)
            
            # URL 변경 확인
            if "main" in page.url.lower() or "home" in page.url.lower():
                print("✅ 로그인 성공! URL 변경 감지")
                login_confirmed = True
                break
            
            # 인증서 창이 사라졌는지 확인
            try:
                dscert_visible = await page.locator("#dscert").is_visible()
                if not dscert_visible:
                    print("✅ 로그인 성공! 인증서 창 사라짐 확인")
                    login_confirmed = True
                    break
            except:
                pass
        
        if login_confirmed:
            print("HomeTax 자동 로그인 성공!")
            
            # 브라우저 포커스 유지
            try:
                await page.bring_to_front()
                print("✅ 브라우저 포커스 유지")
            except Exception as e:
                print(f"⚠️ 브라우저 포커스 설정 실패: {e}")
            
            # 거래처 등록 화면으로 이동
            print("\n=== 거래처 등록 화면으로 이동 ===")
            await page.wait_for_timeout(3000)
            
            # Alert창 닫기 (여러 방법으로 시도)
            alert_closed = False
            alert_close_selectors = [
                "#mf_txppWframe_UTXPPABB29_wframe_btnCloseInvtSpec",
                "[id*='btnCloseInvtSpec']",
                "[title*='닫기']",
                "text=닫기"
            ]
            
            for selector in alert_close_selectors:
                try:
                    print(f"Alert창 닫기 시도: {selector}")
                    close_button = page.locator(selector).first
                    await close_button.wait_for(state="visible", timeout=1000)
                    await close_button.click()
                    print(f"  ✅ Alert창 닫기 완료: {selector}")
                    alert_closed = True
                    await page.wait_for_timeout(2000)
                    break
                except Exception as e:
                    print(f"  ❌ {selector} 실패: {e}")
                    continue
            
            if not alert_closed:
                print("⚠️ Alert창을 자동으로 닫을 수 없습니다. 수동으로 닫아주세요.")
                await page.wait_for_timeout(2000)  # 수동으로 닫을 시간을 줌
            
            # 메뉴 네비게이션 (1초 간격 순차 클릭)
            try:
                print("메뉴 네비게이션: 3단계 순차 클릭...")
                
                # 페이지 로딩 대기
                await page.wait_for_timeout(3000)
                
                # 1단계: 신고/납부 메뉴 클릭
                print("1단계: #mf_wfHeader_wq_uuid_333 클릭")
                await page.locator("#mf_wfHeader_wq_uuid_333").first.click()
                await page.wait_for_timeout(1000)
                print("  ✅ 1단계 완료")
                
                # 2단계: 거래처관리 메뉴 클릭  
                print("2단계: #menuAtag_4601020000 > span 클릭")
                await page.locator("#menuAtag_4601020000 > span").first.click()
                await page.wait_for_timeout(1000)
                print("  ✅ 2단계 완료")
                
                # 3단계: 거래처등록 메뉴 클릭
                print("3단계: #menuAtag_4601020100 > span 클릭")
                await page.locator("#menuAtag_4601020100 > span").first.click()
                await page.wait_for_timeout(1000)
                print("  ✅ 3단계 완료")
                
                print("✅ 메뉴 네비게이션 완료!")
                
                # 거래처 정보관리 화면 로딩 대기
                await page.wait_for_timeout(2000)
                
                # 4단계: 건별등록 버튼 클릭
                print("4단계: #mf_txppWframe_textbox1395 건별등록 클릭")
                await page.locator("#mf_txppWframe_textbox1395").first.click()
                await page.wait_for_timeout(2000)
                print("  ✅ 건별등록 버튼 클릭 완료")
                
            except Exception as e:
                print(f"❌ 메뉴 네비게이션 오류: {e}")
                print("수동으로 메뉴를 선택하세요.")

            # 건별등록 완료 후 사업자번호 입력 필드 확인
            try:
                # 사업자번호 입력 필드가 나타나는지 확인
                await page.wait_for_selector("#mf_txppWframe_txtBsno1", timeout=1000)
                print("✅ 사업자번호 입력 필드 확인됨 - 거래처 등록 화면 진입 완료!")
                return page, browser
            except:
                print("⚠️ 사업자번호 입력 필드를 찾을 수 없습니다. 수동 확인이 필요합니다.")
                return None, browser
        else:
            print("⚠️ 로그인 상태 확인 필요")
            return None, browser
            
    except Exception as e:
        print(f"오류: {e}")
        if 'browser' in locals() and browser:
            try:
                # 모든 페이지 닫기
                pages = browser.contexts[0].pages if browser.contexts else []
                for page in pages:
                    try:
                        await page.close()
                    except:
                        pass
                        
                await browser.close()
                await playwright.stop()
            except Exception as close_error:
                print(f"브라우저 종료 중 오류: {close_error}")
        return None, None

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
    
    # 3. HomeTax 자동 로그인 및 거래처 등록 화면 이동
    print("\n3단계: HomeTax 자동 실행")
    page, browser = await hometax_auto_login()
    
    if page:
        print("\n✅ 모든 단계 완료!")
        print("- 엑셀 파일이 열렸습니다.")
        print(f"- 선택된 행: {excel_selector.selected_rows}")
        print(f"- 처리된 데이터: {len(excel_selector.processed_data)}개 행")
        print("- 엑셀 데이터가 딕셔너리 형태로 변환되었습니다.")
        print("  (이메일 필드는 @ 기준으로 분리됨)")
        print("- HomeTax 거래처 등록 화면에 접속했습니다.")
        
        # 4. 자동 데이터 입력
        print("\n4단계: 거래처 데이터 자동 입력")
        
        for i, processed_row in enumerate(excel_selector.processed_data):
            row_number = processed_row['row_number'] 
            row_data = processed_row['data']
            
            print(f"\n[{i+1}/{len(excel_selector.processed_data)}] 행 {row_number} 데이터 입력")
            
            try:
                success_count, failed_fields = await fill_hometax_form(
                    page, row_data, excel_selector.field_mapping, excel_selector, row_number, is_first_record=(i == 0)
                )
                
                if success_count > 0:
                    print(f"✅ 행 {row_number} 입력 완료 ({success_count}개 필드)")
                
                else:
                    print(f"❌ 행 {row_number} 입력 실패")
                    
            except Exception as e:
                if "SKIP_TO_NEXT_ROW" in str(e):
                    print(f"⚠️ 행 {row_number} 스킵됨: {str(e).split('|')[1] if '|' in str(e) else str(e)}")
                    excel_selector.write_error_to_excel(row_number, "skip")
                    
                    # 마지막 행이 아닌 경우 다음 행을 위한 페이지 준비
                    if i < len(excel_selector.processed_data) - 1:  # 마지막 행이 아닌 경우
                        print(f"  ➡️ 다음 행 ({excel_selector.processed_data[i+1]['row_number']}행) 준비 중...")
                        prepare_success = await prepare_next_registration(page)
                        if not prepare_success:
                            print(f"❌ 행 {row_number} 스킵 후 다음 행 준비 실패")
                    else:
                        print("  ℹ️ 마지막 행이므로 페이지 준비를 생략합니다.")
                    
                    continue  # 다음 행으로 계속
                else:
                    print(f"❌ 행 {row_number} 처리 중 오류: {e}")
                    print("오류가 발생하여 처리를 중단합니다. 브라우저를 확인하세요.")
                    break
        
        print("\n✅ 모든 거래처 데이터 처리 완료!")
        print(f"- 총 {len(excel_selector.processed_data)}개 행 처리됨")
        print("\n추후 개발을 위해 브라우저를 열린 상태로 유지합니다.")
        
        # 브라우저 포커스 유지 및 콘솔 포커스로 복원
        try:
            await page.bring_to_front()
            await page.wait_for_timeout(1000)
            
            # 콘솔로 포커스 복원 시도
            try:
                import win32gui
                console_hwnd = win32gui.GetConsoleWindow()
                if console_hwnd:
                    win32gui.SetForegroundWindow(console_hwnd)
                    print("✅ 작업 완료 후 콘솔 포커스 복원")
            except:
                print("⚠️ 콘솔 포커스 복원 실패 (정상 동작)")
                
        except Exception as e:
            print(f"⚠️ 포커스 제어 실패: {e}")
        
        # 모든 작업 완료 후 자동 종료
        print("\n" + "="*50)
        print("✅ 모든 거래처 등록 작업이 완료되었습니다!")
        print("브라우저를 종료합니다...")
        print("="*50)
        
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
                await playwright.stop()
                print("✅ 브라우저가 정상적으로 종료되었습니다.")
                print("✅ 프로그램이 완료되었습니다.")
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
                await playwright.stop()
                print("브라우저가 정상적으로 종료되었습니다.")
            except Exception as e:
                print(f"브라우저 종료 중 오류: {e}")


if __name__ == "__main__":
    asyncio.run(main())
