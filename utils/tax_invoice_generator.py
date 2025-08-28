# -*- coding: utf-8 -*-
import asyncio
import os
import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import re
from dotenv import load_dotenv
from playwright.async_api import async_playwright

# ========== HomeTax 자동 로그인 기능 (hometax_quick.py에서 복사) ==========

async def hometax_quick_login():
    """
    빠른 홈택스 로그인 자동화 (대기시간 최소화)
    """
    load_dotenv()
    cert_password = os.getenv("PW")
    if not cert_password:
        print("오류: .env 파일에 PW 변수가 설정되지 않았습니다.")
        return None

    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=False, slow_mo=500)
        
        try:
            page = await browser.new_page()
            page.set_default_timeout(10000)  # 10초로 단축
            
            print("홈택스 페이지 이동...")
            await page.goto("https://hometax.go.kr/websquare/websquare.html?w2xPath=/ui/pp/index_pp.xml&menuCd=index3")
            await page.wait_for_load_state('domcontentloaded')  # networkidle → domcontentloaded로 변경
            
            await page.wait_for_timeout(3000)  # 8초 → 3초로 단축
            
            # 빠른 버튼 찾기 - 직접적인 셀렉터부터 시도
            print("공동·금융인증서 버튼 검색...")
            
            button_selectors = [
                "#mf_txppWframe_loginboxFrame_anchor22",  # 정확한 셀렉터
                "#anchor22",
                "a:has-text('공동인증서')",
                "a:has-text('공동·금융인증서')",
                "a:has-text('금융인증서')"
            ]
            
            login_clicked = False
            for selector in button_selectors:
                try:
                    print(f"시도: {selector}")
                    await page.locator(selector).first.click(timeout=2000)
                    print(f"클릭 성공: {selector}")
                    login_clicked = True
                    break
                except:
                    continue
            
            # iframe 내부에서도 빠르게 시도
            if not login_clicked:
                try:
                    iframe = page.frame_locator("#txppIframe")
                    await iframe.locator("a:has-text('공동')").first.click(timeout=2000)
                    login_clicked = True
                    print("iframe 내부 클릭 성공")
                except:
                    pass
            
            if not login_clicked:
                print("자동 클릭 실패 - 수동으로 '공동·금융인증서' 버튼을 클릭하세요")
                await page.wait_for_timeout(10000)  # 10초만 대기
            
            # #dscert iframe 빠른 대기
            print("인증서 창 대기...")
            dscert_found = False
            
            for i in range(10):  # 30회 → 10회로 단축
                try:
                    await page.wait_for_selector("#dscert", timeout=1000)
                    dscert_iframe = page.frame_locator("#dscert")
                    await dscert_iframe.locator("body").wait_for(timeout=1000)
                    print("인증서 창 발견!")
                    dscert_found = True
                    break
                except:
                    await page.wait_for_timeout(1000)
            
            if not dscert_found:
                print("인증서 창을 찾을 수 없습니다.")
                return None
            
            # 인증서 선택 먼저 (추가된 단계)
            print("인증서 선택...")
            try:
                cert_selector = dscert_iframe.locator("#row0dataTable > td:nth-child(1) > a > span").first
                await cert_selector.wait_for(state="visible", timeout=5000)
                await cert_selector.click()
                print("인증서 선택 완료")
                await page.wait_for_timeout(1000)  # 선택 후 잠시 대기
            except Exception as e:
                print(f"인증서 선택 실패: {e}")
                # 다른 방법으로 시도
                try:
                    # 첫 번째 인증서 항목 찾기
                    first_cert = dscert_iframe.locator("td:nth-child(1) > a, tr:first-child td a").first
                    await first_cert.click()
                    print("대체 방법으로 인증서 선택 완료")
                    await page.wait_for_timeout(1000)
                except:
                    print("인증서 선택 실패 - 수동으로 선택하세요")
            
            # 비밀번호 빠른 입력
            print("비밀번호 입력...")
            password_input = dscert_iframe.locator("#input_cert_pw").first
            await password_input.wait_for(state="visible", timeout=3000)
            await password_input.fill(cert_password)
            print("비밀번호 입력 완료")
            
            # 확인 버튼 빠른 클릭
            print("확인 버튼 클릭...")
            await page.wait_for_timeout(500)
            
            # 정확한 확인 버튼 셀렉터 사용
            try:
                confirm_btn = dscert_iframe.locator("#btn_confirm_iframe > span").first
                await confirm_btn.wait_for(state="visible", timeout=3000)
                await confirm_btn.click()
                print("확인 버튼 클릭 완료 (정확한 셀렉터)")
            except Exception as e:
                print(f"정확한 셀렉터 실패: {e}")
                # 대체 방법들 시도
                try:
                    confirm_btn = dscert_iframe.locator("#btn_confirm_iframe").first
                    await confirm_btn.click(timeout=3000)
                    print("확인 버튼 클릭 완료 (대체 방법 1)")
                except:
                    try:
                        confirm_btn = dscert_iframe.locator("button:has-text('확인'), input[value*='확인']").first
                        await confirm_btn.click(timeout=3000)
                        print("확인 버튼 클릭 완료 (대체 방법 2)")
                    except:
                        print("확인 버튼 클릭 실패 - 수동으로 클릭하세요")
            
            # 팝업창 및 Alert 처리 (선택적)
            print("팝업창/Alert 확인 중...")
            
            # 현재 URL 저장 (변수 선언)
            current_initial_url = page.url
            
            # Alert 핸들러 미리 등록 (나타나면 자동 처리)
            dialog_handled = False
            def handle_dialog(dialog):
                nonlocal dialog_handled
                dialog_handled = True
                print(f"Alert 감지 및 처리: '{dialog.message}'")
                asyncio.create_task(dialog.accept())
            
            page.on("dialog", handle_dialog)
            
            # 짧은 시간 동안만 팝업/Alert 확인 (최대 3초)
            popup_found = False
            for check in range(3):
                await page.wait_for_timeout(1000)
                
                # 새로운 팝업창 확인 (context를 통해 접근)
                try:
                    context_pages = page.context.pages
                    if len(context_pages) > 1:  # 메인 페이지 외에 다른 페이지가 있는 경우
                        print(f"새 팝업창 감지: {len(context_pages)}개 페이지 중 {len(context_pages) - 1}개 팝업")
                        popup_found = True
                        
                        # 메인 페이지가 아닌 창들 닫기
                        for popup_page in context_pages:
                            if popup_page != page:
                                try:
                                    await popup_page.close()
                                    print("팝업창 닫기 완료")
                                except:
                                    pass
                        break
                except Exception as e:
                    # 팝업창 확인 중 오류가 발생해도 계속 진행
                    pass
                
                # Alert 처리됨 확인
                if dialog_handled:
                    print("Alert 처리 완료")
                    popup_found = True
                    break
                
                # 로그인이 이미 진행되었는지 확인 (URL 변경)
                if page.url != current_initial_url:
                    print("로그인 진행 중 - 팝업 확인 건너뜀")
                    break
            
            if not popup_found and not dialog_handled:
                print("팝업창/Alert 없음 - 정상 진행")
            
            # Alert 핸들러 제거
            page.remove_listener("dialog", handle_dialog)
            
            # 로그인 완료 정확한 확인
            print("로그인 처리 중...")
            final_initial_url = page.url
            
            login_confirmed = False
            for i in range(15):  # 15초까지 확인
                await page.wait_for_timeout(1000)
                current_url = page.url
                current_title = await page.title()
                
                # URL 변경 확인
                if current_url != final_initial_url:
                    print(f"✅ 로그인 성공! URL 변경 감지")
                    print(f"   새 URL: {current_url}")
                    login_confirmed = True
                    break
                
                # 페이지 제목 확인
                if any(keyword in current_title.lower() for keyword in ['main', 'home', '홈', '메인', '국세청']):
                    print(f"✅ 로그인 성공! 메인페이지 접근: {current_title}")
                    login_confirmed = True
                    break
                
                # 인증서 창이 사라졌는지 확인 (로그인 성공 신호)
                try:
                    dscert_visible = await page.locator("#dscert").is_visible()
                    if not dscert_visible:
                        print("✅ 로그인 성공! 인증서 창 사라짐 확인")
                        login_confirmed = True
                        break
                except:
                    pass
                
                # 로그인 관련 요소 확인
                try:
                    logout_btn = await page.locator("a:has-text('로그아웃'), button:has-text('로그아웃')").count()
                    if logout_btn > 0:
                        print("✅ 로그인 성공! 로그아웃 버튼 확인")
                        login_confirmed = True
                        break
                except:
                    pass
                
                if i % 3 == 2:
                    print(f"   대기 중... ({i + 1}/15초)")
            
            # 최종 상태 확인
            final_url = page.url
            final_title = await page.title()
            
            print(f"\n=== 최종 로그인 결과 ===")
            print(f"최종 URL: {final_url}")
            print(f"최종 제목: {final_title}")
            
            if login_confirmed:
                print("🎉 홈택스 자동 로그인 성공!")
                
                # Alert창 X버튼으로 닫기
                print("\n=== Alert창 X버튼 닫기 ===")
                try:
                    # 정확한 X버튼 클릭
                    close_button = page.locator("#mf_txppWframe_UTXPPABB29_wframe_btnCloseInvtSpec")
                    await close_button.wait_for(state="visible", timeout=5000)
                    await close_button.click()
                    print("   X버튼으로 알림창 닫기 완료")
                    await page.wait_for_timeout(2000)
                    
                except Exception as e:
                    print(f"X버튼 클릭 실패: {e}")
                    # 대체 방법으로 Alert 처리
                    try:
                        await page.evaluate("""
                            if (window.confirm) window.confirm = function() { return true; };
                            if (window.alert) window.alert = function() { return true; };
                        """)
                        print("   JavaScript Alert 무력화 완료")
                    except:
                        pass
                
                # 추가 메뉴 네비게이션
                print("\n=== 메뉴 네비게이션 시작 ===")
                await page.wait_for_timeout(3000)  # 더 긴 안정화 대기
                
                try:
                    # 1단계: #mf_wfHeader_wq_uuid_333 선택 (Alert창 닫은 후 첫 번째 메뉴)
                    print("1단계: 첫 번째 메뉴 선택 (#mf_wfHeader_wq_uuid_333)...")
                    
                    first_menu_selectors = [
                        "#mf_wfHeader_wq_uuid_333",
                        "*[id*='wq_uuid_333']",
                        "*[id*='wfHeader'] *[id*='333']",
                        "a[href*='333'], button[id*='333']"
                    ]
                    
                    first_clicked = False
                    for selector in first_menu_selectors:
                        try:
                            print(f"   시도: {selector}")
                            first_menu = page.locator(selector).first
                            await first_menu.wait_for(state="visible", timeout=3000)
                            await first_menu.click()
                            print(f"   첫 번째 메뉴 클릭 성공: {selector}")
                            first_clicked = True
                            break
                        except:
                            continue
                    
                    if not first_clicked:
                        print("   첫 번째 메뉴를 찾을 수 없습니다 - 수동으로 선택하세요")
                        await page.wait_for_timeout(10000)  # 수동 선택 대기
                    else:
                        await page.wait_for_timeout(2000)
                    
                    # 2단계: #menuAtag_4601020000 > span 선택 (두 번째 메뉴)
                    print("2단계: 두 번째 메뉴 선택 (#menuAtag_4601020000)...")
                    
                    second_menu_selectors = [
                        "#menuAtag_4601020000 > span",
                        "#menuAtag_4601020000",
                        "*[id*='menuAtag'][id*='4601020000'] > span",
                        "*[id*='menuAtag'][id*='4601020000']",
                        "a[href*='4601020000'] > span",
                        "a[href*='4601020000']"
                    ]
                    
                    second_clicked = False
                    for selector in second_menu_selectors:
                        try:
                            print(f"   시도: {selector}")
                            second_menu = page.locator(selector).first
                            await second_menu.wait_for(state="visible", timeout=3000)
                            await second_menu.click()
                            print(f"   두 번째 메뉴 클릭 성공: {selector}")
                            second_clicked = True
                            break
                        except:
                            continue
                    
                    if not second_clicked:
                        print("   두 번째 메뉴를 찾을 수 없습니다 - 수동으로 선택하세요")
                        await page.wait_for_timeout(10000)  # 수동 선택 대기
                    else:
                        await page.wait_for_timeout(2000)
                    
                    # 3단계: #menuAtag_4601020100 > span 선택 (세 번째 메뉴)
                    print("3단계: 세 번째 메뉴 선택 (#menuAtag_4601020100)...")
                    
                    third_menu_selectors = [
                        "#menuAtag_4601020100 > span",
                        "#menuAtag_4601020100",
                        "*[id*='menuAtag'][id*='4601020100'] > span",
                        "*[id*='menuAtag'][id*='4601020100']",
                        "a[href*='4601020100'] > span",
                        "a[href*='4601020100']"
                    ]
                    
                    third_clicked = False
                    for selector in third_menu_selectors:
                        try:
                            print(f"   시도: {selector}")
                            third_menu = page.locator(selector).first
                            await third_menu.wait_for(state="visible", timeout=3000)
                            await third_menu.click()
                            print(f"   세 번째 메뉴 클릭 성공: {selector}")
                            third_clicked = True
                            break
                        except:
                            continue
                    
                    if not third_clicked:
                        print("   세 번째 메뉴를 찾을 수 없습니다 - 수동으로 선택하세요")
                        await page.wait_for_timeout(10000)  # 수동 선택 대기
                    else:
                        await page.wait_for_timeout(2000)
                    
                    # 4단계: #mf_txppWframe_textbox1395 클릭
                    print("4단계: 텍스트박스 클릭 (#mf_txppWframe_textbox1395)...")
                    
                    textbox_selectors = [
                        "#mf_txppWframe_textbox1395",
                        "*[id*='textbox1395']",
                        "*[id*='mf_txppWframe'] *[id*='textbox1395']",
                        "input[id*='textbox1395']"
                    ]
                    
                    textbox_clicked = False
                    for selector in textbox_selectors:
                        try:
                            print(f"   시도: {selector}")
                            textbox = page.locator(selector).first
                            await textbox.wait_for(state="visible", timeout=3000)
                            await textbox.click()
                            print(f"   텍스트박스 클릭 성공: {selector}")
                            textbox_clicked = True
                            break
                        except:
                            continue
                    
                    if not textbox_clicked:
                        print("   텍스트박스를 찾을 수 없습니다 - 수동으로 선택하세요")
                        await page.wait_for_timeout(10000)  # 수동 선택 대기
                    else:
                        await page.wait_for_timeout(2000)
                    
                    # 2-4단계 메뉴 클릭 후 팝업 처리
                    print("5단계: 메뉴 클릭 후 팝업 처리...")
                    await page.wait_for_timeout(2000)  # 팝업이 나타날 시간 대기
                    
                    try:
                        # Alert 대화상자 자동 처리
                        alert_count = 0
                        def handle_second_dialog(dialog):
                            nonlocal alert_count
                            alert_count += 1
                            print(f"   Alert {alert_count} 감지 및 처리: '{dialog.message}'")
                            asyncio.create_task(dialog.accept())
                        
                        page.on("dialog", handle_second_dialog)
                        
                        # 새 팝업창 확인 및 닫기
                        popup_processed = False
                        for check in range(5):  # 5초간 확인
                            await page.wait_for_timeout(1000)
                            
                            # 새로운 팝업창 확인
                            try:
                                context_pages = page.context.pages
                                if len(context_pages) > 1:
                                    print(f"   새 팝업창 감지: {len(context_pages) - 1}개")
                                    
                                    # 메인 페이지가 아닌 모든 창 닫기
                                    for popup_page in context_pages:
                                        if popup_page != page:
                                            try:
                                                await popup_page.close()
                                                print("   새 팝업창 닫기 완료")
                                                popup_processed = True
                                            except:
                                                pass
                            except:
                                pass
                            
                            # Alert 처리됨 확인
                            if alert_count > 0:
                                print(f"   Alert {alert_count}개 처리 완료")
                                popup_processed = True
                        
                        # 알림창 확인 버튼으로 닫기
                        try:
                            print("   알림창 확인 버튼 찾는 중...")
                            notification_confirm = page.locator("#mf_txppWframe_UTEETZZD02_wframe_btnProcess")
                            await notification_confirm.wait_for(state="visible", timeout=3000)
                            await notification_confirm.click()
                            print("   알림창 확인 버튼 클릭 완료")
                            popup_processed = True
                            await page.wait_for_timeout(1000)
                        except Exception as e:
                            print(f"   알림창 확인 버튼 없음 또는 클릭 실패: {e}")
                        
                        # Alert 핸들러 제거
                        page.remove_listener("dialog", handle_second_dialog)
                        
                        if popup_processed:
                            print("   팝업/Alert 처리 완료")
                        else:
                            print("   팝업/Alert 없음 - 정상 진행")
                            
                    except Exception as popup_error:
                        print(f"   팝업 처리 중 오류: {popup_error}")
                    
                    await page.wait_for_timeout(2000)
                    print("✅ 전체 메뉴 네비게이션 완료!")
                    
                    # 엑셀 파일 열기 및 열 선택 단계 추가
                    print("\n=== 6단계: 엑셀 파일 열기 및 열 선택 ===")
                    
                    # 사용자에게 엑셀 파일 선택 요청
                    print("HomeTax에서 엑셀 파일을 업로드하거나 선택해주세요.")
                    print("완료되면 Enter를 눌러주세요.")
                    input("엑셀 파일 선택 완료 후 Enter: ")
                    
                    # 열 선택 받기
                    print("\n=== 처리할 열 선택 ===")
                    print("예시:")
                    print("- 단일 열: A 또는 1")
                    print("- 복수 열: A,C,E 또는 1,3,5")
                    print("- 범위: A-E 또는 1-5")
                    print("- 혼합: A,C-E,G 또는 1,3-5,7")
                    
                    column_selection = input("\n처리할 열을 선택하세요: ").strip()
                    if column_selection:
                        print(f"선택된 열: {column_selection}")
                        # 열 선택 정보를 page 객체에 저장
                        setattr(page, '_selected_columns', column_selection)
                    else:
                        print("열이 선택되지 않았습니다.")
                        setattr(page, '_selected_columns', None)
                    
                except Exception as nav_error:
                    print(f"❌ 메뉴 네비게이션 오류: {nav_error}")
                    print("   수동으로 메뉴를 선택해주세요.")
                
                return page  # 성공 시 page 객체 반환
                
            else:
                print("⚠️  로그인 상태 확인 필요")
                print("   브라우저에서 직접 확인해주세요.")
                return None
            
        except Exception as e:
            print(f"오류: {e}")
            return None
        # finally 블록 제거 - 성공 시 browser를 닫지 않음

# ========== 기존 거래처 정보 처리 기능 ==========

# field_mapping.md 기반 변수 매핑
FIELD_MAPPING = {
    'c_date': '작성일',
    'business_num': '사업자번호', 
    'sub_business_num': '종사업장번호',
    'company_name': '업체명',
    'ceo_name': '대표자',
    'address': '주소',
    'business_type': '업태',
    'business_item': '종목',
    'main_dept_name': '주부서명',
    'main_manager_name': '주담당자명',
    'main_phone': '주전화번호',
    'main_mobile': '주휴대폰',
    'main_fax': '주팩스',
    'main1_email': '주이메일앞',
    'main2_email': '주이메일뒤', 
    'main_memo': '주비고',
    'sub_dept_name': '부부서명',
    'sub_manager_name': '부담당자명',
    'sub_phone': '부전화번호',
    'sub_mobile': '부휴대폰',
    'sub_fax': '부팩스',
    'sub1_email': '부이메일앞',
    'sub2_email': '부이메일뒤',
    'sub_memo': '부비고'
}

class BusinessPartner:
    """거래처 정보를 담는 클래스"""
    def __init__(self):
        self.c_date = ""
        self.business_num = ""
        self.sub_business_num = ""
        self.company_name = ""
        self.ceo_name = ""
        self.address = ""
        self.business_type = ""
        self.business_item = ""
        self.main_dept_name = ""
        self.main_manager_name = ""
        self.main_phone = ""
        self.main_mobile = ""
        self.main_fax = ""
        self.main1_email = ""
        self.main2_email = ""
        self.main_memo = ""
        self.sub_dept_name = ""
        self.sub_manager_name = ""
        self.sub_phone = ""
        self.sub_mobile = ""
        self.sub_fax = ""
        self.sub1_email = ""
        self.sub2_email = ""
        self.sub_memo = ""
    
    def __str__(self):
        return f"""
거래처 정보:
- 업체명: {self.company_name}
- 사업자번호: {self.business_num}
- 대표자: {self.ceo_name}
- 주소: {self.address}
- 업태: {self.business_type}
- 종목: {self.business_item}
- 주담당자: {self.main_manager_name} ({self.main_phone})
        """

class TaxInvoice:
    def __init__(self, supplier_name, supplier_registration_number, customer_name, customer_registration_number, item, quantity, unit_price):
        self.supplier_name = supplier_name
        self.supplier_registration_number = supplier_registration_number
        self.customer_name = customer_name
        self.customer_registration_number = customer_registration_number
        self.item = item
        self.quantity = quantity
        self.unit_price = unit_price
        self.supply_price = quantity * unit_price
        self.tax_amount = int(self.supply_price * 0.1)
        self.total_price = self.supply_price + self.tax_amount

    def __str__(self):
        return f"""
        세금계산서
        공급자: {self.supplier_name} ({self.supplier_registration_number})
        공급받는자: {self.customer_name} ({self.customer_registration_number})
        품목: {self.item}
        수량: {self.quantity}
        단가: {self.unit_price}
        공급가액: {self.supply_price}
        세액: {self.tax_amount}
        합계금액: {self.total_price}
        """

def load_excel_file():
    """
    기본 엑셀 파일을 열거나 파일 선택 다이얼로그를 표시
    """
    default_file = r"C:\Users\man4k\OneDrive\문서\세금계산서.xlsx"
    
    # 기본 파일이 존재하는지 확인
    if os.path.exists(default_file):
        try:
            df = pd.read_excel(default_file)
            print(f"기본 파일을 성공적으로 로드했습니다: {default_file}")
            return df, default_file
        except Exception as e:
            print(f"기본 파일 로드 중 오류 발생: {e}")
            return select_excel_file()
    else:
        print("기본 파일이 존재하지 않습니다. 파일을 선택해주세요.")
        return select_excel_file()

def select_excel_file():
    """
    파일 선택 다이얼로그를 표시하고 선택된 엑셀 파일을 로드
    """
    root = tk.Tk()
    root.withdraw()  # tkinter 메인 창 숨기기
    
    file_path = filedialog.askopenfilename(
        title="세금계산서 엑셀 파일 선택",
        filetypes=[
            ("Excel files", "*.xlsx *.xls"),
            ("All files", "*.*")
        ],
        initialdir=r"C:\Users\man4k\OneDrive\문서"
    )
    
    if file_path:
        try:
            df = pd.read_excel(file_path)
            print(f"파일을 성공적으로 로드했습니다: {file_path}")
            return df, file_path
        except Exception as e:
            messagebox.showerror("오류", f"파일 로드 중 오류 발생:\n{e}")
            return None, None
    else:
        print("파일이 선택되지 않았습니다.")
        return None, None

def parse_row_selection(selection_str):
    """
    행 선택 문자열을 파싱하여 행 번호 리스트를 반환
    예: "2,4,8" → [2, 4, 8]
        "2-8" → [2, 3, 4, 5, 6, 7, 8]
        "2" → [2]
    """
    rows = []
    
    for part in selection_str.split(','):
        part = part.strip()
        if '-' in part:
            # 범위 처리 (예: 2-8)
            start, end = part.split('-')
            start = int(start.strip())
            end = int(end.strip())
            rows.extend(range(start, end + 1))
        else:
            # 단일 행 (예: 2)
            rows.append(int(part))
    
    return sorted(list(set(rows)))  # 중복 제거 및 정렬

def map_excel_to_variables(df, row_index):
    """
    엑셀 데이터의 특정 행을 BusinessPartner 객체의 변수에 매핑
    """
    partner = BusinessPartner()
    
    for var_name, excel_col_name in FIELD_MAPPING.items():
        if excel_col_name in df.columns:
            value = df.loc[row_index, excel_col_name]
            # NaN 값을 빈 문자열로 처리
            if pd.isna(value):
                value = ""
            else:
                value = str(value).strip()
            setattr(partner, var_name, value)
        else:
            print(f"경고: 엑셀에서 '{excel_col_name}' 열을 찾을 수 없습니다.")
    
    return partner

async def input_partner_to_hometax(page, partner):
    """
    거래처 정보를 HomeTax에 자동 입력하는 함수
    """
    try:
        print("HomeTax 거래처 입력 시작...")
        
        # field_mapping.md의 HomeTax 셀렉션명을 사용하여 자동 입력
        hometax_selectors = {
            'business_num': '#mf_txppWframe_txtBsno1',
            'company_name': '#mf_txppWframe_txtTnmNm',
            'ceo_name': '#mf_txppWframe_txtRprs',
            'address': '#mf_txppWframe_edtSplrPfbAdrTop',
            'business_type': '#mf_txppWframe_edtSplrBcNmTop',
            'business_item': '#mf_txppWframe_edtSplrItmNmTop',
            'main_manager_name': '#mf_txppWframe_txtChrgNm',
            'main_phone': '#mf_txppWframe_txtChrgTelNo1',
            'main_mobile': '#mf_txppWframe_txtChrgMpNo2',
            'main_fax': '#mf_txppWframe_txtChrgFaxNo3',
            'main1_email': '#mf_txppWframe_txtChrgEmlAdr1',
            'main2_email': '#mf_txppWframe_txtChrgEmlAdr2',
            'main_memo': '#mf_txppWframe_txtChrgRmrk',
            'sub_manager_name': '#mf_txppWframe_txtSchrgNm',
            'sub_phone': '#mf_txppWframe_txtSchrgTelNo1',
            'sub_mobile': '#mf_txppWframe_txtSchrgMpNo2',
            'sub_fax': '#mf_txppWframe_txtSchrgFaxNo3',
            'sub1_email': '#mf_txppWframe_txtSchrgEmlAdr1',
            'sub2_email': '#mf_txppWframe_txtSchrgEmlAdr2',
            'sub_memo': '#mf_txppWframe_txtSchrgRmrk'
        }
        
        # 각 필드에 데이터 입력
        for field_name, selector in hometax_selectors.items():
            value = getattr(partner, field_name, "")
            if value and value.strip():
                try:
                    element = page.locator(selector).first
                    await element.wait_for(state="visible", timeout=3000)
                    await element.clear()
                    await element.fill(str(value).strip())
                    print(f"✅ {field_name}: {value}")
                    await page.wait_for_timeout(100)  # 입력 간 짧은 대기
                except Exception as e:
                    print(f"⚠️ {field_name} 입력 실패: {e}")
        
        print("✅ HomeTax 거래처 정보 입력 완료!")
        return True
        
    except Exception as e:
        print(f"❌ HomeTax 입력 중 오류: {e}")
        return False

async def process_single_partner(partner, row_num, page=None):
    """
    단일 거래처 정보를 처리하는 함수 (HomeTax 자동 입력 포함)
    """
    print(f"\n=== 행 {row_num} 처리 중 ===")
    print(partner)
    
    if page:
        # HomeTax 페이지가 있으면 자동 입력 시도
        success = await input_partner_to_hometax(page, partner)
        if success:
            print("✅ HomeTax 자동 입력 성공!")
        else:
            print("⚠️ HomeTax 자동 입력 실패 - 수동으로 확인 필요")
    else:
        print("🔍 HomeTax 페이지 없음 - 정보만 출력")
    
    print("처리 완료!\n")

async def process_selected_rows(df, row_selection, page=None):
    """
    선택된 행들을 순차적으로 처리 (HomeTax 자동 입력 포함)
    """
    try:
        selected_rows = parse_row_selection(row_selection)
        print(f"처리할 행: {selected_rows}")
        
        # 각 행을 순차적으로 처리
        for i, row_num in enumerate(selected_rows):
            # 엑셀 행 번호는 1부터 시작하지만 pandas 인덱스는 0부터 시작
            pandas_index = row_num - 1
            
            if pandas_index >= len(df):
                print(f"경고: 행 {row_num}이 데이터 범위를 초과합니다. (최대: {len(df)})")
                continue
            
            # 엑셀 데이터를 변수에 매핑
            partner = map_excel_to_variables(df, pandas_index)
            
            # 단일 거래처 처리 (async 함수 호출)
            await process_single_partner(partner, row_num, page)
            
            # 사용자가 다음 처리를 위해 확인할 수 있도록
            if i < len(selected_rows) - 1:  # 마지막 행이 아닌 경우
                response = input("다음 행을 계속 처리하시겠습니까? (y/n, Enter=y): ").strip()
                if response.lower() == 'n':
                    print("처리를 중단합니다.")
                    break
        
        print("모든 선택된 행 처리 완료!")
        
    except ValueError as e:
        print(f"오류: 행 선택 형식이 올바르지 않습니다. {e}")
    except Exception as e:
        print(f"처리 중 오류 발생: {e}")

async def main():
    """
    메인 실행 함수 (HomeTax 통합 버전)
    """
    print("=== 거래처 정보 처리 프로그램 (HomeTax 통합) ===")
    
    # 엑셀 파일 로드
    df, file_path = load_excel_file()
    
    if df is not None:
        print(f"\n엑셀 파일 정보:")
        print(f"파일 경로: {file_path}")
        print(f"데이터 크기: {df.shape}")
        print("\n엑셀 열 정보:")
        print("사용 가능한 열:", list(df.columns))
        print("\n첫 5행 미리보기:")
        print(df.head())
        
        # HomeTax 로그인 여부 선택
        print("\n=== HomeTax 연동 선택 ===")
        use_hometax = input("HomeTax 자동 로그인 및 입력을 사용하시겠습니까? (y/n): ").strip().lower() == 'y'
        
        page = None
        if use_hometax:
            print("\nHomeTax 로그인 시작...")
            page = await hometax_quick_login()
            if page:
                print("✅ HomeTax 로그인 성공! 자동 입력 모드로 진행합니다.")
            else:
                print("⚠️ HomeTax 로그인 실패. 정보만 출력하는 모드로 진행합니다.")
        
        # 행 선택 입력 받기
        print("\n=== 행 선택 방법 ===")
        print("- 단일 행: 2")
        print("- 복수 행: 2,4,8")  
        print("- 범위: 2-8")
        print("- 혼합: 2,5-7,10")
        
        while True:
            try:
                row_selection = input("\n처리할 행을 선택하세요: ").strip()
                if not row_selection:
                    print("행 선택이 필요합니다.")
                    continue
                
                # 선택된 행들 처리 (HomeTax 연동 포함)
                await process_selected_rows(df, row_selection, page)
                
                # 추가 처리 여부 확인
                response = input("\n다른 행을 더 처리하시겠습니까? (y/n): ").strip()
                if response.lower() != 'y':
                    break
                    
            except KeyboardInterrupt:
                print("\n프로그램을 종료합니다.")
                break
                
        # HomeTax 브라우저 정리 (선택적)
        if page and use_hometax:
            response = input("\nHomeTax 브라우저를 닫으시겠습니까? (y/n): ").strip()
            if response.lower() == 'y':
                try:
                    await page.context.browser.close()
                    print("브라우저를 닫았습니다.")
                except:
                    pass
            else:
                print("브라우저를 열어둡니다. 수동으로 닫아주세요.")
    else:
        print("파일을 로드하지 못했습니다. 프로그램을 종료합니다.")

if __name__ == "__main__":
    asyncio.run(main())
