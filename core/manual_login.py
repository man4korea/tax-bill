# 📁 C:\APP\tax-bill\core\manual_login.py
# Create at 2508312118 Ver1.00
# Update at 2508312148 Ver1.01
# -*- coding: utf-8 -*-
"""
홈택스 수동 로그인 모듈 (Playwright 버전)
import 가능한 모듈로 변경 - 브라우저 세션 연속성 지원
사용자가 수동으로 로그인을 완료할 수 있도록 도움
"""

import asyncio
import time
import sys
import os
from pathlib import Path
from playwright.async_api import async_playwright
import tkinter as tk
from tkinter import messagebox

# Windows 콘솔에서 UTF-8 지원 설정
if sys.platform == "win32":
    # 콘솔 출력 인코딩을 UTF-8로 설정
    os.system("chcp 65001 > nul 2>&1")  # 조용히 실행
    # Python 표준 출력/에러 스트림을 UTF-8로 재설정
    sys.stdout.reconfigure(encoding='utf-8')
    sys.stderr.reconfigure(encoding='utf-8')

# 로그인 모듈 import
sys.path.append(str(Path(__file__).parent))
from hometax_login_module import get_certificate_password


async def manual_login_with_playwright():
    """Playwright를 사용한 홈택스 수동 로그인"""
    
    async with async_playwright() as p:
        print("[MANUAL] Playwright 브라우저 실행 중...")
        
        # 브라우저 실행 (Chrome 사용)
        browser = await p.chromium.launch(
            headless=False,  # 브라우저 창 표시
            args=[
                '--disable-blink-features=AutomationControlled',
                '--disable-dev-shm-usage',
                '--no-sandbox'
            ]
        )
        
        # 새 페이지 생성
        page = await browser.new_page()
        
        # 자동화 감지 우회
        await page.add_init_script("""
            Object.defineProperty(navigator, 'webdriver', {
                get: () => undefined,
            });
        """)
        
        try:
            # 1단계: 홈택스 페이지 열기
            print("[MANUAL] 홈택스 페이지로 이동 중...")
            await page.goto("https://hometax.go.kr/websquare/websquare.html?w2xPath=/ui/pp/index_pp.xml&menuCd=index3")
            
            print("[OK] 홈택스 페이지에 접속했습니다.")
            print(f"페이지 제목: {await page.title()}")
            print(f"현재 URL: {page.url}")
            
            # 1.5단계: 사용자 확인 메시지박스
            print("[USER] 사용자 확인 대기 중...")
            
            # tkinter 창 생성
            root = tk.Tk()
            root.withdraw()  # 메인 창 숨기기
            
            # 메시지박스 표시
            result = messagebox.askquestion(
                "홈택스 수동 로그인", 
                "홈택스 페이지가 열렸습니다.\n\n다음 단계를 진행하시겠습니까?\n- 공동·금융인증서 버튼 클릭\n- 인증서 비밀번호 자동 입력\n- 확인 버튼 클릭",
                icon='question'
            )
            
            root.destroy()
            
            if result == 'no':
                print("[CANCEL] 사용자가 취소했습니다.")
                await browser.close()
                return None, None
            
            print("[OK] 사용자 확인 완료, 로그인 과정을 계속 진행합니다.")
            
            # 2단계: 공동·금융인증서 버튼 클릭
            print("[MANUAL] 공동·금융인증서 버튼 클릭 시도...")
            
            try:
                # 요소가 로드될 때까지 대기
                await page.wait_for_selector("#mf_txppWframe_loginboxFrame_anchor22", timeout=10000)
                await page.click("#mf_txppWframe_loginboxFrame_anchor22")
                
                print("[OK] 공동·금융인증서 버튼 클릭 완료!")
                await page.wait_for_timeout(1000)
                
            except Exception as e:
                print(f"[ERROR] 공동·금융인증서 버튼 클릭 실패: {e}")
                print("요소를 찾을 수 없거나 클릭할 수 없습니다.")
                
            # 3단계: 공인인증서 비밀번호 자동 입력 (선택사항)
            print("[MANUAL] 공인인증서 비밀번호 자동 입력 시도...")
            
            try:
                # 비밀번호 로드
                password = get_certificate_password()
                
                if password:
                    print("[OK] 저장된 비밀번호 발견")
                    
                    # iframe #dscert로 전환 (잠시 대기 후)
                    await page.wait_for_timeout(3000)
                    print("[MANUAL] iframe #dscert로 전환 중...")
                    
                    frame = page.frame("dscert")
                    
                    if frame:
                        # iframe 내부에서 비밀번호 입력 필드 찾기
                        print("[SEARCH] 비밀번호 입력 필드 검색 중...")
                        await frame.wait_for_selector("input[type='password']", timeout=10000)
                        
                        # 비밀번호 입력
                        await frame.fill("input[type='password']", password)
                        
                        print("[OK] 공인인증서 비밀번호 입력 완료!")
                        
                        # 4단계: 공인인증서 확인 버튼 클릭
                        print("[MANUAL] 공인인증서 확인 버튼 클릭...")
                        
                        # 확인 버튼 대기 및 클릭
                        await frame.wait_for_selector("#btn_confirm_iframe > span", timeout=10000)
                        await frame.click("#btn_confirm_iframe > span")
                        
                        print("[OK] 확인 버튼 클릭 성공!")
                        
                        # 로그인 처리 대기
                        print("[WAIT] 로그인 처리 중... (5초 대기)")
                        await page.wait_for_timeout(5000)
                        
                        print(f"[INFO] 현재 URL: {page.url}")
                        print("[SUCCESS] 홈택스 로그인 완료!")
                        
                    else:
                        print("[ERROR] iframe #dscert를 찾을 수 없습니다.")
                        print("[INFO] 수동으로 인증서 비밀번호를 입력하고 확인 버튼을 클릭하세요.")
                        
                        # 사용자가 수동으로 로그인 완료할 때까지 대기
                        print("[WAIT] 수동 로그인 완료 대기 중... (최대 5분)")
                        
                        # 5분간 대기하면서 URL 변화 감지
                        for i in range(300):  # 5분 = 300초
                            await page.wait_for_timeout(1000)
                            current_url = page.url
                            if "index_pp.xml" not in current_url:  # URL이 변했으면 로그인 성공
                                print("[OK] 수동 로그인 완료 감지!")
                                break
                            if i % 30 == 0:  # 30초마다 상태 출력
                                print(f"[WAIT] 대기 중... ({i//60+1}분 경과)")
                
                else:
                    print("[ERROR] 저장된 인증서 비밀번호를 찾을 수 없습니다!")
                    print("[INFO] 수동으로 인증서 비밀번호를 입력하고 로그인을 완료하세요.")
                    
                    # 사용자가 수동으로 로그인 완료할 때까지 대기
                    print("[WAIT] 수동 로그인 완료 대기 중... (최대 5분)")
                    
                    # 5분간 대기하면서 URL 변화 감지
                    for i in range(300):  # 5분 = 300초
                        await page.wait_for_timeout(1000)
                        current_url = page.url
                        if "index_pp.xml" not in current_url:  # URL이 변했으면 로그인 성공
                            print("[OK] 수동 로그인 완료 감지!")
                            break
                        if i % 30 == 0:  # 30초마다 상태 출력
                            print(f"[WAIT] 대기 중... ({i//60+1}분 경과)")
                
                return page, browser
                
            except Exception as e:
                print(f"[ERROR] 인증서 처리 중 오류 발생: {e}")
                print("[INFO] 수동으로 로그인을 완료해주세요.")
                
                # 사용자가 수동으로 로그인 완료할 때까지 대기
                print("[WAIT] 수동 로그인 완료 대기 중... (최대 5분)")
                
                # 5분간 대기하면서 URL 변화 감지
                for i in range(300):  # 5분 = 300초
                    await page.wait_for_timeout(1000)
                    current_url = page.url
                    if "index_pp.xml" not in current_url:  # URL이 변했으면 로그인 성공
                        print("[OK] 수동 로그인 완료 감지!")
                        break
                    if i % 30 == 0:  # 30초마다 상태 출력
                        print(f"[WAIT] 대기 중... ({i//60+1}분 경과)")
                
                return page, browser
                
        except Exception as e:
            print(f"[ERROR] 수동 로그인 중 오류 발생: {e}")
            await browser.close()
            return None, None


async def main():
    """독립 실행용 메인 함수 - 사용자가 직접 실행할 때만 사용"""
    page, browser = await manual_login_with_playwright()
    
    if page and browser:
        print("[SUCCESS] 수동 로그인 성공!")
        print("[INFO] 브라우저를 열린 상태로 유지합니다. (독립 실행 모드)")
        
        # 독립 실행 시에는 무한 대기 (사용자가 수동으로 닫을 때까지)
        try:
            while True:
                await asyncio.sleep(1)
        except KeyboardInterrupt:
            print("[EXIT] 프로그램 종료 중...")
            await browser.close()
    else:
        print("[ERROR] 수동 로그인 실패")


if __name__ == "__main__":
    asyncio.run(main())