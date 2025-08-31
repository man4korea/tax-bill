# 📁 C:\APP\tax-bill\core\auto_login.py
# Create at 2508312118 Ver1.00
# Update at 2508312142 Ver1.01
# -*- coding: utf-8 -*-
"""
홈택스 자동 로그인 모듈 (Playwright 버전)
import 가능한 모듈로 변경 - 브라우저 세션 연속성 지원
"""

import asyncio
import time
import sys
import os
from pathlib import Path
from playwright.async_api import async_playwright

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


async def auto_login_with_playwright():
    """Playwright를 사용한 홈택스 자동 로그인"""
    
    # async with를 사용하지 않고 직접 playwright 인스턴스 생성
    # 이렇게 하면 함수가 끝나도 브라우저가 닫히지 않습니다
    p = await async_playwright().start()
        print("[AUTO] Playwright 브라우저 실행 중...")
        
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
            print("[AUTO] 홈택스 페이지로 이동 중...")
            await page.goto("https://hometax.go.kr/websquare/websquare.html?w2xPath=/ui/pp/index_pp.xml&menuCd=index3")
            
            print("[OK] 홈택스 페이지 열기 완료!")
            print(f"페이지 제목: {await page.title()}")
            print(f"현재 URL: {page.url}")
            
            # 2단계: 공동·금융인증서 버튼 클릭
            print("[AUTO] 공동·금융인증서 버튼 클릭 시도...")
            
            # 버튼이 클릭 가능할 때까지 최대 15초 대기
            print("[WAIT] 버튼 로딩 대기 중...")
            await page.wait_for_selector("#mf_txppWframe_loginboxFrame_anchor22", timeout=15000)
            await page.click("#mf_txppWframe_loginboxFrame_anchor22")
            
            print("[OK] 공동·금융인증서 버튼 클릭 성공!")
            
            # 공인인증서 창이 뜰 때까지 잠시 대기
            await page.wait_for_timeout(3000)
            print("[INFO] 현재 상태 확인...")
            print(f"현재 URL: {page.url}")
            
            # 3단계: 공인인증서 비밀번호 자동 입력
            print("[AUTO] 공인인증서 비밀번호 자동 입력 시작...")
            
            # 비밀번호 로드
            password = get_certificate_password()
            
            if not password:
                print("[ERROR] 저장된 비밀번호를 찾을 수 없습니다!")
                print("[INFO] hometax_cert_manager.py에서 비밀번호를 먼저 저장해주세요.")
                return None, None
            
            print("[OK] 비밀번호 로드 성공")
            
            # iframe #dscert로 전환 시도
            print("[AUTO] iframe #dscert로 전환 중...")
            await page.wait_for_timeout(2000)  # iframe 로딩 대기
            
            frame = page.frame("dscert")
            
            if frame:
                try:
                    # iframe 내부에서 비밀번호 입력 필드 찾기
                    print("[SEARCH] 비밀번호 입력 필드 검색 중...")
                    await frame.wait_for_selector("input[type='password']", timeout=10000)
                    
                    # 비밀번호 입력
                    await frame.fill("input[type='password']", password)
                    print("[OK] 공인인증서 비밀번호 입력 완료!")
                    
                    # 확인 버튼 클릭
                    print("[AUTO] 공인인증서 확인 버튼 클릭...")
                    await frame.wait_for_selector("#btn_confirm_iframe > span", timeout=10000)
                    await frame.click("#btn_confirm_iframe > span")
                    print("[OK] 확인 버튼 클릭 성공!")
                    
                    # 로그인 처리 대기
                    print("[WAIT] 로그인 처리 중... (10초 대기)")
                    await page.wait_for_timeout(10000)
                    
                    # URL 변화 확인 (로그인 성공 여부)
                    current_url = page.url
                    print(f"[INFO] 현재 URL: {current_url}")
                    
                    if "index_pp.xml" not in current_url:
                        print("[SUCCESS] 홈택스 로그인 완료!")
                        return page, browser
                    else:
                        print("[WARNING] 로그인이 완료되지 않았을 수 있습니다.")
                        print("[MANUAL] 수동으로 공인인증서 비밀번호를 입력하고 확인을 눌러주세요.")
                        
                        # 수동 로그인 대기 (최대 3분)
                        for i in range(180):  # 3분 = 180초
                            await page.wait_for_timeout(1000)
                            current_url = page.url
                            if "index_pp.xml" not in current_url:
                                print("[SUCCESS] 수동 로그인 완료 감지!")
                                return page, browser
                            if i % 30 == 0:  # 30초마다 상태 출력
                                print(f"[WAIT] 수동 로그인 대기 중... ({i//60+1}분 경과)")
                        
                        print("[TIMEOUT] 수동 로그인 시간 초과")
                        return page, browser
                        
                except Exception as e:
                    print(f"[ERROR] 자동 입력 실패: {e}")
                    print("[MANUAL] 수동으로 공인인증서 비밀번호를 입력하고 확인을 눌러주세요.")
                    
                    # 수동 로그인 대기
                    for i in range(180):  # 3분
                        await page.wait_for_timeout(1000)
                        current_url = page.url
                        if "index_pp.xml" not in current_url:
                            print("[SUCCESS] 수동 로그인 완료 감지!")
                            return page, browser
                        if i % 30 == 0:
                            print(f"[WAIT] 수동 로그인 대기 중... ({i//60+1}분 경과)")
                    
                    return page, browser
                
            else:
                print("[ERROR] iframe #dscert를 찾을 수 없습니다.")
                print("[MANUAL] 수동으로 공인인증서 비밀번호를 입력하고 확인을 눌러주세요.")
                
                # 수동 로그인 대기
                for i in range(180):  # 3분
                    await page.wait_for_timeout(1000)
                    current_url = page.url
                    if "index_pp.xml" not in current_url:
                        print("[SUCCESS] 수동 로그인 완료 감지!")
                        return page, browser
                    if i % 30 == 0:
                        print(f"[WAIT] 수동 로그인 대기 중... ({i//60+1}분 경과)")
                
                return page, browser
                
        except Exception as e:
            print(f"[ERROR] 자동 로그인 중 오류 발생: {e}")
            await browser.close()
            return None, None


async def main():
    """독립 실행용 메인 함수 - 사용자가 직접 실행할 때만 사용"""
    page, browser = await auto_login_with_playwright()
    
    if page and browser:
        print("[SUCCESS] 자동 로그인 성공!")
        print("[INFO] 브라우저를 열린 상태로 유지합니다. (독립 실행 모드)")
        
        # 독립 실행 시에는 무한 대기 (사용자가 수동으로 닫을 때까지)
        try:
            while True:
                await asyncio.sleep(1)
        except KeyboardInterrupt:
            print("[EXIT] 프로그램 종료 중...")
            await browser.close()
    else:
        print("[ERROR] 자동 로그인 실패")


if __name__ == "__main__":
    asyncio.run(main())