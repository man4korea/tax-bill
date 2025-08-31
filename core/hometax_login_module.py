# 📁 C:\APP\tax-bill\core\hometax_login_module.py
# Create at 2508312118 Ver1.00
# Update at 2508312142 Ver1.01
#-*- coding: utf-8 -*-
"""
홈택스 공통 로그인 모듈 (직접 함수 호출 방식)
브라우저 세션 연속성 지원 - subprocess 제거, 직접 import 방식으로 변경
hometax_partner_registration.py와 hometax_tax_invoice.py에서 공통으로 사용하는 로그인 프로세스
"""

import os
import asyncio
import sys
import base64
from pathlib import Path
from dotenv import load_dotenv
from playwright.async_api import async_playwright
from hometax_security_manager import HomeTaxSecurityManager

# Windows 콘솔에서 UTF-8 지원 설정
if sys.platform == "win32":
    # 콘솔 출력 인코딩을 UTF-8로 설정
    os.system("chcp 65001 > nul 2>&1")  # 조용히 실행
    # Python 표준 출력/에러 스트림을 UTF-8로 재설정
    sys.stdout.reconfigure(encoding='utf-8')
    sys.stderr.reconfigure(encoding='utf-8')


def decrypt_password_from_env(encrypted_config):
    """암호화된 설정에서 비밀번호 복호화 (cert_manager와 동일)"""
    try:
        # 앞뒤 문자열 제거
        if not encrypted_config.startswith("HTC_") or not encrypted_config.endswith("_CFG"):
            return None
        
        middle_part = encrypted_config[4:-4]  # "HTC_"와 "_CFG" 제거
        original_encoded = middle_part[::-1]  # 역순 되돌리기
        decoded = base64.b64decode(original_encoded.encode('utf-8')).decode('utf-8')
        return decoded
    except Exception as e:
        print(f"❌ 비밀번호 복호화 실패: {e}")
        return None


def load_encrypted_config_from_env(env_file):
    """암호화된 설정을 .env 파일에서 로드"""
    try:
        if not env_file.exists():
            return None
        
        with open(env_file, 'r', encoding='utf-8') as f:
            lines = f.readlines()
        
        for line in lines:
            line = line.strip()
            if line.startswith('HTC_CONFIG='):
                encrypted_config = line.split('=', 1)[1].strip()
                password = decrypt_password_from_env(encrypted_config)
                if password:
                    print("🔐 암호화된 설정에서 비밀번호 로드 성공")
                    return password
        
        print("📋 .env에서 암호화된 설정을 찾을 수 없음")
        return None
        
    except Exception as e:
        print(f"❌ 암호화 설정 로드 실패: {e}")
        return None


def get_certificate_password():
    """인증서 비밀번호 가져오기 (새로운 암호화 방식 우선)"""
    try:
        # 1. 새로운 암호화 방식 시도
        env_file = Path(__file__).parent.parent / ".env"
        encrypted_password = load_encrypted_config_from_env(env_file)
        if encrypted_password:
            print("✅ 새로운 암호화 방식으로 비밀번호 로드")
            return encrypted_password
        
        # 2. 레거시 방식 시도 (보안 관리자)
        security_manager = HomeTaxSecurityManager()
        legacy_password = security_manager.load_password_from_env()
        if legacy_password:
            print("✅ 레거시 방식으로 비밀번호 로드 (업그레이드 권장)")
            return legacy_password
        
        # 3. 기본 환경변수 방식 (가장 기본)
        basic_password = os.getenv("PW")
        if basic_password:
            print("✅ 기본 환경변수에서 비밀번호 로드")
            return basic_password
        
        print("❌ 저장된 비밀번호를 찾을 수 없습니다")
        return None
        
    except Exception as e:
        print(f"❌ 비밀번호 로드 중 오류: {e}")
        return None


async def hometax_login_dispatcher(login_complete_callback=None):
    """홈택스 로그인 모드 분기 처리 - 직접 함수 호출 방식"""
    
    # 1. .env 파일에서 설정값 가져오기 (프로젝트 루트)
    env_file = Path(__file__).parent.parent / ".env"
    load_dotenv(env_file)
    
    # 2. auto_login, manual_login 판정
    login_mode = os.getenv("HOMETAX_LOGIN_MODE", "auto")
    print(f"📋 로그인 모드: {login_mode}")
    
    try:
        # 3. 로그인 분기 - 직접 함수 호출로 브라우저 객체 받기
        if login_mode == "auto":
            print("🤖 자동 로그인 모듈 실행 중...")
            from auto_login import auto_login_with_playwright
            page, browser = await auto_login_with_playwright()
        else:  # manual
            print("👤 수동 로그인 모듈 실행 중...")  
            from manual_login import manual_login_with_playwright
            page, browser = await manual_login_with_playwright()
        
        # 4. 로그인 결과 확인 및 콜백 실행
        if page and browser:
            print("✅ 로그인 성공!")
            
            # 콜백 함수가 있으면 실행
            if login_complete_callback:
                print("🔄 콜백 함수 실행 중...")
                result = await login_complete_callback(page, browser)
                return result
            else:
                print("✅ 로그인 완료 - 브라우저 객체 반환")
                return page, browser
        else:
            print("❌ 로그인 실패")
            # fallback 시도
            print("🔄 fallback 방식으로 재시도...")
            return await _fallback_manual_login(login_complete_callback)
        
    except Exception as e:
        print(f"❌ 로그인 디스패처 오류: {e}")
        # fallback으로 기존 방식 시도
        print("🔄 fallback 방식으로 재시도...")
        return await _fallback_manual_login(login_complete_callback)



async def _fallback_manual_login(callback=None):
    """기존 방식 fallback 로그인"""
    print("[FALLBACK] 기존 Playwright 방식으로 fallback 실행...")
    
    try:
        from playwright.async_api import async_playwright
        
        async with async_playwright() as p:
            browser = await p.chromium.launch(headless=False)
            page = await browser.new_page()
            
            # 홈택스 페이지 열기
            await page.goto("https://hometax.go.kr/websquare/websquare.html?w2xPath=/ui/pp/index_pp.xml&menuCd=index3")
            print("[OK] 홈택스 페이지 열기 완료 (fallback)")
            
            if callback:
                await callback(page, browser)
            
            return page, browser
            
    except Exception as e:
        print(f"[ERROR] Fallback 로그인 오류: {e}")
        return None, None


async def main():
    """메인 실행 함수 - 테스트용"""
    print("[TEST] 홈택스 로그인 모듈 테스트 실행")
    
    # 테스트 콜백 함수
    async def test_callback(page=None, browser=None):
        print("[OK] 테스트 콜백 함수 실행됨")
        if page:
            print(f"현재 URL: {page.url}")
        if browser:
            print("브라우저 객체 전달됨")
    
    # 로그인 디스패처 실행
    result = await hometax_login_dispatcher(test_callback)
    print(f"로그인 결과: {result}")


if __name__ == "__main__":
    asyncio.run(main())
