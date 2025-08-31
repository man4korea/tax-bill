#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
MagicLine 감지 로직 테스트
개선된 hometax_login_module.py의 MagicLine 감지 기능 검증
"""

import asyncio
import sys
import os

# 상위 디렉터리의 core 모듈을 import 할 수 있도록 경로 추가
sys.path.append(os.path.join(os.path.dirname(__file__), '..', 'core'))

from playwright.async_api import async_playwright


async def test_magicline_detection():
    """MagicLine 감지 로직 테스트"""
    print("🧪 MagicLine 감지 로직 테스트 시작...")
    
    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=False)
        page = await browser.new_page()
        
        try:
            # 홈택스 페이지 로드
            print("📄 홈택스 페이지 로딩...")
            await page.goto("https://hometax.go.kr/websquare/websquare.html?w2xPath=/ui/pp/index_pp.xml&menuCd=index3", timeout=60000)
            
            # DOM 로드 상태 확인
            print("⏳ DOM 로드 상태 확인...")
            await page.wait_for_load_state("domcontentloaded")
            await page.wait_for_load_state("networkidle")
            print("✅ DOM 로드 완료")
            
            # 개선된 MagicLine 감지 로직 테스트
            magicline_detected = False
            
            # 방법 1: window.magicline 객체 감지
            print("🔍 방법 1: window.magicline 객체 감지 시도...")
            try:
                await page.wait_for_function(
                    "() => typeof window.magicline !== 'undefined' && window.magicline && typeof window.magicline.AGENT_VER !== 'undefined'",
                    timeout=8000
                )
                magicline_detected = True
                print("✅ 방법 1 성공: window.magicline 객체 감지 완료")
            except Exception as e1:
                print(f"❌ 방법 1 실패: {str(e1)}")
                
                # 방법 2: 스크립트 태그 감지
                print("🔍 방법 2: MagicLine 스크립트 태그 감지 시도...")
                try:
                    await page.wait_for_selector("script[src*='magicline'], script[src*='MagicLine'], script[id*='magicline']", timeout=5000)
                    magicline_detected = True
                    print("✅ 방법 2 성공: MagicLine 스크립트 태그 감지 완료")
                except Exception as e2:
                    print(f"❌ 방법 2 실패: {str(e2)}")
                    
                    # 방법 3: DOM 요소 감지
                    print("🔍 방법 3: 보안 프로그램 DOM 요소 감지 시도...")
                    try:
                        await page.wait_for_function(
                            "() => document.querySelector('#magicline_div, .magicline, [id*=\"magicline\"]') !== null || document.querySelector('embed[type*=\"security\"], object[type*=\"security\"]') !== null",
                            timeout=3000
                        )
                        magicline_detected = True
                        print("✅ 방법 3 성공: MagicLine 관련 DOM 요소 감지 완료")
                    except Exception as e3:
                        print(f"❌ 방법 3 실패: {str(e3)}")
            
            # 결과 출력
            if magicline_detected:
                print("🎉 MagicLine 감지 성공!")
            else:
                print("⚠️ 모든 방법 실패 - MagicLine 감지 불가")
            
            # 페이지 정보 수집 (디버깅용)
            print("\n📊 페이지 디버깅 정보:")
            try:
                scripts = await page.evaluate("() => Array.from(document.querySelectorAll('script')).map(s => s.src || s.id || 'inline').filter(s => s)")
                print(f"   - 스크립트 개수: {len(scripts)}")
                print(f"   - MagicLine 관련 스크립트: {[s for s in scripts if 'magicline' in s.lower()]}")
                
                window_magicline = await page.evaluate("() => typeof window.magicline")
                print(f"   - window.magicline 타입: {window_magicline}")
                
                if window_magicline != 'undefined':
                    magicline_props = await page.evaluate("() => window.magicline ? Object.keys(window.magicline) : []")
                    print(f"   - window.magicline 속성: {magicline_props}")
                
            except Exception as e:
                print(f"   - 디버깅 정보 수집 실패: {str(e)}")
            
        except Exception as e:
            print(f"❌ 테스트 중 오류 발생: {str(e)}")
        
        finally:
            print("🔚 테스트 완료 - 브라우저를 5초 후 닫습니다...")
            await asyncio.sleep(5)
            await browser.close()


if __name__ == "__main__":
    asyncio.run(test_magicline_detection())