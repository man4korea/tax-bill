# 📁 C:\APP\tax-bill\core\utils\extract_logo.py
# Create at 2508312118 Ver1.00
# -*- coding: utf-8 -*-
"""
HomeTax 페이지에서 로고 이미지를 추출하는 스크립트
"""

import asyncio
import os
from pathlib import Path
from playwright.async_api import async_playwright
import requests
from PIL import Image
import io

async def extract_hometax_logo():
    """HomeTax 페이지에서 로고를 추출"""
    
    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=True)
        page = await browser.new_page()
        
        try:
            print("HomeTax 페이지 접속 중...")
            await page.goto("https://hometax.go.kr/websquare/websquare.html?w2xPath=/ui/pp/index_pp.xml&menuCd=index3")
            await page.wait_for_load_state('networkidle')
            await page.wait_for_timeout(3000)
            
            print("로고 이미지 검색 중...")
            
            # 여러 가능한 로고 셀렉터들
            logo_selectors = [
                "img[src*='logo']",
                "img[src*='hometax']",
                "img[alt*='hometax']",
                "img[alt*='국세청']",
                "img[alt*='홈택스']",
                ".logo img",
                "#logo img",
                "img[src*='nts']",  # National Tax Service
                "header img",
                ".header img",
                "img[src*='main']"
            ]
            
            found_logos = []
            
            for selector in logo_selectors:
                try:
                    elements = await page.locator(selector).all()
                    for element in elements:
                        src = await element.get_attribute('src')
                        alt = await element.get_attribute('alt') or ""
                        
                        if src:
                            # 상대경로를 절대경로로 변환
                            if src.startswith('/'):
                                src = f"https://hometax.go.kr{src}"
                            elif src.startswith('./'):
                                src = src.replace('./', 'https://hometax.go.kr/')
                            elif not src.startswith('http'):
                                src = f"https://hometax.go.kr/{src}"
                            
                            found_logos.append({
                                'src': src,
                                'alt': alt,
                                'selector': selector
                            })
                            print(f"발견된 로고: {src}")
                except Exception as e:
                    continue
            
            if not found_logos:
                print("로고를 찾을 수 없어 페이지 스크린샷을 촬영합니다...")
                # 전체 페이지 스크린샷
                await page.screenshot(path="hometax_page.png", full_page=True)
                print("스크린샷 저장: hometax_page.png")
                return None
            
            # 첫 번째 로고 이미지 다운로드
            logo_info = found_logos[0]
            logo_url = logo_info['src']
            
            print(f"로고 다운로드 중: {logo_url}")
            
            # 이미지 다운로드
            try:
                # 브라우저 컨텍스트에서 이미지 다운로드
                response = await page.context.request.get(logo_url)
                if response.status == 200:
                    image_data = await response.body()
                    
                    # PNG로 저장
                    with open("hometax_logo.png", "wb") as f:
                        f.write(image_data)
                    
                    print("로고 다운로드 완료: hometax_logo.png")
                    return "hometax_logo.png"
                else:
                    print(f"이미지 다운로드 실패: HTTP {response.status}")
                    return None
                    
            except Exception as e:
                print(f"이미지 다운로드 오류: {e}")
                return None
                
        except Exception as e:
            print(f"페이지 접속 오류: {e}")
            return None
        
        finally:
            await browser.close()

def create_ico_from_image(image_path, ico_path="hometax_icon.ico"):
    """이미지에서 ICO 파일 생성"""
    
    if not os.path.exists(image_path):
        print(f"이미지 파일을 찾을 수 없습니다: {image_path}")
        return None
    
    try:
        print(f"ICO 파일 생성 중: {image_path} -> {ico_path}")
        
        # 이미지 열기
        with Image.open(image_path) as img:
            # RGBA 모드로 변환 (투명도 지원)
            img = img.convert("RGBA")
            
            # 여러 크기의 아이콘 생성
            sizes = [(16, 16), (32, 32), (48, 48), (64, 64), (128, 128), (256, 256)]
            icons = []
            
            for size in sizes:
                # 비율을 유지하며 리사이즈
                resized = img.resize(size, Image.Resampling.LANCZOS)
                icons.append(resized)
            
            # ICO 파일로 저장
            icons[0].save(
                ico_path,
                format='ICO',
                sizes=[(icon.width, icon.height) for icon in icons]
            )
            
        print(f"ICO 파일 생성 완료: {ico_path}")
        return ico_path
        
    except Exception as e:
        print(f"ICO 파일 생성 오류: {e}")
        return None

def create_default_hometax_icon():
    """기본 HomeTax 아이콘 생성 (로고를 찾을 수 없는 경우)"""
    
    try:
        print("기본 HomeTax 아이콘 생성 중...")
        
        # 128x128 기본 이미지 생성
        img = Image.new('RGBA', (128, 128), (0, 0, 0, 0))
        
        # 간단한 세금계산서 아이콘 그리기 (여기서는 기본 사각형)
        from PIL import ImageDraw, ImageFont
        
        draw = ImageDraw.Draw(img)
        
        # 배경 사각형 (국세청 색상 느낌)
        draw.rectangle([10, 10, 118, 118], fill=(41, 128, 185), outline=(52, 73, 94), width=2)
        
        # 텍스트 추가
        try:
            # 시스템 폰트 사용
            font = ImageFont.truetype("malgun.ttf", 20)  # 맑은 고딕
        except:
            font = ImageFont.load_default()
        
        # "HT" 텍스트 (HomeTax)
        text = "HT"
        bbox = draw.textbbox((0, 0), text, font=font)
        text_width = bbox[2] - bbox[0]
        text_height = bbox[3] - bbox[1]
        
        x = (128 - text_width) // 2
        y = (128 - text_height) // 2
        
        draw.text((x, y), text, fill="white", font=font)
        
        # PNG로 저장
        img.save("hometax_logo_default.png")
        
        # ICO로 변환
        return create_ico_from_image("hometax_logo_default.png", "hometax_icon.ico")
        
    except Exception as e:
        print(f"기본 아이콘 생성 오류: {e}")
        return None

async def main():
    """메인 함수"""
    print("=== HomeTax 로고 추출 및 ICO 파일 생성 ===")
    
    # Pillow 라이브러리 확인
    try:
        from PIL import Image
        print("Pillow 라이브러리 확인됨")
    except ImportError:
        print("Pillow 라이브러리가 필요합니다.")
        print("설치 명령: pip install Pillow")
        return
    
    # 1. HomeTax 페이지에서 로고 추출
    logo_path = await extract_hometax_logo()
    
    # 2. ICO 파일 생성
    if logo_path and os.path.exists(logo_path):
        ico_path = create_ico_from_image(logo_path)
    else:
        print("로고 추출에 실패하여 기본 아이콘을 생성합니다.")
        ico_path = create_default_hometax_icon()
    
    if ico_path and os.path.exists(ico_path):
        print(f"\n최종 결과: {ico_path}")
        print("이 파일을 build_app.spec의 icon 파라미터에 사용하세요.")
    else:
        print("\n아이콘 생성에 실패했습니다.")

if __name__ == "__main__":
    asyncio.run(main())