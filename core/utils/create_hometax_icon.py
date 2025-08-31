# 📁 C:\APP\tax-bill\core\utils\create_hometax_icon.py
# Create at 2508312118 Ver1.00
# -*- coding: utf-8 -*-
"""
HomeTax 아이콘 생성 (간단한 버전)
"""

import requests
from PIL import Image, ImageDraw, ImageFont
import io

def download_korea_logo():
    """한국 정부 로고 다운로드"""
    logo_url = "https://hometax.go.kr/css/comm/bpr_portal_images/logo_korea.png?postfix=2025_08_26"
    
    try:
        print(f"로고 다운로드: {logo_url}")
        response = requests.get(logo_url, timeout=10)
        
        if response.status_code == 200:
            with open("hometax_logo.png", "wb") as f:
                f.write(response.content)
            print("로고 다운로드 완료: hometax_logo.png")
            return "hometax_logo.png"
        else:
            print(f"다운로드 실패: HTTP {response.status_code}")
            return None
    except Exception as e:
        print(f"다운로드 오류: {e}")
        return None

def create_hometax_icon():
    """HomeTax 전용 아이콘 생성"""
    
    try:
        print("HomeTax 아이콘 생성 중...")
        
        # 128x128 이미지 생성
        img = Image.new('RGBA', (128, 128), (0, 0, 0, 0))
        draw = ImageDraw.Draw(img)
        
        # 배경 원 (국세청 색상)
        draw.ellipse([8, 8, 120, 120], fill=(41, 128, 185), outline=(31, 97, 141), width=3)
        
        # 내부 사각형 (문서 느낌)
        draw.rectangle([25, 20, 103, 100], fill="white", outline=(52, 73, 94), width=2)
        
        # 문서 라인들
        for i in range(3):
            y = 35 + i * 15
            draw.line([35, y, 93, y], fill=(52, 73, 94), width=2)
        
        # 텍스트 추가
        try:
            font = ImageFont.truetype("arial.ttf", 16)
        except:
            try:
                font = ImageFont.truetype("malgun.ttf", 16)
            except:
                font = ImageFont.load_default()
        
        # "세금" 텍스트
        text = "세금"
        bbox = draw.textbbox((0, 0), text, font=font)
        text_width = bbox[2] - bbox[0]
        x = (128 - text_width) // 2
        draw.text((x, 105), text, fill="white", font=font, stroke_width=1, stroke_fill=(31, 97, 141))
        
        # PNG로 저장
        img.save("hometax_logo_custom.png")
        print("커스텀 로고 생성 완료: hometax_logo_custom.png")
        
        return "hometax_logo_custom.png"
        
    except Exception as e:
        print(f"아이콘 생성 오류: {e}")
        return None

def create_ico_file(png_path, ico_path="hometax_icon.ico"):
    """PNG를 ICO로 변환"""
    
    try:
        print(f"ICO 변환: {png_path} -> {ico_path}")
        
        with Image.open(png_path) as img:
            # RGBA 모드로 변환
            if img.mode != 'RGBA':
                img = img.convert('RGBA')
            
            # 여러 크기 생성
            sizes = [(16, 16), (32, 32), (48, 48), (64, 64), (128, 128)]
            icons = []
            
            for size in sizes:
                resized = img.resize(size, Image.Resampling.LANCZOS)
                icons.append(resized)
            
            # ICO 저장
            icons[0].save(
                ico_path,
                format='ICO',
                sizes=[(icon.width, icon.height) for icon in icons]
            )
            
        print(f"ICO 파일 생성 완료: {ico_path}")
        return ico_path
        
    except Exception as e:
        print(f"ICO 변환 오류: {e}")
        return None

def main():
    """메인 함수"""
    print("=== HomeTax 아이콘 생성 ===")
    
    # 1. 먼저 정부 로고 다운로드 시도
    logo_path = download_korea_logo()
    
    # 2. 다운로드에 실패하면 커스텀 아이콘 생성
    if not logo_path:
        print("로고 다운로드 실패, 커스텀 아이콘 생성")
        logo_path = create_hometax_icon()
    
    # 3. ICO 파일 생성
    if logo_path:
        ico_path = create_ico_file(logo_path)
        
        if ico_path:
            print(f"\n성공! 아이콘 파일: {ico_path}")
            print("build_app.spec에서 이 파일을 사용하세요.")
        else:
            print("\nICO 파일 생성 실패")
    else:
        print("\n아이콘 생성 실패")

if __name__ == "__main__":
    main()