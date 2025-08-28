# -*- coding: utf-8 -*-
"""
오프라인 라이센스 시스템
하드웨어 기반 라이센스 키 생성 및 검증
"""

import hashlib
import platform
import subprocess
import os
import base64
import json
from datetime import datetime, timedelta
import uuid

class OfflineLicenseManager:
    def __init__(self):
        self.secret_key = "HomeTax2025SecretKey!@#"  # 실제 배포시에는 더 복잡하게
        
    def get_hardware_id(self):
        """고유한 하드웨어 ID 생성"""
        try:
            # CPU 정보
            cpu_info = platform.processor()
            
            # 메인보드 시리얼 넘버 (Windows)
            try:
                motherboard = subprocess.check_output(
                    'wmic baseboard get serialnumber', 
                    shell=True, text=True
                ).split('\n')[1].strip()
            except:
                motherboard = "unknown"
            
            # 하드디스크 시리얼 넘버
            try:
                disk_serial = subprocess.check_output(
                    'wmic diskdrive get serialnumber', 
                    shell=True, text=True
                ).split('\n')[1].strip()
            except:
                disk_serial = "unknown"
            
            # MAC 주소
            try:
                mac_address = ':'.join(['{:02x}'.format((uuid.getnode() >> ele) & 0xff) 
                                      for ele in range(0,8*6,8)][::-1])
            except:
                mac_address = "unknown"
            
            # 조합하여 고유 ID 생성
            hardware_string = f"{cpu_info}_{motherboard}_{disk_serial}_{mac_address}"
            hardware_id = hashlib.sha256(hardware_string.encode()).hexdigest()[:16]
            
            return hardware_id.upper()
            
        except Exception as e:
            print(f"하드웨어 ID 생성 오류: {e}")
            return "UNKNOWN_HARDWARE"
    
    def generate_license_key(self, hardware_id, days_valid=365, user_info=""):
        """특정 하드웨어용 라이센스 키 생성 (개발자가 수동으로 생성)"""
        
        # 만료일 설정
        expire_date = datetime.now() + timedelta(days=days_valid)
        
        # 라이센스 정보
        license_data = {
            "hardware_id": hardware_id,
            "expire_date": expire_date.strftime("%Y-%m-%d"),
            "user_info": user_info,
            "product": "HomeTax_System",
            "version": "1.0.0"
        }
        
        # JSON 문자열로 변환
        license_json = json.dumps(license_data, sort_keys=True)
        
        # 서명 생성 (해시 + 비밀키)
        signature = hashlib.sha256((license_json + self.secret_key).encode()).hexdigest()[:8]
        
        # 최종 라이센스 키 (Base64 인코딩)
        license_full = f"{license_json}:{signature}"
        license_key = base64.b64encode(license_full.encode()).decode()
        
        return license_key
    
    def verify_license_key(self, license_key):
        """라이센스 키 검증"""
        try:
            # Base64 디코딩
            license_full = base64.b64decode(license_key.encode()).decode()
            
            # JSON과 서명 분리
            parts = license_full.rsplit(':', 1)
            if len(parts) != 2:
                return False, "잘못된 라이센스 키 형식"
            
            license_json, signature = parts
            
            # 서명 검증
            expected_signature = hashlib.sha256((license_json + self.secret_key).encode()).hexdigest()[:8]
            if signature != expected_signature:
                return False, "라이센스 키가 변조되었습니다"
            
            # JSON 파싱
            license_data = json.loads(license_json)
            
            # 하드웨어 ID 확인
            current_hardware_id = self.get_hardware_id()
            if license_data["hardware_id"] != current_hardware_id:
                return False, f"이 PC에서 사용할 수 없는 라이센스입니다\n현재 PC ID: {current_hardware_id}"
            
            # 만료일 확인
            expire_date = datetime.strptime(license_data["expire_date"], "%Y-%m-%d")
            if datetime.now() > expire_date:
                return False, f"라이센스가 만료되었습니다 (만료일: {license_data['expire_date']})"
            
            return True, f"유효한 라이센스입니다 (만료일: {license_data['expire_date']})"
            
        except Exception as e:
            return False, f"라이센스 검증 오류: {e}"
    
    def save_license_to_registry(self, license_key):
        """라이센스를 Windows 레지스트리에 저장"""
        try:
            import winreg
            
            # 레지스트리 키 열기/생성
            registry_key = winreg.CreateKey(winreg.HKEY_CURRENT_USER, "SOFTWARE\\HomeTax_System")
            
            # 라이센스 키 저장 (암호화)
            encrypted_license = base64.b64encode(license_key.encode()).decode()
            winreg.SetValueEx(registry_key, "License", 0, winreg.REG_SZ, encrypted_license)
            
            winreg.CloseKey(registry_key)
            return True
            
        except Exception as e:
            print(f"레지스트리 저장 오류: {e}")
            return False
    
    def load_license_from_registry(self):
        """Windows 레지스트리에서 라이센스 로드"""
        try:
            import winreg
            
            registry_key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, "SOFTWARE\\HomeTax_System")
            encrypted_license, _ = winreg.QueryValueEx(registry_key, "License")
            winreg.CloseKey(registry_key)
            
            # 복호화
            license_key = base64.b64decode(encrypted_license.encode()).decode()
            return license_key
            
        except Exception as e:
            return None

def generate_key_for_user():
    """사용자용 키 생성 도구 (개발자가 사용)"""
    print("=== HomeTax 시스템 라이센스 키 생성기 ===")
    print("이 도구는 개발자만 사용합니다.\n")
    
    license_manager = OfflineLicenseManager()
    
    print("1. 사용자의 하드웨어 ID를 입력하세요:")
    hardware_id = input("하드웨어 ID: ").strip().upper()
    
    print("\n2. 사용자 정보를 입력하세요 (선택사항):")
    user_info = input("사용자명/회사명: ").strip()
    
    print("\n3. 라이센스 유효 기간을 선택하세요:")
    print("1) 30일")
    print("2) 1년 (365일)")
    print("3) 영구 (99년)")
    print("4) 사용자 정의")
    
    choice = input("\n선택 (1-4): ").strip()
    
    if choice == "1":
        days = 30
    elif choice == "2":
        days = 365
    elif choice == "3":
        days = 365 * 99
    elif choice == "4":
        days = int(input("유효 기간 (일): "))
    else:
        days = 365
    
    # 라이센스 키 생성
    license_key = license_manager.generate_license_key(hardware_id, days, user_info)
    
    print(f"\n=== 생성된 라이센스 키 ===")
    print(f"하드웨어 ID: {hardware_id}")
    print(f"사용자 정보: {user_info}")
    print(f"유효 기간: {days}일")
    print(f"\n라이센스 키:")
    print(f"{license_key}")
    print(f"\n이 키를 사용자에게 제공하세요.")

def check_current_license():
    """현재 PC의 라이센스 상태 확인"""
    print("=== 현재 PC 라이센스 상태 확인 ===")
    
    license_manager = OfflineLicenseManager()
    
    # 현재 하드웨어 ID 표시
    hardware_id = license_manager.get_hardware_id()
    print(f"현재 PC 하드웨어 ID: {hardware_id}")
    
    # 저장된 라이센스 확인
    stored_license = license_manager.load_license_from_registry()
    if stored_license:
        print("\n저장된 라이센스를 찾았습니다.")
        is_valid, message = license_manager.verify_license_key(stored_license)
        print(f"검증 결과: {message}")
    else:
        print("\n저장된 라이센스가 없습니다.")
        
        # 새 라이센스 키 입력 받기
        license_key = input("\n라이센스 키를 입력하세요: ").strip()
        if license_key:
            is_valid, message = license_manager.verify_license_key(license_key)
            print(f"검증 결과: {message}")
            
            if is_valid:
                if license_manager.save_license_to_registry(license_key):
                    print("라이센스가 저장되었습니다.")
                else:
                    print("라이센스 저장에 실패했습니다.")

if __name__ == "__main__":
    print("HomeTax 시스템 라이센스 관리")
    print("1. 라이센스 키 생성 (개발자용)")
    print("2. 현재 PC 라이센스 확인")
    
    choice = input("\n선택 (1-2): ").strip()
    
    if choice == "1":
        generate_key_for_user()
    elif choice == "2":
        check_current_license()
    else:
        print("잘못된 선택입니다.")