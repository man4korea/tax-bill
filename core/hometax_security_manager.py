# 📁 C:\APP\tax-bill\core\hometax_security_manager.py
# Create at 2508312118 Ver1.00
# -*- coding: utf-8 -*-
"""
HomeTax 보안 관리 시스템
비밀번호 암호화/복호화 및 보안 저장 관리
"""

import base64
import json
from pathlib import Path
from cryptography.fernet import Fernet
from cryptography.hazmat.primitives import hashes
from cryptography.hazmat.primitives.kdf.pbkdf2 import PBKDF2HMAC


class HomeTaxSecurityManager:
    def __init__(self):
        self.salt = b'hometax_salt_2024_secure'  # 보안 강화된 salt
        self.iterations = 100000  # PBKDF2 반복 횟수
        
    def generate_key_from_password(self, master_password="hometax_default_key_2024"):
        """마스터 비밀번호로부터 암호화 키 생성"""
        kdf = PBKDF2HMAC(
            algorithm=hashes.SHA256(),
            length=32,
            salt=self.salt,
            iterations=self.iterations,
        )
        key = base64.urlsafe_b64encode(kdf.derive(master_password.encode()))
        return key
    
    def encrypt_password(self, password, master_password="hometax_default_key_2024"):
        """비밀번호 암호화"""
        try:
            key = self.generate_key_from_password(master_password)
            f = Fernet(key)
            encrypted_password = f.encrypt(password.encode())
            # Base64로 인코딩하여 문자열로 반환
            return base64.urlsafe_b64encode(encrypted_password).decode('utf-8')
        except Exception as e:
            print(f"비밀번호 암호화 오류: {e}")
            return None
    
    def decrypt_password(self, encrypted_password, master_password="hometax_default_key_2024"):
        """비밀번호 복호화"""
        try:
            # Base64 디코딩
            encrypted_data = base64.urlsafe_b64decode(encrypted_password.encode('utf-8'))
            
            key = self.generate_key_from_password(master_password)
            f = Fernet(key)
            decrypted_password = f.decrypt(encrypted_data)
            return decrypted_password.decode('utf-8')
        except Exception as e:
            print(f"비밀번호 복호화 오류: {e}")
            return None
    
    def save_encrypted_password_to_env(self, password):
        """암호화된 비밀번호를 .env 파일에 저장"""
        try:
            # 비밀번호 암호화
            encrypted_pw = self.encrypt_password(password)
            if not encrypted_pw:
                return False
            
            env_file = Path(".env")
            env_content = ""
            if env_file.exists():
                with open(env_file, 'r', encoding='utf-8') as f:
                    env_content = f.read()
            
            # 기존 라인들을 분석하여 업데이트 (암호화된 비밀번호만 관리)
            lines = env_content.split('\n') if env_content else []
            pw_encrypted_found = False
            
            new_lines = []
            for line in lines:
                if line.startswith('PW='):
                    # 평문 PW= 라인 완전 제거 (보안 강화)
                    continue
                elif line.startswith('PW_ENCRYPTED='):
                    # 암호화된 비밀번호 라인 업데이트
                    new_lines.append(f'PW_ENCRYPTED={encrypted_pw}')
                    pw_encrypted_found = True
                else:
                    new_lines.append(line)
            
            # 암호화된 비밀번호 라인이 없으면 추가
            if not pw_encrypted_found:
                new_lines.append(f'PW_ENCRYPTED={encrypted_pw}')
            
            # 파일 저장
            with open(env_file, 'w', encoding='utf-8') as f:
                f.write('\n'.join(new_lines))
            
            print("[OK] 암호화된 비밀번호가 .env 파일에 저장되었습니다.")
            return True
            
        except Exception as e:
            print(f"[ERROR] .env 파일 저장 오류: {e}")
            return False
    
    def load_password_from_env(self):
        """환경 변수 파일에서 비밀번호 로드 (암호화된 것 우선)"""
        try:
            env_file = Path(".env")
            if not env_file.exists():
                return None
            
            with open(env_file, 'r', encoding='utf-8') as f:
                lines = f.readlines()
            
            encrypted_password = None
            
            for line in lines:
                line = line.strip()
                if line.startswith('PW_ENCRYPTED='):
                    encrypted_password = line.split('PW_ENCRYPTED=', 1)[1]
                    break  # 암호화된 비밀번호만 찾으면 루프 종료
            
            # 암호화된 비밀번호만 사용 (보안 강화)
            if encrypted_password:
                password = self.decrypt_password(encrypted_password)
                if password:
                    print("[OK] 암호화된 비밀번호 로드 성공")
                    return password
                else:
                    print("[ERROR] 암호화된 비밀번호 복호화 실패")
                    return None
            
            return None
            
        except Exception as e:
            print(f"[ERROR] 환경 변수 파일 로드 오류: {e}")
            return None
    
    def migrate_plaintext_to_encrypted(self):
        """더 이상 사용하지 않음 - 레거시 함수 (평문 지원 중단)"""
        print("[INFO] 평문 비밀번호 지원이 중단되었습니다.")
        print("[INFO] 암호화된 비밀번호만 사용 가능합니다.")
        return False
    
    def validate_password_security(self):
        """비밀번호 저장 보안 상태 검증 (암호화된 비밀번호만 지원)"""
        try:
            env_file = Path(".env")
            if not env_file.exists():
                return {"status": "no_env", "secure": False, "message": ".env 파일이 존재하지 않습니다"}
            
            with open(env_file, 'r', encoding='utf-8') as f:
                content = f.read()
            
            has_encrypted = 'PW_ENCRYPTED=' in content
            
            if has_encrypted:
                return {"status": "secure", "secure": True, "message": "비밀번호가 안전하게 암호화되어 저장되었습니다"}
            else:
                return {"status": "no_password", "secure": False, "message": "암호화된 비밀번호가 저장되지 않았습니다"}
                
        except Exception as e:
            return {"status": "error", "secure": False, "message": f"검증 중 오류: {e}"}


def main():
    """테스트 및 마이그레이션 실행"""
    print("🔐 HomeTax 보안 관리 시스템")
    print("=" * 50)
    
    security_manager = HomeTaxSecurityManager()
    
    # 현재 보안 상태 검증
    status = security_manager.validate_password_security()
    print(f"현재 상태: {status['message']}")
    
    if not status['secure']:
        if status['status'] in ['plaintext', 'mixed']:
            print("\n🔄 자동 마이그레이션을 시작합니다...")
            security_manager.migrate_plaintext_to_encrypted()
        else:
            print("ℹ️ 마이그레이션이 필요하지 않습니다.")
    
    # 최종 상태 재검증
    final_status = security_manager.validate_password_security()
    print(f"\n최종 상태: {final_status['message']}")
    
    # 비밀번호 로드 테스트
    print("\n🧪 비밀번호 로드 테스트:")
    password = security_manager.load_password_from_env()
    if password:
        print("✅ 비밀번호 로드 성공")
    else:
        print("❌ 비밀번호 로드 실패")


if __name__ == "__main__":
    main()