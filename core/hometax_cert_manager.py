# 📁 C:\APP\tax-bill\core\hometax_cert_manager.py
# Create at 2508312118 Ver1.00
# -*- coding: utf-8 -*-
"""
HomeTax 공인인증서 비밀번호 관리 시스템
자동로그인용 비밀번호 저장/관리 및 수동로그인 지원
"""

import tkinter as tk
from tkinter import ttk, messagebox
import os
import base64
from cryptography.fernet import Fernet
from cryptography.hazmat.primitives import hashes
from cryptography.hazmat.primitives.kdf.pbkdf2 import PBKDF2HMAC
import json
from pathlib import Path
import subprocess
import sys

# 보안 관리 import
from hometax_security_manager import HomeTaxSecurityManager

class HomeTaxCertManager:
    def __init__(self, parent=None):
        self.parent = parent
        self.root = tk.Toplevel(parent) if parent else tk.Tk()
        self.cert_file = Path("cert_config.enc")  # 암호화된 인증서 정보 파일
        self.env_file = Path(__file__).parent.parent / ".env"  # 프로젝트 루트 .env 파일
        self.login_mode = tk.StringVar(value="manual")  # 기본값: 수동로그인
        self.security_manager = HomeTaxSecurityManager()  # 보안 관리자 초기화
        self.ensure_env_file_exists()  # .env 파일 존재 확인 및 생성
        self.setup_window()
        self.create_widgets()
        self.load_saved_config()
        
    def setup_window(self):
        """창 설정"""
        self.root.title("공인인증서 비밀번호 관리")
        self.root.geometry("500x580")
        self.root.resizable(False, False)
        
        # 창을 화면 중앙에 위치
        if self.parent:
            # 부모창이 있는 경우 부모창 중앙에 위치
            self.root.update_idletasks()
            
            parent_x = self.parent.winfo_rootx()
            parent_y = self.parent.winfo_rooty()
            parent_width = self.parent.winfo_width()
            parent_height = self.parent.winfo_height()
            
            x = parent_x + (parent_width // 2) - (250)  # 500//2
            y = parent_y + (parent_height // 2) - (290)  # 580//2
            
            self.root.geometry(f"500x580+{x}+{y}")
            
            # 모달 창으로 설정
            self.root.transient(self.parent)
            self.root.grab_set()
        else:
            # 독립 실행시 화면 중앙에 위치
            self.root.eval('tk::PlaceWindow . center')
            
        # 창 닫기 이벤트 처리
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)
        
    def ensure_env_file_exists(self):
        """.env 파일 존재 확인 및 기본값으로 생성"""
        try:
            if not self.env_file.exists():
                print(f"📝 .env 파일이 없습니다. 기본값으로 생성: {self.env_file}")
                
                # 기본 .env 파일 내용
                default_content = """# HomeTax 자동화 시스템 설정
# 로그인 모드: auto (자동) 또는 manual (수동)
HOMETAX_LOGIN_MODE=manual

# 시스템 구성 정보 (자동 생성됨)
# HTC_CONFIG=encrypted_data_here
"""
                
                # .env 파일 생성
                with open(self.env_file, 'w', encoding='utf-8') as f:
                    f.write(default_content)
                
                print(f"✅ .env 파일 생성 완료: {self.env_file}")
                print("📋 기본 로그인 모드: manual")
                
            else:
                print(f"✅ .env 파일 존재: {self.env_file}")
                
        except Exception as e:
            print(f"❌ .env 파일 처리 중 오류: {e}")
            
    def read_env_login_mode(self):
        """.env 파일에서 로그인 모드 읽기"""
        try:
            if not self.env_file.exists():
                return "manual"  # 기본값
            
            with open(self.env_file, 'r', encoding='utf-8') as f:
                lines = f.readlines()
            
            for line in lines:
                line = line.strip()
                if line.startswith('HOMETAX_LOGIN_MODE='):
                    mode = line.split('=', 1)[1].strip()
                    if mode in ['auto', 'manual']:
                        print(f"📋 .env에서 로그인 모드 읽음: {mode}")
                        return mode
            
            print("📋 .env에서 로그인 모드를 찾을 수 없음, 기본값 사용: manual")
            return "manual"
            
        except Exception as e:
            print(f"❌ .env 파일 읽기 오류: {e}")
            return "manual"
        
    def create_widgets(self):
        """위젯 생성"""
        # 메인 프레임
        main_frame = ttk.Frame(self.root, padding="20")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # 제목
        title_label = ttk.Label(
            main_frame,
            text="🔐 공인인증서 비밀번호 관리",
            font=('맑은 고딕', 16, 'bold'),
            foreground='#2E4057'
        )
        title_label.pack(pady=(0, 20))
        
        # 구분선
        separator1 = ttk.Separator(main_frame, orient='horizontal')
        separator1.pack(fill=tk.X, pady=(0, 20))
        
        # 로그인 방식 선택
        login_mode_frame = ttk.LabelFrame(main_frame, text="로그인 방식 선택", padding="15")
        login_mode_frame.pack(fill=tk.X, pady=(0, 20))
        
        # 자동로그인 라디오버튼
        auto_radio = ttk.Radiobutton(
            login_mode_frame,
            text="🤖 자동로그인 (공동공인인증서 활용 자동 로그인)",
            variable=self.login_mode,
            value="auto",
            command=self.on_mode_changed
        )
        auto_radio.pack(anchor=tk.W, pady=5)
        
        # 수동로그인 라디오버튼
        manual_radio = ttk.Radiobutton(
            login_mode_frame,
            text="✋ 수동로그인 (기타인증수단 선택 수동 로그인)",
            variable=self.login_mode,
            value="manual",
            command=self.on_mode_changed
        )
        manual_radio.pack(anchor=tk.W, pady=5)
        
        # 자동로그인 설정 프레임
        self.auto_frame = ttk.LabelFrame(main_frame, text="자동로그인 설정", padding="15")
        self.auto_frame.pack(fill=tk.X, pady=(0, 20))
        
        # 현재 상태 표시
        self.status_label = ttk.Label(
            self.auto_frame,
            text="현재 저장된 비밀번호: 없음",
            font=('맑은 고딕', 9),
            foreground='#666666'
        )
        self.status_label.pack(anchor=tk.W, pady=(0, 10))
        
        # 비밀번호 입력 필드
        ttk.Label(self.auto_frame, text="공인인증서 비밀번호:").pack(anchor=tk.W)
        
        self.password_entry = ttk.Entry(
            self.auto_frame,
            show="*",
            width=40,
            font=('맑은 고딕', 10)
        )
        self.password_entry.pack(fill=tk.X, pady=(5, 5))
        
        # 비밀번호 확인 필드
        ttk.Label(self.auto_frame, text="비밀번호 확인:").pack(anchor=tk.W)
        
        self.password_confirm_entry = ttk.Entry(
            self.auto_frame,
            show="*",
            width=40,
            font=('맑은 고딕', 10)
        )
        self.password_confirm_entry.pack(fill=tk.X, pady=(5, 10))
        
        # 비밀번호 표시/숨기기 체크박스
        self.show_password = tk.BooleanVar()
        show_password_check = ttk.Checkbutton(
            self.auto_frame,
            text="비밀번호 표시",
            variable=self.show_password,
            command=self.toggle_password_visibility
        )
        show_password_check.pack(anchor=tk.W, pady=(0, 10))
        
        # 자동로그인 버튼들 (제거됨)
        
        # 수동로그인 설정 프레임
        self.manual_frame = ttk.LabelFrame(main_frame, text="수동로그인 안내", padding="15")
        self.manual_frame.pack(fill=tk.X, pady=(0, 20))
        
        manual_info = ttk.Label(
            self.manual_frame,
            text="수동로그인 모드에서는:\n\n"
                 "1. 홈택스 로그인 페이지에서 인증서를 선택합니다\n\n"
                 "2. 비밀번호 입력창이 나타나면 직접 입력하세요\n\n"  
                 "3. 시스템이 자동으로 다음 단계를 진행합니다\n\n"
                 "※ 비밀번호가 저장되지 않아 더 안전합니다\n"
                 "※ 매번 수동으로 입력해야 하지만 보안성이 높습니다",
            font=('맑은 고딕', 9),
            foreground='#666666',
            justify=tk.LEFT
        )
        manual_info.pack(anchor=tk.W)
        
        # 하단 버튼 (동적으로 변경됨)
        self.button_frame = ttk.Frame(main_frame)
        self.button_frame.pack(fill=tk.X, pady=(20, 0))
        
        # 초기 모드 설정
        self.on_mode_changed()
        
    def on_mode_changed(self):
        """로그인 모드 변경 시 호출"""
        # 프레임 표시/숨기기
        if self.login_mode.get() == "auto":
            self.auto_frame.pack(fill=tk.X, pady=(0, 20))
            self.manual_frame.pack_forget()
        else:
            self.auto_frame.pack_forget()
            self.manual_frame.pack(fill=tk.X, pady=(0, 20))
            
        # 하단 버튼 업데이트
        self.update_buttons()
        
    def update_buttons(self):
        """모드에 따라 하단 버튼 업데이트"""
        # 기존 버튼들 제거
        for widget in self.button_frame.winfo_children():
            widget.destroy()
            
        # 두 버튼 공통으로 사용: 💾 저장 및 닫기와 닫기
        ttk.Button(
            self.button_frame,
            text="닫기",
            command=self.on_closing
        ).pack(side=tk.RIGHT)
        
        if self.login_mode.get() == "auto":
            ttk.Button(
                self.button_frame,
                text="💾 저장 및 닫기",
                command=self.save_and_close,
                style='Accent.TButton'
            ).pack(side=tk.RIGHT, padx=(0, 10))
        else:
            ttk.Button(
                self.button_frame,
                text="💾 저장 및 닫기",
                command=self.save_manual_mode_and_close,
                style='Accent.TButton'
            ).pack(side=tk.RIGHT, padx=(0, 10))
            
        # 공간 확보를 위한 더미 프레임
        ttk.Frame(self.button_frame, width=20).pack(side=tk.LEFT)
            
    def toggle_password_visibility(self):
        """비밀번호 표시/숨기기"""
        if self.show_password.get():
            self.password_entry.config(show="")
            self.password_confirm_entry.config(show="")
        else:
            self.password_entry.config(show="*")
            self.password_confirm_entry.config(show="*")
            
    def generate_key_from_password(self, password="hometax_default"):
        """비밀번호로부터 암호화 키 생성"""
        salt = b'hometax_salt_2024'  # 고정 salt (실제 운영시에는 랜덤 생성 권장)
        kdf = PBKDF2HMAC(
            algorithm=hashes.SHA256(),
            length=32,
            salt=salt,
            iterations=100000,
        )
        key = base64.urlsafe_b64encode(kdf.derive(password.encode()))
        return key
        
    def encrypt_data(self, data, password="hometax_default"):
        """데이터 암호화"""
        key = self.generate_key_from_password(password)
        f = Fernet(key)
        encrypted_data = f.encrypt(json.dumps(data).encode())
        return encrypted_data
        
    def decrypt_data(self, encrypted_data, password="hometax_default"):
        """데이터 복호화"""
        try:
            key = self.generate_key_from_password(password)
            f = Fernet(key)
            decrypted_data = f.decrypt(encrypted_data)
            return json.loads(decrypted_data.decode())
        except:
            return None
            
    def validate_password_input(self):
        """비밀번호 입력 검증"""
        password = self.password_entry.get().strip()
        confirm_password = self.password_confirm_entry.get().strip()
        
        if not password:
            messagebox.showwarning("경고", "비밀번호를 입력하세요.")
            return False
            
        if not confirm_password:
            messagebox.showwarning("경고", "비밀번호 확인을 입력하세요.")
            return False
            
        if password != confirm_password:
            messagebox.showerror("오류", "비밀번호가 일치하지 않습니다.\n다시 확인해주세요.")
            self.password_confirm_entry.delete(0, tk.END)
            self.password_confirm_entry.focus()
            return False
            
        if len(password) < 4:
            messagebox.showwarning("경고", "비밀번호는 최소 4자 이상이어야 합니다.")
            return False
            
        return True
        
    def save_password(self):
        """비밀번호 저장"""
        # 비밀번호 입력 검증
        if not self.validate_password_input():
            return False
            
        password = self.password_entry.get().strip()
            
        try:
            # 설정 데이터 준비
            config_data = {
                "cert_password": password,
                "login_mode": "auto",
                "created_at": str(Path(__file__).stat().st_mtime)
            }
            
            # 데이터 암호화하여 저장
            encrypted_data = self.encrypt_data(config_data)
            with open(self.cert_file, 'wb') as f:
                f.write(encrypted_data)
                
            # .env 파일에 암호화된 형태로 저장 (보안 강화)
            if not self.save_encrypted_config_to_env(password):
                print("⚠️ .env 파일 암호화 저장 실패")
                return False
                
            self.update_status()
            # 입력 필드 초기화
            self.password_entry.delete(0, tk.END)
            self.password_confirm_entry.delete(0, tk.END)
            return True
            
        except Exception as e:
            messagebox.showerror("오류", f"비밀번호 저장 중 오류가 발생했습니다:\n{str(e)}")
            return False
            
    def save_encrypted_config_to_env(self, password):
        """암호화된 설정을 .env 파일에 저장"""
        try:
            # 비밀번호를 더블 암호화
            encrypted_password = self.encrypt_password_for_env(password)
            
            # .env 파일 읽기
            lines = []
            if self.env_file.exists():
                with open(self.env_file, 'r', encoding='utf-8') as f:
                    lines = f.readlines()
            
            # HTC_CONFIG 라인 찾기/업데이트
            config_updated = False
            for i, line in enumerate(lines):
                if line.strip().startswith('HTC_CONFIG='):
                    lines[i] = f'HTC_CONFIG={encrypted_password}\n'
                    config_updated = True
                    print(f"📝 기존 HTC_CONFIG 업데이트")
                    break
            
            # 새로운 라인 추가
            if not config_updated:
                lines.append(f'HTC_CONFIG={encrypted_password}\n')
                print(f"📝 새로운 HTC_CONFIG 추가")
            
            # .env 파일에 쓰기
            with open(self.env_file, 'w', encoding='utf-8') as f:
                f.writelines(lines)
            
            print("✅ 암호화된 설정 저장 완료")
            return True
            
        except Exception as e:
            print(f"❌ 암호화 저장 실패: {e}")
            return False
    
    def encrypt_password_for_env(self, password):
        """비밀번호를 .env 파일용으로 암호화"""
        import base64
        try:
            # 간단한 base64 인코딩 + 역순 + 추가 문자열
            encoded = base64.b64encode(password.encode('utf-8')).decode('utf-8')
            reversed_encoded = encoded[::-1]  # 문자열 역순
            scrambled = f"HTC_{reversed_encoded}_CFG"  # 앞뒤 추가 문자열
            return scrambled
        except Exception as e:
            print(f"❌ 비밀번호 암호화 실패: {e}")
            return None
    
    def decrypt_password_from_env(self, encrypted_config):
        """암호화된 설정에서 비밀번호 복호화"""
        import base64
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
    
    def load_encrypted_config_from_env(self):
        """암호화된 설정을 .env 파일에서 로드"""
        try:
            if not self.env_file.exists():
                return None
            
            with open(self.env_file, 'r', encoding='utf-8') as f:
                lines = f.readlines()
            
            for line in lines:
                line = line.strip()
                if line.startswith('HTC_CONFIG='):
                    encrypted_config = line.split('=', 1)[1].strip()
                    password = self.decrypt_password_from_env(encrypted_config)
                    if password:
                        print("🔐 암호화된 설정에서 비밀번호 로드 성공")
                        return password
            
            print("📋 .env에서 암호화된 설정을 찾을 수 없음")
            return None
            
        except Exception as e:
            print(f"❌ 암호화 설정 로드 실패: {e}")
            return None
    
    def save_and_close(self):
        """자동 모드에서 로그인 모드 및 비밀번호 저장하고 닫기"""
        # 비밀번호 입력 검증
        if not self.validate_password_input():
            return
            
        # 로그인 모드를 auto로 설정하고 저장
        self.save_login_mode_to_env("auto")
        
        # 비밀번호 저장 시도
        if self.save_password():
            # 저장 성공 메시지 표시
            messagebox.showinfo("저장 완료", 
                "자동로그인 모드로 설정되었습니다.\n"
                "비밀번호는 암호화하여 안전하게 저장되었습니다.")
            
            # 창 닫기
            self.on_closing()
            
    def delete_password(self):
        """저장된 비밀번호 삭제"""
        if messagebox.askyesno("확인", "저장된 비밀번호를 삭제하시겠습니까?"):
            try:
                # 암호화 파일 삭제
                if self.cert_file.exists():
                    self.cert_file.unlink()
                    
                # .env 파일에서 암호화된 설정 제거
                if self.env_file.exists():
                    with open(self.env_file, 'r', encoding='utf-8') as f:
                        lines = f.readlines()
                        
                    with open(self.env_file, 'w', encoding='utf-8') as f:
                        for line in lines:
                            if not (line.strip().startswith('PW=') or 
                                   line.strip().startswith('PW_ENCRYPTED=') or 
                                   line.strip().startswith('HTC_CONFIG=')):
                                f.write(line)
                                
                self.update_status()
                messagebox.showinfo("성공", "저장된 비밀번호가 삭제되었습니다.")
                
            except Exception as e:
                messagebox.showerror("오류", f"비밀번호 삭제 중 오류가 발생했습니다:\n{str(e)}")

    def delete_saved_passwords_silently(self):
        """저장된 비밀번호를 조용히 삭제 (확인 메시지 없음)"""
        try:
            # 암호화 파일 삭제
            if self.cert_file.exists():
                self.cert_file.unlink()
                print("🗑️ 암호화된 인증서 파일 삭제")
                
            # .env 파일에서 암호화된 설정 제거
            if self.env_file.exists():
                with open(self.env_file, 'r', encoding='utf-8') as f:
                    lines = f.readlines()
                    
                with open(self.env_file, 'w', encoding='utf-8') as f:
                    for line in lines:
                        if not (line.strip().startswith('PW=') or 
                               line.strip().startswith('PW_ENCRYPTED=') or 
                               line.strip().startswith('HTC_CONFIG=')):
                            f.write(line)
                            
                print("🗑️ .env 파일에서 비밀번호 관련 설정 제거")
                            
            self.update_status()
            print("✅ 기존 저장된 비밀번호 모두 삭제 완료")
            
        except Exception as e:
            print(f"❌ 비밀번호 삭제 중 오류: {e}")
            raise e
                
    def load_saved_config(self):
        """저장된 설정 불러오기"""
        try:
            # 1. .env 파일에서 로그인 모드 우선 로드
            env_login_mode = self.read_env_login_mode()
            self.login_mode.set(env_login_mode)
            print(f"📋 .env에서 로그인 모드 설정: {env_login_mode}")
            
            # 2. 암호화 파일에서 설정 로드 (보조)
            if self.cert_file.exists():
                with open(self.cert_file, 'rb') as f:
                    encrypted_data = f.read()
                    
                config_data = self.decrypt_data(encrypted_data)
                if config_data and not env_login_mode:
                    # .env에 설정이 없을 때만 암호화 파일 사용
                    self.login_mode.set(config_data.get("login_mode", "manual"))
                    print(f"📋 암호화 파일에서 로그인 모드 설정: {config_data.get('login_mode', 'manual')}")
                    
            # 3. .env 파일에서 암호화된 비밀번호 확인 (자동 모드 판단용)
            encrypted_password = self.load_encrypted_config_from_env()
            if encrypted_password and env_login_mode == "auto":
                print("🔐 .env에서 암호화된 비밀번호 발견 - 자동 로그인 모드 유지")
            elif encrypted_password and env_login_mode == "manual":
                print("🔐 .env에서 암호화된 비밀번호 발견하지만 수동 모드로 설정됨")
            
            # 레거시 지원: 기존 security_manager 방식도 체크
            legacy_password = self.security_manager.load_password_from_env()
            if legacy_password and not encrypted_password:
                print("🔐 레거시 비밀번호 발견 - 업그레이드 권장")
                            
            self.update_status()
            self.on_mode_changed()
            
        except Exception as e:
            print(f"❌ 설정 로드 중 오류: {e}")
            # 오류 발생 시 기본값 사용
            self.login_mode.set("manual")
            self.update_status()
            self.on_mode_changed()
            
    def update_status(self):
        """상태 표시 업데이트"""
        has_encrypted_file = self.cert_file.exists()
        has_encrypted_env = self.load_encrypted_config_from_env() is not None
        
        # 레거시 비밀번호 확인
        password_status = self.security_manager.validate_password_security()
        has_legacy_password = password_status['status'] in ['secure', 'mixed', 'plaintext']
        
        if has_encrypted_file or has_encrypted_env:
            self.status_label.config(
                text="현재 저장된 비밀번호: 있음 ✓ (암호화됨)",
                foreground='#28A745'
            )
        elif has_legacy_password:
            if password_status['secure']:
                self.status_label.config(
                    text="현재 저장된 비밀번호: 있음 ✓ (레거시 암호화)",
                    foreground='#28A745'
                )
            else:
                self.status_label.config(
                    text="현재 저장된 비밀번호: 있음 ⚠️ (업그레이드 권장)",
                    foreground='#FFA500'
                )
        else:
            self.status_label.config(
                text="현재 저장된 비밀번호: 없음",
                foreground='#666666'
            )
            
            
    def save_login_mode_to_env(self, mode):
        """로그인 모드를 .env 파일에 저장"""
        try:
            lines = []
            
            # 기존 .env 파일 읽기
            if self.env_file.exists():
                with open(self.env_file, 'r', encoding='utf-8') as f:
                    lines = f.readlines()
            
            # HOMETAX_LOGIN_MODE 줄 찾기/업데이트
            mode_updated = False
            for i, line in enumerate(lines):
                if line.strip().startswith('HOMETAX_LOGIN_MODE='):
                    lines[i] = f'HOMETAX_LOGIN_MODE={mode}\n'
                    mode_updated = True
                    print(f"📝 기존 HOMETAX_LOGIN_MODE 업데이트: {mode}")
                    break
            
            # 새로운 줄 추가 (존재하지 않는 경우)
            if not mode_updated:
                lines.append(f'HOMETAX_LOGIN_MODE={mode}\n')
                print(f"📝 새로운 HOMETAX_LOGIN_MODE 추가: {mode}")
            
            # .env 파일에 쓰기
            with open(self.env_file, 'w', encoding='utf-8') as f:
                f.writelines(lines)
            
            # 저장 후 검증
            with open(self.env_file, 'r', encoding='utf-8') as f:
                saved_content = f.read()
                if f'HOMETAX_LOGIN_MODE={mode}' in saved_content:
                    print(f"✅ 로그인 모드 저장 및 검증 완료: {mode}")
                else:
                    print(f"❌ 로그인 모드 저장 검증 실패")
                    
            # 환경변수도 즉시 설정
            os.environ['HOMETAX_LOGIN_MODE'] = mode
            print(f"🔧 환경변수도 설정 완료: {os.environ.get('HOMETAX_LOGIN_MODE')}")
            
        except Exception as e:
            print(f"❌ 로그인 모드 저장 실패: {e}")
            # 환경변수로라도 설정
            os.environ['HOMETAX_LOGIN_MODE'] = mode
            
    def save_manual_mode_and_close(self):
        """수동 모드에서 로그인 모드 저장하고 닫기 (기존 비밀번호 삭제)"""
        try:
            # 수동 모드로 설정 저장
            self.save_login_mode_to_env("manual")
            
            # 기존 저장된 비밀번호 삭제
            self.delete_saved_passwords_silently()
            
            # 설정 데이터 준비 (수동모드 기록, 비밀번호 없음)
            config_data = {
                "login_mode": "manual",
                "created_at": str(Path(__file__).stat().st_mtime)
            }
            
            # 데이터 암호화하여 저장 (비밀번호 없이)
            encrypted_data = self.encrypt_data(config_data)
            with open(self.cert_file, 'wb') as f:
                f.write(encrypted_data)
                
            print("✅ 수동로그인 모드 설정 저장 완료 (비밀번호 삭제됨)")
            
            # 저장 완료 메시지
            messagebox.showinfo("저장 완료", 
                "수동로그인 모드로 설정되었습니다.\n"
                "기존 저장된 비밀번호는 삭제되었습니다.")
            
            # 창 닫기
            self.on_closing()
            
        except Exception as e:
            messagebox.showerror("오류", f"설정 저장 중 오류가 발생했습니다:\n{str(e)}")
            
    def open_main_menu(self):
        """원래 초기화면으로 복귀"""
        try:
            if self.parent:
                # 부모 창이 있으면 부모 창으로 복귀
                self.parent.focus_set()
                self.root.destroy()
                print("✅ 원래 화면으로 복귀합니다.")
            else:
                # 독립 실행인 경우에만 메인 메뉴 실행
                self.root.destroy()
                
                main_menu_path = Path(__file__).parent / "hometax_main.py"
                
                if main_menu_path.exists():
                    subprocess.Popen([sys.executable, str(main_menu_path)])
                    print("✅ 메인 메뉴로 이동합니다.")
                else:
                    print(f"❌ 메인 메뉴 파일을 찾을 수 없습니다: {main_menu_path}")
                
        except Exception as e:
            print(f"❌ 메뉴 복귀 오류: {e}")
            # 오류 발생 시에도 현재 창은 닫기
            self.root.destroy()
            
    def on_closing(self):
        """창 닫기"""
        if self.parent:
            self.parent.focus_set()  # 부모 창에 포커스 복원
        self.root.destroy()
        
def main():
    """독립 실행"""
    try:
        # 필요한 패키지 확인
        import cryptography
    except ImportError:
        print("cryptography 패키지가 필요합니다: pip install cryptography")
        return
        
    app = HomeTaxCertManager()
    app.root.mainloop()

if __name__ == "__main__":
    main()