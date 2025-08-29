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

# 보안 관리자 import
from hometax_security_manager import HomeTaxSecurityManager

class HomeTaxCertManager:
    def __init__(self, parent=None):
        self.parent = parent
        self.root = tk.Toplevel(parent) if parent else tk.Tk()
        self.cert_file = Path("cert_config.enc")  # 암호화된 인증서 정보 파일
        self.login_mode = tk.StringVar(value="auto")  # 기본값: 자동로그인
        self.security_manager = HomeTaxSecurityManager()  # 보안 관리자 초기화
        self.setup_window()
        self.create_widgets()
        self.load_saved_config()
        
    def setup_window(self):
        """창 설정"""
        self.root.title("공인인증서 비밀번호 관리")
        self.root.geometry("500x480")
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
            y = parent_y + (parent_height // 2) - (240)  # 480//2
            
            self.root.geometry(f"500x480+{x}+{y}")
            
            # 모달 창으로 설정
            self.root.transient(self.parent)
            self.root.grab_set()
        else:
            # 독립 실행시 화면 중앙에 위치
            self.root.eval('tk::PlaceWindow . center')
            
        # 창 닫기 이벤트 처리
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)
        
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
            text="🤖 자동로그인 (비밀번호 저장)",
            variable=self.login_mode,
            value="auto",
            command=self.on_mode_changed
        )
        auto_radio.pack(anchor=tk.W, pady=5)
        
        # 수동로그인 라디오버튼
        manual_radio = ttk.Radiobutton(
            login_mode_frame,
            text="✋ 수동로그인 (비밀번호 직접 입력)",
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
        self.password_entry.pack(fill=tk.X, pady=(5, 10))
        
        # 비밀번호 표시/숨기기 체크박스
        self.show_password = tk.BooleanVar()
        show_password_check = ttk.Checkbutton(
            self.auto_frame,
            text="비밀번호 표시",
            variable=self.show_password,
            command=self.toggle_password_visibility
        )
        show_password_check.pack(anchor=tk.W, pady=(0, 10))
        
        # 자동로그인 버튼들
        auto_buttons_frame = ttk.Frame(self.auto_frame)
        auto_buttons_frame.pack(fill=tk.X)
        
        ttk.Button(
            auto_buttons_frame,
            text="💾 비밀번호 저장",
            command=self.save_password
        ).pack(side=tk.LEFT, padx=(0, 10))
        
        ttk.Button(
            auto_buttons_frame,
            text="🗑️ 저장된 비밀번호 삭제",
            command=self.delete_password
        ).pack(side=tk.LEFT)
        
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
            
        if self.login_mode.get() == "auto":
            # 자동로그인 모드: 저장 및 닫기, 닫기 버튼
            ttk.Button(
                self.button_frame,
                text="닫기",
                command=self.on_closing
            ).pack(side=tk.RIGHT)
            
            ttk.Button(
                self.button_frame,
                text="💾 저장 및 닫기",
                command=self.save_and_close,
                style='Accent.TButton'
            ).pack(side=tk.RIGHT, padx=(0, 10))
            
            # 공간 확보를 위한 더미 프레임
            ttk.Frame(self.button_frame, width=20).pack(side=tk.LEFT)
            
        else:
            # 수동로그인 모드: 시작, 닫기 버튼
            ttk.Button(
                self.button_frame,
                text="닫기", 
                command=self.on_closing
            ).pack(side=tk.RIGHT)
            
            ttk.Button(
                self.button_frame,
                text="🚀 HomeTax 자동화 시작",
                command=self.start_automation,
                style='Accent.TButton'
            ).pack(side=tk.RIGHT, padx=(0, 10))
            
    def toggle_password_visibility(self):
        """비밀번호 표시/숨기기"""
        if self.show_password.get():
            self.password_entry.config(show="")
        else:
            self.password_entry.config(show="*")
            
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
            
    def save_password(self):
        """비밀번호 저장"""
        password = self.password_entry.get().strip()
        
        if not password:
            messagebox.showwarning("경고", "비밀번호를 입력하세요.")
            return
            
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
            if not self.security_manager.save_encrypted_password_to_env(password):
                print("⚠️ .env 파일 암호화 저장 실패 - 레거시 모드로 저장")
                # 백업 저장 방식 (레거시 지원)
                env_file = Path(".env")
                env_content = ""
                if env_file.exists():
                    with open(env_file, 'r', encoding='utf-8') as f:
                        env_content = f.read()
                        
                lines = env_content.split('\n')
                pw_found = False
                for i, line in enumerate(lines):
                    if line.startswith('PW='):
                        lines[i] = f'PW={password}'
                        pw_found = True
                        break
                        
                if not pw_found:
                    lines.append(f'PW={password}')
                    
                with open(env_file, 'w', encoding='utf-8') as f:
                    f.write('\n'.join(lines))
                
            self.update_status()
            messagebox.showinfo("성공", "비밀번호가 안전하게 저장되었습니다.")
            self.password_entry.delete(0, tk.END)
            return True
            
        except Exception as e:
            messagebox.showerror("오류", f"비밀번호 저장 중 오류가 발생했습니다:\n{str(e)}")
            return False
            
    def save_and_close(self):
        """비밀번호 저장하고 창 닫기"""
        password = self.password_entry.get().strip()
        
        if not password:
            messagebox.showwarning("경고", "비밀번호를 입력하세요.")
            return
            
        # 비밀번호 저장 시도
        if self.save_password():
            # 저장 성공시 창 닫기
            self.on_closing()
            
    def delete_password(self):
        """저장된 비밀번호 삭제"""
        if messagebox.askyesno("확인", "저장된 비밀번호를 삭제하시겠습니까?"):
            try:
                # 암호화 파일 삭제
                if self.cert_file.exists():
                    self.cert_file.unlink()
                    
                # .env 파일에서 암호화된 비밀번호와 평문 비밀번호 모두 제거
                env_file = Path(".env")
                if env_file.exists():
                    with open(env_file, 'r', encoding='utf-8') as f:
                        lines = f.readlines()
                        
                    with open(env_file, 'w', encoding='utf-8') as f:
                        for line in lines:
                            if not (line.strip().startswith('PW=') or line.strip().startswith('PW_ENCRYPTED=')):
                                f.write(line)
                                
                self.update_status()
                messagebox.showinfo("성공", "저장된 비밀번호가 삭제되었습니다.")
                
            except Exception as e:
                messagebox.showerror("오류", f"비밀번호 삭제 중 오류가 발생했습니다:\n{str(e)}")
                
    def load_saved_config(self):
        """저장된 설정 불러오기"""
        try:
            # 암호화 파일에서 설정 로드
            if self.cert_file.exists():
                with open(self.cert_file, 'rb') as f:
                    encrypted_data = f.read()
                    
                config_data = self.decrypt_data(encrypted_data)
                if config_data:
                    self.login_mode.set(config_data.get("login_mode", "auto"))
                    
            # .env 파일에서도 확인 (암호화된 것 우선, fallback으로 평문 지원)
            password = self.security_manager.load_password_from_env()
            if password:
                # .env에 비밀번호가 있으면 자동로그인 모드로 설정
                if not self.cert_file.exists():
                    self.login_mode.set("auto")
                            
            self.update_status()
            self.on_mode_changed()
            
        except Exception as e:
            print(f"설정 로드 중 오류: {e}")
            
    def update_status(self):
        """상태 표시 업데이트"""
        has_encrypted = self.cert_file.exists()
        
        # 보안 관리자를 통해 비밀번호 상태 확인
        password_status = self.security_manager.validate_password_security()
        has_env_password = password_status['status'] in ['secure', 'mixed', 'plaintext']
        
        if has_encrypted or has_env_password:
            if password_status['secure']:
                self.status_label.config(
                    text="현재 저장된 비밀번호: 있음 ✓ (암호화됨)",
                    foreground='#28A745'
                )
            else:
                self.status_label.config(
                    text="현재 저장된 비밀번호: 있음 ⚠️ (보안 개선 필요)",
                    foreground='#FFA500'
                )
        else:
            self.status_label.config(
                text="현재 저장된 비밀번호: 없음",
                foreground='#666666'
            )
            
    def start_automation(self):
        """HomeTax 자동화 시작"""
        try:
            # 선택된 모드에 따라 환경 설정
            mode = self.login_mode.get()
            
            if mode == "auto":
                # 자동로그인 모드 - 저장된 비밀번호 확인
                has_password = False
                
                if self.cert_file.exists():
                    with open(self.cert_file, 'rb') as f:
                        encrypted_data = f.read()
                    config_data = self.decrypt_data(encrypted_data)
                    has_password = config_data is not None
                    
                if not has_password:
                    # .env 파일에서 비밀번호 확인 (보안 관리자 사용)
                    env_password = self.security_manager.load_password_from_env()
                    if env_password:
                        has_password = True
                                    
                if not has_password:
                    messagebox.showwarning("경고", "자동로그인을 위해서는 먼저 비밀번호를 저장해야 합니다.")
                    return
                    
            # HomeTax 자동화 실행
            script_path = Path(__file__).parent / "hometax_quick.py"
            
            if script_path.exists():
                # 새로운 프로세스로 실행
                subprocess.Popen([sys.executable, str(script_path)], 
                               env={**os.environ, 'HOMETAX_LOGIN_MODE': mode})
                
                messagebox.showinfo("시작", f"HomeTax 자동화가 시작되었습니다.\n모드: {mode}")
                
                # 창 닫기
                self.on_closing()
            else:
                messagebox.showerror("오류", f"자동화 스크립트를 찾을 수 없습니다:\n{script_path}")
                
        except Exception as e:
            messagebox.showerror("오류", f"자동화 시작 중 오류가 발생했습니다:\n{str(e)}")
            
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