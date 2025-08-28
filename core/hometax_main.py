# -*- coding: utf-8 -*-
"""
HomeTax 전자세금계산서 시스템 - 메인 화면
전자세금계산서 발행, 거래처 관리, 조회 등의 기능을 통합한 메인 시스템
"""

import tkinter as tk
from tkinter import ttk, messagebox
import subprocess
import sys
import os
from pathlib import Path

class HomeTaxMainSystem:
    def __init__(self):
        self.root = tk.Tk()
        self.setup_main_window()
        self.create_widgets()
        
    def setup_main_window(self):
        """메인 창 설정"""
        self.root.title("HomeTax 전자세금계산서 시스템")
        self.root.geometry("800x600")
        self.root.resizable(True, True)
        
        # 창을 화면 중앙에 위치
        self.root.eval('tk::PlaceWindow . center')
        
        # 최소 크기 설정
        self.root.minsize(700, 500)
        
        # 아이콘 설정 (옵션)
        try:
            # 기본 시스템 아이콘 사용
            self.root.iconbitmap(default=True)
        except:
            pass
    
    def create_widgets(self):
        """위젯 생성"""
        # 메인 프레임
        main_frame = ttk.Frame(self.root, padding="30")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # 제목
        title_label = ttk.Label(
            main_frame, 
            text="HomeTax 전자세금계산서 시스템",
            font=('맑은 고딕', 24, 'bold'),
            foreground='#2E4057'
        )
        title_label.pack(pady=(0, 40))
        
        # 부제목
        subtitle_label = ttk.Label(
            main_frame,
            text="전자세금계산서 발행 및 거래처 관리 통합 시스템",
            font=('맑은 고딕', 12),
            foreground='#666666'
        )
        subtitle_label.pack(pady=(0, 50))
        
        # 메뉴 버튼들을 담을 프레임
        menu_frame = ttk.Frame(main_frame)
        menu_frame.pack(expand=True, fill=tk.BOTH)
        
        # 그리드 설정 (2x3 레이아웃)
        for i in range(2):
            menu_frame.columnconfigure(i, weight=1)
        for i in range(3):
            menu_frame.rowconfigure(i, weight=1)
        
        # 버튼 스타일 설정
        style = ttk.Style()
        style.configure(
            'MenuButton.TButton',
            font=('맑은 고딕', 12, 'bold'),
            padding=(20, 15)
        )
        
        # 메뉴 버튼들
        self.create_menu_buttons(menu_frame)
        
        # 하단 정보
        info_frame = ttk.Frame(main_frame)
        info_frame.pack(side=tk.BOTTOM, fill=tk.X, pady=(30, 0))
        
        # 버전 정보
        version_label = ttk.Label(
            info_frame,
            text="Version 1.0.0 | HomeTax Automation System",
            font=('맑은 고딕', 9),
            foreground='#999999'
        )
        version_label.pack(side=tk.LEFT)
        
        # 상태 표시
        self.status_label = ttk.Label(
            info_frame,
            text="시스템 준비 완료",
            font=('맑은 고딕', 9),
            foreground='#28A745'
        )
        self.status_label.pack(side=tk.RIGHT)
    
    def create_menu_buttons(self, parent):
        """메뉴 버튼들 생성"""
        
        # 1. 전자세금계산서 자동발행
        btn_auto_issue = ttk.Button(
            parent,
            text="📄 전자세금계산서\n자동발행",
            style='MenuButton.TButton',
            command=self.run_auto_issue
        )
        btn_auto_issue.grid(row=0, column=0, padx=20, pady=15, sticky='nsew')
        
        # 2. 거래처 등록관리
        btn_partner_mgmt = ttk.Button(
            parent,
            text="🏢 거래처\n등록관리",
            style='MenuButton.TButton',
            command=self.run_partner_management
        )
        btn_partner_mgmt.grid(row=0, column=1, padx=20, pady=15, sticky='nsew')
        
        # 3. 거래명세서 조회
        btn_transaction_inquiry = ttk.Button(
            parent,
            text="📊 거래명세서\n조회",
            style='MenuButton.TButton',
            command=self.run_transaction_inquiry
        )
        btn_transaction_inquiry.grid(row=1, column=0, padx=20, pady=15, sticky='nsew')
        
        # 4. 세금계산서 조회
        btn_tax_invoice_inquiry = ttk.Button(
            parent,
            text="🔍 세금계산서\n조회",
            style='MenuButton.TButton',
            command=self.run_tax_invoice_inquiry
        )
        btn_tax_invoice_inquiry.grid(row=1, column=1, padx=20, pady=15, sticky='nsew')
        
        # 5. 공인인증서 비밀번호 관리 (전체 폭)
        btn_cert_mgmt = ttk.Button(
            parent,
            text="🔐 공인인증서 비밀번호 관리",
            style='MenuButton.TButton',
            command=self.run_cert_management
        )
        btn_cert_mgmt.grid(row=2, column=0, columnspan=2, padx=20, pady=15, sticky='nsew')
        
        # 각 버튼에 툴팁 스타일 효과 추가
        self.add_button_effects(btn_auto_issue, "HomeTax에서 세금계산서를 자동으로 발행합니다")
        self.add_button_effects(btn_partner_mgmt, "거래처 정보를 등록하고 관리합니다")
        self.add_button_effects(btn_transaction_inquiry, "거래명세서를 조회하고 확인합니다")
        self.add_button_effects(btn_tax_invoice_inquiry, "발행된 세금계산서를 조회합니다")
        self.add_button_effects(btn_cert_mgmt, "공인인증서 비밀번호를 안전하게 관리합니다")
    
    def add_button_effects(self, button, tooltip_text):
        """버튼에 마우스 호버 효과 및 툴팁 추가"""
        def on_enter(event):
            button.configure(cursor="hand2")
            self.status_label.configure(text=tooltip_text, foreground='#007BFF')
            
        def on_leave(event):
            button.configure(cursor="")
            self.status_label.configure(text="시스템 준비 완료", foreground='#28A745')
            
        button.bind("<Enter>", on_enter)
        button.bind("<Leave>", on_leave)
    
    def update_status(self, message, color='#28A745'):
        """상태 메시지 업데이트"""
        self.status_label.configure(text=message, foreground=color)
        self.root.update()
    
    def run_program(self, program_path, program_name):
        """외부 프로그램 실행"""
        if not os.path.exists(program_path):
            messagebox.showerror(
                "파일 오류",
                f"{program_name} 파일을 찾을 수 없습니다.\n경로: {program_path}"
            )
            return
            
        try:
            self.update_status(f"{program_name} 실행 중...", '#007BFF')
            
            # Python 스크립트 실행
            subprocess.Popen([sys.executable, program_path], 
                           cwd=os.path.dirname(program_path))
            
            self.update_status(f"{program_name} 실행 완료", '#28A745')
            
        except Exception as e:
            messagebox.showerror(
                "실행 오류",
                f"{program_name} 실행 중 오류가 발생했습니다.\n오류: {str(e)}"
            )
            self.update_status("실행 오류 발생", '#DC3545')
    
    def run_auto_issue(self):
        """전자세금계산서 자동발행 실행"""
        program_path = r"C:\APP\tax-bill\hometax_quick.py"
        self.run_program(program_path, "전자세금계산서 자동발행")
    
    def run_partner_management(self):
        """거래처 등록관리 실행"""
        program_path = r"C:\APP\tax-bill\hometax_excel_integration.py"
        self.run_program(program_path, "거래처 등록관리")
    
    def run_transaction_inquiry(self):
        """거래명세서 조회 실행 (미구현)"""
        messagebox.showinfo(
            "기능 안내",
            "거래명세서 조회 기능은 개발 예정입니다.\n추후 업데이트에서 제공됩니다."
        )
        self.update_status("거래명세서 조회 - 개발 예정", '#FFC107')
    
    def run_tax_invoice_inquiry(self):
        """세금계산서 조회 실행 (미구현)"""
        messagebox.showinfo(
            "기능 안내",
            "세금계산서 조회 기능은 개발 예정입니다.\n추후 업데이트에서 제공됩니다."
        )
        self.update_status("세금계산서 조회 - 개발 예정", '#FFC107')
    
    def run_cert_management(self):
        """공인인증서 비밀번호 관리 실행 (미구현)"""
        messagebox.showinfo(
            "기능 안내",
            "공인인증서 비밀번호 관리 기능은 개발 예정입니다.\n현재는 .env 파일을 통해 관리됩니다."
        )
        self.update_status("인증서 관리 - 개발 예정", '#FFC107')
    
    def run(self):
        """프로그램 실행"""
        self.root.mainloop()

def main():
    """메인 함수"""
    try:
        app = HomeTaxMainSystem()
        app.run()
    except Exception as e:
        messagebox.showerror("시스템 오류", f"시스템 실행 중 오류가 발생했습니다.\n오류: {str(e)}")

if __name__ == "__main__":
    main()