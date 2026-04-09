import tkinter as tk
from tkinter import ttk, messagebox
import json
from pathlib import Path

class AdminSettingsPopup(tk.Toplevel):
    def __init__(self, parent):
        super().__init__(parent)
        self.title("Administrator Settings")
        self.geometry("400x500")
        self.config_path = Path("config.json")
        self.grab_set() # Modal
        
        self.authenticated = False
        self.build_auth_ui()

    def build_auth_ui(self):
        self.auth_frame = ttk.Frame(self, padding=40)
        self.auth_frame.pack(fill="both", expand=True)
        
        ttk.Label(self.auth_frame, text="🔒 관리자 암호를 입력하세요", font=("Segoe UI", 11, "bold")).pack(pady=20)
        
        self.pw_var = tk.StringVar()
        pw_entry = ttk.Entry(self.auth_frame, textvariable=self.pw_var, show="*", font=("Segoe UI", 12), justify="center")
        pw_entry.pack(fill="x", pady=10)
        pw_entry.focus_set()
        
        ttk.Button(self.auth_frame, text="승인", command=self.check_auth).pack(pady=20)
        
        # Bind Enter key
        self.bind("<Return>", lambda e: self.check_auth())

    def check_auth(self):
        if self.pw_var.get() == "3867":
            self.authenticated = True
            self.auth_frame.destroy()
            self.build_settings_ui()
        else:
            messagebox.showerror("오류", "암호가 틀렸습니다.")
            self.pw_var.set("")

    def build_settings_ui(self):
        main = ttk.Frame(self, padding=20)
        main.pack(fill="both", expand=True)
        
        ttk.Label(main, text="🛠️ 기능 제한 설정 (EasyMatch Pro)", font=("Segoe UI", 12, "bold")).pack(pady=(0, 20))
        
        # Load Current Config
        with open(self.config_path, 'r', encoding='utf-8') as f:
            self.config = json.load(f)
            
        self.toggles = {}
        features = {
            "batch_tab": "폴더 일괄 처리(Batch) 탭 숨기기",
            "cleaner_tab": "데이터 세척 탭 숨기기",
            "cloud_source": "GitHub 클라우드 소스 숨기기",
            "google_sheets": "구글 스프레드시트 연동 숨기기"
        }
        
        for key, label in features.items():
            var = tk.BooleanVar(value=self.config['locked_features'].get(key, False))
            self.toggles[key] = var
            ttk.Checkbutton(main, text=label, variable=var).pack(anchor="w", pady=5)

        ttk.Separator(main, orient="horizontal").pack(fill="x", pady=20)
        
        ttk.Label(main, text="* 설정 변경 후 프로그램을 재시작해야 적용됩니다.", foreground="gray", font=("Segoe UI", 8)).pack()
        
        btn_frame = ttk.Frame(main)
        btn_frame.pack(side="bottom", fill="x", pady=10)
        ttk.Button(btn_frame, text="저장 및 닫기", command=self.save_and_close).pack(fill="x")

    def save_and_close(self):
        for key, var in self.toggles.items():
            self.config['locked_features'][key] = var.get()
            
        with open(self.config_path, 'w', encoding='utf-8') as f:
            json.dump(self.config, f, indent=4, ensure_ascii=False)
            
        messagebox.showinfo("완료", "설정이 저장되었습니다. 프로그램을 다시 시작해 주세요.")
        self.destroy()
