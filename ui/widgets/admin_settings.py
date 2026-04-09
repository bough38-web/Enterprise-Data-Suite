import tkinter as tk
from tkinter import ttk, messagebox
import json
from pathlib import Path
from utils.license_manager import LicenseManager
from utils.telemetry import TelemetryManager

class AdminSettingsPopup(tk.Toplevel):
    def __init__(self, parent):
        super().__init__(parent)
        self.title("Administrator Settings")
        self.geometry("450x700")
        self.config_path = Path("config.json")
        self.grab_set() # Modal
        
        self.authenticated = False
        self.build_auth_ui()

    def build_auth_ui(self):
        self.auth_frame = ttk.Frame(self, padding=40)
        self.auth_frame.pack(fill="both", expand=True)
        
        ttk.Label(self.auth_frame, text="[ AUTH ] 관리자 암호를 입력하세요", font=("System", 11, "bold")).pack(pady=20)
        
        self.pw_var = tk.StringVar()
        pw_entry = ttk.Entry(self.auth_frame, textvariable=self.pw_var, show="*", font=("System", 12), justify="center")
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
        
        ttk.Label(main, text="[ CONFIG ] 기능 제한 설정 (EasyMatch Pro)", font=("System", 12, "bold")).pack(pady=(0, 20))
        
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
        
        ttk.Label(main, text="* 설정 변경 후 프로그램을 재시작해야 적용됩니다.", foreground="gray", font=("System", 8)).pack()
        
        btn_frame = ttk.Frame(main)
        btn_frame.pack(side="bottom", fill="x", pady=10)

        # License Management Section
        ttk.Separator(main, orient="horizontal").pack(fill="x", pady=20)
        ttk.Label(main, text="[ LICENSE ] 라이센스 관리 (Key Generator)", font=("System", 11, "bold")).pack(pady=(0, 10))
        
        # Current Machine Info
        mid = LicenseManager.get_machine_id()
        u_info = self.config.get('user_info', {})
        u_str = f"{u_info.get('name', '미등록')} ({u_info.get('email', '-')})"
        
        ttk.Label(main, text=f"등록 사용자: {u_str}", foreground="#00E5FF", font=("System", 9, "bold")).pack(anchor="w")
        ttk.Label(main, text=f"이 기기 ID: {mid}", foreground="gray", font=("System", 8)).pack(anchor="w")
        
        # Key Generator
        gen_frame = ttk.LabelFrame(main, text="키 생성기", padding=10)
        gen_frame.pack(fill="x", pady=10)
        
        ttk.Label(gen_frame, text="타겟 기기 ID 입력:").pack(anchor="w")
        target_id_var = tk.StringVar()
        ttk.Entry(gen_frame, textvariable=target_id_var).pack(fill="x", pady=5)
        
        def generate_custom_key():
            target = target_id_var.get().strip()
            if not target: return
            new_key = LicenseManager.generate_key(target)
            self.clipboard_clear()
            self.clipboard_append(new_key)
            messagebox.showinfo("생성 완료", f"라이센스 키가 생성되어 클립보드에 복사되었습니다:\n\n{new_key}")
            
        ttk.Button(gen_frame, text="라이센스 키 생성 및 복사", command=generate_custom_key).pack(fill="x", pady=5)

        def reset_license_data():
            if messagebox.askyesno("주의", "현재 등록된 라이센스를 삭제할까요?"):
                self.config['license_key'] = ""
                with open(self.config_path, 'w', encoding='utf-8') as f:
                    json.dump(self.config, f, indent=4, ensure_ascii=False)
                messagebox.showinfo("완료", "라이센스가 초기화되었습니다.")

        ttk.Button(main, text="[ RESET ] 이 기기 라이센스 초기화", command=reset_license_data).pack(fill="x", pady=10)

        # Telemetry Section
        ttk.Separator(main, orient="horizontal").pack(fill="x", pady=15)
        ttk.Label(main, text="[ MONITOR ] 원격 사용현황 모니터링 (Telemetry)", font=("System", 11, "bold")).pack(pady=(0, 10))
        
        tel_cfg = self.config.get('telemetry', {})
        self.tel_enabled = tk.BooleanVar(value=tel_cfg.get('enabled', False))
        ttk.Checkbutton(main, text="사용현황 서버(구글시트)로 전송 활성화", variable=self.tel_enabled).pack(anchor="w")
        
        ttk.Label(main, text="모니터링 Webhook URL:", font=("System", 9)).pack(anchor="w", pady=(10, 0))
        self.tel_url = tk.StringVar(value=tel_cfg.get('url', ''))
        ttk.Entry(main, textvariable=self.tel_url).pack(fill="x", pady=5)
        
        def test_tel():
            url = self.tel_url.get().strip()
            if not url: return
            try:
                if TelemetryManager.test_ping(url):
                    messagebox.showinfo("성공", "연결 성공! 구글 시트를 확인하세요.")
                else:
                    messagebox.showerror("실패", "서버 응답이 200(OK)이 아닙니다.")
            except Exception as e:
                messagebox.showerror("오류", str(e))
                
        ttk.Button(main, text="테스트 핑(Test Ping) 전송", command=test_tel).pack(fill="x", pady=5)

        ttk.Button(btn_frame, text="저장 및 닫기", command=self.save_and_close).pack(fill="x")

    def save_and_close(self):
        for key, var in self.toggles.items():
            self.config['locked_features'][key] = var.get()
            
        self.config['telemetry'] = {
            "enabled": self.tel_enabled.get(),
            "url": self.tel_url.get().strip()
        }
            
        with open(self.config_path, 'w', encoding='utf-8') as f:
            json.dump(self.config, f, indent=4, ensure_ascii=False)
            
        messagebox.showinfo("완료", "설정이 저장되었습니다. 프로그램을 다시 시작해 주세요.")
        self.destroy()
