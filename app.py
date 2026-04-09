import tkinter as tk
from tkinter import ttk, messagebox
import sys
import os
import json
import sv_ttk
from pathlib import Path
from PIL import Image, ImageTk

# Add local directories to path to ensure imports work correctly
sys.path.append(os.path.dirname(os.path.abspath(__file__)))

from ui.tabs.match_tab import MatchTab
from ui.tabs.batch_tab import BatchTab
from ui.tabs.cleaner_tab import CleanerTab
from ui.tabs.stats_tab import StatsTab
from ui.widgets.admin_settings import AdminSettingsPopup
from utils.license_manager import LicenseManager
from utils.telemetry import TelemetryManager

class SafeSplash(tk.Toplevel):
    def __init__(self, callback):
        super().__init__()
        self.callback = callback
        self.title("EasyMatch Loading...")
        self.geometry("600x450")
        
        # Center screen
        sw = self.winfo_screenwidth()
        sh = self.winfo_screenheight()
        self.geometry(f"600x450+{sw//2-300}+{sh//2-225}")
        
        self.setup_ui()
        self.start_loading()

    def setup_ui(self):
        self.container = tk.Frame(self, bg="#001F3F")
        self.container.pack(fill="both", expand=True)
        
        # Logo
        try:
            logo_img = Image.open("assets/logo.png")
            logo_img = logo_img.resize((200, 200), Image.Resampling.LANCZOS)
            self.logo_photo = ImageTk.PhotoImage(logo_img)
            ttk.Label(self.container, image=self.logo_photo, background="#001F3F").pack(pady=(40, 10))
        except:
            ttk.Label(self.container, text="[ LOAD ]", font=("System", 70, "bold"), foreground="#00E5FF", background="#001F3F").pack(pady=40)
            
        ttk.Label(self.container, text="EasyMatch", font=("System", 32, "bold"), foreground="white", background="#001F3F").pack()
        ttk.Label(self.container, text="Powered by Advanced Data IQ", font=("System", 10), foreground="#00E5FF", background="#001F3F").pack(pady=10)
        
        self.progress = ttk.Progressbar(self.container, mode='determinate', length=400)
        self.progress.pack(pady=20)
        self.progress_val = 0

    def start_loading(self):
        if self.progress_val < 100:
            self.progress_val += 4
            self.progress['value'] = self.progress_val
            self.after(50, self.start_loading)
        else:
            self.after(500, self.finish)

    def finish(self):
        self.destroy()
        self.master.after(200, self.callback)

class EasyMatchPro(tk.Tk):
    def __init__(self):
        super().__init__()
        self.withdraw() # Hide until splash finish
        self.config_path = Path("config.json")
        self.load_config()
        
        self.title(f"{self.config['branding']['name']} {self.config['branding']['version']}")
        self.geometry("1450x900")
        
        # Start Theme later
        self.after_idle(lambda: self.apply_theme(self.config['branding'].get('theme', 'dark')))
        
        # Start Splash
        SafeSplash(self.launch_main)

    def load_config(self):
        if not self.config_path.exists():
            self.config = {
                "locked_features": {"batch_tab": False, "cleaner_tab": False, "cloud_source": False, "google_sheets": False},
                "branding": {"name": "EasyMatch", "version": "v3.9 Pro Official", "theme": "dark"},
                "telemetry": {"enabled": False, "url": ""},
                "registered_sources": {"github_url": "", "github_token": "", "google_sheets_url": "", "google_sheet_names": ""}
            }
        else:
            try:
                with open(self.config_path, 'r', encoding='utf-8') as f:
                    self.config = json.load(f)
                    # Migrating old configs
                    if 'registered_sources' not in self.config:
                        self.config['registered_sources'] = {"github_url": "", "github_token": "", "google_sheets_url": "", "google_sheet_names": ""}
            except:
                self.config = {
                    "locked_features": {"batch_tab": False, "cleaner_tab": False, "cloud_source": False, "google_sheets": False},
                    "branding": {"name": "EasyMatch", "version": "v3.9 Pro Official", "theme": "dark"},
                    "telemetry": {"enabled": False, "url": ""},
                    "registered_sources": {"github_url": "", "github_token": "", "google_sheets_url": "", "google_sheet_names": ""}
                }


    def apply_theme(self, theme_name):
        self.config['branding']['theme'] = theme_name
        
        # Base Theme (Standard sv-ttk)
        base = "dark" if theme_name in ["dark", "cosmic", "graphite"] else "light"
        sv_ttk.set_theme(base)
        
        style = ttk.Style(self)
        style.configure("Header.TLabel", font=("System", 18, "bold"))
        
        # Professional Palette Definition
        palettes = {
            "dark": {"bg": "#1E1E1E", "accent": "#4A90E2"},     # VS Code inspired
            "light": {"bg": "#F8F9FA", "accent": "#4A90E2"},    # Modern Clean
            "cosmic": {"bg": "#0F172A", "accent": "#818CF8"},   # Midnight Blue (Soft Indigo)
            "graphite": {"bg": "#18181B", "accent": "#71717A"}  # Zinc Grey
        }
        
        p = palettes.get(theme_name, palettes['dark'])
        cfg_bg = p['bg']
        cfg_accent = p['accent']
        
        # Apply Logic
        self.configure(bg=cfg_bg)
        style.configure("Accent.TButton", background=cfg_accent, foreground="white")
        style.map("Accent.TButton", background=[("active", cfg_accent), ("pressed", cfg_accent)])
        
        # Header/Notebook Tweaks for Comfort
        style.configure("TNotebook", background=cfg_bg)
        style.configure("TFrame", background=cfg_bg)
        
        # Update StatsTab if it exists
        if hasattr(self, 'tab_stats'):
            self.tab_stats.update_theme(theme_name)



    def launch_main(self):
        # 1. Check License first
        if not self.check_license_status():
            self.show_license_gate()
        else:
            self.start_app()

    def check_license_status(self):
        """Verify the stored license key against the current machine ID."""
        saved_key = self.config.get('license_key')
        if not saved_key: return False
        
        mid = LicenseManager.get_machine_id()
        return LicenseManager.verify_key(mid, saved_key)

    def show_license_gate(self):
        mid = LicenseManager.get_machine_id()
        gate = tk.Toplevel(self)
        gate.title("EasyMatch License Verification")
        gate.geometry("500x550")
        gate.resizable(False, False)
        gate.grab_set()
        
        # Center Gate
        sw = self.winfo_screenwidth()
        sh = self.winfo_screenheight()
        gate.geometry(f"500x550+{sw//2-250}+{sh//2-275}")
        
        # UI
        main = ttk.Frame(gate, padding=30)
        main.pack(fill="both", expand=True)
        
        ttk.Label(main, text="기기 인증 및 라이센스 등록", font=("System", 14, "bold")).pack(pady=(0, 20))
        
        ttk.Label(main, text="아래 '기기 고유 ID'를 복사하여 관리자에게 전달하세요.", font=("System", 9)).pack(anchor="w")
        
        id_frame = ttk.Frame(main)
        id_frame.pack(fill="x", pady=10)
        id_entry = ttk.Entry(id_frame, font=("Consolas", 10), justify="center")
        id_entry.insert(0, mid)
        id_entry.config(state="readonly")
        id_entry.pack(fill="x", side="left", expand=True)
        
        def copy_id():
            self.clipboard_clear()
            self.clipboard_append(mid)
            messagebox.showinfo("복사 완료", "기기 ID가 클립보드에 복사되었습니다.")
        
        ttk.Button(id_frame, text="복사", width=5, command=copy_id).pack(side="right", padx=5)

        ttk.Separator(main, orient="horizontal").pack(fill="x", pady=15)
        
        ttk.Label(main, text="사용자 성함:", font=("System", 10, "bold")).pack(anchor="w")
        name_var = tk.StringVar()
        ttk.Entry(main, textvariable=name_var, font=("System", 11)).pack(fill="x", pady=2)
        
        ttk.Label(main, text="이메일 주소 (AS 및 기술지원 용):", font=("System", 10, "bold")).pack(anchor="w", pady=(10, 0))
        email_var = tk.StringVar()
        ttk.Entry(main, textvariable=email_var, font=("System", 11)).pack(fill="x", pady=2)

        ttk.Label(main, text="라이센스 키 입력:", font=("System", 10, "bold")).pack(anchor="w", pady=(10, 0))
        key_var = tk.StringVar()
        key_entry = ttk.Entry(main, textvariable=key_var, font=("Consolas", 11), justify="center")
        key_entry.pack(fill="x", pady=5)
        
        def verify():
            provided = key_var.get().strip()
            u_name = name_var.get().strip()
            u_email = email_var.get().strip()
            
            if not u_name or not u_email:
                messagebox.showwarning("경고", "성함과 이메일을 입력하세요.")
                return

            if LicenseManager.verify_key(mid, provided):
                self.config['license_key'] = provided
                self.config['user_info'] = {"name": u_name, "email": u_email}
                
                # Ping Activation
                tel_cfg = self.config.get('telemetry', {})
                if tel_cfg.get('enabled'):
                    TelemetryManager.log_event(tel_cfg.get('url'), "LICENSE_ACTIVATE", user_info=self.config['user_info'])
                
                with open(self.config_path, 'w', encoding='utf-8') as f:
                    json.dump(self.config, f, indent=4, ensure_ascii=False)
                messagebox.showinfo("성공", "라이센스가 성공적으로 등록되었습니다!")
                gate.destroy()
                self.start_app()
            else:
                messagebox.showerror("오류", "유효하지 않은 라이센스 키입니다.")

        ttk.Button(main, text="라이센스 등록 및 시작", command=verify, style="Accent.TButton").pack(fill="x", pady=20)
        
        # If gate is closed without license, exit app
        gate.protocol("WM_DELETE_WINDOW", sys.exit)

    def start_app(self):
        # Small delay for macOS ARM window mapping stability
        self.after(100, self._show_main)

    def _show_main(self):
        self.deiconify()
        self.build_ui()
        
        # Ping App Start (Moved here for macOS stability)
        tel_cfg = self.config.get('telemetry', {})
        if tel_cfg.get('enabled'):
            TelemetryManager.log_event(tel_cfg.get('url'), "APP_START", user_info=self.config.get('user_info'))


    def build_ui(self):
        # Header
        header_container = ttk.Frame(self, padding=(20, 10))
        header_container.pack(fill="x")
        
        header = ttk.Frame(header_container)
        header.pack(side="left")
        ttk.Label(header, text="이지매치", style="Header.TLabel").pack(side="left")
        ttk.Label(header, text="EasyMatch Pro", foreground="#00E5FF", font=("System", 11, "bold")).pack(side="left", padx=10, pady=(5,0))
        
        # Settings & Theme Buttons
        btn_frame = ttk.Frame(header_container)
        btn_frame.pack(side="right")
        
        # Theme Dropdown
        theme_menu = tk.Menubutton(btn_frame, text="THEME ▼", relief="flat", font=("System", 9), direction="below")
        theme_menu.menu = tk.Menu(theme_menu, tearoff=0)
        theme_menu["menu"] = theme_menu.menu
        
        for t in ["Dark", "Light", "Cosmic", "Graphite"]:
            theme_menu.menu.add_command(label=t, command=lambda x=t.lower(): self.apply_theme(x))
        
        theme_menu.pack(side="left", padx=5)
        ttk.Button(btn_frame, text="ADMIN", width=7, command=self.open_admin_settings).pack(side="right", padx=5)


        # Main Notebook
        self.notebook = ttk.Notebook(self)
        self.notebook.pack(fill="both", expand=True, padx=15, pady=10)

        # Tabs (Passing config for registered sources)
        self.tab_match = MatchTab(self.notebook, self.config)
        self.notebook.add(self.tab_match, text="  스마트 매칭  ")


        self.tab_stats = StatsTab(self.notebook)
        self.notebook.add(self.tab_stats, text="  데이터 인사이트  ")

        # Feature Locking
        if not self.config['locked_features'].get('batch_tab'):
            self.tab_batch = BatchTab(self.notebook)
            self.notebook.add(self.tab_batch, text="  일괄 처리(Batch)  ")

        if not self.config['locked_features'].get('cleaner_tab'):
            self.tab_cleaner = CleanerTab(self.notebook)
            self.notebook.add(self.tab_cleaner, text="  데이터 세척(Cleaner)  ")

        self.tab_match.register_on_load(self.on_data_loaded)

        # Footer & Pulse System
        footer = ttk.Frame(self, padding=(20, 10))
        footer.pack(fill="x", side="bottom")
        
        # Pulse (Satisfaction) Bar
        pulse_frame = ttk.Frame(footer)
        pulse_frame.pack(side="left")
        
        ttk.Label(pulse_frame, text="EasyMatch Pulse:", font=("System", 9, "italic")).pack(side="left", padx=(0, 10))
        
        self.like_btn = ttk.Button(pulse_frame, text="LIKE", width=10, command=self.send_pulse_like)
        self.like_btn.pack(side="left", padx=2)
        
        ttk.Button(pulse_frame, text="REVIEW", width=12, command=self.open_feedback_popup).pack(side="left", padx=2)
        
        ttk.Label(footer, text="Professional Edition | Optimized for Extreme Data", 
                  font=("System", 9), foreground="gray").pack(side="right")

    def send_pulse_like(self):
        tel_cfg = self.config.get('telemetry', {})
        if tel_cfg.get('enabled'):
            TelemetryManager.log_event(tel_cfg.get('url'), "PULSE_LIKE", user_info=self.config.get('user_info'))
            self.like_btn.config(text="THANK YOU", state="disabled")
            messagebox.showinfo("EasyMatch Pulse", "개발자에게 큰 힘이 됩니다. 감사합니다!")

    def open_feedback_popup(self):
        # Custom Feedback Dialog to avoid macOS simpledialog crashes
        top = tk.Toplevel(self)
        top.title("EasyMatch Feedback")
        top.geometry("400x250")
        top.transient(self)
        top.grab_set()
        
        main = ttk.Frame(top, padding=20)
        main.pack(fill="both", expand=True)
        
        ttk.Label(main, text="개발자에게 의견을 보내주세요(100자 내외):", font=("System", 10)).pack(anchor="w", pady=(0, 10))
        
        txt = tk.Text(main, height=5, font=("System", 11))
        txt.pack(fill="both", expand=True, pady=5)
        txt.focus_set()
        
        def commit():
            msg = txt.get("1.0", "end-1c").strip()
            if not msg:
                top.destroy()
                return
            
            tel_cfg = self.config.get('telemetry', {})
            if tel_cfg.get('enabled'):
                TelemetryManager.log_event(tel_cfg.get('url'), "PULSE_REVIEW", {"msg": msg}, user_info=self.config.get('user_info'))
            
            top.destroy()
            messagebox.showinfo("EasyMatch Pulse", "소중한 의견 감사합니다! 품질 개선에 반영하겠습니다.")
            
        btn_pk = ttk.Frame(main)
        btn_pk.pack(fill="x", pady=(10, 0))
        ttk.Button(btn_pk, text="취소", command=top.destroy).pack(side="right", padx=5)
        ttk.Button(btn_pk, text="전송", command=commit).pack(side="right", padx=5)

    def open_admin_settings(self):
        AdminSettingsPopup(self)

    def on_data_loaded(self, df):
        if hasattr(self, 'tab_stats'):
            self.tab_stats.set_data(df)

if __name__ == "__main__":
    import logging
    import traceback
    
    logging.basicConfig(filename="easymatch_debug.log", level=logging.ERROR, format="%(asctime)s - %(message)s")

    try:
        app = EasyMatchPro()
        app.mainloop()
    except Exception as e:
        traceback.print_exc()
