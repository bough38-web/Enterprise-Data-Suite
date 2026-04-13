import tkinter as tk
from tkinter import ttk, messagebox
import json
from pathlib import Path
import webbrowser
import os
from utils.license_manager import LicenseManager
from utils.telemetry import TelemetryManager
from utils.update_manager import UpdateManager

class AdminSettingsPopup(tk.Toplevel):
    def __init__(self, parent):
        super().__init__(parent)
        self.title("Administrator Settings")
        self.geometry("480x850") # Slightly wider and taller
        
        # Inherit standardized fonts from parent
        self.fonts = getattr(parent, 'fonts', {
            "h1": ("Segoe UI Variable Text", 13, "bold"),
            "h2": ("Segoe UI Variable Text", 11, "bold"),
            "normal": ("Segoe UI Variable Text", 10)
        })

        # Inherit stable config path from parent app
        if hasattr(parent, 'config_path'):
            self.config_path = parent.config_path
        else:
            self.config_path = Path("config.json")
            
        self.grab_set() # Modal

        
        self.authenticated = False
        self.build_auth_ui()

    def build_auth_ui(self):
        self.auth_frame = ttk.Frame(self, padding=40)
        self.auth_frame.pack(fill="both", expand=True)
        
        ttk.Label(self.auth_frame, text="[ AUTH ] 관리자 암호를 입력하세요", font=self.fonts["h2"]).pack(pady=20)
        
        self.pw_var = tk.StringVar()
        pw_entry = ttk.Entry(self.auth_frame, textvariable=self.pw_var, show="*", font=self.fonts["normal"], justify="center")
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
        # 1. Action Buttons Frame (Fixed at bottom)
        btn_frame = ttk.Frame(self, padding=10)
        btn_frame.pack(side="bottom", fill="x")
        ttk.Button(btn_frame, text="저장 및 닫기", command=self.save_and_close, style="Accent.TButton").pack(fill="x")

        # 2. Scrollable Content Area
        from ui.widgets.components import ScrollableFrame
        self.scroll_area = ScrollableFrame(self)
        self.scroll_area.pack(fill="both", expand=True)
        
        main = ttk.Frame(self.scroll_area.scrollable_frame, padding=25)
        main.pack(fill="both", expand=True)
        
        ttk.Label(main, text="[ CONFIG ] 기능 제한 설정 (EasyMatch Pro)", font=self.fonts["h2"]).pack(pady=(0, 20))
        
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
        
        # License Management Section
        ttk.Separator(main, orient="horizontal").pack(fill="x", pady=20)
        ttk.Label(main, text="[ LICENSE ] 라이센스 관리 (Key Generator)", font=self.fonts["h2"]).pack(pady=(0, 10))
        
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
        ttk.Label(main, text="[ MONITOR ] 원격 사용현황 모니터링 (Telemetry)", font=self.fonts["h2"]).pack(pady=(0, 10))
        
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

        # Registered Sources Section
        ttk.Separator(main, orient="horizontal").pack(fill="x", pady=20)
        ttk.Label(main, text="[ SOURCES ] 고정 소스 관리 (GitHub / Google)", font=self.fonts["h2"]).pack(pady=(0, 10))
        
        reg_frame = ttk.Frame(main)
        reg_frame.pack(fill="x")
        
        reg = self.config.get('registered_sources', {})
        
        ttk.Label(reg_frame, text="기본 GitHub URL:", font=("System", 9)).pack(anchor="w")
        self.reg_github = tk.StringVar(value=reg.get('github_url', ''))
        ttk.Entry(reg_frame, textvariable=self.reg_github).pack(fill="x", pady=2)
        
        ttk.Label(reg_frame, text="기본 구글 시트 URL:", font=("System", 9)).pack(anchor="w", pady=(5, 0))
        self.reg_google = tk.StringVar(value=reg.get('google_sheets_url', ''))
        ttk.Entry(reg_frame, textvariable=self.reg_google).pack(fill="x", pady=2)
        
        ttk.Label(reg_frame, text="GitHub Access Token (업로드용):", font=("System", 9, "bold")).pack(anchor="w", pady=(5, 0))
        self.reg_token = tk.StringVar(value=reg.get('github_token', ''))
        ttk.Entry(reg_frame, textvariable=self.reg_token, show="*").pack(fill="x", pady=2)
        
        ttk.Label(reg_frame, text="원격 프리셋(Preset) URL (Raw JSON):", font=("System", 9, "bold")).pack(anchor="w", pady=(5, 0))
        self.reg_presets = tk.StringVar(value=reg.get('remote_presets_url', ''))
        ttk.Entry(reg_frame, textvariable=self.reg_presets).pack(fill="x", pady=2)
        
        ttk.Label(reg_frame, text="원격 업데이트(Update) URL (Manifest JSON):", font=("System", 9, "bold")).pack(anchor="w", pady=(5, 0))
        self.reg_update = tk.StringVar(value=reg.get('remote_update_url', ''))
        ttk.Entry(reg_frame, textvariable=self.reg_update).pack(fill="x", pady=2)
        
        def force_sync_presets():
            import sys, os
            url = self.reg_presets.get().strip()
            if not url:
                messagebox.showwarning("입력 필요", "동기화할 URL이 비어있습니다.")
                return
            
            from utils.preset_manager import PresetManager
            # Determine presets path
            if getattr(sys, 'frozen', False):
                p = Path(os.path.dirname(sys.executable)) / "presets.json"
            else:
                # Go up to root from ui/widgets/
                p = Path(os.path.abspath(os.path.join(os.path.dirname(__file__), "../../presets.json")))
            
            pm = PresetManager(p)
            token = self.config.get('registered_sources', {}).get('github_token', '')
            success, msg, _count = pm.sync_from_remote(url, token)
            if success:
                messagebox.showinfo("동기화 성공", msg)
            else:
                messagebox.showerror("동기화 실패", msg)

        ttk.Button(reg_frame, text="지금 즉시 프리셋 동기화 (Force Sync)", command=force_sync_presets).pack(fill="x", pady=(5, 10))
        
        def manual_check_update():
            url = self.reg_update.get().strip()
            if not url:
                messagebox.showwarning("입력 필요", "업데이트 URL이 비어있습니다.")
                return
            
            manifest = UpdateManager.get_remote_manifest(url)
            if manifest:
                ver = manifest.get('version', 'unknown')
                notes = manifest.get('release_notes', '-')
                messagebox.showinfo("업데이트 확인", f"서버 최신 버전: {ver}\n\n[업데이트 내용]\n{notes}")
            else:
                messagebox.showerror("오류", "업데이트 정보를 가져오지 못했습니다.")

        ttk.Button(reg_frame, text="프로그램 업데이트 확인 (Check Version)", command=manual_check_update).pack(fill="x", pady=5)
        
        ttk.Label(main, text="기타 관리 설정", font=self.fonts["h2"]).pack(pady=(20, 10))
        
        # Manual Button
        ttk.Button(main, text="운영 매뉴얼 보기", command=self.open_manual).pack(fill="x", pady=5)
        
        # Theme Selection
        theme_frame = ttk.Frame(main)
        theme_frame.pack(fill="x", pady=5)
        ttk.Label(theme_frame, text="프로그램 테마 설정:", font=("System", 9)).pack(side="left")
        
        self.theme_var = tk.StringVar(value=self.config.get('branding', {}).get('theme', 'forest'))
        theme_options = ["royal", "forest", "arctic", "crimson", "dark", "light"]
        self.theme_combo = ttk.Combobox(theme_frame, textvariable=self.theme_var, values=theme_options, state="readonly", width=15)
        self.theme_combo.pack(side="left", padx=10)

        self.admin_contact = tk.StringVar(value=self.config.get('admin_contact', 'bough38@gmail.com'))
        ent_contact = ttk.Entry(main, textvariable=self.admin_contact)
        ent_contact.pack(fill="x", pady=2)
        
        # [ NETWORK ] Advanced Settings
        ttk.Separator(main, orient="horizontal").pack(fill="x", pady=20)
        ttk.Label(main, text="[ NETWORK ] 네트워크 고급 설정", font=self.fonts["h2"]).pack(pady=(0, 10))
        
        ttk.Label(main, text="HTTP/HTTPS 프록시 주소 (선택):", font=("System", 9)).pack(anchor="w")
        self.net_proxy = tk.StringVar(value=self.config.get('network', {}).get('proxy', ''))
        self.ent_proxy = ttk.Entry(main, textvariable=self.net_proxy)
        self.ent_proxy.pack(fill="x", pady=2)
        
        self.net_verify = tk.BooleanVar(value=self.config.get('network', {}).get('ssl_verify', True))
        ttk.Checkbutton(main, text="SSL 인증서 검증 활성화 (권장)", variable=self.net_verify).pack(anchor="w", pady=5)
        
        ttk.Label(main, text="※ 사내 보안망에서 업로드 오류 시 'SSL 검증'을 해제해 보세요.", 
                  font=("System", 8), foreground="gray").pack(anchor="w")

        # macOS clipboard fix for Toplevel entries
        import sys
        if sys.platform == "darwin":
            for entry in [ent_contact, self.theme_combo, self.reg_github, self.reg_google, self.reg_token, self.reg_presets, self.reg_update, self.tel_url, self.ent_proxy]:
                if hasattr(entry, 'bind'):
                    entry.bind("<Command-v>", lambda e: e.widget.event_generate("<<Paste>>"))
                    entry.bind("<Command-c>", lambda e: e.widget.event_generate("<<Copy>>"))
                    entry.bind("<Command-a>", lambda e: (e.widget.selection_range(0, 'end'), e.widget.icursor('end')))

    def open_manual(self):
        """Open the local operational manual in the default browser."""
        try:
            # Look for manual.html in the current directory
            manual_path = os.path.join(os.getcwd(), "manual.html")
            if os.path.exists(manual_path):
                webbrowser.open(f"file://{manual_path}")
            else:
                messagebox.showerror("오류", "매뉴얼 파일(manual.html)을 찾을 수 없습니다.\n루트 폴더에 파일이 있는지 확인하세요.")
        except Exception as e:
            messagebox.showerror("오류", f"매뉴얼을 여는 중 오류가 발생했습니다: {e}")

    def save_and_close(self):
        for key, var in self.toggles.items():
            self.config['locked_features'][key] = var.get()
            
        self.config['telemetry'] = {
            "enabled": self.tel_enabled.get(),
            "url": self.tel_url.get().strip()
        }
        
        # Save Registered Sources (with automatic whitespace sanitation)
        self.config['registered_sources']['github_url'] = self.reg_github.get().strip().split('\n')[0].strip()
        self.config['registered_sources']['github_token'] = self.reg_token.get().strip().split('\n')[0].strip()
        self.config['registered_sources']['google_sheets_url'] = self.reg_google.get().strip().split('\n')[0].strip()
        self.config['registered_sources']['remote_presets_url'] = self.reg_presets.get().strip().split('\n')[0].strip()
        self.config['registered_sources']['remote_update_url'] = self.reg_update.get().strip().split('\n')[0].strip()
        
        # Save Admin Contact
        self.config['admin_contact'] = self.admin_contact.get().strip()
        
        # Save Network Settings
        if 'network' not in self.config: self.config['network'] = {}
        self.config['network']['proxy'] = self.net_proxy.get().strip()
        self.config['network']['ssl_verify'] = self.net_verify.get()
        
        # Save Theme
        if 'branding' not in self.config: self.config['branding'] = {}
        self.config['branding']['theme'] = self.theme_var.get()
            
        with open(self.config_path, 'w', encoding='utf-8') as f:
            json.dump(self.config, f, indent=4, ensure_ascii=False)
            
        # Update parent's config in memory immediately
        if hasattr(self.master, 'config'):
            self.master.config.update(self.config)
            
        messagebox.showinfo("완료", "설정이 저장 및 동기화되었습니다.")
        self.destroy()

