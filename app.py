import tkinter as tk
from tkinter import ttk, messagebox
import sys
import os
import json
import sv_ttk
import webbrowser
from pathlib import Path
from PIL import Image, ImageTk
import threading
from utils.update_manager import UpdateManager

# Windows High DPI Awareness Optimization (Per-Monitor V2)
if sys.platform == "win32":
    try:
        from ctypes import windll, c_int, byref
        # Per-Monitor DPI Aware V2 (PROCESS_PER_MONITOR_DPI_AWARE_V2 = -4)
        # This is the modern standard for Windows 10/11 for pixel-perfect rendering
        try:
            windll.user32.SetProcessDpiAwarenessContext(-4)
        except Exception:
            try:
                # Fallback to older Per-Monitor Aware (2)
                windll.shcore.SetProcessDpiAwareness(2)
            except Exception:
                try:
                    # Fallback to System DPI Aware (1)
                    windll.shcore.SetProcessDpiAwareness(1)
                except Exception:
                    # Final fallback
                    windll.user32.SetProcessDPIAware()
    except Exception:
        pass

# Add local directories to path to ensure imports work correctly
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
sys.path.append(BASE_DIR)

def get_resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

def get_config_path(filename):
    """ Get path for persistent user data (next to EXE/Script) """
    if getattr(sys, 'frozen', False):
        # Running as EXE
        base_path = os.path.dirname(sys.executable)
    else:
        # Running as script
        base_path = os.path.dirname(os.path.abspath(__file__))
    return Path(os.path.join(base_path, filename))

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
        
        # Professional Typography for Splash
        splash_font_main = ("Segoe UI", 32, "bold") if sys.platform == "win32" else ("Helvetica", 32, "bold")
        splash_font_sub = ("Segoe UI", 10) if sys.platform == "win32" else ("Helvetica", 10)
        splash_accent = "#00E5FF"
        splash_bg = "#001F3F"

        try:
            logo_path = get_resource_path("assets/logo.png")
            logo_img = Image.open(logo_path)
            # Higher quality scaling
            logo_img = logo_img.resize((220, 220), Image.Resampling.LANCZOS)
            self.logo_photo = ImageTk.PhotoImage(logo_img)
            ttk.Label(self.container, image=self.logo_photo, background=splash_bg).pack(pady=(50, 10))
        except:
            ttk.Label(self.container, text="EM", font=("Segoe UI", 80, "bold"), foreground=splash_accent, background=splash_bg).pack(pady=40)
            
        ttk.Label(self.container, text="EasyMatch Pro", font=splash_font_main, foreground="white", background=splash_bg).pack()
        ttk.Label(self.container, text="Powered by Advanced Data IQ", font=splash_font_sub, foreground=splash_accent, background=splash_bg).pack(pady=(5, 20))
        
        self.progress = ttk.Progressbar(self.container, mode='determinate', length=420)
        self.progress.pack(pady=10)
        
        self.status_label = ttk.Label(self.container, text="Initializing components...", font=splash_font_sub, foreground="gray", background=splash_bg)
        self.status_label.pack()

        self.progress_val = 0

    def start_loading(self):
        if self.progress_val < 100:
            self.progress_val += 4
            self.progress['value'] = self.progress_val
            
            # Dynamic status text
            if self.progress_val < 30: self.status_label.config(text="Loading system resources...")
            elif self.progress_val < 60: self.status_label.config(text="Linking data engines...")
            elif self.progress_val < 90: self.status_label.config(text="Optimizing workspace...")
            else: self.status_label.config(text="Ready to launch!")
            
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
        self.config_path = get_config_path("config.json")
        self.load_config()
        
        self.title(f"{self.config['branding']['name']} {self.config['branding']['version']}")
        self.apply_dpi_scaling()
        self.optimize_window_geometry()
        
        # Apply Theme Immediately (Synchronous for correct first-render on Windows)
        self.apply_theme(self.config['branding'].get('theme', 'dark'))
        
        # Set Window Icon
        try:
            icon_path = get_resource_path("assets/logo.png")
            if os.path.exists(icon_path):
                img = Image.open(icon_path)
                self.icon_photo = ImageTk.PhotoImage(img)
                self.iconphoto(True, self.icon_photo)
        except Exception:
            pass

        # Force redraw to ensure theme engine is ready
        self.update_idletasks()
        
        # macOS Clipboard/Selection Shortcuts Fix
        if sys.platform == "darwin":
            self.bind_all("<Command-v>", lambda e: e.widget.event_generate("<<Paste>>"))
            self.bind_all("<Command-c>", lambda e: e.widget.event_generate("<<Copy>>"))
            self.bind_all("<Command-x>", lambda e: e.widget.event_generate("<<Cut>>"))
            self.bind_all("<Command-a>", lambda e: self._macos_select_all(e))

        # Start Splash
        SafeSplash(self.launch_main)

    def _macos_select_all(self, event):
        """Helper to handle Command-A for entries and text widgets on macOS."""
        widget = event.widget
        if isinstance(widget, (ttk.Entry, tk.Entry)):
            widget.selection_range(0, 'end')
            widget.icursor('end')
        elif isinstance(widget, tk.Text):
            widget.tag_add("sel", "1.0", "end")
        return "break"

    def load_config(self):
        if not self.config_path.exists():
            self.config = {
                "locked_features": {"batch_tab": False, "cleaner_tab": False, "cloud_source": False, "google_sheets": False},
                "branding": {"name": "EasyMatch", "version": "v3.9 Pro Official", "theme": "dark"},
                "telemetry": {"enabled": False, "url": ""},
                "registered_sources": {"github_url": "", "github_token": "", "google_sheets_url": "", "google_sheet_names": ""},
                "admin_contact": "bough38@gmail.com"
            }
        else:
            try:
                with open(self.config_path, 'r', encoding='utf-8') as f:
                    self.config = json.load(f)
                    # Migrating old configs
                    if 'registered_sources' not in self.config:
                        self.config['registered_sources'] = {"github_url": "", "github_token": "", "google_sheets_url": "", "google_sheet_names": ""}
                    if 'admin_contact' not in self.config:
                        self.config['admin_contact'] = "bough38@gmail.com"
            except:
                self.config = {
                    "locked_features": {"batch_tab": False, "cleaner_tab": False, "cloud_source": False, "google_sheets": False},
                    "branding": {"name": "EasyMatch", "version": "v3.9 Pro Official", "theme": "light"},
                    "telemetry": {"enabled": False, "url": ""},
                    "registered_sources": {"github_url": "", "github_token": "", "google_sheets_url": "", "google_sheet_names": ""}
                }


    def optimize_window_geometry(self):
        """Detect screen resolution and set optimal window size and position."""
        try:
            sw = self.winfo_screenwidth()
            sh = self.winfo_screenheight()
            
            # Calculate targets based on screen size
            w = int(sw * 0.8)
            h = int(sh * 0.8)
            
            # Refined Clamp limits for diverse screen resolutions (Laptop/Desktop)
            # Min: 1200x800 (Safer for smaller laptops), Max: 1600x950
            w = max(1200, min(w, 1600))
            h = max(750, min(h, 950))
            
            # Ensure it doesn't exceed actual screen resolution if very low (rare)
            w = min(w, sw - 40)
            h = min(h, sh - 80)
            
            # Center it
            x = (sw - w) // 2
            y = (sh - h) // 2
            
            self.geometry(f"{w}x{h}+{x}+{y}")
        except Exception:
            self.geometry("1400x900")

    def apply_dpi_scaling(self):
        """Forge Tkinter to match system DPI scaling for maximum sharpness on Windows."""
        # Standardize Font Constants for Premium Aesthetic
        # 'Segoe UI Variable' for Windows 11 sharpness, 'Inter' for SF-like cross-platform look
        is_win = sys.platform == "win32"
        self.FONT_FAMILY = "Segoe UI Variable Text" if is_win else "Inter"
        if not is_win and sys.platform == "darwin":
            self.FONT_FAMILY = ".AppleSystemUIFont"

        if is_win:
            try:
                from ctypes import windll
                h_dc = windll.user32.GetDC(0)
                dpi = windll.gdi32.GetDeviceCaps(h_dc, 88) # 88 is LOGPIXELSX
                windll.user32.ReleaseDC(0, h_dc)
                
                # Logic: Tkinter's point system is based on 72 DPI.
                # Setting scaling factor to DPI / 72.0 ensures 'points' scale perfectly to pixels.
                # We subtract a tiny correction factor on Windows to match the requested 'smaller' feel.
                self.scaling_factor = (dpi / 72.0) * 0.95 # Slight reduction for compactness
                self.tk.call('tk', 'scaling', self.scaling_factor)
            except Exception as e:
                self.scaling_factor = 1.33
                self.tk.call('tk', 'scaling', 1.33)
        else:
            self.scaling_factor = 1.0 # Mac handles it natively well
            
        # Define Global Semantic Fonts
        # Sizes are slightly reduced (e.g., 10 instead of 11) for 'Smaller' feel on Windows
        base_size = 10 if is_win else 13
        self.fonts = {
            "h1": (self.FONT_FAMILY, base_size + 3, "bold"),
            "h2": (self.FONT_FAMILY, base_size + 1, "bold"),
            "normal": (self.FONT_FAMILY, base_size),
            "small": (self.FONT_FAMILY, base_size - 1),
            "mono": ("Consolas" if is_win else "Menlo", base_size)
        }

    def apply_theme(self, theme_name):
        self.config['branding']['theme'] = theme_name
        
        # Base Theme (Standard sv-ttk as engine)
        base = "dark" if theme_name in ["dark", "cosmic", "graphite", "deep_ocean"] else "light"
        sv_ttk.set_theme(base)
        
        # Apply Custom Styling (Expert Techniques)
        style = ttk.Style()
        
        # Configure fonts globally without overriding sv-ttk colors
        style.configure(".", font=self.fonts["normal"])
        
        # Semantic Styles
        style.configure("Header.TLabel", font=self.fonts["h1"])
        if theme_name == "cosmic":
            style.configure("Header.TLabel", foreground="#00E5FF")
        elif theme_name == "graphite":
            style.configure("Header.TLabel", foreground="#BB86FC")
        elif theme_name == "deep_ocean":
            style.configure("Header.TLabel", foreground="#4FD1C5")
        elif theme_name == "light":
            style.configure("Header.TLabel", foreground="#007AFF")
        else:
            style.configure("Header.TLabel", foreground="#0078D4")
            
        style.configure("Accent.TButton", font=self.fonts["h2"])
        
        # Labelframe Label Fonts
        style.configure("TLabelframe.Label", font=self.fonts["h2"])
        
        # Luxury Two-Tone Header/Footer styling
        # 7 Distinct Theme Profiles (Expert Top 7)
        if theme_name == "light":
            header_bg = "#F8FAFC" # Brilliant bright gray
            header_fg = "#0F172A" # Almost black
            header_sub = "#2563EB" # Royal Blue
            footer_bg = "#F1F5F9"
            footer_fg = "#475569"
        elif theme_name == "cosmic":
            header_bg = "#2E1065" # Deep violet/purple
            header_fg = "#FDF8FF" # Bright white
            header_sub = "#D946EF" # Cyberpunk Pink/Magenta
            footer_bg = "#3B0764"
            footer_fg = "#E9D5FF"
        elif theme_name == "graphite":
            header_bg = "#334155" # Slate gray
            header_fg = "#F8FAFC"
            header_sub = "#F59E0B" # Premium Gold
            footer_bg = "#1E293B"
            footer_fg = "#94A3B8"
        elif theme_name == "dracula":
            header_bg = "#282A36" # Dracula Dark
            header_fg = "#F8F8F2" 
            header_sub = "#FF79C6" # Dracula Pink
            footer_bg = "#44475A"
            footer_fg = "#6272A4"
        elif theme_name == "nord":
            header_bg = "#2E3440" # Nord Polar Night
            header_fg = "#ECEFF4" # Nord Snow Storm
            header_sub = "#88C0D0" # Nord Frost
            footer_bg = "#3B4252"
            footer_fg = "#D8DEE9"
        elif theme_name == "oceanic":
            header_bg = "#082F49" # Deep Ocean Blue
            header_fg = "#F0F9FF"
            header_sub = "#2DD4BF" # Vibrant Teal
            footer_bg = "#0C4A6E"
            footer_fg = "#7DD3FC"
        else: # "dark"
            header_bg = "#18181B" # Midnight dark
            header_fg = "#FFFFFF"
            header_sub = "#00E5FF" # Neon Cyan
            footer_bg = "#18181B"
            footer_fg = "#A1A1AA"
            
        style.configure("Header.TFrame", background=header_bg)
        style.configure("Header.TLabel", background=header_bg, foreground=header_fg, font=self.fonts["h1"])
        style.configure("HeaderSub.TLabel", background=header_bg, foreground=header_sub, font=("System", 11, "bold"))
        
        style.configure("Footer.TFrame", background=footer_bg)
        style.configure("Footer.TLabel", background=footer_bg, foreground=footer_fg, font=("System", 9, "italic"))
        
        if hasattr(self, 'theme_menu'):
            try:
                self.theme_menu.config(bg=header_bg, fg=header_fg)
            except: pass
        
        # Expert Fine-tuning for Trees and Notebooks
        style.configure("Treeview", font=self.fonts["small"], rowheight=int(28 * self.scaling_factor))
        style.configure("TNotebook.Tab", font=self.fonts["normal"], padding=(10, 5))
        
        cfg_bg = "#FAFAFA" if base == "light" else "#1c1c1c"
        self.configure(bg=cfg_bg)
        
        # Propagate to all open Toplevel windows
        for child in self.winfo_children():
            if isinstance(child, tk.Toplevel):
                child.configure(bg=cfg_bg)
                # If the child has a .container or .main (common in our popups)
                for sub in child.winfo_children():
                    if isinstance(sub, (tk.Frame, ttk.Frame)):
                        try: sub.configure(style="TFrame")
                        except: pass

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
        
        # Apply theme colors to the Toplevel itself
        theme_name = self.config['branding'].get('theme', 'dark')
        palettes = {
            "dark": {"bg": "#1E1E1E"},
            "light": {"bg": "#F8F9FA"},
            "cosmic": {"bg": "#0F172A"},
            "graphite": {"bg": "#18181B"}
        }
        bg_color = palettes.get(theme_name, palettes['dark'])['bg']
        gate.configure(bg=bg_color)
        
        # UI
        main = ttk.Frame(gate, padding=30)
        main.pack(fill="both", expand=True)
        
        ttk.Label(main, text="기기 인증 및 라이센스 등록", font=("System", 14, "bold")).pack(pady=(0, 20))
        
        admin_email = self.config.get('admin_contact', 'bough38@gmail.com')
        ttk.Label(main, text=f"아래 ID를 복사하여 관리자({admin_email})에게 라이센스를 요청하세요.", font=("System", 9), foreground="#007ACC").pack(anchor="w")
        
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

        def open_mail():
            subject = "EasyMatch 라이센스 요청"
            body = f"안녕하세요, 관리자님.\n\nEasyMatch 라이센스 발급을 요청합니다.\n\n[기기 고유 ID]\n{mid}\n\n감사합니다."
            import urllib.parse
            encoded_body = urllib.parse.quote(body)
            encoded_subject = urllib.parse.quote(subject)
            webbrowser.open(f"mailto:{admin_email}?subject={encoded_subject}&body={encoded_body}")

        ttk.Button(main, text="관리자에게 메일 보내기 (자동 ID 포함)", command=open_mail).pack(fill="x", pady=(0, 10))

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
        
        # Auto-sync Presets
        self.trigger_auto_sync()
        
        # Check for Program Updates
        self.after(5000, self.check_for_program_updates) # Delay slightly for better startup experience


    def build_ui(self):
        # Header (Luxury Two-Tone)
        self.header_container = ttk.Frame(self, padding=(20, 10), style="Header.TFrame")
        self.header_container.pack(fill="x")
        
        # Spacer for centering
        spacer_l = ttk.Frame(self.header_container, style="Header.TFrame")
        spacer_l.pack(side="left", expand=True)
        
        header = ttk.Frame(self.header_container, style="Header.TFrame")
        header.pack(side="left")
        ttk.Label(header, text="이지매치", style="Header.TLabel").pack(side="left")
        ttk.Label(header, text="EasyMatch Pro", style="HeaderSub.TLabel").pack(side="left", padx=10, pady=(5,0))
        
        # Right aligned buttons
        btn_frame = ttk.Frame(self.header_container, style="Header.TFrame")
        btn_frame.pack(side="left", expand=True, anchor="e")
        
        # Theme Dropdown (Dynamic colors handled in apply_theme)
        self.theme_menu = tk.Menubutton(btn_frame, text="THEME ▼", relief="flat", font=("System", 9), direction="below")
        self.theme_menu.menu = tk.Menu(self.theme_menu, tearoff=0)
        self.theme_menu["menu"] = self.theme_menu.menu
        
        # 7 Expert Themes
        theme_options = ["Dark", "Light", "Cosmic", "Graphite", "Dracula", "Nord", "Oceanic"]
        for t in theme_options:
            self.theme_menu.menu.add_command(label=t, command=lambda x=t.lower(): self.apply_theme(x))
        
        self.theme_menu.pack(side="left", padx=5)
        ttk.Button(btn_frame, text="ADMIN", width=7, command=self.open_admin_settings).pack(side="left", padx=5)


        # Main Notebook
        self.notebook = ttk.Notebook(self)
        self.notebook.pack(fill="both", expand=True, padx=15, pady=10)

        # Tabs (Passing config for registered sources)
        self.tab_match = MatchTab(self.notebook, self.config, self.config_path)
        self.notebook.add(self.tab_match, text="  스마트 매칭  ")


        self.tab_stats = StatsTab(self.notebook)
        self.notebook.add(self.tab_stats, text="  데이터 인사이트  ")

        # Feature Locking
        if not self.config['locked_features'].get('batch_tab'):
            self.tab_batch = BatchTab(self.notebook, self.config)
            self.notebook.add(self.tab_batch, text="  일괄 처리(Batch)  ")

        if not self.config['locked_features'].get('cleaner_tab'):
            self.tab_cleaner = CleanerTab(self.notebook)
            self.notebook.add(self.tab_cleaner, text="  데이터 세척(Cleaner)  ")

        self.tab_match.register_on_load(self.on_data_loaded)

        # Footer & Pulse System (Two-Tone Footer)
        self.footer = ttk.Frame(self, padding=(20, 10), style="Footer.TFrame")
        self.footer.pack(fill="x", side="bottom")
        
        # Pulse (Satisfaction) Bar
        pulse_frame = ttk.Frame(self.footer, style="Footer.TFrame")
        pulse_frame.pack(side="left")
        
        ttk.Label(pulse_frame, text="EasyMatch Pulse:", style="Footer.TLabel").pack(side="left", padx=(0, 10))
        
        self.like_btn = ttk.Button(pulse_frame, text="LIKE", width=10, command=self.send_pulse_like)
        self.like_btn.pack(side="left", padx=2)

    def trigger_auto_sync(self):
        url = self.config.get('registered_sources', {}).get('remote_presets_url', '')
        if not url: return
        
        def task():
            try:
                # We can use either tab's preset_manager since they point to same file
                # If BatchTab is locked, it won't exist. So we use MatchTab.
                token = self.config.get('registered_sources', {}).get('github_token', '')
                success, _msg, count = self.tab_match.preset_manager.sync_from_remote(url, token)
                if success and count > 0:
                    # Notify tabs to reload
                    self.after(0, self.tab_match.load_presets)
                    if hasattr(self, 'tab_batch'):
                        self.after(0, self.tab_batch.load_presets)
            except:
                pass 
        
        threading.Thread(target=task, daemon=True).start()

    def check_for_program_updates(self):
        url = self.config.get('registered_sources', {}).get('remote_update_url', '')
        if not url: return
        
        def task():
            try:
                manifest = UpdateManager.get_remote_manifest(url)
                if not manifest: return
                
                remote_ver = manifest.get('version', '')
                current_ver = self.config['branding'].get('version', '')
                
                if UpdateManager.is_newer(current_ver, remote_ver):
                    notes = manifest.get('release_notes', '새로운 기능 및 안정성 개선')
                    msg = f"새로운 버전({remote_ver})이 출시되었습니다.\n\n[업데이트 내용]\n{notes}\n\n지금 업데이트하시겠습니까?"
                    
                    if messagebox.askyesno("업데이트 알림", msg):
                        self.after(0, lambda: self.perform_full_update(manifest['download_url'], remote_ver))
            except Exception as e:
                print(f"Auto-update check failed: {e}")

        threading.Thread(target=task, daemon=True).start()

    def perform_full_update(self, download_url, new_version):
        # Create progress popup
        progress_win = tk.Toplevel(self)
        progress_win.title("업데이트 다운로드 중...")
        progress_win.geometry("400x150")
        progress_win.resizable(False, False)
        progress_win.grab_set()
        
        # Center progress win
        sw = self.winfo_screenwidth()
        sh = self.winfo_screenheight()
        progress_win.geometry(f"400x150+{sw//2-200}+{sh//2-75}")
        
        lbl = ttk.Label(progress_win, text=f"이지매치 {new_version} 다운로드 중...", padding=20)
        lbl.pack()
        
        prog_var = tk.DoubleVar()
        pb = ttk.Progressbar(progress_win, variable=prog_var, maximum=1.0, length=300)
        pb.pack(pady=10)

        def download_task():
            try:
                # Target path for new exe
                if getattr(sys, 'frozen', False):
                    current_exe = sys.executable
                    new_exe = current_exe.replace(".exe", "_new.exe")
                    if "_new" not in new_exe: new_exe = current_exe + ".new"
                else:
                    # Dev mode - just mock it or download to temp
                    current_exe = os.path.abspath(__file__)
                    new_exe = current_exe + ".new"

                success = UpdateManager.download_file(download_url, new_exe, lambda p: prog_var.set(p))
                
                if success:
                    progress_win.destroy()
                    if messagebox.showinfo("업데이트 준비 완료", "다운로드가 완료되었습니다. 프로그램을 종료하고 업데이트를 적용합니다."):
                        if sys.platform == "win32" and getattr(sys, 'frozen', False):
                            UpdateManager.apply_update_windows(current_exe, new_exe)
                            sys.exit(0)
                        else:
                            # Non-windows or non-frozen: just tell user
                            messagebox.showinfo("수동 적용 필요", f"업데이트 파일이 다운로드되었습니다: {new_exe}\n직접 교체해 주세요.")
                else:
                    progress_win.destroy()
                    messagebox.showerror("오류", "다운로드 중 오류가 발생했습니다.")
            except Exception as e:
                progress_win.destroy()
                messagebox.showerror("오류", str(e))

        threading.Thread(target=download_task, daemon=True).start()
        
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
