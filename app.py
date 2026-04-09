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

class AnimatedSplash(tk.Toplevel):
    def __init__(self, callback):
        super().__init__()
        self.callback = callback
        self.title("EasyMatch Loading...")
        self.geometry("600x400")
        self.overrideredirect(True) # No border
        
        # Center screen
        sw = self.winfo_screenwidth()
        sh = self.winfo_screenheight()
        self.geometry(f"600x400+{sw//2-300}+{sh//2-200}")
        
        self.attributes("-topmost", True)
        self.setup_ui()
        
        # Animation
        self.alpha = 0.0
        self.attributes("-alpha", self.alpha)
        self.fade_in()

    def setup_ui(self):
        self.container = tk.Frame(self, bg="#001F3F", highlightthickness=2, highlightbackground="#00E5FF")
        self.container.pack(fill="both", expand=True)
        
        # Logo
        try:
            logo_img = Image.open("assets/logo.png")
            logo_img = logo_img.resize((200, 200), Image.Resampling.LANCZOS)
            self.logo_photo = ImageTk.PhotoImage(logo_img)
            ttk.Label(self.container, image=self.logo_photo, background="#001F3F").pack(pady=(40, 10))
        except:
            ttk.Label(self.container, text="🔍", font=("Segoe UI", 80), foreground="#00E5FF", background="#001F3F").pack(pady=40)
            
        ttk.Label(self.container, text="EasyMatch", font=("Segoe UI", 32, "bold"), foreground="white", background="#001F3F").pack()
        ttk.Label(self.container, text="Powered by Advanced Data IQ", font=("Segoe UI", 10), foreground="#00E5FF", background="#001F3F").pack(pady=10)
        
        self.progress = ttk.Progressbar(self.container, mode='determinate', length=400)
        self.progress.pack(pady=20)
        self.progress_val = 0

    def fade_in(self):
        if self.alpha < 1.0:
            self.alpha += 0.05
            self.attributes("-alpha", self.alpha)
            self.progress_val += 5
            self.progress['value'] = self.progress_val
            self.after(30, self.fade_in)
        else:
            self.after(1500, self.fade_out)

    def fade_out(self):
        if self.alpha > 0.0:
            self.alpha -= 0.05
            self.attributes("-alpha", self.alpha)
            self.after(30, self.fade_out)
        else:
            self.destroy()
            self.callback()

class EasyMatchPro(tk.Tk):
    def __init__(self):
        super().__init__()
        self.withdraw() # Hide until splash finish
        self.config_path = Path("config.json")
        self.load_config()
        
        self.title(f"{self.config['branding']['name']} {self.config['branding']['version']}")
        self.geometry("1450x900")
        
        self.setup_theme()
        
        # Start Splash
        Splash = AnimatedSplash(self.launch_main)

    def load_config(self):
        if not self.config_path.exists():
            self.config = {
                "locked_features": {"batch_tab": False, "cleaner_tab": False, "cloud_source": False, "google_sheets": False},
                "branding": {"name": "EasyMatch", "version": "v3.5 Pro", "theme": "dark"}
            }
        else:
            with open(self.config_path, 'r', encoding='utf-8') as f:
                self.config = json.load(f)

    def setup_theme(self):
        sv_ttk.set_theme(self.config['branding'].get('theme', 'dark'))
        style = ttk.Style(self)
        style.configure("Header.TLabel", font=("Segoe UI", 18, "bold"))
        style.configure("Accent.TButton", font=("Segoe UI", 10, "bold"), foreground="white")

    def launch_main(self):
        self.deiconify()
        self.build_ui()

    def build_ui(self):
        # Header
        header_container = ttk.Frame(self, padding=(20, 10))
        header_container.pack(fill="x")
        
        header = ttk.Frame(header_container)
        header.pack(side="left")
        ttk.Label(header, text="이지매치", style="Header.TLabel").pack(side="left")
        ttk.Label(header, text="EasyMatch Pro", foreground="#00E5FF", font=("Segoe UI", 10, "bold")).pack(side="left", padx=10, pady=(5,0))
        
        # Settings & Theme Buttons
        btn_frame = ttk.Frame(header_container)
        btn_frame.pack(side="right")
        
        ttk.Button(btn_frame, text="🌓", width=3, command=sv_ttk.toggle_theme).pack(side="left", padx=5)
        ttk.Button(btn_frame, text="⚙️", width=3, command=self.open_admin_settings).pack(side="right", padx=5)

        # Main Notebook
        self.notebook = ttk.Notebook(self)
        self.notebook.pack(fill="both", expand=True, padx=15, pady=10)

        # Tabs
        self.tab_match = MatchTab(self.notebook)
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

        # Footer
        footer = ttk.Frame(self, padding=10)
        footer.pack(fill="x", side="bottom")
        ttk.Label(footer, text="EasyMatch Professional Edition | All Rights Reserved", 
                  font=("Segoe UI", 8), foreground="gray").pack(side="right")

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
