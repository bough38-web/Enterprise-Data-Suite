import tkinter as tk
from tkinter import ttk
import sys
import os
import sv_ttk

# Add local directories to path to ensure imports work correctly
sys.path.append(os.path.dirname(os.path.abspath(__file__)))

from ui.tabs.match_tab import MatchTab
from ui.tabs.batch_tab import BatchTab
from ui.tabs.cleaner_tab import CleanerTab
from ui.tabs.stats_tab import StatsTab

class EnterpriseDataSuite(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Enterprise Data Management Suite v.3.0 Pro")
        self.geometry("1400x950")
        
        self.setup_theme()
        self.build_ui()

    def setup_theme(self):
        # Apply Sun Valley modern theme
        sv_ttk.set_theme("light")
        
        style = ttk.Style(self)
        style.configure("Header.TLabel", font=("Segoe UI", 16, "bold"))
        style.configure("Accent.TButton", font=("Segoe UI", 10, "bold"))

    def toggle_theme(self):
        sv_ttk.toggle_theme()

    def build_ui(self):
        # Header with Theme Toggle
        header_container = ttk.Frame(self, padding=(15, 10))
        header_container.pack(fill="x")
        
        header = ttk.Frame(header_container)
        header.pack(side="left")
        ttk.Label(header, text="Enterprise Data Suite", style="Header.TLabel").pack(side="left")
        ttk.Label(header, text="v3.0 Pro", foreground="#0078D4", font=("Segoe UI", 10, "bold")).pack(side="left", padx=5)
        
        theme_btn = ttk.Button(header_container, text="🌓 테마 전환 (Light/Dark)", command=self.toggle_theme)
        theme_btn.pack(side="right")

        # Tab Control (Notebook)
        self.notebook = ttk.Notebook(self)
        self.notebook.pack(fill="both", expand=True, padx=10, pady=10)

        # Tab 1: Smart Match & Extract
        self.tab_match = MatchTab(self.notebook)
        self.notebook.add(self.tab_match, text="  스마트 매칭 & 추출  ")

        # Tab 2: Data Insights (New)
        self.tab_stats = StatsTab(self.notebook)
        self.notebook.add(self.tab_stats, text="  데이터 인사이트 (통계)  ")

        # Tab 3: Batch Processing
        self.tab_batch = BatchTab(self.notebook)
        self.notebook.add(self.tab_batch, text="  폴더 일괄 처리 (Batch)  ")

        # Tab 4: Data Quality & Cleaner
        self.tab_cleaner = CleanerTab(self.notebook)
        self.notebook.add(self.tab_cleaner, text="  데이터 품질 관리  ")

        # Listen for data changes in MatchTab to update StatsTab
        self.tab_match.register_on_load(self.on_data_loaded)

        # Footer
        footer = ttk.Frame(self, relief="flat", padding=5)
        footer.pack(fill="x", side="bottom")
        ttk.Label(footer, text="© 2024 Advanced Data Extraction Tool | Optimized for 900k+ Rows", 
                  font=("Segoe UI", 8), foreground="gray").pack(side="right")

    def on_data_loaded(self, df):
        """When data is loaded in MatchTab, sync it to the StatsTab automatically."""
        if hasattr(self, 'tab_stats'):
            self.tab_stats.set_data(df)

if __name__ == "__main__":
    import logging
    import traceback
    
    logging.basicConfig(
        filename="data_suite_debug.log",
        level=logging.ERROR,
        format="%(asctime)s - %(levelname)s - %(message)s",
        encoding="utf-8"
    )

    try:
        app = EnterpriseDataSuite()
        app.mainloop()
    except Exception as e:
        error_msg = traceback.format_exc()
        logging.error(f"Critical Exception:\n{error_msg}")
        
        try:
            from tkinter import messagebox
            messagebox.showerror("시스템 오류", f"프로그램 실행 중 치명적인 오류가 발생했습니다.\n{e}")
        except:
            print(f"CRITICAL ERROR:\n{error_msg}")
