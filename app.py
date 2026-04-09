import tkinter as tk
from tkinter import ttk
import sys
import os

# Add local directories to path to ensure imports work correctly
sys.path.append(os.path.dirname(os.path.abspath(__file__)))

from ui.tabs.match_tab import MatchTab
from ui.tabs.batch_tab import BatchTab
from ui.tabs.cleaner_tab import CleanerTab

class EnterpriseDataSuite(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("다목적 엑셀 데이터 매니지먼트 스위트 v.2.0")
        self.geometry("1400x900")
        
        self.setup_theme()
        self.build_ui()

    def setup_theme(self):
        style = ttk.Style(self)
        
        # Try to use a modern theme if available
        try:
            if sys.platform == "darwin":
                style.theme_use("aqua")
            else:
                style.theme_use("vista")
        except:
            pass

        # Custom Styles
        style.configure("Accent.TButton", font=("Segoe UI", 10, "bold"))
        style.configure("Header.TLabel", font=("Segoe UI", 14, "bold"))

    def build_ui(self):
        # Header
        header = ttk.Frame(self, padding=15)
        header.pack(fill="x")
        ttk.Label(header, text="Enterprise Data Management Suite", style="Header.TLabel").pack(side="left")
        ttk.Label(header, text="Built for Efficiency", foreground="gray", font=("Segoe UI", 9, "italic")).pack(side="left", padx=15, pady=5)

        # Tab Control (Notebook)
        self.notebook = ttk.Notebook(self)
        self.notebook.pack(fill="both", expand=True, padx=10, pady=10)

        # Tab 1: Smart Match & Extract
        self.tab_match = MatchTab(self.notebook)
        self.notebook.add(self.tab_match, text="  스마트 매칭 & 추출  ")

        # Tab 2: Batch Processing
        self.tab_batch = BatchTab(self.notebook)
        self.notebook.add(self.tab_batch, text="  폴더 일괄 처리 (Batch)  ")

        # Tab 3: Data Quality & Cleaner
        self.tab_cleaner = CleanerTab(self.notebook)
        self.notebook.add(self.tab_cleaner, text="  데이터 품질 관리  ")

        # Footer / Status bar
        footer = ttk.Frame(self, relief="flat", padding=5)
        footer.pack(fill="x", side="bottom")
        ttk.Label(footer, text="© 2024 Advanced Data Extraction Tool | Logic Powered by Pandas & Xlwings", font=("Segoe UI", 8), foreground="gray").pack(side="right")

if __name__ == "__main__":
    import logging
    import traceback
    
    # Configure logging
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
        
        # Show a fallback error message if Tkinter fails
        try:
            from tkinter import messagebox
            messagebox.showerror("시스템 오류", f"프로그램 실행 중 치명적인 오류가 발생했습니다.\n자세한 내용은 'data_suite_debug.log' 파일을 확인해 주세요.\n\n오류 내용: {e}")
        except:
            print(f"CRITICAL ERROR:\n{error_msg}")
