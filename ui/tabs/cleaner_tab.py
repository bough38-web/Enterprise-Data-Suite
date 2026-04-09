import tkinter as tk
from tkinter import ttk, messagebox
import pandas as pd
import threading
from utils.excel_io import ExcelHandler
from ui.widgets.components import create_help_btn

class CleanerTab(ttk.Frame):
    def __init__(self, parent):
        super().__init__(parent)
        self.df = None
        self.build_ui()

    def build_ui(self):
        main = ttk.Frame(self, padding=20)
        main.pack(fill="both", expand=True)

        ttk.Label(main, text="데이터 품질 관리 (Data Quality & Cleaning)", font=("Segoe UI", 12, "bold")).pack(anchor="w", pady=(0, 20))

        # Source Selection
        src_frame = ttk.LabelFrame(main, text="작업 데이터 소스", padding=15)
        src_frame.pack(fill="x", pady=(0, 20))
        
        ttk.Button(src_frame, text="활성 엑셀 데이터 가져오기", command=self.load_active).pack(side="left", padx=5)
        self.info_var = tk.StringVar(value="데이터를 불러오세요.")
        ttk.Label(src_frame, textvariable=self.info_var).pack(side="left", padx=20)

        # Cleaning Tools
        tools_frame = ttk.Frame(main)
        tools_frame.pack(fill="both", expand=True)

        # 1. Deduplication
        dedup_lf = ttk.LabelFrame(tools_frame, text="중복 제거", padding=15)
        dedup_lf.grid(row=0, column=0, sticky="nsew", padx=5, pady=5)
        create_help_btn(dedup_lf, "중복 제거 가이드", 
            "• 기준 컬럼: 중복 여부를 판단할 기둥이 되는 컬럼을 선택합니다.\n"
            "• 선택하지 않으면 행 전체가 동일한 경우에만 삭제됩니다.").place(relx=1.0, x=-5, y=-5, anchor="ne")
        
        ttk.Label(dedup_lf, text="기준 컬럼:").pack(anchor="w")
        self.dedup_col_combo = ttk.Combobox(dedup_lf, state="readonly")
        self.dedup_col_combo.pack(fill="x", pady=5)
        ttk.Button(dedup_lf, text="중복 행 삭제 실행", command=self.run_dedup).pack(fill="x", pady=5)

        # 2. String Standardizer
        std_lf = ttk.LabelFrame(tools_frame, text="텍스트 표준화", padding=15)
        std_lf.grid(row=0, column=1, sticky="nsew", padx=5, pady=5)
        create_help_btn(std_lf, "텍스트 정리 가이드", 
            "• 공백 제거: 모든 텍스트 컬럼의 앞뒤 빈 칸을 없앱니다.\n"
            "• 전화번호 정규화: '010-1234-5678'을 '01012345678'로 통일하여 매칭율을 높입니다.").place(relx=1.0, x=-5, y=-5, anchor="ne")
        
        self.trim_space = tk.BooleanVar(value=True)
        ttk.Checkbutton(std_lf, text="앞뒤 공백 제거 (Trim)", variable=self.trim_space).pack(anchor="w")
        
        self.clean_tel = tk.BooleanVar(value=False)
        ttk.Checkbutton(std_lf, text="전화번호 형식 정규화 (숫자만 추출)", variable=self.clean_tel).pack(anchor="w")
        
        ttk.Button(std_lf, text="텍스트 정리 실행", command=self.run_text_std).pack(fill="x", pady=5)

        # 3. Missing Values
        nan_lf = ttk.LabelFrame(tools_frame, text="결측치(NaN) 처리", padding=15)
        nan_lf.grid(row=1, column=0, sticky="nsew", padx=5, pady=5)
        
        ttk.Label(nan_lf, text="빈 칸을 다음으로 채우기:").pack(anchor="w")
        self.nan_fill_var = tk.StringVar(value="0")
        ttk.Entry(nan_lf, textvariable=self.nan_fill_var).pack(fill="x", pady=5)
        ttk.Button(nan_lf, text="결측치 채우기 실행", command=self.run_nan_fill).pack(fill="x", pady=5)

        # Result Action
        ttk.Button(main, text="정리된 데이터를 엑셀로 내보내기", command=self.export_result, style="Accent.TButton").pack(pady=20, ipadx=20)

    def load_active(self):
        try:
            app = ExcelHandler.detect_special_sheets(None)[0] # Just take first sheet
            self.df = app.used_range.options(pd.DataFrame, header=1, index=False).value
            cols = self.df.columns.tolist()
            self.dedup_col_combo['values'] = cols
            self.info_var.set(f"데이터 로드됨 ({len(self.df)}행)")
        except Exception as e:
            messagebox.showerror("오류", f"데이터 로드 실패: {e}")

    def run_dedup(self):
        if self.df is None: return
        col = self.dedup_col_combo.get()
        before = len(self.df)
        if col:
            self.df = self.df.drop_duplicates(subset=[col])
        else:
            self.df = self.df.drop_duplicates()
        after = len(self.df)
        messagebox.showinfo("완료", f"중복 제거 완료: {before} -> {after} 행")
        self.info_var.set(f"데이터 상태 ({after}행)")

    def run_text_std(self):
        if self.df is None: return
        for col in self.df.select_dtypes(include=['object']).columns:
            if self.trim_space.get():
                self.df[col] = self.df[col].astype(str).str.strip()
            if self.clean_tel.get():
                if '전화' in col or '휴대폰' in col or '연락처' in col:
                    self.df[col] = self.df[col].astype(str).str.replace(r'[^0-9]', '', regex=True)
        messagebox.showinfo("완료", "텍스트 표준화 처리가 완료되었습니다.")

    def run_nan_fill(self):
        if self.df is None: return
        val = self.nan_fill_var.get()
        self.df = self.df.fillna(val)
        messagebox.showinfo("완료", f"결측치를 '{val}'(으)로 채웠습니다.")

    def export_result(self):
        if self.df is None: return
        try:
            sheet_name = ExcelHandler.write_to_active_excel(self.df, "데이터정리")
            messagebox.showinfo("성공", f"데이터를 엑셀로 내보냈습니다. (시트: {sheet_name})")
        except Exception as e:
            messagebox.showerror("오류", f"내보내기 실패: {e}")
