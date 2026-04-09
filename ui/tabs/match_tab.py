import tkinter as tk
from tkinter import ttk, messagebox, filedialog, simpledialog
import pandas as pd
from pathlib import Path
import json
import threading
import os

from utils.excel_io import ExcelHandler
from utils.data_engine import DataEngine
from ui.widgets.components import ScrollableFrame, ValueFilterPopup, SheetSelectPopup, create_help_btn

class MatchTab(ttk.Frame):
    def __init__(self, parent):
        super().__init__(parent)
        self.parent = parent
        self.df_left = None
        self.df_right = None
        self.left_path = "ActiveExcel"
        self.right_path = "None"
        self.col_vars = {}
        self.active_filters = []
        self.presets_file = Path("presets.json")
        
        self.build_ui()
        self.load_presets_list()

    def build_ui(self):
        ctrl_frame = ttk.Frame(self, padding=10)
        ctrl_frame.pack(side="left", fill="y")
        
        # Data Source
        load_lf = ttk.LabelFrame(ctrl_frame, text="데이터 소스", padding=10)
        load_lf.pack(fill="x", pady=(0, 10))
        create_help_btn(load_lf, "데이터 소스 가이드", 
            "• 활성 엑셀: 현재 열려있는 엑셀 창에서 데이터를 가져옵니다.\n"
            "• 원본 파일: 추출 및 매칭을 할 주 데이터를 선택합니다.\n"
            "• 참조 파일: 매칭 시 기준이 되는 데이터를 선택합니다. (생략 가능)").place(relx=1.0, x=-5, y=-5, anchor="ne")
        
        ttk.Button(load_lf, text="활성 엑셀 연동", command=self.load_active).pack(fill="x", pady=2)
        ttk.Button(load_lf, text="원본 파일 열기", command=lambda: self.load_file('left')).pack(fill="x", pady=2)
        ttk.Button(load_lf, text="참조 파일 열기", command=lambda: self.load_file('right')).pack(fill="x", pady=2)
        
        self.info_var = tk.StringVar(value="대기 중...")
        ttk.Label(load_lf, textvariable=self.info_var, foreground="blue").pack(pady=5)

        # Options
        opt_lf = ttk.LabelFrame(ctrl_frame, text="추출 옵션", padding=10)
        opt_lf.pack(fill="x", pady=(0, 10))
        create_help_btn(opt_lf, "옵션 가이드", 
            "• 대상 필터: 시설/요금구분 컬럼이 '대상'인 행만 추출합니다.\n"
            "• 직접 파일 저장: 50만 행 이상의 대용량 작업 시 권장합니다. "
            "엑셀 창을 거치지 않고 파일로 즉시 저장하여 매우 빠르고 안전합니다.").place(relx=1.0, x=-5, y=-5, anchor="ne")
        
        self.mode_var = tk.StringVar(value="keep")
        ttk.Radiobutton(opt_lf, text="유지", variable=self.mode_var, value="keep").pack(side="left", padx=5)
        ttk.Radiobutton(opt_lf, text="삭제", variable=self.mode_var, value="delete").pack(side="left", padx=5)
        
        self.auto_target = tk.BooleanVar(value=True)
        ttk.Checkbutton(opt_lf, text="대상 필터", variable=self.auto_target).pack(side="left", padx=5)

        # Large Data Mode
        self.direct_save = tk.BooleanVar(value=False)
        self.ds_check = ttk.Checkbutton(opt_lf, text="직접 파일 저장 (대용량 권장)", variable=self.direct_save)
        self.ds_check.pack(fill="x", pady=5)

        # Multi Filter
        filter_lf = ttk.LabelFrame(ctrl_frame, text="필터링 조건 설정", padding=10)
        filter_lf.pack(fill="both", expand=True, pady=(0, 10))
        create_help_btn(filter_lf, "필터 가이드", 
            "• In: 선택한 값들이 포함된 행만 추출합니다.\n"
            "• Ex: 선택한 값들을 제외하고 추출합니다.\n"
            "• '추가' 버튼을 눌러 여러 개의 조건을 동시에 적용할 수 있습니다. (AND 조건)").place(relx=1.0, x=-5, y=-5, anchor="ne")
        
        f_top = ttk.Frame(filter_lf)
        f_top.pack(fill="x")
        self.filter_col_var = tk.StringVar()
        self.filter_combo = ttk.Combobox(f_top, textvariable=self.filter_col_var, state="readonly", width=15)
        self.filter_combo.pack(side="left", padx=2)
        
        self.filter_mode = tk.StringVar(value="include")
        ttk.Radiobutton(f_top, text="In", variable=self.filter_mode, value="include").pack(side="left")
        ttk.Radiobutton(f_top, text="Ex", variable=self.filter_mode, value="exclude").pack(side="left")
        
        ttk.Button(f_top, text="추가", command=self.add_filter_rule).pack(side="right", padx=2)
        
        self.filter_listbox = tk.Listbox(filter_lf, height=4)
        self.filter_listbox.pack(fill="both", expand=True, pady=5)
        
        f_btns = ttk.Frame(filter_lf)
        f_btns.pack(fill="x")
        ttk.Button(f_btns, text="삭제", command=self.remove_filter_rule).pack(side="left", padx=2)
        ttk.Button(f_btns, text="초기화", command=self.clear_all_filters).pack(side="left", padx=2)

        # Presets
        preset_lf = ttk.LabelFrame(ctrl_frame, text="프리셋", padding=10)
        preset_lf.pack(fill="x", pady=(0, 10))
        self.preset_list = ttk.Combobox(preset_lf, state="readonly")
        self.preset_list.pack(fill="x", pady=2)
        
        p_btns = ttk.Frame(preset_lf)
        p_btns.pack(fill="x")
        ttk.Button(p_btns, text="저장", command=self.save_preset).pack(side="left", expand=True, fill="x", padx=1)
        ttk.Button(p_btns, text="로드", command=self.load_preset).pack(side="left", expand=True, fill="x", padx=1)

        # Action
        self.run_btn = ttk.Button(ctrl_frame, text="추출 및 저장 실행", command=self.run_process)
        self.run_btn.pack(fill="x", pady=10)

        # Columns
        col_frame = ttk.LabelFrame(self, text="컬럼 선택", padding=10)
        col_frame.pack(side="right", fill="both", expand=True)
        self.scroll_frame = ScrollableFrame(col_frame)
        self.scroll_frame.pack(fill="both", expand=True)
        
        btn_bar = ttk.Frame(col_frame)
        btn_bar.pack(fill="x", pady=5)
        ttk.Button(btn_bar, text="전체 선택", command=self.select_all_cols).pack(side="left", padx=2)
        ttk.Button(btn_bar, text="전체 해제", command=self.unselect_all_cols).pack(side="left", padx=2)

    def load_active(self):
        try:
            self.set_info("연결 중...")
            left_ws, right_ws = ExcelHandler.detect_special_sheets(None)
            self.df_left = left_ws.used_range.options(pd.DataFrame, header=1, index=False).value
            self.left_path = f"ActiveSheet: {left_ws.name}"
            
            if right_ws and right_ws != left_ws:
                self.df_right = right_ws.used_range.options(pd.DataFrame, header=1, index=False).value
                self.right_path = f"ActiveSheet: {right_ws.name}"
            else:
                self.df_right = None
            
            self.check_size()
            self.refresh_cols()
            self.set_info(f"연동: {left_ws.name}")
        except Exception as e:
            messagebox.showerror("오류", str(e))

    def load_file(self, side):
        path = filedialog.askopenfilename(filetypes=[("Excel/CSV", "*.xlsx *.xls *.csv")])
        if not path: return
        try:
            sheets = ExcelHandler.get_sheet_names(path)
            sheet_name = None
            if len(sheets) > 1:
                popup = SheetSelectPopup(self, "시트 선택", sheets)
                if popup.result: sheet_name = popup.result
            
            df = ExcelHandler.read_file(path, sheet_name)
            if side == 'left': 
                self.df_left = df
                self.left_path = path
            else: 
                self.df_right = df
                self.right_path = path
            
            self.check_size()
            self.refresh_cols()
            self.set_info(f"{'원본' if side=='left' else '참조'} 로드 완료")
        except Exception as e:
            messagebox.showerror("오류", str(e))

    def check_size(self):
        if self.df_left is not None and len(self.df_left) > 500000:
            self.direct_save.set(True)
            messagebox.showinfo("알림", f"데이터가 매우 큽니다({len(self.df_left):,}행).\n안정적인 처리를 위해 '직접 파일 저장' 모드가 선택되었습니다.")

    def refresh_cols(self):
        if self.df_left is None: return
        for w in self.scroll_frame.scrollable_frame.winfo_children(): w.destroy()
        self.col_vars = {}
        cols = self.df_left.columns.tolist()
        self.filter_combo['values'] = cols
        for idx, col in enumerate(cols):
            var = tk.BooleanVar(value=True)
            self.col_vars[col] = var
            ttk.Checkbutton(self.scroll_frame.scrollable_frame, text=str(col), variable=var).grid(row=idx//3, column=idx%3, sticky="w", padx=10, pady=5)

    def add_filter_rule(self):
        col = self.filter_col_var.get()
        if not col or self.df_left is None: return
        vals = self.df_left[col].dropna().unique().tolist()
        prev = next((f['values'] for f in self.active_filters if f['column'] == col), [])
        popup = ValueFilterPopup(self, f"[{col}] 필터", vals, prev)
        if popup.result:
            self.active_filters = [f for f in self.active_filters if f['column'] != col]
            self.active_filters.append({'column': col, 'values': popup.result, 'mode': self.filter_mode.get()})
            self.update_filter_listbox()

    def remove_filter_rule(self):
        sel = self.filter_listbox.curselection()
        if not sel: return
        del self.active_filters[sel[0]]
        self.update_filter_listbox()

    def clear_all_filters(self):
        self.active_filters = []
        self.update_filter_listbox()

    def update_filter_listbox(self):
        self.filter_listbox.delete(0, tk.END)
        for f in self.active_filters:
            m = "포함" if f['mode'] == 'include' else "제외"
            v = f"{f['values'][0]}" if len(f['values']) == 1 else f"{f['values'][0]} 외 {len(f['values'])-1}건"
            self.filter_listbox.insert(tk.END, f"[{f['column']}] {m}: {v}")

    def save_preset(self):
        name = simpledialog.askstring("프리셋 저장", "이름:")
        if not name: return
        data = {"columns": [c for c, v in self.col_vars.items() if v.get()], "mode": self.mode_var.get(), "auto_target": self.auto_target.get(), "active_filters": self.active_filters}
        presets = {}
        if self.presets_file.exists():
            with open(self.presets_file, 'r', encoding='utf-8') as f: presets = json.load(f)
        presets[name] = data
        with open(self.presets_file, 'w', encoding='utf-8') as f: json.dump(presets, f, ensure_ascii=False, indent=4)
        self.load_presets_list()

    def load_preset(self):
        name = self.preset_list.get()
        if not name or not self.presets_file.exists(): return
        with open(self.presets_file, 'r', encoding='utf-8') as f:
            data = json.load(f).get(name)
            if not data: return
            self.mode_var.set(data.get("mode", "keep"))
            self.auto_target.set(data.get("auto_target", True))
            self.active_filters = data.get("active_filters", [])
            self.update_filter_listbox()
            sel = set(data.get("columns", []))
            for c, v in self.col_vars.items(): v.set(c in sel)

    def load_presets_list(self):
        if self.presets_file.exists():
            with open(self.presets_file, 'r', encoding='utf-8') as f:
                self.preset_list['values'] = list(json.load(f).keys())

    def run_process(self):
        if self.df_left is None: return
        
        def task():
            try:
                self.set_info("처리 중...")
                df_res = self.df_left.copy()
                
                if self.df_right is not None:
                    key = DataEngine.auto_find_key(self.df_left, self.df_right)
                    matches = DataEngine.auto_match_columns(self.df_left, self.df_right)
                    df_res = DataEngine.perform_matching(self.df_left, self.df_right, key, matches)
                
                df_res = DataEngine.apply_filters(df_res, {"auto_target": self.auto_target.get(), "custom_filters": self.active_filters})
                df_res = DataEngine.select_columns(df_res, [c for c, v in self.col_vars.items() if v.get()], self.mode_var.get())
                df_res = DataEngine.add_source_info(df_res, self.left_path)
                
                if self.direct_save.get():
                    out_path = filedialog.asksaveasfilename(defaultextension=".csv", filetypes=[("CSV", "*.csv"), ("Excel", "*.xlsx")])
                    if not out_path: return
                    ExcelHandler.save_to_file(df_res, out_path)
                    res_msg = f"파일 저장 완료: {os.path.basename(out_path)}"
                else:
                    sheet_name = ExcelHandler.write_to_active_excel(df_res, "추출결과")
                    res_msg = f"엑셀 완료 (시트: {sheet_name})"
                
                self.set_info("완료")
                messagebox.showinfo("성공", f"{res_msg}\n처리 행: {len(df_res):,}건")
            except Exception as e:
                messagebox.showerror("오류", str(e))
                self.set_info("실패")

        threading.Thread(target=task, daemon=True).start()

    def set_info(self, msg):
        self.info_var.set(msg)
        self.update_idletasks()

    def select_all_cols(self):
        for v in self.col_vars.values(): v.set(True)
    def unselect_all_cols(self):
        for v in self.col_vars.values(): v.set(False)
