import tkinter as tk
from tkinter import ttk, messagebox, filedialog, simpledialog
import pandas as pd
from pathlib import Path
import json
import threading
import os
import sys

from utils.excel_io import ExcelHandler
from utils.data_engine import DataEngine
from utils.telemetry import TelemetryManager
from ui.widgets.components import ScrollableFrame, ValueFilterPopup, SheetSelectPopup, create_help_btn

class MatchTab(ttk.Frame):
    def __init__(self, parent, config=None):
        super().__init__(parent)
        self.parent = parent
        self.config = config or {}

        self.df_left = None
        self.df_right = None
        self.left_path = "ActiveExcel"
        self.right_path = "None"
        self.col_vars = {}
        self.active_filters = []
        # Resolve stable path for presets.json (next to EXE/Script)
        if getattr(sys, 'frozen', False):
            base_path = os.path.dirname(sys.executable)
        else:
            base_path = os.path.dirname(os.path.abspath(__file__))
            # If in ui/tabs folder, go up to project root
            if "ui" in base_path:
                base_path = os.path.abspath(os.path.join(base_path, "../.."))
        
        self.presets_file = Path(os.path.join(base_path, "presets.json"))

        self.on_load_callback = None
        
        # Cloud Integration State
        self.cloud_headers = []
        
        self.build_ui()
        self.load_registered_sources()
        self.load_presets_list()


    def register_on_load(self, callback):
        """Register a function to be called whenever df_left is updated."""
        self.on_load_callback = callback

    def build_ui(self):
        # Main Layout: Sidebar (Left) and Data (Right)
        ctrl_frame = ttk.Frame(self, padding=0)
        ctrl_frame.pack(side="left", fill="y")

        # Use ScrollableFrame for the left sidebar to handle small screens
        self.sidebar = ScrollableFrame(ctrl_frame)
        self.sidebar.pack(fill="both", expand=True)
        self.sidebar.canvas.config(width=340) # Fixed width for consistency
        
        # All widgets now go to self.sidebar.scrollable_frame
        sf = self.sidebar.scrollable_frame
        p_frame = ttk.Frame(sf, padding=15)
        p_frame.pack(fill="both", expand=True)

        # 1. Data Source
        load_lf = ttk.LabelFrame(p_frame, text="데이터 소스", padding=10)
        load_lf.pack(fill="x", pady=(0, 10))
        create_help_btn(load_lf, "데이터 소스 가이드", 
            "- 활성 엑셀: 현재 열려있는 엑셀 창에서 데이터를 가져옵니다.\n"
            "- 원본 파일: 추출 및 매칭을 할 주 데이터를 선택합니다.\n"
            "- 참조 파일: 매칭 시 기준이 되는 데이터를 선택합니다. (생략 가능)").place(relx=1.0, x=-5, y=-5, anchor="ne")
        
        ttk.Button(load_lf, text="활성 엑셀 연동", command=self.load_active).pack(fill="x", pady=2)
        ttk.Button(load_lf, text="원본 파일 열기", command=lambda: self.load_file('left')).pack(fill="x", pady=2)
        ttk.Button(load_lf, text="참조 파일 열기", command=lambda: self.load_file('right')).pack(fill="x", pady=2)
        
        self.info_var = tk.StringVar(value="대기 중...")
        ttk.Label(load_lf, textvariable=self.info_var, foreground="#4A90E2").pack(pady=5)

        # 2. Cloud Source (GitHub Raw)
        cloud_lf = ttk.LabelFrame(p_frame, text="클라우드 소스 (GitHub Raw)", padding=10)
        cloud_lf.pack(fill="x", pady=(0, 10))
        
        cloud_title_frame = ttk.Frame(cloud_lf)
        cloud_title_frame.pack(fill="x")
        ttk.Label(cloud_title_frame, text="GitHub URL:").pack(side="left")
        ttk.Button(cloud_title_frame, text="저장", width=5, command=lambda: self.save_source_config('github')).pack(side="right")
        
        create_help_btn(cloud_lf, "클라우드 가이드", 
            "- GitHub Raw URL을 입력하세요.\n"
            "- '헤더 확인' 후 필요한 컬럼만 선택하여 다운로드할 수 있습니다.\n"
            "- Private 저장소라면 Token이 필요합니다.").place(relx=1.0, x=-5, y=-5, anchor="ne")
        
        self.cloud_url = tk.StringVar()
        self.cloud_token = tk.StringVar()
        
        ttk.Label(cloud_lf, text="URL:").pack(anchor="w")
        ttk.Entry(cloud_lf, textvariable=self.cloud_url).pack(fill="x", pady=2)
        
        ttk.Label(cloud_lf, text="Token:").pack(anchor="w")
        ttk.Entry(cloud_lf, textvariable=self.cloud_token, show="*").pack(fill="x", pady=2)
        
        c_btns = ttk.Frame(cloud_lf)
        c_btns.pack(fill="x", pady=5)
        ttk.Button(c_btns, text="헤더 확인", command=self.peek_cloud).pack(side="left", expand=True, fill="x", padx=2)
        ttk.Button(c_btns, text="선택 다운로드", command=self.download_cloud).pack(side="left", expand=True, fill="x", padx=2)

        # 3. Google Sheets Source
        gs_lf = ttk.LabelFrame(p_frame, text="구글 스프레드시트 연동", padding=10)
        gs_lf.pack(fill="x", pady=(0, 10))
        
        gs_title_frame = ttk.Frame(gs_lf)
        gs_title_frame.pack(fill="x")
        ttk.Label(gs_title_frame, text="시트 주소:").pack(side="left")
        ttk.Button(gs_title_frame, text="저장", width=5, command=lambda: self.save_source_config('google')).pack(side="right")
        
        create_help_btn(gs_lf, "구글 시트 가이드", 
            "- 시트 주소를 입력하세요.\n"
            "- 특정 시트(탭)를 여러 개 불러오려면 쉼표(,)로 구분하세요.").place(relx=1.0, x=-5, y=-5, anchor="ne")
        
        self.gs_url = tk.StringVar()
        self.gs_sheet_names = tk.StringVar()
        
        ttk.Label(gs_lf, text="시트 주소:").pack(anchor="w")
        ttk.Entry(gs_lf, textvariable=self.gs_url).pack(fill="x", pady=2)
        
        ttk.Label(gs_lf, text="탭 이름 (쉼표 구분):").pack(anchor="w")
        # Fixed the truncated line from previous edit
        ttk.Entry(gs_lf, textvariable=self.gs_sheet_names).pack(fill="x", pady=2)
        
        gs_btns = ttk.Frame(gs_lf)
        gs_btns.pack(fill="x", pady=5)
        ttk.Button(gs_btns, text="헤더 확인", command=self.peek_google).pack(side="left", expand=True, fill="x", padx=2)
        ttk.Button(gs_btns, text="선택 다운로드", command=self.download_google).pack(side="left", expand=True, fill="x", padx=2)

        # 4. Options
        opt_lf = ttk.LabelFrame(p_frame, text="추출 옵션", padding=10)
        opt_lf.pack(fill="x", pady=(0, 10))
        
        self.mode_var = tk.StringVar(value="keep")
        ttk.Radiobutton(opt_lf, text="유지", variable=self.mode_var, value="keep").pack(side="left", padx=5)
        ttk.Radiobutton(opt_lf, text="삭제", variable=self.mode_var, value="delete").pack(side="left", padx=5)
        
        self.auto_target = tk.BooleanVar(value=True)
        ttk.Checkbutton(opt_lf, text="대상 필터", variable=self.auto_target).pack(side="left", padx=5)

        self.fuzzy_match = tk.BooleanVar(value=False)
        ttk.Checkbutton(opt_lf, text="유사도", variable=self.fuzzy_match).pack(side="left", padx=5)

        self.direct_save = tk.BooleanVar(value=False)
        ttk.Checkbutton(opt_lf, text="직접 파일 저장", variable=self.direct_save).pack(fill="x", pady=5)

        # 5. Multi Filter
        filter_lf = ttk.LabelFrame(p_frame, text="필터링 조건 설정", padding=10)
        filter_lf.pack(fill="x", pady=(0, 10))
        
        f_top = ttk.Frame(filter_lf)
        f_top.pack(fill="x")
        self.filter_col_var = tk.StringVar()
        self.filter_combo = ttk.Combobox(f_top, textvariable=self.filter_col_var, state="readonly", width=12)
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

        # 6. Presets
        pre_lf = ttk.LabelFrame(p_frame, text="추출 프리셋", padding=10)
        pre_lf.pack(fill="x", pady=(0, 10))
        
        pre_top = ttk.Frame(pre_lf)
        pre_top.pack(fill="x")
        self.preset_var = tk.StringVar()
        self.preset_combo = ttk.Combobox(pre_top, textvariable=self.preset_var, state="readonly")
        self.preset_combo.pack(side="left", expand=True, fill="x", padx=(0, 5))
        self.preset_combo.bind("<<ComboboxSelected>>", self.apply_preset)
        
        ttk.Button(pre_top, text="적용", command=self.apply_preset, width=5).pack(side="right")

        pre_btns = ttk.Frame(pre_lf)
        pre_btns.pack(fill="x", pady=(5, 0))
        ttk.Button(pre_btns, text="+ 저장", command=self.save_current_preset).pack(side="left", expand=True, fill="x", padx=2)
        ttk.Button(pre_btns, text="관리", command=self.manage_presets_ui).pack(side="left", expand=True, fill="x", padx=2)

        # 7. Execution Button
        self.run_btn = ttk.Button(p_frame, text="추출 및 저장 실행", command=self.run_process, style="Accent.TButton")
        self.run_btn.pack(fill="x", pady=(10, 20))

        # 8. Column Selection (Right Pane - Not in scrollable area)
        col_frame = ttk.LabelFrame(self, text="컬럼 선택", padding=10)
        col_frame.pack(side="right", fill="both", expand=True, padx=(10, 0))
        
        self.scroll_frame = ScrollableFrame(col_frame)
        self.scroll_frame.pack(fill="both", expand=True)
        
        btn_bar = ttk.Frame(col_frame)
        btn_bar.pack(fill="x", pady=5)
        ttk.Button(btn_bar, text="전체 선택", command=self.select_all_cols).pack(side="left", padx=2)
        ttk.Button(btn_bar, text="전체 해제", command=self.unselect_all_cols).pack(side="left", padx=2)

    def peek_cloud(self):

        url = self.cloud_url.get().strip()
        if not url: return
        try:
            self.set_info("헤더 확인 중...")
            self.cloud_headers = ExcelHandler.peek_headers_from_url(url, self.cloud_token.get())
            
            # Show headers in the column selection area immediately
            for w in self.scroll_frame.scrollable_frame.winfo_children(): w.destroy()
            self.col_vars = {}
            for idx, col in enumerate(self.cloud_headers):
                var = tk.BooleanVar(value=True)
                self.col_vars[col] = var
                ttk.Checkbutton(self.scroll_frame.scrollable_frame, text=str(col), variable=var).grid(row=idx//4, column=idx%4, sticky="w", padx=15, pady=6)

            
            self.filter_combo['values'] = self.cloud_headers
            self.set_info("클라우드 헤더 로드 완료")
        except Exception as e:
            messagebox.showerror("클라우드 오류", f"헤더를 읽을 수 없습니다: {e}")
            self.set_info("실패")

    def download_cloud(self):
        url = self.cloud_url.get().strip()
        if not url: return
        selected_cols = [c for c, v in self.col_vars.items() if v.get()]
        if not selected_cols:
            messagebox.showwarning("경고", "다운로드할 컬럼을 선택하세요.")
            return

        def task():
            try:
                self.set_info("클라우드 다운로드 중...")
                df = ExcelHandler.read_from_url(url, usecols=selected_cols, token=self.cloud_token.get())
                self.df_left = df
                self.left_path = f"Cloud: {url.split('/')[-1]}"
                self.check_size()
                if self.on_load_callback: self.on_load_callback(self.df_left)
                self.set_info(f"다운로드 완료 ({len(df):,}행)")
            except Exception as e:
                messagebox.showerror("다운로드 오류", str(e))
                self.set_info("실패")

        threading.Thread(target=task, daemon=True).start()

    def load_active(self):
        try:
            self.set_info("연결 중...")
            import xlwings as xw
            app = xw.apps.active
            if not app:
                raise Exception("실행 중인 엑셀이 없습니다.")
            wb = app.books.active
            left_ws, right_ws = ExcelHandler.detect_special_sheets(wb)

            self.df_left = left_ws.used_range.options(pd.DataFrame, header=1, index=False).value
            self.left_path = f"ActiveSheet: {left_ws.name}"
            
            if right_ws and right_ws != left_ws:
                self.df_right = right_ws.used_range.options(pd.DataFrame, header=1, index=False).value
                self.right_path = f"ActiveSheet: {right_ws.name}"
            else:
                self.df_right = None
            
            self.check_size()
            self.refresh_cols()
            if self.on_load_callback: self.on_load_callback(self.df_left)
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
            if self.on_load_callback: self.on_load_callback(self.df_left)
            self.set_info(f"{'원본' if side=='left' else '참조'} 로드 완료")
        except Exception as e:
            messagebox.showerror("오류", str(e))

    def check_size(self):
        if self.df_left is not None and len(self.df_left) > 500000:
            self.direct_save.set(True)

    def refresh_cols(self):
        if self.df_left is None: return
        for w in self.scroll_frame.scrollable_frame.winfo_children(): w.destroy()
        self.col_vars = {}
        cols = self.df_left.columns.tolist()
        self.filter_combo['values'] = cols
        for idx, col in enumerate(cols):
            var = tk.BooleanVar(value=True)
            self.col_vars[col] = var
            ttk.Checkbutton(self.scroll_frame.scrollable_frame, text=str(col), variable=var).grid(row=idx//4, column=idx%4, sticky="w", padx=15, pady=6)


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

    def load_presets_list(self):
        if self.presets_file.exists():
            with open(self.presets_file, 'r', encoding='utf-8') as f:
                self.preset_combo['values'] = list(json.load(f).keys())

    def save_current_preset(self):
        name = simpledialog.askstring("프리셋 저장", "새 프리셋 이름을 입력하세요:")
        if not name: return
        
        # Capture state
        current_cols = [c for c, v in self.col_vars.items() if v.get()]
        preset_data = {
            "columns": current_cols,
            "filters": self.active_filters,
            "mode": self.mode_var.get(),
            "auto_target": self.auto_target.get()
        }
        
        presets = {}
        if self.presets_file.exists():
            with open(self.presets_file, 'r', encoding='utf-8') as f:
                presets = json.load(f)
        
        presets[name] = preset_data
        with open(self.presets_file, 'w', encoding='utf-8') as f:
            json.dump(presets, f, indent=4, ensure_ascii=False)
            
        self.load_presets_list()
        messagebox.showinfo("완료", f"'{name}' 프리셋이 저장되었습니다.")

    def apply_preset(self, event=None):
        name = self.preset_var.get()
        if not name or not self.presets_file.exists(): return
        
        with open(self.presets_file, 'r', encoding='utf-8') as f:
            presets = json.load(f)
            p = presets.get(name)
            if not p: return
            
            # Apply Columns
            for col, var in self.col_vars.items():
                var.set(col in p.get('columns', []))
            
            # Apply Filters
            self.active_filters = p.get('filters', [])
            self.update_filter_listbox()
            self.mode_var.set(p.get('mode', 'keep'))
            self.auto_target.set(p.get('auto_target', True))
            
        self.set_info(f"프리셋 적용: {name}")

    def manage_presets_ui(self):
        if not self.presets_file.exists(): return
        
        popup = tk.Toplevel(self)
        popup.title("프리셋 관리")
        popup.geometry("300x400")
        popup.grab_set()
        
        ttk.Label(popup, text="저장된 프리셋 목록", font=("System", 10, "bold")).pack(pady=10)
        
        lb = tk.Listbox(popup)
        lb.pack(fill="both", expand=True, padx=10, pady=5)
        
        with open(self.presets_file, 'r', encoding='utf-8') as f:
            presets = json.load(f)
            for name in presets.keys():
                lb.insert(tk.END, name)
        
        def delete_selected():
            sel = lb.curselection()
            if not sel: return
            name = lb.get(sel[0])
            if messagebox.askyesno("삭제 확인", f"'{name}' 프리셋을 삭제할까요?"):
                del presets[name]
                with open(self.presets_file, 'w', encoding='utf-8') as f:
                    json.dump(presets, f, indent=4, ensure_ascii=False)
                lb.delete(sel[0])
                self.load_presets_list()
        
        ttk.Button(popup, text="삭제", command=delete_selected).pack(fill="x", padx=10, pady=10)
        ttk.Button(popup, text="닫기", command=popup.destroy).pack(fill="x", padx=10, pady=(0, 10))

    def run_process(self):
        if self.df_left is None: return
        
        def task():
            try:
                self.set_info("처리 중...")
                df_res = self.df_left.copy()
                
                if self.df_right is not None:
                    key = DataEngine.auto_find_key(self.df_left, self.df_right)
                    matches = DataEngine.auto_match_columns(self.df_left, self.df_right)
                    
                    if self.fuzzy_match.get():
                        df_res = DataEngine.perform_fuzzy_matching(self.df_left, self.df_right, key)
                    else:
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
                
                # Ping Task Complete
                try:
                    master = self.winfo_toplevel()
                    if hasattr(master, 'config'):
                        tel_cfg = master.config.get('telemetry', {})
                        if tel_cfg.get('enabled'):
                            # Mask filename for privacy
                            fname = os.path.basename(self.left_path)
                            mask_name = fname[:3] + "***" + os.path.splitext(fname)[1]
                            TelemetryManager.log_event(tel_cfg.get('url'), "TASK_COMPLETE", {
                                "rows": len(df_res),
                                "file": mask_name,
                                "fuzzy": self.fuzzy_match.get()
                            })
                except: pass

            except Exception as e:
                messagebox.showerror("오류", str(e))
                self.set_info("실패")

        threading.Thread(target=task, daemon=True).start()

    def load_registered_sources(self):
        """Pre-fill URLs from registered config."""
        reg = self.config.get('registered_sources', {})
        self.cloud_url.set(reg.get('github_url', ""))
        self.cloud_token.set(reg.get('github_token', ""))
        self.gs_url.set(reg.get('google_sheets_url', ""))
        self.gs_sheet_names.set(reg.get('google_sheet_names', ""))

    def save_source_config(self, source_type):
        """Save Cloud/Google URLs with password authentication."""
        # Custom PW prompt to avoid basic dialog limitations
        pw = simpledialog.askstring("관리자 인증", "설정을 저장하려면 관리자 암호를 입력하세요:", show="*")
        if pw != "3867":
            messagebox.showerror("오류", "암호가 틀렸습니다. 저장 권한이 없습니다.")
            return

        reg = self.config.get('registered_sources', {})
        if source_type == 'github':
            reg['github_url'] = self.cloud_url.get().strip()
            reg['github_token'] = self.cloud_token.get().strip()
        else:
            reg['google_sheets_url'] = self.gs_url.get().strip()
            reg['google_sheet_names'] = self.gs_sheet_names.get().strip()
        
        # Persist to JSON
        try:
            with open("config.json", 'w', encoding='utf-8') as f:
                json.dump(self.config, f, indent=4, ensure_ascii=False)
            messagebox.showinfo("완료", "소스 설정이 안전하게 등록 및 저장되었습니다.")
        except Exception as e:
            messagebox.showerror("오류", f"저장 실패: {e}")

    def set_info(self, msg):
        self.info_var.set(msg)
        self.update_idletasks()

    def select_all_cols(self):
        for v in self.col_vars.values(): v.set(True)
    def unselect_all_cols(self):
        for v in self.col_vars.values(): v.set(False)

    def peek_google(self):
        url = self.gs_url.get().strip()
        if not url: return
        names = [n.strip() for n in self.gs_sheet_names.get().split(',') if n.strip()]
        target_name = names[0] if names else None
        
        try:
            self.set_info("구글 헤더 확인 중...")
            self.cloud_headers = ExcelHandler.peek_google_sheet_headers(url, target_name)
            
            # Show headers
            for w in self.scroll_frame.scrollable_frame.winfo_children(): w.destroy()
            self.col_vars = {}
            for idx, col in enumerate(self.cloud_headers):
                var = tk.BooleanVar(value=True)
                self.col_vars[col] = var
                ttk.Checkbutton(self.scroll_frame.scrollable_frame, text=str(col), variable=var).grid(row=idx//4, column=idx%4, sticky="w", padx=15, pady=6)

            
            self.filter_combo['values'] = self.cloud_headers
            self.set_info("구글 헤더 로드 완료")
        except Exception as e:
            messagebox.showerror("구글 오류", f"데이터를 읽을 수 없습니다. (공개 여부 확인 필요)\n{e}")

    def download_google(self):
        url = self.gs_url.get().strip()
        if not url: return
        names = [n.strip() for n in self.gs_sheet_names.get().split(',') if n.strip()]
        selected_cols = [c for c, v in self.col_vars.items() if v.get()]
        
        def task():
            try:
                self.set_info("구글 시트 로딩 중...")
                dfs = []
                if not names:
                    dfs.append(ExcelHandler.read_google_sheet(url, usecols=selected_cols))
                else:
                    for name in names:
                        self.set_info(f"시트 로드 중: {name}")
                        dfs.append(ExcelHandler.read_google_sheet(url, sheet_name=name, usecols=selected_cols))
                
                # Combine multiple sheets if provided
                if len(dfs) > 1:
                    self.df_left = pd.concat(dfs, ignore_index=True)
                else:
                    self.df_left = dfs[0]
                
                self.left_path = f"GoogleSheet: {url.split('/')[-1][:10]}..."
                self.check_size()
                self.refresh_cols()
                if self.on_load_callback: self.on_load_callback(self.df_left)
                self.set_info(f"구글 연동 완료 ({len(self.df_left):,}행)")
            except Exception as e:
                messagebox.showerror("다운로드 오류", str(e))
                self.set_info("실패")

        threading.Thread(target=task, daemon=True).start()
