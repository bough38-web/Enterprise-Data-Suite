import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import pandas as pd
import xlwings as xw
from pathlib import Path


def normalize(x):
    if pd.isna(x):
        return ""
    return str(x).strip()


def normalize_col(x):
    return str(x).replace(" ", "").replace("_", "").strip()


def safe_sheet_name(name):
    invalid = ['\\', '/', '?', '*', '[', ']', ':']
    for ch in invalid:
        name = name.replace(ch, "_")
    return name[:31]


def df_to_excel_rows(df: pd.DataFrame):
    df2 = df.copy().fillna("")
    return [df2.columns.tolist()] + df2.astype(object).values.tolist()


def detect_sheets(wb):
    left = None
    right = None

    for sht in wb.sheets:
        try:
            headers = sht.range("A1").expand("right").value
            if not headers:
                continue

            if not isinstance(headers, list):
                headers = [headers]

            headers = [str(h).strip() for h in headers if h not in [None, ""]]
            header_set = set(headers)

            if {"관리본부명", "시설구분", "요금구분"}.issubset(header_set):
                left = sht

            if "추천자명" in header_set or "추천자유형" in header_set:
                right = sht
        except Exception:
            continue

    if left is None:
        left = wb.sheets[0]

    if right is None:
        right = wb.sheets[1] if len(wb.sheets) >= 2 else wb.sheets[0]

    return left, right


def auto_key(df1, df2):
    priority = ["계약번호", "서비스(소)"]

    for col in priority:
        if col in df1.columns and col in df2.columns:
            return col

    common = [c for c in df1.columns if c in df2.columns]
    if common:
        return common[0]

    return df1.columns[0]


def auto_match(df1, df2):
    left_map = {normalize_col(c): c for c in df1.columns}
    right_map = {normalize_col(c): c for c in df2.columns}

    result = {}
    for norm_name, right_col in right_map.items():
        if norm_name in left_map:
            result[right_col] = left_map[norm_name]

    return result


class SheetSelectPopup:
    def __init__(self, parent, title, sheet_names):
        self.result = None
        self.top = tk.Toplevel(parent)
        self.top.title(title)
        self.top.geometry("320x150")
        self.top.transient(parent)
        self.top.grab_set()

        frame = ttk.Frame(self.top, padding=12)
        frame.pack(fill="both", expand=True)

        ttk.Label(frame, text="시트를 선택하세요").pack(anchor="w", pady=(0, 8))

        self.sheet_var = tk.StringVar()
        combo = ttk.Combobox(frame, textvariable=self.sheet_var, values=sheet_names, state="readonly", width=32)
        combo.pack(fill="x", pady=(0, 12))
        if sheet_names:
            self.sheet_var.set(sheet_names[0])

        btn_frame = ttk.Frame(frame)
        btn_frame.pack(fill="x")

        ttk.Button(btn_frame, text="취소", command=self.cancel).pack(side="right", padx=4)
        ttk.Button(btn_frame, text="확인", command=self.apply).pack(side="right", padx=4)

        self.top.wait_window()

    def apply(self):
        self.result = self.sheet_var.get().strip()
        self.top.destroy()

    def cancel(self):
        self.result = None
        self.top.destroy()


class ValueFilterPopup:
    def __init__(self, parent, title, values, selected_values):
        self.top = tk.Toplevel(parent)
        self.top.title(title)
        self.top.geometry("430x560")
        self.top.transient(parent)
        self.top.grab_set()

        self.result = None
        self.vars = {}

        main = ttk.Frame(self.top, padding=10)
        main.pack(fill="both", expand=True)

        ttk.Label(main, text=title, font=("맑은 고딕", 10, "bold")).pack(anchor="w", pady=(0, 8))

        search_frame = ttk.Frame(main)
        search_frame.pack(fill="x", pady=(0, 8))

        ttk.Label(search_frame, text="검색").pack(side="left")
        self.search_var = tk.StringVar()
        self.search_var.trace_add("write", self.refresh_list)
        ttk.Entry(search_frame, textvariable=self.search_var, width=28).pack(side="left", padx=6)

        btn_frame = ttk.Frame(main)
        btn_frame.pack(fill="x", pady=(0, 8))
        ttk.Button(btn_frame, text="전체선택", command=self.select_all).pack(side="left", padx=4)
        ttk.Button(btn_frame, text="전체해제", command=self.unselect_all).pack(side="left", padx=4)

        wrap = ttk.Frame(main)
        wrap.pack(fill="both", expand=True)

        self.canvas = tk.Canvas(wrap, highlightthickness=0)
        self.scrollbar = ttk.Scrollbar(wrap, orient="vertical", command=self.canvas.yview)
        self.inner = ttk.Frame(self.canvas)

        self.inner.bind(
            "<Configure>",
            lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all"))
        )

        self.canvas.create_window((0, 0), window=self.inner, anchor="nw")
        self.canvas.configure(yscrollcommand=self.scrollbar.set)

        self.canvas.pack(side="left", fill="both", expand=True)
        self.scrollbar.pack(side="right", fill="y")

        self.all_values = list(values)
        selected_set = set(selected_values or [])

        for v in self.all_values:
            self.vars[v] = tk.BooleanVar(value=(v in selected_set))

        self.refresh_list()

        bottom = ttk.Frame(main)
        bottom.pack(fill="x", pady=(10, 0))

        ttk.Button(bottom, text="취소", command=self.cancel).pack(side="right", padx=4)
        ttk.Button(bottom, text="적용", command=self.apply).pack(side="right", padx=4)

        self.top.wait_window()

    def refresh_list(self, *args):
        for w in self.inner.winfo_children():
            w.destroy()

        keyword = self.search_var.get().strip()
        filtered = [v for v in self.all_values if (not keyword or keyword in v)]

        cols_per_row = 2
        for idx, value in enumerate(filtered):
            chk = ttk.Checkbutton(self.inner, text=value, variable=self.vars[value])
            r = idx // cols_per_row
            c = idx % cols_per_row
            chk.grid(row=r, column=c, sticky="w", padx=8, pady=4)

    def select_all(self):
        keyword = self.search_var.get().strip()
        for value in self.all_values:
            if not keyword or keyword in value:
                self.vars[value].set(True)

    def unselect_all(self):
        keyword = self.search_var.get().strip()
        for value in self.all_values:
            if not keyword or keyword in value:
                self.vars[value].set(False)

    def apply(self):
        self.result = [v for v, var in self.vars.items() if var.get()]
        self.top.destroy()

    def cancel(self):
        self.result = None
        self.top.destroy()


class App:
    def __init__(self, root):
        self.root = root
        self.root.title("엑셀 자동 매칭 + 조건 추출 Pro")
        self.root.geometry("1320x960")

        self.df_left = None
        self.df_right = None
        self.ws_left = None
        self.ws_right = None
        self.col_vars = {}
        self.filter_value_map = {}
        self.running = False

        self.left_source_name = ""
        self.right_source_name = ""

        self.setup_style()
        self.build_ui()

    def setup_style(self):
        style = ttk.Style()
        try:
            style.theme_use("vista")
        except Exception:
            pass

        style.configure("Title.TLabel", font=("맑은 고딕", 12, "bold"))

    def build_ui(self):
        main = ttk.Frame(self.root, padding=12)
        main.pack(fill="both", expand=True)

        ttk.Label(main, text="엑셀 자동 매칭 + 조건 추출 Pro", style="Title.TLabel").pack(anchor="w", pady=(0, 10))

        # 불러오기
        load_frame = ttk.LabelFrame(main, text="데이터 불러오기", padding=10)
        load_frame.pack(fill="x", pady=(0, 10))

        ttk.Button(load_frame, text="활성 엑셀 불러오기", command=self.load_from_active_excel).pack(side="left", padx=5)
        ttk.Button(load_frame, text="원본 파일 불러오기", command=lambda: self.load_file_data("left")).pack(side="left", padx=5)
        ttk.Button(load_frame, text="참조 파일 불러오기", command=lambda: self.load_file_data("right")).pack(side="left", padx=5)
        ttk.Button(load_frame, text="실행", command=self.run).pack(side="left", padx=12)

        self.status_var = tk.StringVar(value="활성 엑셀 또는 xlsx/csv 파일을 불러오세요.")
        ttk.Label(load_frame, textvariable=self.status_var).pack(side="left", padx=16)

        # 데이터 상태
        source_frame = ttk.LabelFrame(main, text="현재 데이터 상태", padding=10)
        source_frame.pack(fill="x", pady=(0, 10))

        self.left_info_var = tk.StringVar(value="원본: 미선택")
        self.right_info_var = tk.StringVar(value="참조: 미선택")
        ttk.Label(source_frame, textvariable=self.left_info_var).pack(anchor="w", pady=2)
        ttk.Label(source_frame, textvariable=self.right_info_var).pack(anchor="w", pady=2)

        # 옵션
        option_frame = ttk.LabelFrame(main, text="컬럼 처리 옵션", padding=10)
        option_frame.pack(fill="x", pady=(0, 10))

        self.mode = tk.StringVar(value="keep")
        ttk.Radiobutton(option_frame, text="선택 컬럼만 유지", variable=self.mode, value="keep").pack(side="left", padx=8)
        ttk.Radiobutton(option_frame, text="선택 컬럼 삭제", variable=self.mode, value="delete").pack(side="left", padx=8)

        self.auto_filter = tk.BooleanVar(value=False)
        ttk.Checkbutton(option_frame, text="시설구분/요금구분 대상만", variable=self.auto_filter).pack(side="left", padx=20)

        ttk.Button(option_frame, text="전체선택", command=self.select_all).pack(side="left", padx=8)
        ttk.Button(option_frame, text="전체해제", command=self.unselect_all).pack(side="left", padx=4)

        # 값 필터
        filter_frame = ttk.LabelFrame(main, text="특정 컬럼 값 추출", padding=10)
        filter_frame.pack(fill="x", pady=(0, 10))

        ttk.Label(filter_frame, text="필터 컬럼").grid(row=0, column=0, sticky="w", padx=4, pady=4)

        self.filter_col_var = tk.StringVar()
        self.filter_col_combo = ttk.Combobox(filter_frame, textvariable=self.filter_col_var, state="readonly", width=24)
        self.filter_col_combo.grid(row=0, column=1, sticky="w", padx=4, pady=4)
        self.filter_col_combo.bind("<<ComboboxSelected>>", self.on_filter_column_changed)

        ttk.Button(filter_frame, text="값 선택", command=self.open_value_selector).grid(row=0, column=2, sticky="w", padx=8)

        self.filter_mode = tk.StringVar(value="include")
        ttk.Radiobutton(filter_frame, text="선택값만 포함", variable=self.filter_mode, value="include").grid(row=0, column=3, sticky="w", padx=8)
        ttk.Radiobutton(filter_frame, text="선택값 제외", variable=self.filter_mode, value="exclude").grid(row=0, column=4, sticky="w", padx=8)

        self.selected_values_var = tk.StringVar(value="선택된 값 없음")
        ttk.Label(filter_frame, textvariable=self.selected_values_var, foreground="blue").grid(
            row=1, column=0, columnspan=5, sticky="w", padx=4, pady=6
        )

        # 진행상황
        progress_frame = ttk.LabelFrame(main, text="진행상황", padding=10)
        progress_frame.pack(fill="x", pady=(0, 10))

        self.progress = ttk.Progressbar(progress_frame, mode="determinate", maximum=100)
        self.progress.pack(fill="x", pady=(0, 8))

        self.summary_var = tk.StringVar(value="대기 중")
        ttk.Label(progress_frame, textvariable=self.summary_var).pack(anchor="w")

        # 로그
        log_frame = ttk.LabelFrame(main, text="처리 로그", padding=10)
        log_frame.pack(fill="x", pady=(0, 10))

        self.log_text = tk.Text(log_frame, height=8, wrap="word")
        self.log_text.pack(fill="x")
        self.log_text.insert("end", "[시작] 프로그램 준비 완료\n")
        self.log_text.config(state="disabled")

        # 컬럼 선택
        col_frame = ttk.LabelFrame(main, text="컬럼 선택", padding=10)
        col_frame.pack(fill="both", expand=True)

        self.col_wrap = ttk.Frame(col_frame)
        self.col_wrap.pack(fill="both", expand=True)

        self.canvas = tk.Canvas(self.col_wrap, highlightthickness=0)
        self.v_scrollbar = ttk.Scrollbar(self.col_wrap, orient="vertical", command=self.canvas.yview)
        self.h_scrollbar = ttk.Scrollbar(self.col_wrap, orient="horizontal", command=self.canvas.xview)

        self.frame = ttk.Frame(self.canvas)

        self.frame.bind(
            "<Configure>",
            lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all"))
        )

        self.canvas_window = self.canvas.create_window((0, 0), window=self.frame, anchor="nw")
        self.canvas.configure(yscrollcommand=self.v_scrollbar.set, xscrollcommand=self.h_scrollbar.set)

        self.canvas.pack(side="left", fill="both", expand=True)
        self.v_scrollbar.pack(side="right", fill="y")
        self.h_scrollbar.pack(side="bottom", fill="x")

        self.canvas.bind("<Configure>", self.on_canvas_configure)

    def log(self, msg):
        self.log_text.config(state="normal")
        self.log_text.insert("end", f"{msg}\n")
        self.log_text.see("end")
        self.log_text.config(state="disabled")
        self.root.update_idletasks()

    def set_progress(self, value, msg=None):
        self.progress["value"] = value
        if msg:
            self.summary_var.set(msg)
            self.log(msg)
        self.root.update_idletasks()

    def on_canvas_configure(self, event):
        self.canvas.itemconfig(self.canvas_window, width=event.width)

    def update_source_labels(self):
        left_rows = len(self.df_left) if self.df_left is not None else 0
        right_rows = len(self.df_right) if self.df_right is not None else 0
        left_cols = len(self.df_left.columns) if self.df_left is not None else 0
        right_cols = len(self.df_right.columns) if self.df_right is not None else 0

        self.left_info_var.set(f"원본: {self.left_source_name or '미선택'} / {left_rows:,}행 / {left_cols:,}열")
        self.right_info_var.set(f"참조: {self.right_source_name or '미선택'} / {right_rows:,}행 / {right_cols:,}열")

    def refresh_ui_after_load(self):
        if self.df_left is not None:
            self.load_columns(self.df_left)
            cols = list(self.df_left.columns)
            self.filter_col_combo["values"] = cols
            self.filter_col_var.set("관리본부명" if "관리본부명" in cols else (cols[0] if cols else ""))
            self.on_filter_column_changed()

        self.update_source_labels()

    def read_table_file(self, path):
        ext = Path(path).suffix.lower()

        if ext == ".csv":
            df = pd.read_csv(path, dtype=object, encoding="utf-8-sig")
            return df, None

        if ext in [".xlsx", ".xlsm", ".xls"]:
            xl = pd.ExcelFile(path)
            sheet_names = xl.sheet_names

            if len(sheet_names) == 1:
                sheet_name = sheet_names[0]
            else:
                popup = SheetSelectPopup(self.root, "시트 선택", sheet_names)
                if popup.result is None:
                    return None, None
                sheet_name = popup.result

            df = pd.read_excel(path, sheet_name=sheet_name, dtype=object)
            return df, sheet_name

        raise Exception("지원하지 않는 파일 형식입니다. xlsx, xls, xlsm, csv만 가능합니다.")

    def load_from_active_excel(self):
        try:
            self.set_progress(5, "활성 엑셀 연결 중...")
            app = xw.apps.active
            wb = app.books.active

            self.ws_left, self.ws_right = detect_sheets(wb)
            self.df_left = self.ws_left.used_range.options(pd.DataFrame, header=1, index=False).value
            self.df_right = self.ws_right.used_range.options(pd.DataFrame, header=1, index=False).value

            self.left_source_name = f"활성엑셀/{self.ws_left.name}"
            self.right_source_name = f"활성엑셀/{self.ws_right.name}"

            self.status_var.set(f"원본 시트: {self.ws_left.name} / 참조 시트: {self.ws_right.name}")
            self.refresh_ui_after_load()
            self.set_progress(100, "활성 엑셀 불러오기 완료")

        except Exception as e:
            messagebox.showerror("오류", f"활성 엑셀 불러오기 실패\n{e}")
            self.set_progress(0, "불러오기 실패")

    def load_file_data(self, side):
        path = filedialog.askopenfilename(
            title="파일 선택",
            filetypes=[
                ("지원 파일", "*.xlsx *.xls *.xlsm *.csv"),
                ("Excel 파일", "*.xlsx *.xls *.xlsm"),
                ("CSV 파일", "*.csv"),
            ]
        )
        if not path:
            return

        try:
            self.set_progress(10, f"{'원본' if side == 'left' else '참조'} 파일 읽는 중...")
            df, sheet_name = self.read_table_file(path)

            if df is None:
                self.set_progress(0, "불러오기 취소")
                return

            source_name = Path(path).name
            if sheet_name:
                source_name = f"{source_name}/{sheet_name}"

            if side == "left":
                self.df_left = df
                self.left_source_name = source_name
            else:
                self.df_right = df
                self.right_source_name = source_name

            self.status_var.set("파일 불러오기 완료")
            self.refresh_ui_after_load()
            self.set_progress(100, f"{'원본' if side == 'left' else '참조'} 파일 불러오기 완료")

        except UnicodeDecodeError:
            try:
                df = pd.read_csv(path, dtype=object, encoding="cp949")
                if side == "left":
                    self.df_left = df
                    self.left_source_name = Path(path).name
                else:
                    self.df_right = df
                    self.right_source_name = Path(path).name

                self.status_var.set("파일 불러오기 완료")
                self.refresh_ui_after_load()
                self.set_progress(100, f"{'원본' if side == 'left' else '참조'} 파일 불러오기 완료")
            except Exception as e:
                messagebox.showerror("오류", f"CSV 인코딩 읽기 실패\n{e}")
                self.set_progress(0, "불러오기 실패")
        except Exception as e:
            messagebox.showerror("오류", f"파일 불러오기 실패\n{e}")
            self.set_progress(0, "불러오기 실패")

    def load_columns(self, df):
        for w in self.frame.winfo_children():
            w.destroy()

        self.col_vars = {}
        cols_per_row = 4
        max_text_len = 24

        for idx, col in enumerate(df.columns):
            var = tk.BooleanVar(value=True)

            display_text = str(col)
            if len(display_text) > max_text_len:
                display_text = display_text[:max_text_len] + "..."

            chk = ttk.Checkbutton(self.frame, text=display_text, variable=var)

            row = idx // cols_per_row
            col_idx = idx % cols_per_row
            chk.grid(row=row, column=col_idx, sticky="w", padx=10, pady=6)

            self.col_vars[col] = var

    def select_all(self):
        for v in self.col_vars.values():
            v.set(True)

    def unselect_all(self):
        for v in self.col_vars.values():
            v.set(False)

    def on_filter_column_changed(self, event=None):
        col = self.filter_col_var.get().strip()

        if not col or self.df_left is None or col not in self.df_left.columns:
            self.selected_values_var.set("선택된 값 없음")
            return

        if col not in self.filter_value_map:
            self.filter_value_map[col] = []

        selected = self.filter_value_map.get(col, [])
        if selected:
            preview = ", ".join(selected[:8])
            if len(selected) > 8:
                preview += f" 외 {len(selected) - 8}건"
            self.selected_values_var.set(f"[{col}] 선택값: {preview}")
        else:
            self.selected_values_var.set(f"[{col}] 선택된 값 없음")

    def open_value_selector(self):
        try:
            if self.df_left is None:
                messagebox.showwarning("안내", "먼저 원본 데이터를 불러오세요.")
                return

            col = self.filter_col_var.get().strip()
            if not col or col not in self.df_left.columns:
                messagebox.showwarning("안내", "필터 컬럼을 선택해주세요.")
                return

            values = sorted([
                normalize(v) for v in self.df_left[col].dropna().tolist()
                if normalize(v) != ""
            ])
            values = list(dict.fromkeys(values))

            popup = ValueFilterPopup(
                self.root,
                f"{col} 값 선택",
                values,
                self.filter_value_map.get(col, [])
            )

            if popup.result is not None:
                self.filter_value_map[col] = popup.result
                self.on_filter_column_changed()

        except Exception as e:
            messagebox.showerror("오류", f"값 선택창 열기 실패\n{e}")

    def create_result_sheet(self, wb, base_name):
        base_name = safe_sheet_name(base_name)
        existing = [s.name for s in wb.sheets]

        if base_name not in existing:
            return wb.sheets.add(after=wb.sheets[-1], name=base_name)

        idx = 2
        while True:
            new_name = safe_sheet_name(f"{base_name}_{idx}")
            if new_name not in existing:
                return wb.sheets.add(after=wb.sheets[-1], name=new_name)
            idx += 1

    def run(self):
        try:
            if self.running:
                return

            if self.df_left is None:
                messagebox.showwarning("안내", "원본 데이터를 먼저 불러오세요.")
                return

            if self.df_right is None:
                messagebox.showwarning("안내", "참조 데이터를 먼저 불러오세요.")
                return

            self.running = True
            self.set_progress(0, "처리 시작")

            df_left = self.df_left.copy()
            df_right = self.df_right.copy()

            self.set_progress(10, "기준컬럼 찾는 중...")
            key = auto_key(df_left, df_right)
            self.log(f"기준컬럼: {key}")

            self.set_progress(20, "매칭 컬럼 분석 중...")
            col_map = auto_match(df_left, df_right)
            self.log(f"자동 매칭 컬럼 수: {len(col_map)}")

            left_key = df_left[key].map(normalize)
            right_key = df_right[key].map(normalize)

            self.set_progress(35, "컬럼 매칭 중...")
            total_match_cols = max(1, len(col_map))
            for i, right_col in enumerate(col_map.keys(), start=1):
                lookup = dict(zip(right_key, df_right[right_col]))
                df_left[f"매칭_{right_col}"] = left_key.map(lookup)
                prog = 35 + int((i / total_match_cols) * 20)
                self.progress["value"] = prog
                self.summary_var.set(f"컬럼 매칭 중... {i}/{total_match_cols}")
                self.root.update_idletasks()

            applied_filters = []

            self.set_progress(60, "조건 필터 적용 중...")
            if self.auto_filter.get():
                if "시설구분" in df_left.columns:
                    before = len(df_left)
                    df_left = df_left[df_left["시설구분"].map(normalize) == "대상"]
                    self.log(f"시설구분=대상 적용: {before:,} → {len(df_left):,}")
                    applied_filters.append("시설구분_대상")

                if "요금구분" in df_left.columns:
                    before = len(df_left)
                    df_left = df_left[df_left["요금구분"].map(normalize) == "대상"]
                    self.log(f"요금구분=대상 적용: {before:,} → {len(df_left):,}")
                    applied_filters.append("요금구분_대상")

            filter_col = self.filter_col_var.get().strip()
            selected_values = self.filter_value_map.get(filter_col, [])

            if filter_col and selected_values and filter_col in df_left.columns:
                before = len(df_left)
                series = df_left[filter_col].map(normalize)

                if self.filter_mode.get() == "include":
                    df_left = df_left[series.isin(selected_values)]
                    applied_filters.append(f"{filter_col}_포함")
                else:
                    df_left = df_left[~series.isin(selected_values)]
                    applied_filters.append(f"{filter_col}_제외")

                self.log(f"{filter_col} 값 조건 적용: {before:,} → {len(df_left):,}")
                self.log(f"선택값 수: {len(selected_values)}")

            self.set_progress(75, "컬럼 유지/삭제 적용 중...")
            selected_cols = [c for c, v in self.col_vars.items() if v.get()]

            if self.mode.get() == "keep":
                if not selected_cols:
                    messagebox.showwarning("안내", "유지할 컬럼을 하나 이상 선택해주세요.")
                    self.running = False
                    return
                df_left = df_left[selected_cols]
                self.log(f"유지 컬럼 수: {len(selected_cols)}")
            else:
                delete_cols = [c for c in selected_cols if c in df_left.columns]
                df_left = df_left.drop(columns=delete_cols, errors="ignore")
                self.log(f"삭제 컬럼 수: {len(delete_cols)}")

            self.set_progress(85, "결과 저장 준비 중...")
            app = xw.apps.active
            wb = app.books.active

            if applied_filters:
                sheet_name = "추출_" + "_".join(applied_filters[:2])
            else:
                sheet_name = "추출결과"

            ws_result = self.create_result_sheet(wb, sheet_name)

            self.set_progress(92, "엑셀에 결과 쓰는 중...")
            rows = df_to_excel_rows(df_left)
            ws_result.range("A1").value = rows

            self.set_progress(97, "서식 적용 중...")
            try:
                ws_result.used_range.columns.autofit()
                ws_result.range("A1").expand("right").api.Font.Bold = True
                ws_result.range("A1").expand("right").color = (220, 230, 241)
                ws_result.activate()
            except Exception:
                pass

            self.set_progress(100, f"완료 / {len(df_left):,}행 / {len(df_left.columns):,}열 / 시트: {ws_result.name}")
            self.status_var.set(f"완료: {len(df_left):,}행 / {len(df_left.columns):,}열 / 시트: {ws_result.name}")
            self.log("결과 저장 완료")

            messagebox.showinfo("완료", f"처리가 완료되었습니다.\n저장 시트: {ws_result.name}\n행 수: {len(df_left):,}")

        except Exception as e:
            messagebox.showerror("오류", f"실행 실패\n{e}")
            self.set_progress(0, "실행 실패")
        finally:
            self.running = False


if __name__ == "__main__":
    root = tk.Tk()
    app = App(root)
    root.mainloop()
