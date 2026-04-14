import tkinter as tk
from tkinter import ttk, messagebox, filedialog, simpledialog
import pandas as pd
from pathlib import Path
import json
import threading
import os
import sys

import xlwings as xw
from utils.excel_io import ExcelHandler
from utils.data_engine import DataEngine
from utils.telemetry import TelemetryManager
from ui.widgets.components import ScrollableFrame, ValueFilterPopup, SheetSelectPopup, create_help_btn
from utils.preset_manager import PresetManager
from utils.github_sync import GitHubSync

class MatchTab(ttk.Frame):
    def __init__(self, parent, config=None, config_path=None):
        super().__init__(parent)
        self.parent = parent
        self.config = config or {}
        self.config_path = config_path

        self.df_left = None
        self.df_right = None
        self.left_path = "ActiveExcel"
        self.right_path = "None"
        self.col_vars = {}
        self.active_filters = []
        self.replacements = [] # [{'column': c, 'find': a, 'replace': b, 'exact': True}]
        
        self.expert_options = {
            "trim_whitespace": tk.BooleanVar(value=True),
            "remove_all_whitespace": tk.BooleanVar(value=False),
            "format_phone": tk.BooleanVar(value=False),
            "drop_duplicates": tk.BooleanVar(value=False),
            "drop_empty_rows": tk.BooleanVar(value=True),
            "to_upper": tk.BooleanVar(value=False),
            "to_lower": tk.BooleanVar(value=False),
            "mask_id": tk.BooleanVar(value=False),
            "extract_email": tk.BooleanVar(value=False),
            "remove_special_chars": tk.BooleanVar(value=False),
            "normalize_numeric": tk.BooleanVar(value=False)
        }
        # Resolve stable path for presets.json (next to EXE/Script)
        if getattr(sys, 'frozen', False):
            base_path = os.path.dirname(sys.executable)
        else:
            base_path = os.path.dirname(os.path.abspath(__file__))
            # If in ui/tabs folder, go up to project root
            if "ui" in base_path:
                base_path = os.path.abspath(os.path.join(base_path, "../.."))
        
        self.presets_file = Path(os.path.join(base_path, "presets.json"))
        self.preset_manager = PresetManager(self.presets_file)

        self.on_load_callback = None
        
        # Cloud Integration State
        self.cloud_headers = []
        
        # Fallback fonts if not available from master
        try:
            self.fonts = self.winfo_toplevel().fonts
        except AttributeError:
            self.fonts = {
                "h1": ("System", 14, "bold"),
                "h2": ("System", 11, "bold"),
                "normal": ("System", 9),
                "small": ("System", 8)
            }
            
        self.build_ui()
        self.load_registered_sources()
        self.load_presets_list()

        # Bind resize for responsive column layout
        self.scroll_frame.bind("<Configure>", self.on_layout_change)


    def register_on_load(self, callback):
        """Register a function to be called whenever df_left is updated."""
        self.on_load_callback = callback

    def build_ui(self):
        # Main Layout: Sidebar (Left) and Data (Right)
        ctrl_frame = ttk.Frame(self, padding=0)
        ctrl_frame.pack(side="left", fill="y")

        # 1. NEW: Sticky Top Frame for fixed 'Run' button
        sticky_top = ttk.Frame(ctrl_frame, padding=(15, 10, 15, 0))
        sticky_top.pack(fill="x")
        
        self.run_btn = ttk.Button(sticky_top, text="추출 및 저장 실행", command=self.run_process, style="Accent.TButton")
        self.run_btn.pack(fill="x")
        
        ttk.Separator(ctrl_frame, orient="horizontal").pack(fill="x", pady=(10, 0))

        # 2. Scrollable Settings Area
        # horizontal=True for horizontal scrolling at the bottom
        self.sidebar = ScrollableFrame(ctrl_frame, horizontal=True)
        self.sidebar.pack(fill="both", expand=True)
        
        # Increase width significantly for macOS to fit buttons beautifully
        is_mac = sys.platform == "darwin"
        sidebar_width = 460 if is_mac else 420
        self.sidebar.canvas.config(width=sidebar_width)
        
        # All widgets now go to self.sidebar.scrollable_frame
        sf = self.sidebar.scrollable_frame
        # Increased padding for a more spacious, premium feel
        p_frame = ttk.Frame(sf, padding=(20, 15, 20, 20))
        p_frame.pack(fill="both", expand=True)

        # 1. Title & Branding (More Compact)
        header = ttk.Frame(p_frame, padding=(0, 0, 0, 10))
        header.pack(fill="x")
        ttk.Label(header, text="✨ Smart Matrix & Match", font=self.fonts.get("h1"), foreground="#0078D4").pack(side="left")

        # 2. Source Selection (Condensed Padding)
        load_lf = ttk.LabelFrame(p_frame, text="입력 소스 관리", padding=12, labelanchor="n")
        load_lf.pack(fill="x", pady=(0, 12))
        create_help_btn(load_lf, "데이터 소스 가이드", 
            "- 활성 엑셀: 현재 열려있는 엑셀 창에서 데이터를 가져옵니다.\n"
            "- 원본 파일: 추출 및 매칭을 할 주 데이터를 선택합니다.\n"
            "- 참조 파일: 매칭 시 기준이 되는 데이터를 선택합니다. (생략 가능)").place(relx=1.0, x=-5, y=-5, anchor="ne")
        
        def add_load_row(parent, text, cmd, upload_type):
            row = ttk.Frame(parent)
            row.pack(fill="x", pady=3)
            # Use fixed width for the cloud button, remaining space for primary button
            ttk.Button(row, text="(Cloud)", width=9, command=lambda: self.secure_upload_handler(upload_type)).pack(side="right", padx=(6, 0))
            ttk.Button(row, text=text, command=cmd).pack(side="left", expand=True, fill="x")

        add_load_row(load_lf, "활성 엑셀 연동", self.load_active, 'active')
        add_load_row(load_lf, "원본 파일 열기", lambda: self.load_file('left'), 'left')
        add_load_row(load_lf, "참조 파일 열기", lambda: self.load_file('right'), 'right')
        
        self.info_var = tk.StringVar(value="대기 중...")
        ttk.Label(load_lf, textvariable=self.info_var, foreground="#4A90E2").pack(pady=5, anchor="center")

        # 3. Cloud Source (GitHub Raw)
        cloud_lf = ttk.LabelFrame(p_frame, text="클라우드 소스 (GitHub Raw)", padding=12, labelanchor="n")
        cloud_lf.pack(fill="x", pady=(0, 15))
        
        cloud_title_frame = ttk.Frame(cloud_lf)
        cloud_title_frame.pack(fill="x", pady=(0, 5))
        ttk.Label(cloud_title_frame, text="GitHub 연동 설정", font=self.fonts.get("h2")).pack(side="left")
        
        btn_grp_c = ttk.Frame(cloud_title_frame)
        btn_grp_c.pack(side="right")
        ttk.Button(btn_grp_c, text="📂 탐색", width=8, command=self.open_cloud_explorer).pack(side="left", padx=2)
        ttk.Button(btn_grp_c, text="수정", width=6, command=lambda: self.unlock_source_config('github')).pack(side="left", padx=2)
        ttk.Button(btn_grp_c, text="저장", width=6, command=lambda: self.save_source_config('github')).pack(side="left", padx=2)
        
        create_help_btn(cloud_lf, "전용 클라우드 매뉴얼 (GitHub)", 
            "### 1. GitHub 저장소 준비\n"
            "- GitHub.com 로그인 후 새로운 저장소(Repository)를 생성합니다.\n"
            "- 보안을 위해 'Private' 설정을 권장합니다.\n\n"
            "### 2. 액세스 토큰 (PAT) 발급 및 등록\n"
            "- Settings > Developer Settings > Personal Access Tokens (classic) 선택\n"
            "- 필요한 권한(Scopes): **repo** (Private 저장소 접근 시 필수) 체크\n"
            "- 생성된 토큰을 앱의 [GitHub Token] 입력란에 저장하세요.\n\n"
            "### 3. 클라우드 연동 및 워크플로우\n"
            "- **업로드**: 추출 완료 후 결과 창의 [Cloud 업로드] 버튼으로 결과를 전송합니다.\n"
            "- **탐색 및 로드**: [📂 탐색] 버튼을 눌러 날짜별로 업로드된 파일을 확인하세요.\n"
            "- **데이터 활용**: 탐색기에서 데이터 더블클릭 -> [헤더 확인] -> [선택 다운로드]\n\n"
            "### 4. 주의사항\n"
            "- GitHub Raw URL은 'https://raw.githubusercontent.com/...' 형식을 권장합니다.\n"
            "- '탐색기'를 사용하면 주소를 수동으로 딸 필요 없이 자동 입력됩니다.").place(relx=1.0, x=-5, y=-5, anchor="ne")
        
        self.cloud_url = tk.StringVar()
        self.cloud_token = tk.StringVar()
        
        ttk.Label(cloud_lf, text="Raw URL:").pack(anchor="center")
        self.ent_c_url = ttk.Entry(cloud_lf, textvariable=self.cloud_url, justify="center")
        self.ent_c_url.pack(fill="x", pady=2)
        
        ttk.Label(cloud_lf, text="Personal Access Token:").pack(anchor="center")
        self.ent_c_token = ttk.Entry(cloud_lf, textvariable=self.cloud_token, show="*", justify="center")
        self.ent_c_token.pack(fill="x", pady=2)
        
        c_btns = ttk.Frame(cloud_lf)
        c_btns.pack(fill="x", pady=(10, 2))
        ttk.Button(c_btns, text="헤더 확인", command=self.peek_cloud).pack(side="left", expand=True, fill="x", padx=4)
        ttk.Button(c_btns, text="선택 다운로드", command=self.download_cloud).pack(side="left", expand=True, fill="x", padx=4)

        # 4. Google Sheets Source
        gs_lf = ttk.LabelFrame(p_frame, text="구글 스프레드시트 연동", padding=12, labelanchor="n")
        gs_lf.pack(fill="x", pady=(0, 15))
        
        gs_title_frame = ttk.Frame(gs_lf)
        gs_title_frame.pack(fill="x", pady=(0, 5))
        ttk.Label(gs_title_frame, text="Google Sheets 설정", font=self.fonts.get("h2")).pack(side="left")
        
        btn_grp_g = ttk.Frame(gs_title_frame)
        btn_grp_g.pack(side="right")
        ttk.Button(btn_grp_g, text="수정", width=6, command=lambda: self.unlock_source_config('google')).pack(side="left", padx=2)
        ttk.Button(btn_grp_g, text="저장", width=6, command=lambda: self.save_source_config('google')).pack(side="left", padx=2)
        
        create_help_btn(gs_lf, "구글 시트 가이드", 
            "- 시트 주소를 입력하세요.\n"
            "- 특정 시트(탭)를 여러 개 불러오려면 쉼표(,)로 구분하세요.").place(relx=1.0, x=-5, y=-5, anchor="ne")
        
        self.gs_url = tk.StringVar()
        self.gs_sheet_names = tk.StringVar()
        
        ttk.Label(gs_lf, text="스프레드시트 주소(URL):").pack(anchor="center")
        self.ent_g_url = ttk.Entry(gs_lf, textvariable=self.gs_url, justify="center")
        self.ent_g_url.pack(fill="x", pady=2)
        
        ttk.Label(gs_lf, text="대상 시트(탭) 목록:").pack(anchor="center")
        self.ent_g_names = ttk.Entry(gs_lf, textvariable=self.gs_sheet_names, justify="center")
        self.ent_g_names.pack(fill="x", pady=2)
        
        gs_btns = ttk.Frame(gs_lf)
        gs_btns.pack(fill="x", pady=(10, 2))
        ttk.Button(gs_btns, text="시트 목록", command=self.fetch_google_sheets).pack(side="left", expand=True, fill="x", padx=3)
        ttk.Button(gs_btns, text="헤더 확인", command=self.peek_google).pack(side="left", expand=True, fill="x", padx=3)
        ttk.Button(gs_btns, text="선택 다운로드", command=self.download_google).pack(side="left", expand=True, fill="x", padx=3)

        # 5. Options
        opt_lf = ttk.LabelFrame(p_frame, text="추출 옵션", padding=10, labelanchor="n")
        opt_lf.pack(fill="x", pady=(0, 10))
        
        opt_top = ttk.Frame(opt_lf)
        opt_top.pack(fill="x", pady=2)
        self.mode_var = tk.StringVar(value="keep")
        ttk.Radiobutton(opt_top, text="유지", variable=self.mode_var, value="keep").pack(side="left", expand=True)
        ttk.Radiobutton(opt_top, text="삭제", variable=self.mode_var, value="delete").pack(side="left", expand=True)
        
        opt_mid = ttk.Frame(opt_lf)
        opt_mid.pack(fill="x", pady=2)
        self.auto_target = tk.BooleanVar(value=True)
        ttk.Checkbutton(opt_mid, text="대상 필터", variable=self.auto_target).pack(side="left", expand=True)

        self.fuzzy_match = tk.BooleanVar(value=False)
        ttk.Checkbutton(opt_lf, text="유사도", variable=self.fuzzy_match).pack(side="left", padx=5)

        self.direct_save = tk.BooleanVar(value=False)
        ttk.Checkbutton(opt_lf, text="직접 파일 저장", variable=self.direct_save).pack(fill="x", pady=5)

        # 6. Filters
        filter_lf = ttk.LabelFrame(p_frame, text="조건 필터 관리", padding=12, labelanchor="n")
        filter_lf.pack(fill="x", pady=(0, 12), expand=True)
        
        f_top = ttk.Frame(filter_lf)
        f_top.pack(fill="x")
        self.filter_col_var = tk.StringVar()
        self.filter_combo = ttk.Combobox(f_top, textvariable=self.filter_col_var, state="readonly", width=12, justify="center")
        self.filter_combo.pack(side="left", padx=2, expand=True)
        
        self.filter_mode = tk.StringVar(value="include")
        ttk.Radiobutton(f_top, text="In", variable=self.filter_mode, value="include").pack(side="left", expand=True)
        ttk.Radiobutton(f_top, text="Ex", variable=self.filter_mode, value="exclude").pack(side="left", expand=True)
        
        ttk.Button(f_top, text="추가", command=self.add_filter_rule).pack(side="right", padx=2)
        
        self.filter_listbox = tk.Listbox(filter_lf, height=4)
        self.filter_listbox.pack(fill="both", expand=True, pady=5)
        
        f_btns = ttk.Frame(filter_lf)
        f_btns.pack(fill="x")
        ttk.Button(f_btns, text="삭제", command=self.remove_filter_rule).pack(side="left", padx=2)
        ttk.Button(f_btns, text="초기화", command=self.clear_all_filters).pack(side="left", padx=2)

        # 6-1. Replace Rules (치환 기능)
        rep_lf = ttk.LabelFrame(p_frame, text="값 치환 및 변경", padding=12, labelanchor="n")
        rep_lf.pack(fill="x", pady=(0, 12), expand=True)
        
        ttk.Label(rep_lf, text="특정 값을 찾아 다른 값으로 일괄 변경합니다.").pack(anchor="w", pady=(0,5))
        
        # Add a button to open a replacement management popup (since it's complex)
        ttk.Button(rep_lf, text="🔠 치환 규칙 관리", command=self.manage_replacements_ui).pack(fill="x", pady=2)
        
        self.rep_count_var = tk.StringVar(value="현재 규칙: 0 개")
        ttk.Label(rep_lf, textvariable=self.rep_count_var).pack(pady=2)

        # 6-2. Expert Data Cleaning Options
        exp_lf = ttk.LabelFrame(p_frame, text="💡 전문가 데이터 클리닝 (Expert)", padding=12, labelanchor="n")
        exp_lf.pack(fill="x", pady=(0, 12), expand=True)
        
        create_help_btn(exp_lf, "전문가 기법 안내", 
            "데이터를 자동으로 정제해주는 고급 기능들입니다.\n클릭 한 번으로 복잡한 엑셀 수식을 대체할 수 있습니다.").place(relx=1.0, x=-5, y=-5, anchor="ne")

        def add_exp_cb(parent, text, var_key):
            ttk.Checkbutton(parent, text=text, variable=self.expert_options[var_key]).pack(anchor="w", pady=2)
            
        r1 = ttk.Frame(exp_lf)
        r1.pack(fill="x")
        r2 = ttk.Frame(exp_lf)
        r2.pack(fill="x", pady=(5,0))
        
        add_exp_cb(r1, "앞뒤 공백 무조건 제거", "trim_whitespace")
        add_exp_cb(r1, "모든 공백 제거 (사이 띄어쓰기 포함)", "remove_all_whitespace")
        add_exp_cb(r1, "전화번호 하이픈(-) 010 자동 포맷", "format_phone")
        add_exp_cb(r1, "비어있는 줄(Empty) 자동 삭제", "drop_empty_rows")
        add_exp_cb(r1, "중복 데이터(행) 무조건 삭제", "drop_duplicates")
        
        add_exp_cb(r2, "주민번호/사업자번호 별표(*) 마스킹", "mask_id")
        add_exp_cb(r2, "이메일 주소만 텍스트에서 강제 추출", "extract_email")
        add_exp_cb(r2, "영문 모두 대문자로 변경", "to_upper")
        add_exp_cb(r2, "특수기호 제거 (알파벳, 글자, 숫자만)", "remove_special_chars")
        add_exp_cb(r2, "금액/숫자 콤마(,) 제거 및 숫자화", "normalize_numeric")

        # 7. Presets
        pre_lf = ttk.LabelFrame(p_frame, text="추출 프리셋", padding=12, labelanchor="n")
        pre_lf.pack(fill="x", pady=(0, 12))
        
        pre_top = ttk.Frame(pre_lf)
        pre_top.pack(fill="x")
        self.preset_var = tk.StringVar()
        self.preset_combo = ttk.Combobox(pre_top, textvariable=self.preset_var, state="readonly", justify="center")
        self.preset_combo.pack(side="left", expand=True, fill="x", padx=(0, 5))
        self.preset_combo.bind("<<ComboboxSelected>>", self.apply_preset)
        
        ttk.Button(pre_top, text="적용", command=self.apply_preset, width=5).pack(side="left", padx=2)
        ttk.Button(pre_top, text="(Sync)", command=self.sync_presets, width=6).pack(side="left", padx=2)

        pre_btns = ttk.Frame(pre_lf)
        pre_btns.pack(fill="x", pady=(8, 2))
        ttk.Button(pre_btns, text="+ 새 프리셋 저장", command=self.save_current_preset).pack(side="left", expand=True, fill="x", padx=4)
        ttk.Button(pre_btns, text="프리셋 관리", command=self.manage_presets_ui).pack(side="left", expand=True, fill="x", padx=4)

        # 8. Column Selection (Right Pane - Not in scrollable area)
        col_frame = ttk.LabelFrame(self, text="컬럼 선택", padding=10, labelanchor="n")
        col_frame.pack(side="right", fill="both", expand=True, padx=(10, 0))
        
        # 8-1. Quick Favorite Columns (유지컬럼)
        fav_frame = ttk.Frame(col_frame)
        fav_frame.pack(fill="x", pady=(0, 10))
        ttk.Label(fav_frame, text="⭐ 단골 컬럼 (유지컬럼) 스마트 선택", font=self.fonts.get("h2")).pack(side="left")
        
        fb = ttk.Frame(fav_frame)
        fb.pack(side="right")
        
        ttk.Button(fb, text="★ 저장된 유지컬럼 자동 체크", command=self.load_favorite_columns, style="Accent.TButton").pack(side="left", padx=2)
        ttk.Button(fb, text="현재 체크된 컬럼을 유지컬럼으로 덮어쓰기", command=self.save_favorite_columns).pack(side="left", padx=2)
            
        ttk.Separator(col_frame, orient="horizontal").pack(fill="x", pady=(0,5))
        
        self.scroll_frame = ScrollableFrame(col_frame, horizontal=True)
        self.scroll_frame.pack(fill="both", expand=True)
        
        btn_bar = ttk.Frame(col_frame)
        btn_bar.pack(fill="x", pady=5)
        ttk.Button(btn_bar, text="전체 선택", command=self.select_all_cols).pack(side="left", padx=2)
        ttk.Button(btn_bar, text="전체 해제", command=self.unselect_all_cols).pack(side="left", padx=2)
        ttk.Button(btn_bar, text="+ 수동 컬럼", command=self.add_manual_column, style="Accent.TButton").pack(side="right", padx=2)

    def open_cloud_explorer(self):
        """Open a popup to browse files in the GitHub uploads/ directory."""
        master = self.winfo_toplevel()
        url = master.config['registered_sources'].get('github_url', '')
        token = master.config['registered_sources'].get('github_token', '')
        
        if not url:
            from tkinter import messagebox
            messagebox.showwarning("설정 필요", "관리자 설정에서 GitHub 저장소 URL을 먼저 등록해 주세요.")
            return

        from ui.widgets.components import CloudExplorerPopup
        explorer = CloudExplorerPopup(self, token, url)
        
        if explorer.result:
            self.cloud_url.set(explorer.result)
            self.peek_cloud()

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
                ttk.Checkbutton(self.scroll_frame.scrollable_frame, text=str(col), variable=var, style="Small.TCheckbutton").grid(row=0, column=0)
            
            self.regrid_columns()

            
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
            self.set_info("활성 엑셀 연동 중...")
            app = xw.apps.active
            if not app:
                raise Exception("열려있는 엑셀 앱이 없습니다.")
                
            if not app.books:
                raise Exception("열려있는 엑셀 문서(작업창)가 없습니다.")
                
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
            # Use dynamic layout helper
            self.place_col_checkbox(idx, col, var)

    def on_layout_change(self, event=None):
        """Respond to window resizing by re-gridding columns."""
        if not self.col_vars: return
        self.regrid_columns()

    def regrid_columns(self):
        """Recalculate and apply grid positions for all column checkboxes."""
        if not self.col_vars: return
        
        parent = self.scroll_frame.scrollable_frame
        children = parent.winfo_children()
        if not children: return
        
        # Calculate optimal columns based on width
        width = self.scroll_frame.winfo_width()
        if width < 100: width = 800 # Fallback for initial render
        
        # Heuristic: ~150px per column for 'Small' font checkboxes
        num_cols = max(1, min(10, width // 140))
        
        for idx, child in enumerate(children):
            if isinstance(child, ttk.Checkbutton):
                child.grid(row=idx//num_cols, column=idx%num_cols, sticky="w", padx=10, pady=4)

    def place_col_checkbox(self, idx, col, var):
        """Initial placement helper."""
        # This will be refined by regrid_columns via <Configure> event
        ttk.Checkbutton(self.scroll_frame.scrollable_frame, text=str(col), variable=var, style="Small.TCheckbutton").grid(
            row=idx//6, column=idx%6, sticky="w", padx=10, pady=4
        )


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
            if not f.get('values'):
                v = "값 없음"
            else:
                v = f"{f['values'][0]}" if len(f['values']) == 1 else f"{f['values'][0]} 외 {len(f['values'])-1}건"
            self.filter_listbox.insert(tk.END, f"[{f['column']}] {m}: {v}")

    def load_presets_list(self):
        presets = self.preset_manager.load_all()
        self.preset_combo['values'] = list(presets.keys())

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
        
        self.preset_manager.add_preset(name, preset_data)
        self.load_presets_list()
        messagebox.showinfo("완료", f"'{name}' 프리셋이 저장되었습니다.")

    def apply_preset(self, event=None):
        name = self.preset_var.get()
        if not name: return
        
        presets = self.preset_manager.load_all()
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

    def sync_presets(self):
        url = self.config.get('registered_sources', {}).get('remote_presets_url', '')
        if not url:
            messagebox.showwarning("경고", "원격 프리셋 URL이 설정되어 있지 않습니다.\n관리자 설정에서 등록하세요.")
            return
        
        def task():
            try:
                self.set_info("프리셋 동기화 중...")
                token = self.config.get('registered_sources', {}).get('github_token', '')
                success, msg, _count = self.preset_manager.sync_from_remote(url, token)
                if success:
                    self.load_presets_list()
                    messagebox.showinfo("동기화 성공", msg)
                    self.set_info("동기화 완료")
                else:
                    messagebox.showerror("동기화 실패", msg)
                    self.set_info("실패")
            except Exception as e:
                messagebox.showerror("오류", str(e))
                self.set_info("실패")
        
        threading.Thread(target=task, daemon=True).start()

    def manage_presets_ui(self):
        presets = self.preset_manager.load_all()
        if not presets: return
        
        popup = tk.Toplevel(self)
        popup.title("프리셋 관리")
        popup.geometry("300x400")
        popup.grab_set()
        
        # Center popup
        sw = self.winfo_screenwidth()
        sh = self.winfo_screenheight()
        popup.geometry(f"300x400+{sw//2-150}+{sh//2-200}")

        ttk.Label(popup, text="저장된 프리셋 목록", font=("System", 10, "bold")).pack(pady=10)
        
        lb = tk.Listbox(popup)
        lb.pack(fill="both", expand=True, padx=10, pady=5)
        
        for name in presets.keys():
            lb.insert(tk.END, name)
        
        def delete_selected():
            sel = lb.curselection()
            if not sel: return
            name = lb.get(sel[0])
            if messagebox.askyesno("삭제 확인", f"'{name}' 프리셋을 삭제할까요?"):
                if self.preset_manager.delete_preset(name):
                    lb.delete(sel[0])
                    self.load_presets_list()
        
        ttk.Button(popup, text="삭제", command=delete_selected).pack(fill="x", padx=10, pady=10)
        ttk.Button(popup, text="닫기", command=popup.destroy).pack(fill="x", padx=10, pady=(0, 10))

    def toggle_fav_col(self, col_name):
        pass # Replaced by load_favorite_columns / save_favorite_columns

    def save_favorite_columns(self):
        """Save currently checked columns to config as 'favorite_columns'."""
        current_cols = [c for c, v in self.col_vars.items() if v.get()]
        if not current_cols:
            messagebox.showwarning("오류", "선택된 컬럼이 없습니다.")
            return
            
        master = self.winfo_toplevel()
        if hasattr(master, 'config'):
            master.config['favorite_columns'] = current_cols
            try:
                with open(self.config_path, 'w', encoding='utf-8') as f:
                    json.dump(master.config, f, indent=4, ensure_ascii=False)
                messagebox.showinfo("저장 완료", f"현재 선택된 {len(current_cols)}개의 컬럼이 '유지컬럼'으로 영구 저장되었습니다.\n\n[저장된 목록]\n{', '.join(current_cols)}")
            except Exception as e:
                messagebox.showerror("저장 실패", f"설정 파일 저장 중 오류: {e}")

    def load_favorite_columns(self):
        """Check all columns that are in the saved 'favorite_columns' list with EXPERT Smart Matching."""
        import difflib
        
        if not self.col_vars:
            messagebox.showinfo("안내", "먼저 데이터를 불러와 컬럼 목록을 활성화해주세요.")
            return
            
        master = self.winfo_toplevel()
        favs = getattr(master, 'config', {}).get('favorite_columns', [])
        
        if not favs:
            messagebox.showinfo("정보", "저장된 유지컬럼이 없습니다. 먼저 '유지컬럼 덮어쓰기'를 통해 저장해 주세요.")
            return
            
        self.unselect_all_cols()
        match_count = 0
        all_cols = list(self.col_vars.keys())
        
        for fav in favs:
            # 1. Strict exact match
            if fav in self.col_vars:
                if not self.col_vars[fav].get():
                    self.col_vars[fav].set(True)
                    match_count += 1
            else:
                # 2. Expert Fuzzy Match (using difflib for high precision)
                close_matches = difflib.get_close_matches(fav, all_cols, n=1, cutoff=0.6)
                if close_matches:
                    target = close_matches[0]
                    if not self.col_vars[target].get():
                        self.col_vars[target].set(True)
                        match_count += 1
                else:
                    # 3. Fallback: case-insensitive/space-stripped contains
                    fav_clean = str(fav).lower().replace(" ","")
                    for c in all_cols:
                        c_clean = str(c).lower().replace(" ","")
                        if fav_clean in c_clean or c_clean in fav_clean:
                            if not self.col_vars[c].get():
                                self.col_vars[c].set(True)
                                match_count += 1
                                break
                
        if match_count > 0:
            self.set_info(f"유지컬럼 스마트 로드 완료 ({match_count}건 매칭)")
            messagebox.showinfo("완료", f"데이터 상에서 {match_count}개의 유지컬럼이 스마트 매칭(Fuzzy Match)으로 자동 선택되었습니다.")
        else:
            messagebox.showwarning("매칭 실패", "현재 데이터에서 저장된 유지컬럼과 유사한 항목을 찾지 못했습니다.")

    def manage_replacements_ui(self):
        """UI to manage find & replace rules."""
        popup = tk.Toplevel(self)
        popup.title("데이터 치환 규칙 관리")
        popup.geometry("400x500")
        popup.grab_set()
        
        sw = self.winfo_screenwidth()
        sh = self.winfo_screenheight()
        popup.geometry(f"400x500+{sw//2-200}+{sh//2-250}")
        
        ttk.Label(popup, text="[ 특정 텍스트 일괄 치환 ]", font=self.fonts.get("h2")).pack(pady=10)
        
        lb = tk.Listbox(popup, height=8)
        lb.pack(fill="both", expand=True, padx=10, pady=5)
        
        def refresh_list():
            lb.delete(0, tk.END)
            for r in self.replacements:
                m = "정확히 일치" if r.get('exact') else "포함"
                lb.insert(tk.END, f"[{r['column']}] '{r['find']}' -> '{r['replace']}' ({m})")
            self.rep_count_var.set(f"현재 규칙: {len(self.replacements)} 개")
                
        refresh_list()
        
        def delete_rule():
            sel = lb.curselection()
            if not sel: return
            del self.replacements[sel[0]]
            refresh_list()
            
        ttk.Button(popup, text="선택 규칙 삭제", command=delete_rule).pack(fill="x", padx=10, pady=2)
        ttk.Separator(popup, orient="horizontal").pack(fill="x", pady=10)
        
        # Add Area
        add_f = ttk.Frame(popup, padding=10)
        add_f.pack(fill="x")
        
        ttk.Label(add_f, text="대상 컬럼:").grid(row=0, column=0, sticky="w")
        col_cbo = ttk.Combobox(add_f, values=list(self.col_vars.keys()), state="readonly")
        col_cbo.grid(row=0, column=1, sticky="ew", pady=2, padx=5)
        
        ttk.Label(add_f, text="찾을 값(A):").grid(row=1, column=0, sticky="w")
        ent_find = ttk.Entry(add_f)
        ent_find.grid(row=1, column=1, sticky="ew", pady=2, padx=5)
        
        ttk.Label(add_f, text="바꿀 값(B):").grid(row=2, column=0, sticky="w")
        ent_rep = ttk.Entry(add_f)
        ent_rep.grid(row=2, column=1, sticky="ew", pady=2, padx=5)
        
        exact_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(add_f, text="셀 전체가 정확히 일치할 때만 변경", variable=exact_var).grid(row=3, column=0, columnspan=2, sticky="w", pady=5)
        
        def add_rule():
            c = col_cbo.get()
            f = ent_find.get()
            r = ent_rep.get()
            if not c or not f:
                messagebox.showwarning("입력 필요", "컬럼과 찾을 값을 입력하세요.")
                return
            self.replacements.append({"column": c, "find": f, "replace": r, "exact": exact_var.get()})
            refresh_list()
            ent_find.delete(0, tk.END)
            ent_rep.delete(0, tk.END)
            
        ttk.Button(add_f, text="규칙 추가", command=add_rule, style="Accent.TButton").grid(row=4, column=0, columnspan=2, sticky="ew", pady=10)
        ttk.Button(popup, text="닫기", command=popup.destroy).pack(fill="x", padx=10, pady=(0, 10))

    def run_process(self):
        if self.df_left is None: return
        
        def task():
            try:
                self.set_info("처리 중...")
                df_res = self.df_left.copy()
                
                # Matching
                if self.df_right is not None:
                    # In existing code, it was self.df_left.copy() at line 512
                    # We'll use the working copy df_res
                    key = DataEngine.auto_find_key(df_res, self.df_right)
                    matches = DataEngine.auto_match_columns(df_res, self.df_right)
                    
                    if self.fuzzy_match.get():
                        df_res = DataEngine.perform_fuzzy_matching(df_res, self.df_right, key)
                    else:
                        df_res = DataEngine.perform_matching(df_res, self.df_right, key, matches)
                
                # Apply Filters with Diagnostics
                df_res, diag = DataEngine.apply_filters(df_res, {"auto_target": self.auto_target.get(), "custom_filters": self.active_filters})
                
                # Apply Replacements
                if self.replacements:
                    df_res = DataEngine.apply_replacements(df_res, self.replacements)
                    
                # Apply Expert Options
                active_experts = [k for k, v in self.expert_options.items() if v.get()]
                if active_experts:
                    df_res = DataEngine.apply_expert_filters(df_res, active_experts)
                
                # Select Columns
                df_res = DataEngine.select_columns(df_res, [c for c, v in self.col_vars.items() if v.get()], self.mode_var.get())
                df_res = DataEngine.add_source_info(df_res, self.left_path)
                
                if self.direct_save.get():
                    out_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel", "*.xlsx"), ("CSV", "*.csv")])
                    if not out_path:
                        self.set_info("저장 취소")
                        return
                    ExcelHandler.save_to_file(df_res, out_path)
                    res_msg = f"파일 저장 완료: {os.path.basename(out_path)}"
                else:
                    try:
                        sheet_name = ExcelHandler.write_to_active_excel(df_res, "추출결과")
                        res_msg = f"엑셀 완료 (시트: {sheet_name})"
                    except Exception as export_err:
                        if messagebox.askyesno("엑셀 연동 실패", 
                            f"열려있는 엑셀에 데이터를 직접 쓸 수 없습니다.\n오류: {export_err}\n\n결과물을 파일로 대신 저장하시겠습니까?"):
                            out_path = filedialog.asksaveasfilename(
                                defaultextension=".xlsx",
                                filetypes=[("Excel files", "*.xlsx"), ("CSV files", "*.csv")]
                            )
                            if out_path:
                                ExcelHandler.save_to_file(df_res, out_path)
                                res_msg = f"파일 저장 완료: {os.path.basename(out_path)}"
                            else:
                                raise Exception("작업이 취소되었습니다.")
                        else:
                            raise export_err
                
                self.set_info("완료")
                
                # Show results with Cloud Sync option
                self.after(0, lambda: self.show_result_popup(res_msg, diag, out_path if self.direct_save.get() else None))
                
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

    def show_result_popup(self, res_msg, diag, saved_path):
        popup = tk.Toplevel(self)
        popup.title("추출 작업 완료")
        popup.geometry("450x450")
        popup.resizable(False, False)
        popup.grab_set()
        
        # Center
        sw = self.winfo_screenwidth()
        sh = self.winfo_screenheight()
        popup.geometry(f"450x450+{sw//2-225}+{sh//2-225}")
        
        main = ttk.Frame(popup, padding=20)
        main.pack(fill="both", expand=True)
        
        ttk.Label(main, text=">> 추출이 성공적으로 완료되었습니다", font=("System", 12, "bold"), foreground="#0078D4").pack(pady=(0, 10))
        ttk.Label(main, text=res_msg, font=("System", 10)).pack(pady=(0, 20))
        
        diag_frame = ttk.LabelFrame(main, text="[ 처리 상세 내역 ]", padding=15)
        diag_frame.pack(fill="x", pady=(0, 25))
        
        stats = [
            (f"원본 데이터:", f"{diag['initial']:,} 건"),
            (f"대상 필터 제거:", f"-{diag['auto_target_removed']:,} 건"),
            (f"조건 필터 제거:", f"-{diag['custom_filter_removed']:,} 건"),
            (f"최종 추출 수량:", f"{diag['final']:,} 건")
        ]
        
        for lbl, val in stats:
            f = ttk.Frame(diag_frame)
            f.pack(fill="x", pady=2)
            ttk.Label(f, text=lbl).pack(side="left")
            ttk.Label(f, text=val, font=("System", 9, "bold")).pack(side="right")

        # Cloud Sync Section
        btn_frame = ttk.Frame(main)
        btn_frame.pack(fill="x", side="bottom")

        def upload_to_cloud():
            if not saved_path or not os.path.exists(saved_path):
                messagebox.showwarning("오류", "먼저 결과물을 파일로 저장해야 클라우드 업로드가 가능합니다.\n(직접 저장 옵션을 사용하세요)")
                return
            
            master = self.winfo_toplevel()
            url = master.config['registered_sources'].get('github_url', '')
            token = master.config['registered_sources'].get('github_token', '')
            
            if not url or not token:
                messagebox.showwarning("설정 필요", "관리자 설정에서 GitHub URL과 토큰을 먼저 등록해 주세요.")
                return

            def upload_task():
                try:
                    self.set_info("클라우드 업로드 중...")
                    success, msg = GitHubSync.upload_file(token, url, saved_path)
                    if success:
                        messagebox.showinfo("업로드 성공", msg)
                    else:
                        messagebox.showerror("업로드 실패", msg)
                except Exception as e:
                    messagebox.showerror("오류", str(e))
                finally:
                    self.set_info("완료")

            threading.Thread(target=upload_task, daemon=True).start()

        sync_btn = ttk.Button(btn_frame, text="(Cloud) 결과를 깃허브에 업로드 (Sync)", style="Accent.TButton", command=upload_to_cloud)
        sync_btn.pack(fill="x", pady=5)
        
        ttk.Button(btn_frame, text="닫기", command=popup.destroy).pack(fill="x", pady=5)

    def secure_upload_handler(self, upload_type):
        """Security-verified upload handler for any loaded file or active excel."""
        # 1. Identity Verification
        pwd = simpledialog.askstring("동기화 보안 확인", "전송 비밀번호를 입력하세요:", show="*")
        if pwd != "3867": # Admin/Sync password
            if pwd: messagebox.showerror("인증 실패", "비밀번호가 올바르지 않습니다.")
            return

        # 2. Path Determination
        path_to_upload = None
        if upload_type == 'left': path_to_upload = self.left_path
        elif upload_type == 'right': path_to_upload = self.right_path
        elif upload_type == 'active':
            if self.df_left is None:
                messagebox.showwarning("데이터 없음", "업로드할 활성 데이터가 없습니다.\n먼저 '활성 엑셀 연동'을 실행해 주세요.")
                return
            # Save to temporary file
            try:
                temp_dir = os.path.join(os.path.abspath(os.curdir), "temp")
                os.makedirs(temp_dir, exist_ok=True)
                path_to_upload = os.path.join(temp_dir, f"ActiveExcel_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx")
                ExcelHandler.save_to_file(self.df_left, path_to_upload)
            except Exception as e:
                messagebox.showerror("오류", f"임시 파일 생성 실패: {e}")
                return

        if not path_to_upload or not os.path.exists(path_to_upload):
            messagebox.showwarning("파일 없음", "업로드할 파일을 찾을 수 없습니다.")
            return

        # 3. Perform Upload
        master = self.winfo_toplevel()
        url = master.config['registered_sources'].get('github_url', '')
        token = master.config['registered_sources'].get('github_token', '')

        if not url or not token:
            messagebox.showwarning("설정 필요", "관리자 설정에서 GitHub URL과 토큰을 먼저 등록해 주세요.")
            return

        def upload_task():
            try:
                self.set_info("클라우드 전송 중...")
                # Pass network config for proxy/SSL bypass
                net_cfg = master.config.get('network', {})
                success, msg = GitHubSync.upload_file(token, url, path_to_upload, network_config=net_cfg)
                if success:
                    messagebox.showinfo("전송 성공", f"보안 전송 완료!\n{msg}")
                else:
                    messagebox.showerror("전송 실패", msg)
            except Exception as e:
                messagebox.showerror("오류", str(e))
            finally:
                self.set_info("완료")

        threading.Thread(target=upload_task, daemon=True).start()

    def load_registered_sources(self):
        """Pre-fill URLs from registered config and lock fields."""
        reg = self.config.get('registered_sources', {})
        self.cloud_url.set(reg.get('github_url', ""))
        self.cloud_token.set(reg.get('github_token', ""))
        self.gs_url.set(reg.get('google_sheets_url', ""))
        self.gs_sheet_names.set(reg.get('google_sheet_names', ""))
        
        # Initially lock if there is data
        for ent in [self.ent_c_url, self.ent_c_token, self.ent_g_url, self.ent_g_names]:
            if ent.get():
                ent.config(state="disabled")

    def unlock_source_config(self, source_type):
        """Unlock fields with password."""
        pw = simpledialog.askstring("관리자 인증", "수정 권한을 얻으려면 관리자 암호를 입력하세요:", show="*")
        if pw == "3867":
            if source_type == 'github':
                self.ent_c_url.config(state="normal")
                self.ent_c_token.config(state="normal")
            else:
                self.ent_g_url.config(state="normal")
                self.ent_g_names.config(state="normal")
            messagebox.showinfo("인증 성공", "입력창이 활성화되었습니다.")
        else:
            messagebox.showerror("오류", "암호가 틀렸습니다.")

    def save_source_config(self, source_type):
        """Save Cloud/Google URLs with password authentication."""
        pw = simpledialog.askstring("관리자 인증", "설정을 저장하려면 관리자 암호를 입력하세요:", show="*")
        if pw != "3867":
            messagebox.showerror("오류", "암호가 틀렸습니다. 저장 권한이 없습니다.")
            return

        reg = self.config.get('registered_sources', {})
        if source_type == 'github':
            reg['github_url'] = self.cloud_url.get().strip()
            reg['github_token'] = self.cloud_token.get().strip()
            self.ent_c_url.config(state="disabled")
            self.ent_c_token.config(state="disabled")
        else:
            reg['google_sheets_url'] = self.gs_url.get().strip()
            reg['google_sheet_names'] = self.gs_sheet_names.get().strip()
            self.ent_g_url.config(state="disabled")
            self.ent_g_names.config(state="disabled")
        
        # Persist to JSON
        try:
            path = self.config_path or "config.json"
            with open(path, 'w', encoding='utf-8') as f:
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
            if not self.cloud_headers:
                raise ValueError("가져올 헤더가 없습니다. 시트 내용을 확인하세요.")
            
            # Show headers
            for w in self.scroll_frame.scrollable_frame.winfo_children(): w.destroy()
            self.col_vars = {}
            for idx, col in enumerate(self.cloud_headers):
                var = tk.BooleanVar(value=True)
                self.col_vars[col] = var
                ttk.Checkbutton(self.scroll_frame.scrollable_frame, text=str(col), variable=var, style="Small.TCheckbutton").grid(row=0, column=0)
            
            self.regrid_columns()

            
            self.filter_combo['values'] = self.cloud_headers
            self.set_info("구글 헤더 로드 완료")
        except Exception as e:
            messagebox.showerror("구글 오류", f"데이터를 읽을 수 없습니다. (공개 여부 확인 필요)\n{e}")

    def fetch_google_sheets(self):
        url = self.gs_url.get().strip()
        if not url:
            messagebox.showwarning("입력 필요", "스프레드시트 주소를 먼저 입력하세요.")
            return
            
        try:
            self.set_info("시트 목록 가져오는 중...")
            sheet_names = ExcelHandler.get_google_sheet_list(url)
            if not sheet_names:
                raise ValueError("시트 목록을 가져올 수 없습니다. 주소나 공개 여부를 확인하세요.")
                
            popup = SheetSelectPopup(self.winfo_toplevel(), "시트 선택", sheet_names)
            if popup.result:
                current = self.gs_sheet_names.get().strip()
                if current:
                    if popup.result not in [n.strip() for n in current.split(',')]:
                        self.gs_sheet_names.set(f"{current}, {popup.result}")
                else:
                    self.gs_sheet_names.set(popup.result)
                self.set_info(f"시트 추가됨: {popup.result}")
            else:
                self.set_info("선택 취소")
        except Exception as e:
            messagebox.showerror("구글 오류", str(e))
            self.set_info("실패")

    def add_manual_column(self):
        name = simpledialog.askstring("수동 컬럼 추가", "추가할 컬럼명을 입력하세요:")
        if not name: return
        
        name = name.strip()
        if name in self.col_vars:
            messagebox.showwarning("중복", "이미 존재하는 컬럼명입니다.")
            return
            
        # Add to UI
        idx = len(self.col_vars)
        var = tk.BooleanVar(value=True)
        self.col_vars[name] = var
        
        # Grid placement logic (match existing grid)
        ttk.Checkbutton(self.scroll_frame.scrollable_frame, text=str(name), variable=var, style="Small.TCheckbutton").grid(
            row=idx//6, column=idx%6, sticky="w", padx=10, pady=4
        )
        
        # Update filter combo
        self.filter_combo['values'] = list(self.col_vars.keys())
        self.set_info(f"수동 컬럼 추가: {name}")

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
                if not dfs:
                    raise ValueError("로딩된 데이터가 없습니다.")
                    
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
