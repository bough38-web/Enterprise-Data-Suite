import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import pandas as pd
from pathlib import Path
import json
import threading
import os

from utils.excel_io import ExcelHandler
from utils.data_engine import DataEngine
from ui.widgets.components import create_help_btn

class BatchTab(ttk.Frame):
    def __init__(self, parent):
        super().__init__(parent)
        self.presets_file = Path("presets.json")
        self.build_ui()

    def build_ui(self):
        main = ttk.Frame(self, padding=20)
        main.pack(fill="both", expand=True)

        ttk.Label(main, text="폴더 일괄 처리 (Batch Folder Processing)", font=("System", 12, "bold")).pack(anchor="w", pady=(0, 20))

        # Folder Selection
        f_frame = ttk.LabelFrame(main, text="경로 설정", padding=15)
        f_frame.pack(fill="x", pady=(0, 20))
        create_help_btn(f_frame, "배치 경로 가이드", 
            "- 소스 폴더: 처리할 엑셀/CSV 파일들이 들어있는 폴더입니다.\n"
            "- 결과 저장 폴더: 작업이 끝난 파일들이 저장될 위치입니다.").place(relx=1.0, x=-5, y=-5, anchor="ne")

        self.src_path = tk.StringVar()
        ttk.Label(f_frame, text="소스 폴더:").grid(row=0, column=0, sticky="w")
        ttk.Entry(f_frame, textvariable=self.src_path, width=50).grid(row=0, column=1, padx=5, pady=5)
        ttk.Button(f_frame, text="찾아보기", command=lambda: self.browse_folder(self.src_path)).grid(row=0, column=2)

        self.out_path = tk.StringVar()
        ttk.Label(f_frame, text="결과 저장 폴더:").grid(row=1, column=0, sticky="w")
        ttk.Entry(f_frame, textvariable=self.out_path, width=50).grid(row=1, column=1, padx=5, pady=5)
        ttk.Button(f_frame, text="찾아보기", command=lambda: self.browse_folder(self.out_path)).grid(row=1, column=2)

        # Preset & Options
        opt_frame = ttk.LabelFrame(main, text="실행 규칙 (프리셋)", padding=15)
        opt_frame.pack(fill="x", pady=(0, 20))
        create_help_btn(opt_frame, "배치 옵션 가이드", 
            "- 프리셋: 추출할 컬럼과 필터 규칙을 미리 저장한 프리셋을 선택합니다.\n"
            "- 결과 합치기: 모든 파일의 결과물을 하나의 엑셀 시트로 병합합니다.\n"
            "- 공통 참조: 모든 원본 파일에 대해 공통으로 매칭할 데이터가 있을 경우 선택합니다.").place(relx=1.0, x=-5, y=-5, anchor="ne")

        ttk.Label(opt_frame, text="적용할 프리셋:").grid(row=0, column=0, sticky="w")
        self.preset_combo = ttk.Combobox(opt_frame, state="readonly", width=30)
        self.preset_combo.grid(row=0, column=1, padx=10, pady=5)
        ttk.Button(opt_frame, text="새로고침", command=self.load_presets).grid(row=0, column=2)

        self.merge_all = tk.BooleanVar(value=False)
        ttk.Checkbutton(opt_frame, text="모든 결과를 하나의 파일로 합치기", variable=self.merge_all).grid(row=1, column=0, columnspan=2, pady=10, sticky="w")

        # Reference File (Common for all)
        ttk.Label(opt_frame, text="공통 참조 파일 (선택):").grid(row=2, column=0, sticky="w")
        self.ref_path = tk.StringVar()
        ttk.Entry(opt_frame, textvariable=self.ref_path, width=30).grid(row=2, column=1, padx=10, pady=5)
        ttk.Button(opt_frame, text="찾아보기", command=lambda: self.browse_file(self.ref_path)).grid(row=2, column=2)

        # Progress
        self.prog_var = tk.StringVar(value="준비됨")
        ttk.Label(main, textvariable=self.prog_var).pack(anchor="w")
        self.progress = ttk.Progressbar(main, mode="determinate")
        self.progress.pack(fill="x", pady=10)

        # Action
        self.run_btn = ttk.Button(main, text="일괄 처리 시작", command=self.run_batch)
        self.run_btn.pack(pady=20, ipadx=20, ipady=5)

        self.load_presets()

    def browse_folder(self, var):
        path = filedialog.askdirectory()
        if path: var.set(path)

    def browse_file(self, var):
        path = filedialog.askopenfilename(filetypes=[("Excel", "*.xlsx *.xls")])
        if path: var.set(path)

    def load_presets(self):
        if self.presets_file.exists():
            with open(self.presets_file, 'r', encoding='utf-8') as f:
                presets = json.load(f)
                self.preset_combo['values'] = list(presets.keys())

    def run_batch(self):
        src = self.src_path.get()
        out = self.out_path.get()
        preset_name = self.preset_combo.get()

        if not src or not out or not preset_name:
            messagebox.showwarning("필수 항목 누락", "모든 경로와 프리셋을 선택하세요.")
            return

        def task():
            try:
                self.run_btn.config(state="disabled")
                with open(self.presets_file, 'r', encoding='utf-8') as f:
                    preset = json.load(f)[preset_name]

                files = [f for f in os.listdir(src) if f.lower().endswith(('.xlsx', '.xls', '.csv'))]
                if not files:
                    messagebox.showinfo("정보", "처리할 파일이 없습니다.")
                    return

                # Load Reference if provided
                df_ref = None
                if self.ref_path.get():
                    df_ref = ExcelHandler.read_file(self.ref_path.get())

                all_results = []
                self.progress['maximum'] = len(files)
                
                for i, filename in enumerate(files):
                    self.prog_var.set(f"처리 중 ({i+1}/{len(files)}): {filename}")
                    self.progress['value'] = i + 1
                    self.update_idletasks()

                    df_src = ExcelHandler.read_file(os.path.join(src, filename))
                    
                    # 1. Matching
                    if df_ref is not None:
                        key = DataEngine.auto_find_key(df_src, df_ref)
                        matches = DataEngine.auto_match_columns(df_src, df_ref)
                        df_res = DataEngine.perform_matching(df_src, df_ref, key, matches)
                    else:
                        df_res = df_src.copy()

                    # 2. Filtering
                    f_config = {
                        "auto_target": preset.get("auto_target", True),
                        "custom_filters": preset.get("active_filters", [])
                    }
                    df_res = DataEngine.apply_filters(df_res, f_config)

                    # 3. Column Selection
                    df_res = DataEngine.select_columns(df_res, preset.get("columns", []), preset.get("mode", "keep"))

                    # Add Source Info
                    df_res = DataEngine.add_source_info(df_res, filename)

                    if self.merge_all.get():
                        all_results.append(df_res)
                    else:
                        out_file = os.path.join(out, f"result_{filename}")
                        if out_file.endswith('.csv'):
                            df_res.to_csv(out_file, index=False, encoding='utf-8-sig')
                        else:
                            df_res.to_excel(out_file.replace('.csv', '.xlsx'), index=False)

                if self.merge_all.get() and all_results:
                    final_df = pd.concat(all_results, ignore_index=True)
                    # Use the optimized handler for merged results as well
                    sheet_name = ExcelHandler.write_to_active_excel(final_df, "MergedResult")
                    messagebox.showinfo("작업 완료", f"모든 파일 처리가 완료되었습니다.\n병합 시트: {sheet_name}")
                else:
                    messagebox.showinfo("작업 완료", "모든 파일 처리가 완료되었습니다.")
                self.prog_var.set("완료됨")
            except Exception as e:
                messagebox.showerror("오류", f"배치 처리 중 오류 발생: {e}")
            finally:
                self.run_btn.config(state="normal")

        threading.Thread(target=task, daemon=True).start()
