import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import pandas as pd
from pathlib import Path
import json
import threading
import os
import sys
from datetime import datetime

from utils.excel_io import ExcelHandler
from utils.data_engine import DataEngine
from ui.widgets.components import create_help_btn
from utils.preset_manager import PresetManager
from utils.github_sync import GitHubSync

class BatchTab(ttk.Frame):
    def __init__(self, parent, config=None):
        super().__init__(parent)
        self.config = config or {}
        
        # Resolve stable path for presets.json (next to EXE/Script)
        if getattr(sys, 'frozen', False):
            base_path = os.path.dirname(sys.executable)
        else:
            base_path = os.path.dirname(os.path.abspath(__file__))
            if "ui" in base_path:
                base_path = os.path.abspath(os.path.join(base_path, "../.."))
        
        self.presets_file = Path(os.path.join(base_path, "presets.json"))
        self.preset_manager = PresetManager(self.presets_file)

        self.build_ui()

    def build_ui(self):
        main = ttk.Frame(self, padding=30)
        main.pack(fill="both", expand=True)

        ttk.Label(main, text="폴더 일괄 처리 (Batch Folder Processing)", font=("System", 12, "bold")).pack(anchor="w", pady=(0, 20))

        # Folder Selection
        f_frame = ttk.LabelFrame(main, text="경로 설정", padding=20)
        f_frame.pack(fill="x", pady=(10, 25))
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
        opt_frame = ttk.LabelFrame(main, text="실행 규칙 (프리셋)", padding=20)
        opt_frame.pack(fill="x", pady=(10, 25))
        create_help_btn(opt_frame, "배치 옵션 가이드", 
            "- 프리셋: 추출할 컬럼과 필터 규칙을 미리 저장한 프리셋을 선택합니다.\n"
            "- 결과 합치기: 모든 파일의 결과물을 하나의 엑셀 시트로 병합합니다.\n"
            "- 공통 참조: 모든 원본 파일에 대해 공통으로 매칭할 데이터가 있을 경우 선택합니다.").place(relx=1.0, x=-5, y=-5, anchor="ne")

        ttk.Label(opt_frame, text="적용할 프리셋:").grid(row=0, column=0, sticky="w")
        self.preset_combo = ttk.Combobox(opt_frame, state="readonly", width=30)
        self.preset_combo.grid(row=0, column=1, padx=10, pady=5)
        
        btn_grp = ttk.Frame(opt_frame)
        btn_grp.grid(row=0, column=2)
        ttk.Button(btn_grp, text="새로고침", width=8, command=self.load_presets).pack(side="left", padx=2)
        ttk.Button(btn_grp, text="(Reload)", width=8, command=self.sync_presets).pack(side="left", padx=2)

        self.merge_all = tk.BooleanVar(value=False)
        ttk.Checkbutton(opt_frame, text="모든 결과를 하나의 파일로 합치기", variable=self.merge_all).grid(row=1, column=0, columnspan=2, pady=10, sticky="w")

        # Reference File (Common for all)
        ttk.Label(opt_frame, text="공통 참조 파일 (선택):").grid(row=2, column=0, sticky="w")
        self.ref_path = tk.StringVar()
        ttk.Entry(opt_frame, textvariable=self.ref_path, width=30).grid(row=2, column=1, padx=10, pady=5)
        
        btn_grp_ref = ttk.Frame(opt_frame)
        btn_grp_ref.grid(row=2, column=2)
        ttk.Button(btn_grp_ref, text="찾아보기", width=8, command=lambda: self.browse_file(self.ref_path)).pack(side="left", padx=2)
        ttk.Button(btn_grp_ref, text="(Cloud)", width=8, command=self.secure_upload_handler).pack(side="left", padx=2)

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
        presets = self.preset_manager.load_all()
        self.preset_combo['values'] = list(presets.keys())

    def sync_presets(self):
        url = self.config.get('registered_sources', {}).get('remote_presets_url', '')
        if not url:
            messagebox.showwarning("경고", "원격 프리셋 URL이 설정되어 있지 않습니다.\n관리자 설정에서 등록하세요.")
            return
        
        def task():
            try:
                self.prog_var.set("프리셋 동기화 중...")
                token = self.config.get('registered_sources', {}).get('github_token', '')
                success, msg, _count = self.preset_manager.sync_from_remote(url, token)
                if success:
                    self.load_presets()
                    messagebox.showinfo("동기화 성공", msg)
                    self.prog_var.set("동기화 완료")
                else:
                    messagebox.showerror("동기화 실패", msg)
                    self.prog_var.set("실패")
            except Exception as e:
                messagebox.showerror("오류", str(e))
                self.prog_var.set("실패")
        
        threading.Thread(target=task, daemon=True).start()

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
                presets = self.preset_manager.load_all()
                preset = presets.get(preset_name)
                if not preset:
                    messagebox.showerror("오류", f"'{preset_name}' 프리셋을 찾을 수 없습니다.")
                    return

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
                    df_res, _diag = DataEngine.apply_filters(df_res, f_config)

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
                    try:
                        ExcelHandler.write_to_active_excel(final_df, "MergedResult")
                    except: pass
                    
                    # Also save a temp copy for Cloud Sync if they choose
                    date_str = datetime.now().strftime("%Y%m%d_%H%M%S")
                    temp_dir = os.path.join(os.path.abspath(os.curdir), "temp")
                    os.makedirs(temp_dir, exist_ok=True)
                    temp_out = os.path.join(temp_dir, f"BatchMerged_{date_str}.xlsx")
                    ExcelHandler.save_to_file(final_df, temp_out)
                    
                    self.after(0, lambda: self.show_batch_result_popup(f"배치 완료 (총 {len(files)}개 파일)", temp_out))
                else:
                    self.after(0, lambda: self.show_batch_result_popup(f"배치 완료 (총 {len(files)}개 파일)", None))
                
                self.prog_var.set("완료됨")
            except Exception as e:
                messagebox.showerror("오류", f"배치 처리 중 오류 발생: {e}")
            finally:
                self.run_btn.config(state="normal")

        threading.Thread(target=task, daemon=True).start()

    def show_batch_result_popup(self, summary_msg, saved_path):
        popup = tk.Toplevel(self)
        popup.title("배치 처리 작업 완료")
        popup.geometry("450x300")
        popup.resizable(False, False)
        popup.grab_set()
        
        # Center
        sw = self.winfo_screenwidth()
        sh = self.winfo_screenheight()
        popup.geometry(f"450x300+{sw//2-225}+{sh//2-150}")
        
        main = ttk.Frame(popup, padding=20)
        main.pack(fill="both", expand=True)
        
        ttk.Label(main, text="[OK] 배치 처리 성공", font=("System", 12, "bold"), foreground="#28A745").pack(pady=(0, 10))
        ttk.Label(main, text=summary_msg, font=("System", 10)).pack(pady=(0, 20))
        
        # Cloud Sync Section
        btn_frame = ttk.Frame(main)
        btn_frame.pack(fill="x", side="bottom")

        def upload_to_cloud():
            if not saved_path or not os.path.exists(saved_path):
                messagebox.showwarning("정보", "병합 저장된 파일이 없거나 찾을 수 없어 클라우드 전송을 수행할 수 없습니다.")
                return
            
            master = self.winfo_toplevel()
            url = master.config['registered_sources'].get('github_url', '')
            token = master.config['registered_sources'].get('github_token', '')
            
            if not url or not token:
                messagebox.showwarning("설정 필요", "관리자 설정에서 GitHub URL과 토큰을 먼저 등록해 주세요.")
                return

            def upload_task():
                try:
                    self.prog_var.set("클라우드 업로드 중...")
                    success, msg = GitHubSync.upload_file(token, url, saved_path)
                    if success:
                        messagebox.showinfo("업로드 성공", msg)
                    else:
                        messagebox.showerror("업로드 실패", msg)
                except Exception as e:
                    messagebox.showerror("오류", str(e))
                finally:
                    self.prog_var.set("완료")

            threading.Thread(target=upload_task, daemon=True).start()

        if saved_path:
            sync_btn = ttk.Button(btn_frame, text="(Cloud) 병합 결과 깃허브 업로드 (Sync)", style="Accent.TButton", command=upload_to_cloud)
            sync_btn.pack(fill="x", pady=5)
        
        ttk.Button(btn_frame, text="닫기", command=popup.destroy).pack(fill="x", pady=5)

    def secure_upload_handler(self):
        """Security-verified upload for the common reference file."""
        path_to_upload = self.ref_path.get()
        if not path_to_upload or not os.path.exists(path_to_upload):
            messagebox.showwarning("파일 없음", "업로드할 참조 파일을 먼저 선택해 주세요.")
            return

        # 1. Identity Verification
        pwd = simpledialog.askstring("동기화 보안 확인", "전송 비밀번호를 입력하세요:", show="*")
        if pwd != "3867":
            if pwd: messagebox.showerror("인증 실패", "비밀번호가 올바르지 않습니다.")
            return

        # 2. Perform Upload
        master = self.winfo_toplevel()
        url = master.config['registered_sources'].get('github_url', '')
        token = master.config['registered_sources'].get('github_token', '')

        if not url or not token:
            messagebox.showwarning("설정 필요", "관리자 설정에서 GitHub URL과 토큰을 먼저 등록해 주세요.")
            return

        def upload_task():
            try:
                self.prog_var.set("클라우드 전송 중...")
                success, msg = GitHubSync.upload_file(token, url, path_to_upload)
                if success:
                    messagebox.showinfo("전송 성공", f"보안 전송 완료!\n{msg}")
                else:
                    messagebox.showerror("전송 실패", msg)
            except Exception as e:
                messagebox.showerror("오류", str(e))
            finally:
                self.prog_var.set("완료")

        threading.Thread(target=upload_task, daemon=True).start()
