import pandas as pd
import xlwings as xw
from pathlib import Path
import tkinter as tk
from tkinter import messagebox

def normalize_text(x):
    if pd.isna(x):
        return ""
    return str(x).strip()

def safe_sheet_name(name):
    invalid = ['\\', '/', '?', '*', '[', ']', ':']
    for ch in invalid:
        name = name.replace(ch, "_")
    return name[:31]

class ExcelHandler:
    @staticmethod
    def read_file(path, sheet_name=None):
        ext = Path(path).suffix.lower()
        if ext == ".csv":
            try:
                return pd.read_csv(path, dtype=object, encoding="utf-8-sig", low_memory=False)
            except UnicodeDecodeError:
                return pd.read_csv(path, dtype=object, encoding="cp949", low_memory=False)
        
        if ext in [".xlsx", ".xlsm", ".xls"]:
            return pd.read_excel(path, sheet_name=sheet_name, dtype=object)
        
        raise ValueError(f"Unsupported file format: {ext}")

    @staticmethod
    def get_sheet_names(path):
        ext = Path(path).suffix.lower()
        if ext in [".xlsx", ".xlsm", ".xls"]:
            xl = pd.ExcelFile(path)
            return xl.sheet_names
        return []

    @staticmethod
    def write_to_active_excel(df, sheet_name_base="Result", bold_header=True):
        try:
            app = xw.apps.active
            if not app:
                app = xw.App(visible=True)
            
            # Optimization for Active Excel
            app.screen_updating = False
            app.display_alerts = False
            
            wb = app.books.active
            
            sheet_name = safe_sheet_name(sheet_name_base)
            existing = [s.name for s in wb.sheets]
            
            if sheet_name in existing:
                idx = 2
                while f"{sheet_name}_{idx}" in existing:
                    idx += 1
                sheet_name = f"{sheet_name}_{idx}"
            
            ws = wb.sheets.add(after=wb.sheets[-1], name=sheet_name)
            
            # Optimization: Disable calculation during write
            initial_calc = app.calculation
            app.calculation = 'manual'
            
            try:
                # Write to Excel using chunksize for memory efficiency
                # Passing DataFrame directly to xlwings is better than converting to list
                ws.range("A1").options(index=False, chunksize=5000).value = df
            finally:
                app.calculation = initial_calc
            
            # Formatting
            ws.used_range.columns.autofit()
            if bold_header:
                header_range = ws.range("1:1") # Use row indexing for header
                header_range.api.Font.Bold = True
                header_range.color = (220, 230, 241)
            
            ws.activate()
            # Restore Settings
            app.screen_updating = True
            app.display_alerts = True
            return sheet_name
        except Exception as e:
            try:
                if 'app' in locals():
                    app.screen_updating = True
                    app.display_alerts = True
            except: pass
            raise Exception(f"Excel Export Error: {e}")

    @staticmethod
    def save_to_file(df, path):
        """Direct file saving (CSV or XLSX) - much faster for 900k+ rows."""
        ext = Path(path).suffix.lower()
        if ext == ".csv":
            df.to_csv(path, index=False, encoding="utf-8-sig")
        else:
            # Use XlsxWriter engine if possible for speed
            df.to_excel(path, index=False, engine="xlsxwriter")
        return path

    @staticmethod
    def detect_special_sheets(wb):
        """Logic to detect 'Source' and 'Reference' sheets based on headers."""
        left = None
        right = None
        for sht in wb.sheets:
            try:
                headers = sht.range("A1").expand("right").value
                if not headers: continue
                if not isinstance(headers, list): headers = [headers]
                
                header_set = {str(h).strip() for h in headers if h}
                
                # Heuristics from original code
                if {"관리본부명", "시설구분", "요금구분"}.issubset(header_set):
                    left = sht
                if any(x in header_set for x in ["추천자명", "추천자유형", "계약번호"]):
                    if left != sht:
                        right = sht
            except: continue
        
        if not left: left = wb.sheets[0]
        if not right and len(wb.sheets) > 1: right = wb.sheets[1]
        elif not right: right = wb.sheets[0]
            
        return left, right
