import tkinter as tk
from tkinter import ttk
import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg

class StatsTab(ttk.Frame):
    def __init__(self, parent):
        super().__init__(parent)
        self.df = None
        self.build_ui()

    def build_ui(self):
        main = ttk.Frame(self, padding=20)
        main.pack(fill="both", expand=True)

        header = ttk.Frame(main)
        header.pack(fill="x", pady=(0, 20))
        ttk.Label(header, text="데이터 통계 분석 (Data Insights)", font=("System", 12, "bold")).pack(side="left")

        # Stats Summary
        self.summary_frame = ttk.LabelFrame(main, text="기본 정보", padding=15)
        self.summary_frame.pack(fill="x", pady=(0, 20))
        
        self.stats_label = ttk.Label(self.summary_frame, text="데이터를 먼저 로드해 주세요.", font=("System", 10))
        self.stats_label.pack(anchor="w")

        # Chart Section
        chart_lf = ttk.LabelFrame(main, text="데이터 분포 시각화 (Top 10)", padding=15)
        chart_lf.pack(fill="both", expand=True)

        ctrl = ttk.Frame(chart_lf)
        ctrl.pack(fill="x", pady=(0, 10))
        ttk.Label(ctrl, text="분석 컬럼 선택:").pack(side="left")
        self.col_var = tk.StringVar()
        self.col_combo = ttk.Combobox(ctrl, textvariable=self.col_var, state="readonly")
        self.col_combo.pack(side="left", padx=10)
        ttk.Button(ctrl, text="차트 업데이트", command=self.update_chart).pack(side="left")

        self.chart_container = ttk.Frame(chart_lf)
        self.chart_container.pack(fill="both", expand=True)
        self.canvas = None

    def set_data(self, df):
        self.df = df
        if df is not None:
            # Summary stats
            total_rows = len(df)
            total_cols = len(df.columns)
            missing = df.isnull().sum().sum()
            
            summary_txt = (
                f"- 전체 행 수: {total_rows:,}건\n"
                f"- 컬럼 수: {total_cols}개\n"
                f"- 결측치(빈 칸): {missing:,}개"
            )
            self.stats_label.config(text=summary_txt)
            
            # Update combo
            self.col_combo['values'] = df.columns.tolist()
            if len(df.columns) > 0:
                self.col_combo.current(0)
                self.update_chart()

    def update_chart(self):
        if self.df is None: return
        col = self.col_var.get()
        if not col: return

        # Clear old chart
        for w in self.chart_container.winfo_children():
            w.destroy()

        # Generate Plot
        fig, ax = plt.subplots(figsize=(6, 4))
        
        # Cross-platform Korean Font Fix
        import platform
        sys_plat = platform.system()
        if sys_plat == "Darwin":
            plt.rcParams['font.family'] = 'AppleGothic'
        elif sys_plat == "Windows":
            plt.rcParams['font.family'] = 'Malgun Gothic'
        else:
            plt.rcParams['font.family'] = 'NanumGothic'
            
        plt.rcParams['axes.unicode_minus'] = False

        
        # Apply Theme Colors to Chart (Professional Sync)
        theme = "dark" 
        try:
            root = self.winfo_toplevel()
            if hasattr(root, 'config'):
                theme = root.config['branding'].get('theme', 'dark')
        except: pass
        
        palettes = {
            "dark": {"bg": "#1E1E1E", "accent": "#4A90E2"},
            "light": {"bg": "#F8F9FA", "accent": "#4A90E2"},
            "cosmic": {"bg": "#0F172A", "accent": "#818CF8"},
            "graphite": {"bg": "#18181B", "accent": "#71717A"}
        }
        
        p = palettes.get(theme, palettes['dark'])
        bg_color = p['bg']
        fg_color = "white" if theme in ["dark", "cosmic", "graphite"] else "#333333"
        accent_color = p['accent']
        
        fig.patch.set_facecolor(bg_color)
        ax.set_facecolor(bg_color)
        ax.tick_params(colors=fg_color, which='both', labelsize=8)
        ax.xaxis.label.set_color(fg_color)
        ax.yaxis.label.set_color(fg_color)
        ax.title.set_color(fg_color)
        for spine in ax.spines.values():
            spine.set_edgecolor(fg_color)
            
        counts = self.df[col].value_counts().head(10)
        counts.plot(kind='bar', ax=ax, color=accent_color, alpha=0.8) # Softened bar
        ax.set_title(f"[{col}] 상위 10개 분포", fontsize=11, pad=15, fontweight='bold')
        ax.tick_params(axis='x', rotation=45)
        
        # Adjust layout
        plt.tight_layout()

        # Embed in Tkinter
        self.canvas = FigureCanvasTkAgg(fig, master=self.chart_container)
        self.canvas.draw()
        self.canvas.get_tk_widget().pack(fill="both", expand=True)

    def update_theme(self, theme_name):
        """Called by app.py when theme changes to refresh chart colors."""
        self.update_chart()


