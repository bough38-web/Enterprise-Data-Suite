import tkinter as tk
from tkinter import ttk

def get_app_fonts(widget):
    """Helper to get standardized fonts from the main application."""
    root = widget.winfo_toplevel()
    if hasattr(root, 'fonts'):
        return root.fonts
    # Fallback fonts
    family = "Segoe UI Variable Text" if root.tk.call('tk', 'windowingsystem') == 'win32' else "Inter"
    return {
        "h1": (family, 13, "bold"),
        "h2": (family, 11, "bold"),
        "normal": (family, 10),
        "small": (family, 9),
        "mono": ("Consolas", 10)
    }

def get_scaling_factor(widget):
    """Get the current DPI scaling factor."""
    root = widget.winfo_toplevel()
    if hasattr(root, 'scaling_factor'):
        return root.scaling_factor
    try:
        return float(root.tk.call('tk', 'scaling')) / 1.3333333333333333
    except:
        return 1.0

class ScrollableFrame(ttk.Frame):
    def __init__(self, container, horizontal=True, *args, **kwargs):
        super().__init__(container, *args, **kwargs)
        self.canvas = tk.Canvas(self, highlightthickness=0, bg="#1E1E1E") # Default dark
        self.v_scrollbar = ttk.Scrollbar(self, orient="vertical", command=self.canvas.yview)
        self.h_scrollbar = ttk.Scrollbar(self, orient="horizontal", command=self.canvas.xview)
        self.scrollable_frame = ttk.Frame(self.canvas)
        
        # Function to sync background color with current theme
        def sync_bg():
            try:
                # Try to get background from parent or a known style
                root = self.winfo_toplevel()
                if hasattr(root, 'cget'):
                    self.canvas.configure(bg=root.cget('bg'))
            except: pass
            
        sync_bg()
        self.after(100, sync_bg) # Double check after render

        self.scrollable_frame.bind(
            "<Configure>",
            lambda e: self.canvas.configure(
                scrollregion=self.canvas.bbox("all")
            )
        )

        self.canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")
        self.canvas.configure(yscrollcommand=self.v_scrollbar.set, xscrollcommand=self.h_scrollbar.set)

        self.v_scrollbar.pack(side="right", fill="y")
        if horizontal:
            self.h_scrollbar.pack(side="bottom", fill="x")
        self.canvas.pack(side="left", fill="both", expand=True)

class ValueFilterPopup:
    def __init__(self, parent, title, values, selected_values):
        self.top = tk.Toplevel(parent)
        self.top.title(title)
        
        # Scale geometry based on DPI
        sf = get_scaling_factor(parent)
        w = int(450 * sf)
        h = int(600 * sf)
        self.top.geometry(f"{w}x{h}")
        
        self.top.transient(parent)
        self.top.grab_set()

        # Set theme background immediately
        try: self.top.configure(bg=parent.winfo_toplevel().cget('bg'))
        except: pass

        self.result = None
        self.vars = {}

        fonts = get_app_fonts(self.top)
        main = ttk.Frame(self.top, padding=12)
        main.pack(fill="both", expand=True)

        ttk.Label(main, text=title, font=fonts["h2"]).pack(anchor="w", pady=(0, 10))

        # Search box
        search_frame = ttk.Frame(main)
        search_frame.pack(fill="x", pady=(0, 10))
        ttk.Label(search_frame, text="검색:").pack(side="left")
        self.search_var = tk.StringVar()
        self.search_var.trace_add("write", self.refresh_list)
        ttk.Entry(search_frame, textvariable=self.search_var).pack(side="left", fill="x", expand=True, padx=8)

        # Buttons
        btn_frame = ttk.Frame(main)
        btn_frame.pack(fill="x", pady=(0, 10))
        ttk.Button(btn_frame, text="전체 선택", command=self.select_all).pack(side="left", padx=2)
        ttk.Button(btn_frame, text="전체 해제", command=self.unselect_all).pack(side="left", padx=2)

        # List
        self.scroll_frame = ScrollableFrame(main)
        self.scroll_frame.pack(fill="both", expand=True)

        self.all_values = sorted(list(values))
        selected_set = set(selected_values or [])
        for v in self.all_values:
            self.vars[v] = tk.BooleanVar(value=(v in selected_set))

        self.refresh_list()

        # Bottom buttons
        bottom = ttk.Frame(main)
        bottom.pack(fill="x", pady=(12, 0))
        ttk.Button(bottom, text="취소", command=self.cancel).pack(side="right", padx=4)
        ttk.Button(bottom, text="적용", command=self.apply).pack(side="right", padx=4)

        self.top.wait_window()

    def refresh_list(self, *args):
        for w in self.scroll_frame.scrollable_frame.winfo_children():
            w.destroy()
        
        keyword = self.search_var.get().strip().lower()
        cols = 2
        filtered = [v for v in self.all_values if not keyword or keyword in str(v).lower()]
        
        for idx, val in enumerate(filtered):
            chk = ttk.Checkbutton(self.scroll_frame.scrollable_frame, text=str(val), variable=self.vars[val])
            chk.grid(row=idx//cols, column=idx%cols, sticky="w", padx=10, pady=4)

    def select_all(self):
        for var in self.vars.values(): var.set(True)

    def unselect_all(self):
        for var in self.vars.values(): var.set(False)

    def apply(self):
        self.result = [v for v, var in self.vars.items() if var.get()]
        self.top.destroy()

    def cancel(self):
        self.result = None
        self.top.destroy()

class SheetSelectPopup:
    def __init__(self, parent, title, sheet_names):
        self.result = None
        self.top = tk.Toplevel(parent)
        self.top.title(title)
        
        # Scale geometry based on DPI
        sf = get_scaling_factor(parent)
        w = int(350 * sf)
        h = int(180 * sf)
        self.top.geometry(f"{w}x{h}")
        
        self.top.transient(parent)
        self.top.grab_set()
        
        # Set theme background immediately
        try: self.top.configure(bg=parent.winfo_toplevel().cget('bg'))
        except: pass

        fonts = get_app_fonts(self.top)
        frame = ttk.Frame(self.top, padding=20)
        frame.pack(fill="both", expand=True)

        ttk.Label(frame, text="작업할 시트를 선택하세요:", font=fonts["normal"]).pack(anchor="w", pady=(0, 10))
        
        self.var = tk.StringVar(value=sheet_names[0] if sheet_names else "")
        combo = ttk.Combobox(frame, textvariable=self.var, values=sheet_names, state="readonly")
        combo.pack(fill="x", pady=(0, 20))

        btn_pw = ttk.Frame(frame)
        btn_pw.pack(fill="x")
        ttk.Button(btn_pw, text="취소", command=self.top.destroy).pack(side="right", padx=4)
        ttk.Button(btn_pw, text="확인", command=self.apply).pack(side="right", padx=4)

        self.top.wait_window()

    def apply(self):
        self.result = self.var.get()
        self.top.destroy()

class HelpPopup:
    def __init__(self, parent, title, content):
        self.top = tk.Toplevel(parent)
        self.top.title(f"도움말: {title}")
        
        # Scale geometry based on DPI
        sf = get_scaling_factor(parent)
        w = int(400 * sf)
        h = int(300 * sf)
        self.top.geometry(f"{w}x{h}")
        
        self.top.transient(parent)
        self.top.grab_set()

        # Set theme background immediately
        try:
            bg_color = parent.winfo_toplevel().cget('bg')
            self.top.configure(bg=bg_color)
        except: pass

        main = ttk.Frame(self.top, padding=20)
        main.pack(fill="both", expand=True)

        fonts = get_app_fonts(self.top)
        header = ttk.Frame(main)
        header.pack(fill="x", pady=(0, 10))
        ttk.Label(header, text=title, font=fonts["h1"], foreground="#0078D4").pack(side="left")

        # Use a text widget for multi-line content with tag support
        txt = tk.Text(main, wrap="word", font=fonts["normal"], bg=self.top.cget('bg'), 
                      fg=fonts.get("fg", "black"), relief="flat", padx=10, pady=10)
        
        # Configure tags for basic markdown-like formatting
        txt.tag_configure("bold", font=(fonts["normal"][0], fonts["normal"][1], "bold"))
        txt.tag_configure("h3", font=(fonts["normal"][0], fonts["normal"][1]+1, "bold"), foreground="#0078D4")
        txt.tag_configure("gray", foreground="#666666")
        
        # Simple parser for basic formatting
        lines = content.split('\n')
        for line in lines:
            if line.startswith('### '):
                txt.insert("end", line[4:] + "\n", "h3")
            elif line.startswith('**') and line.endswith('**'):
                txt.insert("end", line[2:-2] + "\n", "bold")
            elif line.startswith('- '):
                txt.insert("end", "  • ", "gray")
                txt.insert("end", line[2:] + "\n")
            else:
                txt.insert("end", line + "\n")
                
        txt.config(state="disabled")
        txt.pack(fill="both", expand=True)

        btn = ttk.Button(main, text="닫기", command=self.top.destroy)
        btn.pack(pady=(15, 0))

        # Center on screen
        self.top.update_idletasks()
        w = self.top.winfo_width()
        h = self.top.winfo_height()
        x = parent.winfo_rootx() + (parent.winfo_width() // 2) - (w // 2)
        y = parent.winfo_rooty() + (parent.winfo_height() // 2) - (h // 2)
        self.top.geometry(f"+{x}+{y}")

def create_help_btn(parent, title, content):
    """Helper to create a consistent '?' button that launches a HelpPopup."""
    fonts = get_app_fonts(parent)
    btn = ttk.Label(parent, text="?", background="#E1E1E1", foreground="#666666", 
                    font=fonts["small"], cursor="hand2", padding=(4, 0))
    btn.bind("<Button-1>", lambda e: HelpPopup(parent.winfo_toplevel(), title, content))
    return btn

class CloudExplorerPopup:
    def __init__(self, parent, token, repo_url):
        self.result = None
        self.token = token
        self.repo_url = repo_url
        
        self.top = tk.Toplevel(parent)
        self.top.title("GitHub 구름 탐색기")
        
        # Scale geometry based on DPI
        sf = get_scaling_factor(parent)
        w = int(500 * sf)
        h = int(600 * sf)
        self.top.geometry(f"{w}x{h}")
        
        self.top.transient(parent)
        self.top.grab_set()
        
        try: self.top.configure(bg=parent.winfo_toplevel().cget('bg'))
        except: pass

        main = ttk.Frame(self.top, padding=20)
        main.pack(fill="both", expand=True)

        fonts = get_app_fonts(self.top)
        header = ttk.Frame(main)
        header.pack(fill="x", pady=(0, 15))
        ttk.Label(header, text="☁️ 업로드된 파일 모아보기", font=fonts["h2"]).pack(side="left")
        ttk.Button(header, text="새로고침", command=self.load_data, width=10).pack(side="right")

        # Treeview
        tree_frame = ttk.Frame(main)
        tree_frame.pack(fill="both", expand=True)
        
        self.tree = ttk.Treeview(tree_frame, columns=("size"), show="tree")
        self.tree.pack(side="left", fill="both", expand=True)
        
        vsb = ttk.Scrollbar(tree_frame, orient="vertical", command=self.tree.yview)
        vsb.pack(side="right", fill="y")
        self.tree.configure(yscrollcommand=vsb.set)
        
        self.tree.bind("<Double-1>", lambda e: self.apply())

        btn_frame = ttk.Frame(main)
        btn_frame.pack(fill="x", pady=(20, 0))
        
        ttk.Button(btn_frame, text="취소", command=self.top.destroy, width=12).pack(side="right", padx=5)
        ttk.Button(btn_frame, text="파일 로드 (Load)", command=self.apply, style="Accent.TButton", width=15).pack(side="right", padx=5)

        self.load_data()
        self.top.wait_window()

    def load_data(self):
        # Clear tree
        for i in self.tree.get_children():
            self.tree.delete(i)
            
        from utils.github_sync import GitHubSync
        
        # Get advanced network config from master
        master = self.top.master.winfo_toplevel()
        net_cfg = {}
        if hasattr(master, 'config'):
            net_cfg = master.config.get('network', {})

        success, data = GitHubSync.list_files(self.token, self.repo_url, network_config=net_cfg)
        
        if not success:
            from tkinter import messagebox
            messagebox.showerror("오류", f"파일 목록을 가져오지 못했습니다:\n{data}")
            return
            
        # Parse flat paths into a hierarchy
        # uploads/YYYY-MM-DD/filename.xlsx
        folders = {}
        for item in data:
            parts = item['path'].split('/')
            if len(parts) >= 2:
                folder_name = parts[1] # YYYY-MM-DD
                file_name = parts[-1]
                
                if folder_name not in folders:
                    folders[folder_name] = self.tree.insert("", "end", text=f"📂 {folder_name}", open=True)
                
                self.tree.insert(folders[folder_name], "end", text=f"📄 {file_name}", values=(item['raw_url'],))

    def apply(self):
        selected = self.tree.selection()
        if not selected:
            return
            
        item = self.tree.item(selected[0])
        # If it's a file (has raw_url in values)
        if item['values']:
            self.result = item['values'][0]
            self.top.destroy()
