import tkinter as tk
from tkinter import ttk

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
        scaling = parent.tk.call('tk', 'scaling')
        w = int(450 * (scaling / 1.33))
        h = int(600 * (scaling / 1.33))
        self.top.geometry(f"{w}x{h}")
        
        self.top.transient(parent)
        self.top.grab_set()

        # Set theme background immediately
        try: self.top.configure(bg=parent.winfo_toplevel().cget('bg'))
        except: pass

        self.result = None
        self.vars = {}

        main = ttk.Frame(self.top, padding=12)
        main.pack(fill="both", expand=True)

        ttk.Label(main, text=title, font=("System", 11, "bold")).pack(anchor="w", pady=(0, 10))

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
        scaling = parent.tk.call('tk', 'scaling')
        w = int(350 * (scaling / 1.33))
        h = int(180 * (scaling / 1.33))
        self.top.geometry(f"{w}x{h}")
        
        self.top.transient(parent)
        self.top.grab_set()
        
        # Set theme background immediately
        try: self.top.configure(bg=parent.winfo_toplevel().cget('bg'))
        except: pass

        frame = ttk.Frame(self.top, padding=20)
        frame.pack(fill="both", expand=True)

        ttk.Label(frame, text="작업할 시트를 선택하세요:", font=("System", 10)).pack(anchor="w", pady=(0, 10))
        
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
        scaling = parent.tk.call('tk', 'scaling')
        w = int(400 * (scaling / 1.33))
        h = int(300 * (scaling / 1.33))
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

        header = ttk.Frame(main)
        header.pack(fill="x", pady=(0, 10))
        ttk.Label(header, text=title, font=("System", 12, "bold"), foreground="#0078D4").pack(side="left")

        # Use a text widget for multi-line content
        txt = tk.Text(main, wrap="word", font=("System", 10), bg="#F5F5F5", relief="flat", padx=10, pady=10)
        txt.insert("1.0", content)
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
    btn = ttk.Label(parent, text="?", background="#E1E1E1", foreground="#666666", 
                    font=("System", 9, "bold"), cursor="hand2", padding=(4, 0))
    btn.bind("<Button-1>", lambda e: HelpPopup(parent.winfo_toplevel(), title, content))
    return btn
