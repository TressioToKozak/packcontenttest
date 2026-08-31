from pathlib import Path
import re

path = Path('pack data/2.5.PY')
text = path.read_text(encoding='utf-8')

text = text.replace(
    'NAVY_DARK = "#2d354d"      # nagłówek / status\nNAVY_LIGHT = "#2d354d"     # gradient\nPANEL_BG = "#F2F4F7"       # panel sterowania / tło główne\nWHITE = "#FFFFFF"\nBTN_BLUE = "#0078D4"\nBTN_GREEN = "#28A745"\nBTN_RED = "#DC3545"\nBTN_GREY = "#6C757D"',
    'NAVY_DARK = "#1E2A44"\nPANEL_BG = "#F4F6F9"\nWHITE = "#FFFFFF"\nTEXT_DARK = "#182230"\nTEXT_MUTED = "#667085"\nBORDER = "#DDE3EA"\nSOFT_BLUE = "#EEF4FF"\nBTN_BLUE = "#1769E0"\nBTN_GREEN = "#0B8F68"\nBTN_RED = "#C73A3A"\nBTN_GREY = "#667085"'
)
text = text.replace('self.root.geometry("1200x700")', 'self.root.geometry("1280x760")')
text = text.replace('self.root.minsize(980, 600)', 'self.root.minsize(1060, 650)')

styles = '''    def _configure_styles(self):
        style = ttk.Style(self.root)
        style.theme_use("clam")
        style.configure("App.TEntry", padding=10, fieldbackground=WHITE, foreground=TEXT_DARK)
        style.configure("Treeview", background=WHITE, fieldbackground=WHITE,
                        foreground=TEXT_DARK, rowheight=36, font=("Segoe UI", 10), borderwidth=0)
        style.configure("Treeview.Heading", background="#F8FAFC", foreground=NAVY_DARK,
                        font=("Segoe UI Semibold", 10), padding=(10, 11), relief="flat", borderwidth=0)
        style.map("Treeview", background=[("selected", "#DCEAFF")], foreground=[("selected", TEXT_DARK)])
        style.map("Treeview.Heading", background=[("active", "#EFF3F8")])

'''
text, n = re.subn(r'    def _configure_styles\(self\):.*?(?=    def build_header\(self\):)', styles, text, flags=re.S)
assert n == 1, n

header = '''    def build_header(self):
        shell = tk.Frame(self.root, bg=WHITE, height=94)
        shell.pack(fill=tk.X)
        shell.pack_propagate(False)

        brand = tk.Frame(shell, bg=WHITE)
        brand.pack(side=tk.LEFT, fill=tk.Y, padx=(28, 0))

        self.logo_img = self.wczytaj_logo(LOGO_PATH)
        if self.logo_img:
            tk.Label(brand, image=self.logo_img, bg=WHITE).pack(side=tk.LEFT, padx=(0, 22), pady=14)

        title = tk.Frame(brand, bg=WHITE)
        title.pack(side=tk.LEFT, pady=18)
        tk.Label(title, text="Pack Contents Reprint", bg=WHITE, fg=TEXT_DARK,
                 font=("Segoe UI Semibold", 21)).pack(anchor="w")
        tk.Label(title, text="Warehouse label management", bg=WHITE, fg=TEXT_MUTED,
                 font=("Segoe UI", 10)).pack(anchor="w", pady=(2, 0))

        right = tk.Frame(shell, bg=WHITE)
        right.pack(side=tk.RIGHT, padx=28)
        tk.Label(right, text="v2.5", bg=WHITE, fg=TEXT_MUTED,
                 font=("Segoe UI", 9)).pack(side=tk.RIGHT, padx=(0, 14), pady=31)
        chip = tk.Frame(right, bg="#F2F8F5", highlightbackground="#D7E9DF", highlightthickness=1)
        chip.pack(side=tk.RIGHT, pady=24)
        tk.Label(chip, text="●  READY", bg="#F2F8F5", fg="#35624A",
                 padx=12, pady=6, font=("Segoe UI Semibold", 9)).pack()

        tk.Frame(self.root, bg=BTN_BLUE, height=3).pack(fill=tk.X)

'''
text, n = re.subn(r'    def build_header\(self\):.*?(?=    def build_dashboard\(self\):)', header, text, flags=re.S)
assert n == 1, n

dashboard = '''    def build_dashboard(self):
        self.main_frame = tk.Frame(self.root, bg=PANEL_BG)
        self.main_frame.pack(fill=tk.BOTH, expand=True, padx=28, pady=(24, 18))

        panel = tk.Frame(self.main_frame, bg=WHITE, highlightbackground=BORDER, highlightthickness=1)
        panel.pack(fill=tk.X, pady=(0, 18))

        heading = tk.Frame(panel, bg=WHITE)
        heading.grid(row=0, column=0, columnspan=7, sticky="ew", padx=20, pady=(18, 14))
        tk.Label(heading, text="Container lookup", bg=WHITE, fg=TEXT_DARK,
                 font=("Segoe UI Semibold", 13)).pack(anchor="w")
        tk.Label(heading, text="Search one container or retrieve all containers assigned to the same order.",
                 bg=WHITE, fg=TEXT_MUTED, font=("Segoe UI", 9)).pack(anchor="w", pady=(3, 0))

        tk.Label(panel, text="CONTAINER", bg=WHITE, fg=TEXT_MUTED,
                 font=("Segoe UI Semibold", 8)).grid(row=1, column=0, sticky="w", padx=(20, 8))
        tk.Label(panel, text="ALL CONTAINERS FOR ORDER", bg=WHITE, fg=TEXT_MUTED,
                 font=("Segoe UI Semibold", 8)).grid(row=1, column=2, sticky="w", padx=(18, 8))

        self.search_entry = ttk.Entry(panel, style="App.TEntry", font=("Segoe UI", 11))
        self.search_entry.grid(row=2, column=0, sticky="ew", padx=(20, 8), pady=(5, 20))
        self.order_entry = ttk.Entry(panel, style="App.TEntry", font=("Segoe UI", 11))
        self.order_entry.grid(row=2, column=2, sticky="ew", padx=(18, 8), pady=(5, 20))
        self.search_entry.bind("<Return>", lambda _event: self.szukaj_container())
        self.order_entry.bind("<Return>", lambda _event: self.szukaj_order_po_containerze())

        self.buttons = []
        def button(text, color, active, command, column, fg=WHITE):
            btn = tk.Button(panel, text=text, command=command, bg=color, fg=fg,
                            activebackground=active, activeforeground=fg, relief="flat",
                            bd=0, padx=17, pady=10, font=("Segoe UI Semibold", 9), cursor="hand2")
            btn.grid(row=2, column=column, padx=8, pady=(5, 20), sticky="ew")
            self.buttons.append(btn)

        button("Search", BTN_BLUE, "#0F5BC8", self.szukaj_container, 1)
        button("All containers", NAVY_DARK, "#131E33", self.szukaj_order_po_containerze, 3)
        button("Print", BTN_GREEN, "#087A59", self.obsluga_druku, 5)
        button("Clear", "#EEF1F5", "#E2E7ED", self.odswiez, 6, TEXT_DARK)

        panel.columnconfigure(0, weight=3)
        panel.columnconfigure(2, weight=3)

'''
text, n = re.subn(r'    def build_dashboard\(self\):.*?(?=    def build_table\(self\):)', dashboard, text, flags=re.S)
assert n == 1, n

table = '''    def build_table(self):
        table_frame = tk.Frame(self.main_frame, bg=WHITE, highlightbackground=BORDER, highlightthickness=1)
        table_frame.pack(fill=tk.BOTH, expand=True)

        heading = tk.Frame(table_frame, bg=WHITE)
        heading.pack(fill=tk.X, padx=18, pady=(15, 12))
        title_block = tk.Frame(heading, bg=WHITE)
        title_block.pack(side=tk.LEFT)
        tk.Label(title_block, text="Results", bg=WHITE, fg=TEXT_DARK,
                 font=("Segoe UI Semibold", 13)).pack(anchor="w")
        tk.Label(title_block, text="Current pack contents", bg=WHITE, fg=TEXT_MUTED,
                 font=("Segoe UI", 9)).pack(anchor="w", pady=(2, 0))

        self.result_count = tk.Label(heading, text="0 rows", bg=SOFT_BLUE, fg=BTN_BLUE,
                                     padx=10, pady=5, font=("Segoe UI Semibold", 9))
        self.result_count.pack(side=tk.RIGHT)
        tk.Frame(table_frame, bg=BORDER, height=1).pack(fill=tk.X)

        grid = tk.Frame(table_frame, bg=WHITE)
        grid.pack(fill=tk.BOTH, expand=True)
        columns = ("Order", "SKU", "Qty", "Container", "Description", "PO")
        self.tree = ttk.Treeview(grid, columns=columns, show="headings")
        widths = {"Order": 145, "SKU": 140, "Qty": 75, "Container": 150, "Description": 390, "PO": 145}
        for col in columns:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=widths[col], minwidth=65,
                             anchor=tk.CENTER if col == "Qty" else tk.W,
                             stretch=col == "Description")
        self.tree.tag_configure("even", background="#FAFBFC")
        self.tree.tag_configure("missing_po", foreground="#A15C00")
        vsb = ttk.Scrollbar(grid, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=vsb.set)
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        vsb.pack(side=tk.RIGHT, fill=tk.Y)
        self.empty_label = tk.Label(grid, bg=WHITE, fg="#8A94A6", font=("Segoe UI", 11))

'''
text, n = re.subn(r'    def build_table\(self\):.*?(?=    def build_statusbar\(self\):)', table, text, flags=re.S)
assert n == 1, n

status = '''    def build_statusbar(self):
        frame = tk.Frame(self.root, bg=WHITE, highlightbackground=BORDER, highlightthickness=1)
        frame.pack(fill=tk.X, side=tk.BOTTOM)
        tk.Label(frame, text="●", bg=WHITE, fg=BTN_GREEN, font=("Segoe UI", 8)).pack(side=tk.LEFT, padx=(24, 0), pady=8)
        self.status_label = tk.Label(frame, text="Ready", bg=WHITE, fg=TEXT_MUTED,
                                     anchor="w", font=("Segoe UI", 9))
        self.status_label.pack(side=tk.LEFT, padx=(7, 24), pady=8)

'''
text, n = re.subn(r'    def build_statusbar\(self\):.*?(?=    def set_status\(self, text, count=None\):)', status, text, flags=re.S)
assert n == 1, n

logo = '''    def wczytaj_logo(self, path):
        candidates = [path, os.path.join(os.path.dirname(os.path.abspath(__file__)), "logo.png")]
        for candidate in candidates:
            try:
                if not os.path.exists(candidate):
                    continue
                img = Image.open(candidate)
                img.thumbnail((185, 56), Image.LANCZOS)
                return ImageTk.PhotoImage(img)
            except Exception as exc:
                print("Error with logo:", exc)
        return None

'''
text, n = re.subn(r'    def wczytaj_logo\(self, path\):.*?(?=    def wczytaj_dane\(self\):)', logo, text, flags=re.S)
assert n == 1, n

old = '    root.iconbitmap(r"D:\\OneDrive - CEVA Logistics\\Europe-CL-ResMed Operations - Pack Contents reprint\\Data\\pack data\\ceva.ico")    '
new = '''    icon_path = r"D:\\OneDrive - CEVA Logistics\\Europe-CL-ResMed Operations - Pack Contents reprint\\Data\\pack data\\ceva.ico"
    if not os.path.exists(icon_path):
        icon_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "Ceva.ico")
    try:
        root.iconbitmap(icon_path)
    except (tk.TclError, OSError):
        pass'''
assert old in text
text = text.replace(old, new, 1)

path.write_text(text, encoding='utf-8')
