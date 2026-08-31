import os
import re
import shutil
import tempfile
import threading
import unicodedata
import uuid

import pandas as pd
import tkinter as tk
from tkinter import messagebox, ttk
import win32print
from PIL import Image, ImageTk


# --- KONFIGURACJA ŚCIEŻEK I DRUKARKI ---
SCIEZKA = r"D:\OneDrive - CEVA Logistics\Europe-CL-ResMed Operations - Pack Contents reprint\Data Pack contents.xlsm"
NAZWA_DRUKARKI = "ts new zebra"
LOGO_PATH = r"D:\OneDrive - CEVA Logistics\Europe-CL-ResMed Operations - Pack Contents reprint\Data\pack data\logo.png"

# --- KOLORY (dashboard magazynowy) ---
NAVY_DARK = "#2d354d"      # nagłówek / status
NAVY_LIGHT = "#2d354d"     # gradient
PANEL_BG = "#F2F4F7"       # panel sterowania / tło główne
WHITE = "#FFFFFF"
BTN_BLUE = "#0078D4"
BTN_GREEN = "#28A745"
BTN_RED = "#DC3545"
BTN_GREY = "#6C757D"



def clean_text(s):
    """
    Normalizuje tekst: unicode NFKC, usuwa NBSP, usuwa końcowe '.0', trim, uppercase.
    Zwraca pusty string dla NaN.
    """
    s = '' if pd.isna(s) else str(s)
    s = unicodedata.normalize('NFKC', s)
    s = s.replace('\xa0', ' ')
    s = s.strip()
    s = re.sub(r'\.0$', '', s)
    return s.upper()


def zpl_text(value):
    """Return printable text that cannot inject Zebra control commands."""
    if pd.isna(value):
        return ""
    return str(value).replace("^", " ").replace("~", " ").replace("\r", " ").replace("\n", " ")


class ExcelSearchApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Pack Contents Reprint")
        self.root.geometry("1200x700")
        self.root.minsize(980, 600)
        self.root.configure(bg=PANEL_BG)
        self.current_data = pd.DataFrame()
        self.spinner_win = None
        self.busy = False
        self._configure_styles()
        self.build_header()
        self.build_dashboard()
        self.build_table()
        self.build_statusbar()
        self._set_empty_state("Enter a container number to search")
        self.root.bind("<Control-p>", lambda _event: self.obsluga_druku())
        self.root.bind("<Escape>", lambda _event: self.odswiez())
        self.root.after(100, self.search_entry.focus_set)

    def _configure_styles(self):
        style = ttk.Style(self.root)
        style.theme_use("clam")
        style.configure("App.TEntry", padding=8, fieldbackground=WHITE)
        style.configure("Treeview", background=WHITE, fieldbackground=WHITE,
                        foreground="#172033", rowheight=32, font=("Segoe UI", 10), borderwidth=0)
        style.configure("Treeview.Heading", background=NAVY_DARK, foreground=WHITE,
                        font=("Segoe UI", 10, "bold"), padding=(8, 10), relief="flat")
        style.map("Treeview", background=[("selected", BTN_BLUE)], foreground=[("selected", WHITE)])
        style.map("Treeview.Heading", background=[("active", "#3D4968")])

    def build_header(self):
        header = tk.Frame(self.root, bg=NAVY_DARK, height=96)
        header.pack(fill=tk.X)
        header.pack_propagate(False)
        self.logo_img = self.wczytaj_logo(LOGO_PATH)
        if self.logo_img:
            tk.Label(header, image=self.logo_img, bg=NAVY_DARK).pack(side=tk.LEFT, padx=(24, 20))
        else:
            tk.Label(header, text="CEVA LOGISTICS", bg=NAVY_DARK, fg=WHITE,
                     font=("Segoe UI", 16, "bold")).pack(side=tk.LEFT, padx=24)
        title = tk.Frame(header, bg=NAVY_DARK)
        title.pack(side=tk.LEFT, pady=19)
        tk.Label(title, text="Pack Contents Reprint", bg=NAVY_DARK, fg=WHITE,
                 font=("Segoe UI", 20, "bold")).pack(anchor="w")
        tk.Label(title, text="Warehouse Label Utility", bg=NAVY_DARK, fg="#C8D2E6",
                 font=("Segoe UI", 10)).pack(anchor="w")
        tk.Label(header, text="OPERATIONAL TOOL  •  v2.5", bg=NAVY_DARK, fg="#C8D2E6",
                 font=("Segoe UI", 9, "bold")).pack(side=tk.RIGHT, padx=24)

    def build_dashboard(self):
        self.main_frame = tk.Frame(self.root, bg=PANEL_BG)
        self.main_frame.pack(fill=tk.BOTH, expand=True, padx=24, pady=(20, 14))
        panel = tk.Frame(self.main_frame, bg=WHITE, highlightbackground="#D8DEE8", highlightthickness=1)
        panel.pack(fill=tk.X, pady=(0, 16))
        tk.Label(panel, text="CONTAINER LOOKUP", bg=WHITE, fg=NAVY_DARK,
                 font=("Segoe UI", 11, "bold")).grid(row=0, column=0, columnspan=6, sticky="w", padx=18, pady=(14, 8))
        tk.Label(panel, text="Container search", bg=WHITE, fg="#465268",
                 font=("Segoe UI", 9, "bold")).grid(row=1, column=0, sticky="w", padx=(18, 8))
        tk.Label(panel, text="Find all containers for order", bg=WHITE, fg="#465268",
                 font=("Segoe UI", 9, "bold")).grid(row=1, column=2, sticky="w", padx=(18, 8))
        self.search_entry = ttk.Entry(panel, style="App.TEntry", font=("Segoe UI", 11))
        self.search_entry.grid(row=2, column=0, sticky="ew", padx=(18, 8), pady=(3, 16))
        self.order_entry = ttk.Entry(panel, style="App.TEntry", font=("Segoe UI", 11))
        self.order_entry.grid(row=2, column=2, sticky="ew", padx=(18, 8), pady=(3, 16))
        self.search_entry.bind("<Return>", lambda _event: self.szukaj_container())
        self.order_entry.bind("<Return>", lambda _event: self.szukaj_order_po_containerze())
        self.buttons = []
        def button(text, color, command, column):
            btn = tk.Button(panel, text=text, command=command, bg=color, fg=WHITE,
                            activebackground=color, activeforeground=WHITE, relief="flat",
                            bd=0, padx=16, pady=9, font=("Segoe UI", 9, "bold"), cursor="hand2")
            btn.grid(row=2, column=column, padx=8, pady=(3, 16), sticky="ew")
            self.buttons.append(btn)
        button("Search", BTN_BLUE, self.szukaj_container, 1)
        button("All containers", BTN_GREEN, self.szukaj_order_po_containerze, 3)
        button("Print", BTN_RED, self.obsluga_druku, 4)
        button("Clear", BTN_GREY, self.odswiez, 5)
        panel.columnconfigure(0, weight=2)
        panel.columnconfigure(2, weight=2)

    def build_table(self):
        table_frame = tk.Frame(self.main_frame, bg=WHITE, highlightbackground="#D8DEE8", highlightthickness=1)
        table_frame.pack(fill=tk.BOTH, expand=True)
        heading = tk.Frame(table_frame, bg=WHITE)
        heading.pack(fill=tk.X)
        tk.Label(heading, text="RESULTS", bg=WHITE, fg=NAVY_DARK,
                 font=("Segoe UI", 11, "bold")).pack(side=tk.LEFT, padx=16, pady=12)
        self.result_count = tk.Label(heading, text="0 rows", bg=WHITE, fg="#68758A", font=("Segoe UI", 9))
        self.result_count.pack(side=tk.RIGHT, padx=16)
        grid = tk.Frame(table_frame, bg=WHITE)
        grid.pack(fill=tk.BOTH, expand=True)
        columns = ("Order", "SKU", "Qty", "Container", "Description", "PO")
        self.tree = ttk.Treeview(grid, columns=columns, show="headings")
        widths = {"Order": 135, "SKU": 135, "Qty": 70, "Container": 145, "Description": 340, "PO": 135}
        for col in columns:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=widths[col], minwidth=60,
                             anchor=tk.CENTER if col == "Qty" else tk.W,
                             stretch=col == "Description")
        self.tree.tag_configure("even", background="#F7F9FC")
        self.tree.tag_configure("missing_po", foreground="#8A5A00")
        vsb = ttk.Scrollbar(grid, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=vsb.set)
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        vsb.pack(side=tk.RIGHT, fill=tk.Y)
        self.empty_label = tk.Label(grid, bg=WHITE, fg="#7A8699", font=("Segoe UI", 11))

    def build_statusbar(self):
        frame = tk.Frame(self.root, bg=NAVY_DARK)
        frame.pack(fill=tk.X, side=tk.BOTTOM)
        self.status_label = tk.Label(frame, text="Ready", bg=NAVY_DARK, fg=WHITE,
                                     anchor="w", font=("Segoe UI", 9))
        self.status_label.pack(fill=tk.X, padx=24, pady=7)

    def set_status(self, text, count=None):
        self.status_label.config(text=f"{text}  •  {count} rows" if count is not None else text)

    def _set_empty_state(self, text=None):
        self.empty_label.place_forget()
        if text:
            self.empty_label.config(text=text)
            self.empty_label.place(relx=.5, rely=.45, anchor="center")
        self.result_count.config(text=f"{len(self.tree.get_children())} rows")

    def _set_busy(self, busy, text="Loading data..."):
        self.busy = busy
        for button in self.buttons:
            button.config(state=tk.DISABLED if busy else tk.NORMAL)
        if busy:
            self.show_spinner(text)
        else:
            self.hide_spinner()

    def show_spinner(self, text="Loading data..."):
        if self.spinner_win is not None:
            return
        self.spinner_win = tk.Toplevel(self.root)
        self.spinner_win.overrideredirect(True)
        self.spinner_win.configure(bg=NAVY_DARK)
        self.spinner_win.transient(self.root)
        w, h = 240, 86
        x = self.root.winfo_x() + (self.root.winfo_width() - w) // 2
        y = self.root.winfo_y() + (self.root.winfo_height() - h) // 2
        self.spinner_win.geometry(f"{w}x{h}+{x}+{y}")
        tk.Label(self.spinner_win, text=text, bg=NAVY_DARK, fg=WHITE,
                 font=("Segoe UI", 10, "bold")).pack(pady=(13, 7))
        bar = ttk.Progressbar(self.spinner_win, mode="indeterminate", length=180)
        bar.pack()
        bar.start(12)

    def hide_spinner(self):
        if self.spinner_win is not None:
            self.spinner_win.destroy()
            self.spinner_win = None

    def wczytaj_logo(self, path):
        try:
            img = Image.open(path)
            img.thumbnail((210, 64), Image.LANCZOS)
            return ImageTk.PhotoImage(img)
        except Exception as exc:
            print("Error with logo:", exc)
            return None

    def wczytaj_dane(self):
        temp_path = None
        try:
            temp_path = tempfile.gettempdir() + f"\\excel_temp_{uuid.uuid4().hex}.xlsx"
            shutil.copy2(SCIEZKA, temp_path)
            df_move = pd.read_excel(temp_path, sheet_name="move task", header=0, dtype=str).iloc[:, :5].copy()
            if df_move.shape[1] < 5:
                raise ValueError("The 'move task' sheet must contain at least five columns.")
            df_move.columns = ["Order", "SKU", "Qty", "Container", "Description"]
            with pd.ExcelFile(temp_path) as xls:
                names = xls.sheet_names
            order_sheet = next((n for n in names if n.strip().lower() == "order header"), None)
            order_sheet = order_sheet or next((n for n in names if n.strip().lower() == "order hedear"), None)
            order_sheet = order_sheet or next((n for n in names if "order" in n.strip().lower() and "header" in n.strip().lower()), None)
            order_sheet = order_sheet or next((n for n in names if "order" in n.strip().lower()), None)
            order_sheet = order_sheet or names[0]
            raw = pd.read_excel(temp_path, sheet_name=order_sheet, header=0, dtype=str)
            if raw.shape[1] >= 2:
                df_order = raw.iloc[:, :2].copy()
                df_order.columns = ["A_Order", "PO"]
            elif raw.shape[1] == 1:
                df_order = pd.DataFrame({"A_Order": raw.iloc[:, 0], "PO": ""})
            else:
                df_order = pd.DataFrame(columns=["A_Order", "PO"])
            df_move["Order"] = df_move["Order"].apply(clean_text)
            df_order["A_Order"] = df_order["A_Order"].apply(clean_text)
            for col in ["SKU", "Container", "Description"]:
                df_move[col] = df_move[col].fillna("").astype(str).str.strip()
            df_order["PO"] = df_order["PO"].fillna("").astype(str).str.strip()
            df_move["Qty"] = pd.to_numeric(df_move["Qty"], errors="coerce").fillna(0)
            df_order = df_order.drop_duplicates(subset=["A_Order"], keep="first")
            return df_move, df_order
        except PermissionError as exc:
            raise RuntimeError("The Excel file is open, locked, or access was denied. Close it and try again.") from exc
        except FileNotFoundError as exc:
            raise RuntimeError(f"Excel file not found: {SCIEZKA}") from exc
        except Exception as exc:
            raise RuntimeError(f"Could not load the Excel file: {exc}") from exc
        finally:
            if temp_path and os.path.exists(temp_path):
                try:
                    os.remove(temp_path)
                except OSError:
                    pass

    @staticmethod
    def przygotuj_wynik(df_filtered, df_order):
        grouped = df_filtered.groupby(["Order", "SKU", "Container", "Description"], as_index=False)["Qty"].sum()
        final = pd.merge(grouped, df_order, left_on="Order", right_on="A_Order", how="left")
        final["PO"] = final["PO"].apply(lambda value: "0" if pd.isna(value) or str(value).strip().lower() in {"", "nan", "none"} else str(value).strip())
        return final

    def wyswietl_i_zapisz(self, df_filtered, df_order):
        final = self.przygotuj_wynik(df_filtered, df_order)
        self.current_data = final
        self.tree.delete(*self.tree.get_children())
        for index, (_, row) in enumerate(final.iterrows()):
            tags = ["even"] if index % 2 else []
            if row["PO"] == "0":
                tags.append("missing_po")
            self.tree.insert("", tk.END, values=(row["Order"], row["SKU"], int(row["Qty"]),
                             row["Container"], row["Description"], row["PO"]), tags=tags)
        self._set_empty_state(None if len(final) else "Container not found")
        self.set_status("Search completed", len(final))

    def _start_search(self, container, all_for_order=False):
        if self.busy:
            return

        # Clear previous results before starting a new lookup so stale data cannot be printed.
        self.current_data = pd.DataFrame()
        self.tree.delete(*self.tree.get_children())
        self._set_empty_state(None)

        if not container:
            self._set_empty_state("Enter a container number to search")
            self.set_status("Enter a container number")
            return
        self._set_busy(True)
        self.set_status("Loading data...")
        def worker():
            try:
                move, order = self.wczytaj_dane()
                move["Container"] = move["Container"].astype(str).str.strip()
                matching = move[move["Container"] == container]
                if all_for_order and not matching.empty:
                    orders = matching["Order"].dropna().unique()
                    if len(orders) > 1:
                        raise ValueError("This container is assigned to more than one Order. Contact support before printing.")
                    matching = move[move["Order"] == orders[0]]
                self.root.after(0, lambda: self._finish_search(matching, order))
            except Exception as exc:
                self.root.after(0, lambda error=exc: self._search_error(error))
        threading.Thread(target=worker, daemon=True).start()

    def _finish_search(self, matching, order):
        try:
            if matching.empty:
                self.current_data = pd.DataFrame()
                self.tree.delete(*self.tree.get_children())
                self._set_empty_state("Container not found")
                self.set_status("No results")
            else:
                self.wyswietl_i_zapisz(matching, order)
        finally:
            self._set_busy(False)

    def _search_error(self, error):
        self.current_data = pd.DataFrame()
        self.tree.delete(*self.tree.get_children())
        self._set_empty_state("Unable to load data")
        self._set_busy(False)
        self.set_status("Excel error")
        messagebox.showerror("Excel error", str(error))

    def szukaj_container(self):
        self._start_search(self.search_entry.get().strip())

    def szukaj_order_po_containerze(self):
        self._start_search(self.order_entry.get().strip(), all_for_order=True)

    def odswiez(self):
        if self.busy:
            return
        self.search_entry.delete(0, tk.END)
        self.order_entry.delete(0, tk.END)
        self.tree.delete(*self.tree.get_children())
        self.current_data = pd.DataFrame()
        self._set_empty_state("Enter a container number to search")
        self.set_status("Ready")
        self.search_entry.focus_set()

    def generuj_zpl(self, container_id, rows):
        n = len(rows)

        # LIMIT 28 SKU
        max_items = 28
        show_more = n > max_items
        rows_to_print = rows.iloc[:max_items]

        r1 = rows_to_print.iloc[0]

        zpl = f"""^XA^LH0,60^LT0^MNY^MFN,N~TA000^MD0^PR6,6^PON^BY2,3.0,10^CVY^XZ^XA^PW833^LL1179^XZ^XA^FO25,34^GB191,0,38,W,0^FS^FT34,60^A0N,23^FVOrderID^FS^FT224,60^A0N,23^FV{zpl_text(r1['Order'])}^FS^FO406,34^GB191,0,38,W,0^FS^FT415,60^A0N,23^FVContainer ID^FS^FT605,60^A0N,23^FV{zpl_text(container_id)}^FS^FO25,72^GB191,0,38,W,0^FS^FT34,98^A0N,23^FVPO Reference^FS^FT224,98^A0N,23^FV{zpl_text(r1['PO'])}^FS^FO215,71^GB573,40,1,B,^FS^FO25,71^GB192,40,1,B,^FS^FO596,33^GB192,40,1,B,^FS^FO406,33^GB192,40,1,B,^FS^FO215,33^GB192,40,1,B,^FS^FO25,33^GB192,40,1,B,^FS^FO24,32^GB765,79,3,B,^FS^FT70,836^A0B,37^FVPACK / CONTENTS LABEL^FS"""

        limit_k1 = min(len(rows_to_print), 14)
        zpl += """^FO102,1007^GB0,130,34,W,0^FS^FT128,1129^A0B,21^FVSKU^FS^FO102,724^GB0,283,34,W,0^FS^FT126,999^A0B,17^FVSKU Description:^FS^FO102,659^GB0,65,34,W,0^FS^FT128,716^A0B,21^FVQTY^FS"""

        for i in range(limit_k1):
            row = rows_to_print.iloc[i]
            v = 166 + (i * 44)
            d_off = 7 if i >= 10 else 10

            zpl += f"^FT{v},1129^A0B,20^FV{zpl_text(row['SKU'])}^FS"
            zpl += f"^FO{v-d_off-12},724^A0B,17^FB283,2,0,L,0^FD{zpl_text(row['Description'])}^FS"
            zpl += f"^FT{v},716^A0B,20^FV{int(row['Qty'])}^FS"
            zpl += f"^FO{134 + (i*44)},658^GB46,68,2,B,^FS^FO{134 + (i*44)},723^GB46,285,2,B,^FS^FO{134 + (i*44)},1006^GB46,132,2,B,^FS"

        zpl += "^FO100,658^GB36,68,2,B,^FS^FO100,723^GB36,285,2,B,^FS^FO100,1006^GB36,132,2,B,^FS"
        bg1 = 648 if limit_k1 == 14 else (78 + (limit_k1 - 1) * 44)
        zpl += f"^FO102,659^GB{bg1},478,0,B,^FS"

        if len(rows_to_print) > 14:
            zpl += "^FO102,505^GB0,130,34,W,0^FS^FT128,627^A0B,21^FVSKU^FS^FO102,223^GB0,283,34,W,0^FS^FT126,497^A0B,17^FVSKU Description:^FS^FO102,157^GB0,65,34,W,0^FS^FT128,215^A0B,21^FVQTY^FS"

            k2_rows = len(rows_to_print) - 14
            for i in range(k2_rows):
                row = rows_to_print.iloc[14 + i]
                v = 166 + (i * 44)

                zpl += f"^FT{v},627^A0B,20^FV{zpl_text(row['SKU'])}^FS"
                zpl += f"^FO{v-18},223^A0B,17^FB283,2,0,L,0^FD{zpl_text(row['Description'])}^FS"
                zpl += f"^FT{v},215^A0B,20^FV{int(row['Qty'])}^FS"
                zpl += f"^FO{134 + (i*44)},156^GB46,68,2,B,^FS^FO{134 + (i*44)},221^GB46,285,2,B,^FS^FO{134 + (i*44)},504^GB46,132,2,B,^FS"

            zpl += "^FO100,156^GB36,68,2,B,^FS^FO100,221^GB36,285,2,B,^FS^FO100,504^GB36,132,2,B,^FS"
            bg2 = 34 + (k2_rows * 44)
            zpl += f"^FO102,157^GB{bg2},478,0,B,^FS"

        # ZAKOŃCZENIE – DWA WARIANTY
        if show_more:
            zpl += "^FS^FT803,488^A0B,28^FVAnd more items...^FS^XZ"
        else:
            zpl += "^FT813,1097^A0B,17^FVPRD - 7.0.0.1^FS^XZ"

        return zpl

    def obsluga_druku(self):
        if self.busy:
            return
        if self.current_data.empty:
            self.set_status("No data to print")
            return
        self._set_busy(True, "Printing...")
        self.set_status("Printing...")
        data = self.current_data.copy()
        def worker():
            handle = None
            document_started = False
            page_started = False
            try:
                zpl = "".join(self.generuj_zpl(container, group) for container, group in data.groupby("Container"))
                handle = win32print.OpenPrinter(NAZWA_DRUKARKI)
                win32print.StartDocPrinter(handle, 1, ("Ceva Labels", None, "RAW"))
                document_started = True
                win32print.StartPagePrinter(handle)
                page_started = True
                win32print.WritePrinter(handle, zpl.encode("utf-8"))
                win32print.EndPagePrinter(handle)
                page_started = False
                win32print.EndDocPrinter(handle)
                document_started = False
                self.root.after(0, lambda: self._print_finished(len(data)))
            except Exception as exc:
                if handle is not None:
                    try:
                        if page_started:
                            win32print.EndPagePrinter(handle)
                        if document_started:
                            win32print.EndDocPrinter(handle)
                    except Exception:
                        pass
                self.root.after(0, lambda error=exc: self._print_error(error))
            finally:
                if handle is not None:
                    try:
                        win32print.ClosePrinter(handle)
                    except Exception:
                        pass
        threading.Thread(target=worker, daemon=True).start()

    def _print_finished(self, count):
        self._set_busy(False)
        self.set_status("Print completed", count)

    def _print_error(self, error):
        self._set_busy(False)
        self.set_status("Printer error")
        messagebox.showerror("Printer error", f"Printer: {NAZWA_DRUKARKI}\nUnable to print the label.\n\nDetails: {error}")


if __name__ == "__main__":
    root = tk.Tk()
    root.iconbitmap(r"D:\OneDrive - CEVA Logistics\Europe-CL-ResMed Operations - Pack Contents reprint\Data\pack data\ceva.ico")    
    app = ExcelSearchApp(root)
    root.mainloop()

