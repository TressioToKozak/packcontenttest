import re
import unicodedata
import pandas as pd
import tkinter as tk
from tkinter import ttk, messagebox
import win32print
from PIL import Image, ImageTk
import subprocess
import shutil, tempfile


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


class ExcelSearchApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Pack Content Reprint")
        self.root.geometry("1200x700")
        self.root.configure(bg=NAVY_DARK)

        self.current_data = pd.DataFrame()
        self.spinner_win = None

        # UI
        self.build_header()
        self.build_dashboard()
        self.build_table()
        self.build_statusbar()

    # ---------- UI BUDOWA ----------

    def build_header(self):
        self.header_height = 90
        self.header_canvas = tk.Canvas(self.root, height=self.header_height,
                                       highlightthickness=0, bd=0)
        self.header_canvas.pack(fill=tk.X)

        self.root.update_idletasks()
        width = self.root.winfo_width()
        self.draw_gradient(width)

        # Logo
        logo_frame = tk.Frame(self.header_canvas, bg=NAVY_DARK)
        self.logo_img = self.wczytaj_logo(LOGO_PATH)
        if self.logo_img:
            tk.Label(logo_frame, image=self.logo_img, bg=NAVY_DARK).pack()
        else:
            tk.Label(logo_frame, text="CEVA LOGISTICS", bg=NAVY_DARK,
                     fg="black", font=("Segoe UI", 20, "bold")).pack()
        self.header_canvas.create_window(20, 5, anchor="nw", window=logo_frame)

        # Przyciski akcji po prawej
        btn_frame = tk.Frame(self.header_canvas, bg="", highlightthickness=0)
        self.header_canvas.create_window(width - 20, self.header_height // 2,
                                         anchor="e", window=btn_frame)

        def action_btn(symbol, cmd):
            return tk.Button(
                btn_frame,
                text=symbol,
                command=cmd,
                bg=WHITE,
                relief="flat",
                font=("Segoe UI", 20),
                cursor="hand2",
                width=2
            )

        self.root.bind("<Configure>", self.on_resize)

    def draw_gradient(self, width):
        h = self.header_height
        r1, g1, b1 = self.header_canvas.winfo_rgb(NAVY_DARK)
        r2, g2, b2 = self.header_canvas.winfo_rgb(NAVY_LIGHT)
        r_ratio = (r2 - r1) / max(width, 1)
        g_ratio = (g2 - g1) / max(width, 1)
        b_ratio = (b2 - b1) / max(width, 1)
        for i in range(width):
            nr = int(r1 + (r_ratio * i))
            ng = int(g1 + (g_ratio * i))
            nb = int(b1 + (b_ratio * i))
            color = f"#{nr//256:02x}{ng//256:02x}{nb//256:02x}"
            self.header_canvas.create_line(i, 0, i, h, fill=color)

    def on_resize(self, event):
        if event.widget == self.root:
            self.draw_gradient(event.width)

    def build_dashboard(self):
        self.main_frame = tk.Frame(self.root, bg=PANEL_BG)
        self.main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=(5, 5))

        control_frame = tk.LabelFrame(
            self.main_frame,
            text=" Control Panel",
            bg=PANEL_BG,
            fg=NAVY_DARK,
            padx=15,
            pady=15,
            font=("Segoe UI", 11, "bold")
        )
        control_frame.pack(fill=tk.X, pady=(0, 10))

        # Container
        tk.Label(control_frame, text="Container :", bg=PANEL_BG,
                 fg="black", font=("Segoe UI", 10)).grid(row=0, column=0, padx=5, pady=5, sticky="w")
        self.search_entry = tk.Entry(control_frame, font=("Segoe UI", 11), width=20)
        self.search_entry.grid(row=0, column=1, padx=5, pady=5)

        # Order z boxa
        tk.Label(control_frame, text="Container :", bg=PANEL_BG,
                 fg="black", font=("Segoe UI", 10)).grid(row=0, column=3, padx=5, pady=5, sticky="w")
        self.order_entry = tk.Entry(control_frame, font=("Segoe UI", 11), width=20)
        self.order_entry.grid(row=0, column=4, padx=5, pady=5)

        def cmd_btn(text, color, cmd):
            return tk.Button(
                control_frame,
                text=text,
                command=cmd,
                bg=color,
                fg=WHITE,
                activebackground=color,
                activeforeground=WHITE,
                relief="flat",
                font=("Segoe UI", 10, "bold"),
                cursor="hand2",
                width=14,
                height=1
            )

        cmd_btn("Search", BTN_BLUE, self.szukaj_container).grid(row=0, column=2, padx=10)
        cmd_btn("All containers", BTN_GREEN, self.szukaj_order_po_containerze).grid(row=0, column=5, padx=10)
        cmd_btn("Print", BTN_RED, self.obsluga_druku).grid(row=0, column=6, padx=20)
        cmd_btn("Refresh", BTN_GREY, self.odswiez).grid(row=0, column=7, padx=10)

    def build_table(self):
        table_frame = tk.Frame(self.main_frame, bg=PANEL_BG)
        table_frame.pack(fill=tk.BOTH, expand=True)

        style = ttk.Style()
        style.theme_use("default")

        style.configure(
            "Treeview",
            background=WHITE,
            foreground="black",
            rowheight=26,
            fieldbackground=WHITE,
            font=("Segoe UI", 10)
        )

        style.configure(
            "Treeview.Heading",
            background=NAVY_DARK,
            foreground=WHITE,
            font=("Segoe UI", 10, "bold")
        )

        style.map("Treeview",
                  background=[("selected", BTN_BLUE)],
                  foreground=[("selected", WHITE)])

        columns = ("Order", "SKU", "Qty", "Container", "Description", "PO")
        self.tree = ttk.Treeview(table_frame, columns=columns, show="headings")

        for col in columns:
            self.tree.heading(col, text=col)
            width = 120
            if col == "Description":
                width = 260
            self.tree.column(col, width=width, anchor=tk.W)

        vsb = ttk.Scrollbar(table_frame, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=vsb.set)

        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        vsb.pack(side=tk.RIGHT, fill=tk.Y)

    def build_statusbar(self):
        self.status_frame = tk.Frame(self.root, bg=NAVY_DARK)
        self.status_frame.pack(fill=tk.X, side=tk.BOTTOM)

        self.status_label = tk.Label(
            self.status_frame,
            text="Done",
            bg=NAVY_DARK,
            fg=WHITE,
            anchor="w",
            font=("Segoe UI", 10)
        )
        self.status_label.pack(fill=tk.X, padx=10, pady=4)

    # ---------- STATUS / SPINNER ----------

    def set_status(self, text, count=None):
        if count is not None:
            self.status_label.config(text=f"{text} | Amount: {count}")
        else:
            self.status_label.config(text=text)
        self.root.update_idletasks()

    def show_spinner(self, text="Searching..."):
        if self.spinner_win is not None:
            return
        self.spinner_win = tk.Toplevel(self.root)
        self.spinner_win.title("")
        self.spinner_win.overrideredirect(True)
        self.spinner_win.configure(bg=NAVY_DARK)

        w, h = 220, 80
        x = self.root.winfo_x() + (self.root.winfo_width() // 2) - w // 2
        y = self.root.winfo_y() + (self.root.winfo_height() // 2) - h // 2
        self.spinner_win.geometry(f"{w}x{h}+{x}+{y}")

        lbl = tk.Label(self.spinner_win, text=text, bg=NAVY_DARK,
                       fg=WHITE, font=("Segoe UI", 10))
        lbl.pack(pady=(10, 5))

        pb = ttk.Progressbar(self.spinner_win, mode="indeterminate", length=160)
        pb.pack(pady=(0, 10))
        pb.start(10)

        self.spinner_win.update_idletasks()

    def hide_spinner(self):
        if self.spinner_win is not None:
            self.spinner_win.destroy()
            self.spinner_win = None

    # ---------- LOGO ----------

    def wczytaj_logo(self, path):
        try:
            img = Image.open(path)
            img = img.resize((260, 80), Image.LANCZOS)
            return ImageTk.PhotoImage(img)
        except Exception as e:
            print("Error with logo:", e)
            return None

    # ---------- LOGIKA DANYCH ----------

    def wczytaj_dane(self):
        import shutil
        import tempfile
        try:
            # Stwórz kopię pliku w katalogu tymczasowym, żeby nie blokować oryginału na ShareDrive/OneDrive
            temp_dir = tempfile.gettempdir()
            import uuid
            temp_path = temp_dir + f"\\excel_temp_{uuid.uuid4().hex}.xlsx"
            shutil.copy2(SCIEZKA, temp_path)

            # Wczytywanie wymuszone jako string (dtype=str)
            df_move = pd.read_excel(temp_path, sheet_name='move task', header=None, dtype=str).iloc[:, 0:5]
            df_move.columns = ['Order', 'SKU', 'Qty', 'Container', 'Description']

            # Odczyt arkuszy i znalezienie właściwego arkusza 'Order header' (case-insensitive)
            xls = pd.ExcelFile(temp_path)
            sheet_names = xls.sheet_names
            order_sheet_name = None
            for s in sheet_names:
                if s.strip().lower() == 'order header':
                    order_sheet_name = s
                    break
            if order_sheet_name is None:
                for s in sheet_names:
                    if 'order' in s.strip().lower() and 'header' in s.strip().lower():
                        order_sheet_name = s
                        break
            if order_sheet_name is None:
                for s in sheet_names:
                    if 'order' in s.strip().lower():
                        order_sheet_name = s
                        break
            if order_sheet_name is None:
                order_sheet_name = sheet_names[0]

            # Wczytaj surowy df_order bez założeń o kolumnach
            df_order_raw = pd.read_excel(temp_path, sheet_name=order_sheet_name, header=None, dtype=str)

            # Spróbuj wybrać dwie pierwsze kolumny, ale jeśli druga kolumna jest pusta, utwórz drugą pustą
            if df_order_raw.shape[1] >= 2:
                df_order = df_order_raw.iloc[:, 0:2].copy()
                df_order.columns = ['A_Order', 'PO']
            else:
                df_order = pd.DataFrame({
                    'A_Order': df_order_raw.iloc[:, 0].astype(str),
                    'PO': [''] * len(df_order_raw)
                })

            # Upewniamy się, że nie ma wartości NaN przed formatowaniem
            df_move['Order'] = df_move['Order'].fillna('')
            df_order['A_Order'] = df_order['A_Order'].fillna('')

            # Zastosuj solidne czyszczenie kluczy (funkcja clean_text powinna być zdefiniowana globalnie)
            df_move['Order'] = df_move['Order'].apply(clean_text)
            df_order['A_Order'] = df_order['A_Order'].apply(clean_text)

            # Formatuj pozostałe kolumny
            for col in ['SKU', 'Container', 'Description']:
                df_move[col] = df_move[col].fillna('').astype(str).str.strip()

            # Konwersja Qty do numeric i uzupełnienie braków zerami
            df_move['Qty'] = pd.to_numeric(df_move['Qty'], errors='coerce').fillna(0)

            # Usuń wiersze nagłówkowe, które mogły wpaść jako dane
            df_move = df_move[~df_move['Order'].str.contains('ORDER', case=False, na=False)]
            df_order = df_order[~df_order['A_Order'].str.contains('ORDER', case=False, na=False)]

            # Zabezpieczenie przed powielaniem wierszy przez duplikaty w sheet2
            df_order = df_order.drop_duplicates(subset=['A_Order'], keep='first')

            return df_move, df_order

        except PermissionError:
            raise Exception("PermissionError: plik jest otwarty lub zablokowany. Zamknij plik w Excelu lub upewnij się, że masz dostęp.")
        except FileNotFoundError:
            raise Exception(f"FileNotFoundError: nie znaleziono pliku pod ścieżką: {SCIEZKA}")
        except Exception as e:
            raise Exception(f"Error with loading Excel File: {e}")
        finally:
            # Spróbuj usunąć plik tymczasowy, jeśli istnieje
            try:
                if 'temp_path' in locals():
                    import os
                    if os.path.exists(temp_path):
                        os.remove(temp_path)
            except Exception:
                pass

    def wyswietl_i_zapisz(self, df_filtered, df_order):
        wynik = df_filtered.groupby(
            ['Order', 'SKU', 'Container', 'Description'],
            as_index=False
        )['Qty'].sum()

        # Merge z użyciem wyczyszczonych wartości tekstowych
        final = pd.merge(wynik, df_order, left_on='Order', right_on='A_Order', how='left')

        # Ostateczne czyszczenie brakujących wartości PO (gdy nie znajdzie pary)
        final['PO'] = final['PO'].apply(
            lambda x: "No Data" if pd.isna(x) or str(x).strip().lower() in ['nan', 'none', ''] else str(x).strip()
        )

        self.current_data = final

        for item in self.tree.get_children():
            self.tree.delete(item)

        for _, row in final.iterrows():
            self.tree.insert(
                "",
                tk.END,
                values=(
                    row['Order'],
                    row['SKU'],
                    int(row['Qty']),
                    row['Container'],
                    row['Description'],
                    row['PO']
                )
            )

        self.set_status("Done", count=len(final))

    # ---------- AKCJE ----------

    def szukaj_container(self):
        cont = self.search_entry.get().strip()
        if not cont:
            messagebox.showinfo("Info", "No container number.")
            return
        self.show_spinner("Loading data...")
        self.set_status("Loading...")
        self.root.update()
        try:
            df_m, df_o = self.wczytaj_dane()
            # Upewnij się, że container też jest czyszczony podobnie (trim)
            df_m['Container'] = df_m['Container'].astype(str).str.strip()
            df_f = df_m[df_m['Container'] == cont]
            if df_f.empty:
                self.hide_spinner()
                self.set_status("No results")
                messagebox.showinfo("Info", "Container not found.")
                return
            self.wyswietl_i_zapisz(df_f, df_o)
        except Exception as e:
            messagebox.showerror("Error", str(e))
            self.set_status("Error")
        finally:
            self.hide_spinner()

    def szukaj_order_po_containerze(self):
        cont = self.order_entry.get().strip()
        if not cont:
            messagebox.showinfo("Info", "No container number for all container.")
            return
        self.show_spinner("Loading data...")
        self.set_status("Loading...")
        self.root.update()
        try:
            df_m, df_o = self.wczytaj_dane()
            df_m['Container'] = df_m['Container'].astype(str).str.strip()
            ord_nrs = df_m[df_m['Container'] == cont]['Order'].unique()
            if len(ord_nrs) == 0:
                self.hide_spinner()
                self.set_status("No results")
                messagebox.showinfo("Info", "Container not found.")
                return
            # ord_nrs są już w formacie wyczyszczonym (uppercase, bez .0)
            df_f = df_m[df_m['Order'] == ord_nrs[0]]
            self.wyswietl_i_zapisz(df_f, df_o)
        except Exception as e:
            messagebox.showerror("Error", str(e))
            self.set_status("Error")
        finally:
            self.hide_spinner()

    def odswiez(self):
        self.search_entry.delete(0, tk.END)
        self.order_entry.delete(0, tk.END)
        for item in self.tree.get_children():
            self.tree.delete(item)
        self.current_data = pd.DataFrame()
        self.set_status("Done")

    # ---------- ZPL / DRUK (Z LIMITEM 28 SKU) ----------

    def generuj_zpl(self, container_id, rows):
        n = len(rows)

        # LIMIT 28 SKU
        max_items = 28
        show_more = n > max_items
        rows_to_print = rows.iloc[:max_items]

        r1 = rows_to_print.iloc[0]

        zpl = f"""^XA^LH0,60^LT0^MNY^MFN,N~TA000^MD0^PR6,6^PON^BY2,3.0,10^CVY^XZ^XA^PW833^LL1179^XZ^XA^FO25,34^GB191,0,38,W,0^FS^FT34,60^A0N,23^FVOrderID^FS^FT224,60^A0N,23^FV{r1['Order']}^FS^FO406,34^GB191,0,38,W,0^FS^FT415,60^A0N,23^FVContainer ID^FS^FT605,60^A0N,23^FV{container_id}^FS^FO25,72^GB191,0,38,W,0^FS^FT34,98^A0N,23^FVPO Reference^FS^FT224,98^A0N,23^FV{r1['PO']}^FS^FO215,71^GB573,40,1,B,^FS^FO25,71^GB192,40,1,B,^FS^FO596,33^GB192,40,1,B,^FS^FO406,33^GB192,40,1,B,^FS^FO215,33^GB192,40,1,B,^FS^FO25,33^GB192,40,1,B,^FS^FO24,32^GB765,79,3,B,^FS^FT70,836^A0B,37^FVPACK / CONTENTS LABEL^FS"""

        limit_k1 = min(len(rows_to_print), 14)
        zpl += """^FO102,1007^GB0,130,34,W,0^FS^FT128,1129^A0B,21^FVSKU^FS^FO102,724^GB0,283,34,W,0^FS^FT126,999^A0B,17^FVSKU Description:^FS^FO102,659^GB0,65,34,W,0^FS^FT128,716^A0B,21^FVQTY^FS"""

        for i in range(limit_k1):
            row = rows_to_print.iloc[i]
            v = 166 + (i * 44)
            d_off = 7 if i >= 10 else 10

            zpl += f"^FT{v},1129^A0B,20^FV{row['SKU']}^FS"
            zpl += f"^FO{v-d_off-12},724^A0B,17^FB283,2,0,L,0^FD{row['Description']}^FS"
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

                zpl += f"^FT{v},627^A0B,20^FV{row['SKU']}^FS"
                zpl += f"^FO{v-18},223^A0B,17^FB283,2,0,L,0^FD{row['Description']}^FS"
                zpl += f"^FT{v},215^A0B,20^FV{int(row['Qty'])}^FS"
                zpl += f"^FO{134 + (i*44)},156^GB46,68,2,B,^FS^FO{134 + (i*44)},221^GB46,285,2,B,^FS^FO{134 + (i*44)},504^GB46,132,2,B,^FS"

            zpl += "^FO100,156^GB36,68,2,B,^FS^FO100,221^GB36,285,2,B,^FS^FO100,504^GB36,132,2,B,^FS"
            bg2 = 34 + (k2_rows * 44)
            zpl += f"^FO102,157^GB{bg2},478,0,B,^FS"

        # ZAKOŃCZENIE – DWA WARIANTY
        if show_more:
            zpl += "^FS^FT803,488^A0B,28^FVAnd more items...^FS^XZ^XA"
        else:
            zpl += "^FT813,1097^A0B,17^FVPRD - 7.0.0.1^FS^XZ"

        return zpl

    def obsluga_druku(self):
        if self.current_data.empty:
            messagebox.showinfo("Info", "No data to be printed")
            return

        self.show_spinner("Printing...")
        self.set_status("Printing...")
        self.root.update()

        pelny_zpl = ""
        grouped = self.current_data.groupby('Container')
        for container_id, group in grouped:
            pelny_zpl += self.generuj_zpl(container_id, group)

        try:
            hPrinter = win32print.OpenPrinter(NAZWA_DRUKARKI)
            try:
                hJob = win32print.StartDocPrinter(hPrinter, 1, ("Ceva Labels", None, "RAW"))
                win32print.StartPagePrinter(hPrinter)
                win32print.WritePrinter(hPrinter, pelny_zpl.encode('utf-8'))
                win32print.EndPagePrinter(hPrinter)
                win32print.EndDocPrinter(hPrinter)
            finally:
                win32print.ClosePrinter(hPrinter)
            self.set_status("Done", count=len(self.current_data))
        except Exception as e:
            messagebox.showerror("Print error", f"Printner: {NAZWA_DRUKARKI}\nError: {e}")
            self.set_status("Print error")
        finally:
            self.hide_spinner()


if __name__ == "__main__":
    root = tk.Tk()
    root.iconbitmap(r"D:\OneDrive - CEVA Logistics\Europe-CL-ResMed Operations - Pack Contents reprint\Data\pack data\ceva.ico")    
    app = ExcelSearchApp(root)
    root.mainloop()

