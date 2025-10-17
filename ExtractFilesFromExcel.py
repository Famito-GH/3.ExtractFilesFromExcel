import os
import sys
import shutil
import pandas as pd
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from datetime import datetime
import re
import threading

# --- Cesty ---
SOURCE_PATH_ORIGINAL    = r"\\NAS\spolecne\1. PRODUKTOVÉ FOTKY\AKTUÁLNÍ"
SOURCE_PATH_PROMO_FOTO  = r"\\NAS\spolecne\00 - PROMO FOTOGRAFIE A VIDEA\fotky"
SOURCE_PATH_PROMO_VIDEA = r"\\NAS\spolecne\00 - PROMO FOTOGRAFIE A VIDEA\videa"

# --- Logger ---
def setup_logger(script_dir):
    log_file = os.path.join(script_dir, "vypis konzole.txt")
    if os.path.exists(log_file):
        os.remove(log_file)
    open(log_file, "w", encoding="utf-8").close()

    def log(msg):
        ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        line = f"[{ts}] {msg}"
        print(line)
        with open(log_file, "a", encoding="utf-8") as f:
            f.write(line + "\n")
    return log

# --- Čištění buněk ---
def clean_cell(val):
    if pd.isna(val):
        return None
    s = str(val).strip()
    if not s or s.lower() in ("nan", "none", "n/a", "na", "-", "null"):
        return None
    return s

# --- Výpočet výstupní cesty ---
def get_output_dir(dest_dir, znacka, kategorie, original_product, root_mode=False, flat_structure=False):
    znacka = clean_cell(znacka)
    kategorie = clean_cell(kategorie)

    if root_mode:
        return dest_dir
    if flat_structure:
        return os.path.join(dest_dir, original_product)

    if znacka and kategorie:
        return os.path.join(dest_dir, znacka, kategorie, original_product)
    elif znacka and not kategorie:
        return os.path.join(dest_dir, znacka, "Nezařazeno", original_product)
    elif not znacka and kategorie:
        return os.path.join(dest_dir, "Nezařazeno", kategorie, original_product)
    else:
        return os.path.join(dest_dir, "Nezařazeno", "Nezařazeno", original_product)

# --- Načtení Excelu ---
def get_mapping_from_excel(excel_path, require_structure=True):
    df = pd.read_excel(excel_path)
    col_kod = col_znacka = col_kategorie = None

    for col in df.columns:
        n = str(col).strip().lower()
        if "kód" in n and not col_kod:
            col_kod = col
        elif "značka" in n and not col_znacka:
            col_znacka = col
        elif "kategorie" in n and not col_kategorie:
            col_kategorie = col

    if not col_kod:
        raise ValueError("Excel musí obsahovat sloupec s kódem produktu.")
    if require_structure and not (col_znacka and col_kategorie):
        raise ValueError("Excel musí obsahovat i sloupce Značka a Kategorie.")

    mapping = {}
    for _, row in df.iterrows():
        kod = clean_cell(row[col_kod])
        if not kod:
            continue
        znacka = clean_cell(row[col_znacka]) if col_znacka else None
        kategorie = clean_cell(row[col_kategorie]) if col_kategorie else None
        mapping[kod] = (znacka, kategorie)
    return mapping

# --- Kopírování fotek podle Excelu ---
def copy_photos_by_excel(source_dir, dest_dir, mapping, flat_structure=False, root_mode=False, log=None):
    import re
    import shutil
    import os

    os.makedirs(dest_dir, exist_ok=True)
    exts = [".jpg",".jpeg",".png",".gif",".bmp",".tif",".tiff",
            ".mp4",".avi",".mov",".mkv",".wmv",".flv"]
    copied_count = 0
    mapping_lower = {str(k).strip().lower(): v for k, v in mapping.items()}

    for root, _, files in os.walk(source_dir):
        for fn in files:
            if not any(fn.lower().endswith(ext) for ext in exts):
                continue

            name_no_ext = os.path.splitext(fn)[0].strip()
            parts = [p.strip() for p in re.split(r",", name_no_ext) if p.strip()]
            detected_products = [re.sub(r"\([^)]*\)", "", p).strip() for p in parts if p.strip()]

            if not detected_products:
                log(f"Přeskočeno (nelze detekovat produkty): {fn}")
                continue

            all_in_excel = all(prod.lower() in mapping_lower for prod in detected_products)
            if not all_in_excel:
                log(f"Přeskočeno (ne všechny produkty nalezeny v Excelu): {fn}")
                continue

            for prod in detected_products:
                prod_key = prod.lower().strip()
                znacka, kategorie = mapping_lower[prod_key]
                original_product = next(
                    (k for k in mapping.keys() if k.strip().lower() == prod_key),
                    prod
                )

                out_dir = get_output_dir(dest_dir, znacka, kategorie, original_product, root_mode, flat_structure)
                os.makedirs(out_dir, exist_ok=True)

                try:
                    shutil.copy2(os.path.join(root, fn), os.path.join(out_dir, fn))
                    log(f"Zkopírován soubor: {fn} -> {out_dir}")
                    copied_count += 1
                except Exception as e:
                    log(f"Chyba při kopírování {fn} -> {out_dir}: {e}")

    log(f"Celkem zkopírováno {copied_count} souborů." if copied_count else "Nenalezeny žádné soubory.")

# --- Kopírování prvního média ---
def copy_first_media(src_dir, dest_dir):
    exts = [".jpg",".jpeg",".png",".gif",".bmp",".tif",".tiff",".mp4",".avi",".mov",".mkv",".wmv",".flv"]
    try:
        files = sorted(os.listdir(src_dir))
    except:
        return
    for fn in files:
        if any(fn.lower().endswith(ext) for ext in exts):
            os.makedirs(dest_dir, exist_ok=True)
            shutil.copy2(os.path.join(src_dir, fn), os.path.join(dest_dir, fn))
            print(f"Zkopírován první soubor: {fn} -> {dest_dir}")
            return

# --- Kopírování složek podle Excelu ---
def copy_folders_with_mapping(source_path, dest_path, mapping, copy_mode, flat_structure=False, root_mode=False):
    unfound = set(mapping.keys())
    for root, dirs, _ in os.walk(source_path):
        for folder in dirs:
            if folder in unfound:
                znacka, kategorie = mapping[folder]
                src_dir = os.path.join(root, folder)
                out_dir = get_output_dir(dest_path, znacka, kategorie, folder, root_mode, flat_structure)

                if copy_mode == "all":
                    shutil.copytree(src_dir, out_dir, dirs_exist_ok=True)
                    print(f"Zkopírována složka: {folder} -> {out_dir}")
                else:
                    copy_first_media(src_dir, out_dir)
                unfound.remove(folder)
    return unfound

# --- GUI aplikace ---
class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.original_title = "Kopírování fotek podle Excelu"
        self.title(self.original_title)
        self.geometry("550x380")
        self.resizable(False, False)
        self.configure(bg="#f5f5f5")
        self.script_dir = os.path.dirname(sys.executable if getattr(sys, 'frozen', False) else __file__)
        self.log = setup_logger(self.script_dir)
        self.create_widgets()

    def create_widgets(self):
        main = ttk.Frame(self, padding=15)
        main.pack(fill="both", expand=True)

        frame_mode = ttk.LabelFrame(main, text="Režim kopírování", padding=10)
        frame_mode.grid(row=0, column=0, sticky="ew", pady=10)
        self.mode_var = tk.StringVar(value="1")
        ttk.Radiobutton(frame_mode, text="Celé složky", variable=self.mode_var, value="1").grid(row=0, column=0, padx=10)
        ttk.Radiobutton(frame_mode, text="První soubor", variable=self.mode_var, value="2").grid(row=0, column=1, padx=10)
        ttk.Radiobutton(frame_mode, text="Podle Excelu", variable=self.mode_var, value="3").grid(row=0, column=2, padx=10)

        frame_source = ttk.LabelFrame(main, text="Zdrojová složka", padding=10)
        frame_source.grid(row=1, column=0, sticky="ew", pady=10)
        self.source_var = tk.StringVar(value="1")
        ttk.Radiobutton(frame_source, text="Promo fotky", variable=self.source_var, value="1").grid(row=0, column=0, padx=10)
        ttk.Radiobutton(frame_source, text="Promo videa", variable=self.source_var, value="2").grid(row=0, column=1, padx=10)
        ttk.Radiobutton(frame_source, text="Vlastní složka", variable=self.source_var, value="3").grid(row=0, column=2, padx=10)
        ttk.Radiobutton(frame_source, text="Produktové fotky", variable=self.source_var, value="4").grid(row=0, column=3, padx=10)

        frame_sort = ttk.LabelFrame(main, text="Třídění souborů", padding=10)
        frame_sort.grid(row=2, column=0, sticky="ew", pady=10)
        self.sort_var = tk.StringVar(value="1")
        ttk.Radiobutton(frame_sort, text="Podle struktury", variable=self.sort_var, value="1").grid(row=0, column=0, padx=10)
        ttk.Radiobutton(frame_sort, text="Plochá struktura", variable=self.sort_var, value="2").grid(row=0, column=1, padx=10)
        ttk.Radiobutton(frame_sort, text="Vše do root složky", variable=self.sort_var, value="3").grid(row=0, column=2, padx=10)

        frame_action = ttk.Frame(main, padding=10)
        frame_action.grid(row=3, column=0, pady=20)
        self.copy_button = ttk.Button(frame_action, text="Spustit kopírování", command=self.start_copy_thread)
        self.copy_button.pack(padx=10, pady=5)

    def start_copy_thread(self):
        self.title("Kopírování fotek podle Excelu - Probíhá kopírování")
        self.copy_button.config(text="Probíhá kopírování", state="disabled")
        self.update_idletasks()
        t = threading.Thread(target=self.run_copy)
        t.start()

    def run_copy(self):
        try:
            script_dir = os.path.dirname(sys.executable if getattr(sys, 'frozen', False) else os.path.abspath(__file__))
            dest_path = os.path.join(script_dir, "foto_folders")

            if os.path.exists(dest_path):
                try:
                    shutil.rmtree(dest_path)
                    print(f"Odstraněna stará složka: {dest_path}")
                except Exception as e:
                    print(f"Nepodařilo se odstranit složku '{dest_path}': {e}")

            os.makedirs(dest_path, exist_ok=True)
            mode = self.mode_var.get()
            src_choice = self.source_var.get()
            sort_choice = self.sort_var.get()

            if src_choice == "1": source_path = SOURCE_PATH_PROMO_FOTO
            elif src_choice == "2": source_path = SOURCE_PATH_PROMO_VIDEA
            elif src_choice == "4": source_path = SOURCE_PATH_ORIGINAL
            else:
                chosen = filedialog.askdirectory(title="Vyberte zdrojovou složku")
                if not chosen: return
                source_path = chosen

            excel_path = None
            for fn in os.listdir(script_dir):
                if fn.lower().endswith((".xlsx", ".xls", ".xlsm")):
                    excel_path = os.path.join(script_dir, fn)
                    break

            if not excel_path or not os.path.isfile(excel_path):
                messagebox.showerror("Chyba", "Excel nebyl nalezen ve složce se skriptem.")
                self.log("Excel nenalezen.")
                return

            require_structure = (sort_choice == "1")
            try:
                mapping = get_mapping_from_excel(excel_path, require_structure=require_structure)
                self.log(f"Načten Excel: {excel_path}")
            except Exception as e:
                messagebox.showerror("Chyba při načítání Excelu", str(e))
                self.log(f"Chyba při načítání Excelu: {e}")
                return

            flat_structure = (sort_choice == "2")
            root_mode = (sort_choice == "3")

            if mode == "3":
                copy_photos_by_excel(source_path, dest_path, mapping, flat_structure, root_mode, log=self.log)
            else:
                copy_mode = "all" if mode == "1" else "first"
                unfound = copy_folders_with_mapping(source_path, dest_path, mapping, copy_mode, flat_structure, root_mode)
                if unfound:
                    uf = os.path.join(self.script_dir, "unfound_folders.txt")
                    with open(uf, "w", encoding="utf-8") as f:
                        for k in unfound:
                            f.write(k + "\n")
                    self.log(f"Některé složky nebyly nalezeny. Seznam uložen: {uf}")
                    messagebox.showwarning("Upozornění", f"Některé složky nebyly nalezeny. Seznam uložen: {uf}")

            self.log("Kopírování dokončeno.")
            messagebox.showinfo("Hotovo", "Kopírování dokončeno.")
        finally:
            self.copy_button.config(text="Spustit kopírování", state="normal")
            self.title(self.original_title)


if __name__ == "__main__":
    app = App()
    app.mainloop()
