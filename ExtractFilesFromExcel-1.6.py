import os
import sys
import shutil
import pandas as pd
import tkinter as tk
from openpyxl import Workbook, load_workbook
from datetime import datetime

SOURCE_PATH_ORIGINAL    = r"\\NAS\spolecne\1. PRODUKTOVÉ FOTKY\AKTUÁLNÍ"
SOURCE_PATH_PROMO_FOTO  = r"\\NAS\spolecne\00 - PROMO FOTOGRAFIE A VIDEA\fotky"
SOURCE_PATH_PROMO_VIDEA = r"\\NAS\spolecne\00 - PROMO FOTOGRAFIE A VIDEA\videa"

# --- Funkce pro logování ---
def setup_logger(script_dir):
    log_file = os.path.join(script_dir, "vypis konzole.txt")
    # Smažeme starý log, pokud existuje
    if os.path.exists(log_file):
        os.remove(log_file)

    def log(msg):
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        line = f"[{timestamp}] {msg}"
        print(line)
        with open(log_file, 'a', encoding='utf-8') as f:
            f.write(line + "\n")
    return log


# --- Ostatní funkce zůstávají stejné ---
def get_mapping_from_excel(excel_path, require_structure=True):
    df = pd.read_excel(excel_path)
    col_kod = col_znacka = col_kategorie = None

    for col in df.columns:
        n = col.strip().lower()
        if "kód" in n and not col_kod:        col_kod = col
        elif "značka" in n and not col_znacka: col_znacka = col
        elif "kategorie" in n and not col_kategorie: col_kategorie = col

    if not col_kod:
        raise ValueError("Excel musí obsahovat sloupec s kódem produktu.")

    if require_structure and not (col_znacka and col_kategorie):
        raise ValueError("Pro třídění podle struktury Excel musí mít i sloupce Značka a Kategorie.")

    mapping = {}
    for _, row in df.iterrows():
        kod = str(row[col_kod]).strip()
        znacka = str(row[col_znacka]).strip() if col_znacka and pd.notna(row[col_znacka]) else None
        kategorie = str(row[col_kategorie]).strip() if col_kategorie and pd.notna(row[col_kategorie]) else None

        if kod:
            if require_structure:
                mapping[kod] = (znacka, kategorie)
            else:
                mapping[kod] = (None, None)
    return mapping

def copy_first_media(src_dir, dest_dir):
    exts = [".jpg",".jpeg",".png",".gif",".bmp",".tif",".tiff",
            ".mp4",".avi",".mov",".mkv",".wmv",".flv"]
    try: files = sorted(os.listdir(src_dir))
    except: return
    for fn in files:
        if any(fn.lower().endswith(ext) for ext in exts):
            os.makedirs(dest_dir, exist_ok=True)
            shutil.copy2(os.path.join(src_dir, fn), os.path.join(dest_dir, fn))
            return

def copy_folders_with_mapping(source_path, dest_path, mapping, copy_mode, flat_structure=False, root_mode=False, log=None):
    unfound = set(mapping.keys())
    for root, dirs, _ in os.walk(source_path):
        for folder in dirs:
            if folder in unfound:
                _, _ = mapping[folder]
                src_dir = os.path.join(root, folder)

                if root_mode:
                    out_dir = dest_path
                elif flat_structure:
                    out_dir = os.path.join(dest_path, folder)
                else:
                    znacka, kategorie = mapping[folder]
                    if not znacka:
                        brand_dir = kategorie
                        category_dir = None
                    else:
                        brand_dir = znacka
                        category_dir = kategorie
                    out_dir = (os.path.join(dest_path, brand_dir, category_dir, folder)
                               if category_dir else os.path.join(dest_path, brand_dir, folder))

                if copy_mode == "all":
                    shutil.copytree(src_dir, out_dir, dirs_exist_ok=True)
                    log(f"Zkopírována CELÁ složka '{folder}' -> {out_dir}")
                else:
                    copy_first_media(src_dir, out_dir)
                    log(f"Zkopírován první soubor ze složky '{folder}' -> {out_dir}")

                unfound.remove(folder)
    return unfound

def delete_folder_contents_safe(folder_path, errors_file_path, log=None):
    if not os.path.isdir(folder_path):
        return
    for item in os.listdir(folder_path):
        path = os.path.join(folder_path, item)
        try:
            if os.path.isdir(path):
                shutil.rmtree(path)
                log(f"Složka odstraněna: {path}")
            else:
                os.remove(path)
                log(f"Soubor odstraněn: {path}")
        except Exception as e:
            log(f"Chyba při mazání '{path}': {e}")
            with open(errors_file_path, 'a', encoding='utf-8') as ef:
                ef.write(f"Chyba při mazání '{path}': {e}\n")

def main():
    if getattr(sys, 'frozen', False):
        script_dir = os.path.dirname(sys.executable)
    else:
        script_dir = os.path.dirname(os.path.abspath(__file__))

    log = setup_logger(script_dir)
    errors_file = os.path.join(script_dir, "delete_errors.txt")
    if os.path.exists(errors_file):
        os.remove(errors_file)

    log("Spuštěn skript.")

    while True:
        print("Zvolte režim kopírování:\n 1 - Celé složky\n 2 - Jen první soubor")
        choice = input("Číslo: ").strip()
        if choice in ("1", "2"):
            copy_mode = "all" if choice == "1" else "first"
            break
        else:
            print("Neplatná volba. Zadejte 1 nebo 2.\n")

    while True:
        print("\nZdroj:\n 1 - Promo fotky\n 2 - Promo videa\n 3 - Vybrat složku dle vlastního výběru\n 4 - Aktuální produktové fotky")
        sc = input("Číslo: ").strip()
        if sc == "1":
            source_path = SOURCE_PATH_PROMO_FOTO
            break
        elif sc == "2":
            source_path = SOURCE_PATH_PROMO_VIDEA
            break
        elif sc == "3":
            root = tk.Tk(); root.withdraw(); root.lift(); root.attributes("-topmost", True)
            from tkinter import filedialog
            chosen = filedialog.askdirectory(title="Vyberte zdrojovou složku")
            root.destroy()
            if chosen:
                source_path = chosen
                break
            else:
                print("Nevybrána žádná složka. Zkuste to znovu.\n")
        elif sc == "4":
            source_path = SOURCE_PATH_ORIGINAL
            break
        else:
            print("Neplatná volba. Zadejte 1, 2, 3 nebo 4.\n")

    while True:
        print("\nChcete třídit produkty podle struktury?\n 1 - Ano\n 2 - Ne\n 3 - Všechny do root složky")
        flat = input("Zadejte číslo: ").strip()
        if flat in ("1","2","3"):
            flat_structure = (flat == "2")
            root_mode = (flat == "3")
            break
        else:
            print("Neplatná volba. Zadejte 1, 2 nebo 3.\n")

    dest_path = os.path.join(script_dir, "foto_folders")
    if os.path.isdir(dest_path):
        delete_folder_contents_safe(dest_path, errors_file, log=log)
    else:
        os.makedirs(dest_path, exist_ok=True)

    excel_path = os.path.join(script_dir, "Export fotek z NAS.xlsx")
    if not os.path.isfile(excel_path):
        log(f"Excel nenalezen: {excel_path}")
        sys.exit(1)
    try:
        require_structure = (flat == "1")
        mapping = get_mapping_from_excel(excel_path, require_structure=require_structure)
        log(f"Načten excel: {excel_path}")
    except Exception as e:
        log(f"Chyba při načítání Excelu: {e}")
        sys.exit(1)

    unfound = copy_folders_with_mapping(source_path, dest_path, mapping, copy_mode, flat_structure, root_mode, log=log)
    if unfound:
        uf = os.path.join(script_dir, "unfound_folders.txt")
        with open(uf, "w", encoding="utf-8") as f:
            for k in unfound:
                f.write(k + "\n")
        log(f"Některé složky nebyly nalezeny. Seznam: {uf}")

    log("Hotovo.")

if __name__ == "__main__":
    main()
