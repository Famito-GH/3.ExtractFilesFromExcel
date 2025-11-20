# -- Verze 1.9 Bez GUI --

import os
import sys
import shutil
import pandas as pd
from datetime import datetime
import re
from PIL import Image
from PIL.IptcImagePlugin import getiptcinfo

# --- Cesty ---
tag_to_find = None
SOURCE_PATH_ORIGINAL    = r"\\NAS\spolecne\1. PRODUKTOVÉ FOTKY\AKTUÁLNÍ"
SOURCE_PATH_PROMO_FOTO  = r"\\NAS\spolecne\00 - PROMO FOTOGRAFIE A VIDEA\fotky"
SOURCE_PATH_PROMO_VIDEA = r"\\NAS\spolecne\00 - PROMO FOTOGRAFIE A VIDEA\videa"

# --- Logger ---
def setup_logger(script_dir):
    log_file = os.path.join(script_dir, "vypis_konzole.txt")
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
def copy_photos_by_excel(source_dir, dest_dir, mapping, flat_structure=False, root_mode=False, log=None, tag_to_find=None):
    exts = [".jpg",".jpeg",".png",".gif",".bmp",".tif",".tiff",
            ".mp4",".avi",".mov",".mkv",".wmv",".flv"]
    os.makedirs(dest_dir, exist_ok=True)
    copied_count = 0
    mapping_lower = {str(k).strip().lower(): v for k, v in mapping.items()}

    for root, _, files in os.walk(source_dir):
        for fn in files:
            if not any(fn.lower().endswith(ext) for ext in exts):
                continue

            filepath = os.path.join(root, fn)
            if tag_to_find and not has_tag(filepath, tag_to_find, log):
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

# --- Kontrola metadat souboru ---
def has_tag(filepath, tag_to_find, log):
    image_exts = [".jpg", ".jpeg", ".tif", ".tiff"]
    if not any(filepath.lower().endswith(ext) for ext in image_exts):
        return False

    tag_lower = tag_to_find.lower()
    file_tags = set()
    
    try:
        with Image.open(filepath) as img:
            iptc = getiptcinfo(img)
            if iptc and (2, 25) in iptc:
                keywords = iptc[(2, 25)]
                decoded_keywords = []
                if isinstance(keywords, list):
                    for kw_bytes in keywords:
                        decoded_keywords.append(kw_bytes.decode('utf-8', errors='ignore'))
                else:
                    decoded_keywords.append(keywords.decode('utf-8', errors='ignore'))
                file_tags.update(k.lower().strip() for k in decoded_keywords)

            exif_data = img.getexif()
            XPKEYWORDS_TAG = 40094
            if exif_data and XPKEYWORDS_TAG in exif_data:
                keywords_bytes = exif_data[XPKEYWORDS_TAG]
                keywords_str = keywords_bytes.decode('utf-16-le').rstrip('\x00')
                tags_from_exif = [t.strip() for t in keywords_str.split(';') if t.strip()]
                file_tags.update(t.lower() for t in tags_from_exif)
    except Exception as e:
        log(f"Varování: Nelze přečíst metadata ze souboru {os.path.basename(filepath)}. Chyba: {e}")
        return False
    
    return tag_lower in file_tags

# --- Kopírování prvního média ---
def copy_first_media(src_dir, dest_dir, log, tag_to_find=None):
    exts = [".jpg",".jpeg",".png",".gif",".bmp",".tif",".tiff",".mp4",".avi",".mov",".mkv",".wmv",".flv"]
    try:
        files = sorted(os.listdir(src_dir))
    except Exception as e:
        log(f"Chyba při čtení složky {src_dir}: {e}")
        return
        
    for fn in files:
        if any(fn.lower().endswith(ext) for ext in exts):
            filepath = os.path.join(src_dir, fn)
            if tag_to_find and not has_tag(filepath, tag_to_find, log):
                log(f"Přeskočeno: soubor '{fn}' neobsahuje štítek '{tag_to_find}'. Hledám dál...")
                continue
            
            os.makedirs(dest_dir, exist_ok=True)
            shutil.copy2(filepath, os.path.join(dest_dir, fn))
            log(f"Zkopírován první soubor '{fn}' ze složky '{os.path.basename(src_dir)}'.")
            return

# --- Kopírování složek podle Excelu ---
def copy_folders_with_mapping(source_path, dest_path, mapping, copy_mode, flat_structure=False, root_mode=False, log=None, tag_to_find=None):
    unfound = set(mapping.keys())
    for root, dirs, _ in os.walk(source_path):
        for folder in list(dirs):
            if folder in unfound:
                znacka, kategorie = mapping[folder]
                src_dir = os.path.join(root, folder)
                out_dir = get_output_dir(dest_path, znacka, kategorie, folder, root_mode, flat_structure)

                if copy_mode == "all":
                    log(f"Kopíruji obsah složky '{folder}' do '{out_dir}' s filtrem na štítek.")
                    files_copied_in_folder = 0
                    for src_root, _, src_files in os.walk(src_dir):
                        relative_path = os.path.relpath(src_root, src_dir)
                        dest_subdir = os.path.join(out_dir, relative_path) if relative_path != '.' else out_dir
                        os.makedirs(dest_subdir, exist_ok=True)
                        
                        for file in src_files:
                            src_file_path = os.path.join(src_root, file)
                            if tag_to_find and not has_tag(src_file_path, tag_to_find, log):
                                continue
                            
                            shutil.copy2(src_file_path, os.path.join(dest_subdir, file))
                            files_copied_in_folder += 1
                    if files_copied_in_folder > 0:
                        log(f"Zkopírováno {files_copied_in_folder} souborů (splňujících filtr) ze složky '{folder}'.")
                else: # copy_mode == "first"
                    copy_first_media(src_dir, out_dir, log, tag_to_find=tag_to_find)
                
                unfound.remove(folder)
                dirs.remove(folder) 
    return unfound

if __name__ == "__main__":
    script_dir = os.path.dirname(sys.executable if getattr(sys, 'frozen', False) else os.path.abspath(__file__))
    log = setup_logger(script_dir)

    print("=== Kopírování fotek podle Excelu ===")

    print("\nZvol režim kopírování:")
    print("1) Celé složky")
    print("2) První soubor")
    print("3) Podle excelu")
    mode = input("Zadejte číslo režimu: ").strip()
    if mode not in ["1", "2", "3"]:
        log("Neplatná volba režimu.")
        sys.exit(1)
    
    tag_to_find = None
    print("\nChcete použít filtr podle štítků?")
    print("1) Ano")
    print("2) Ne")
    tag_choice = input("Zadejte volbu: ").strip()
    if tag_choice == "1":
        tag_to_find = input("Zadejte štítek, podle kterého se bude filtrovat: ").strip()
        if not tag_to_find:
            log("Chyba: Nebyl zadán žádný štítek pro filtrování.")
            sys.exit(1)
    elif tag_choice != "2":
        log("Neplatná volba pro filtrování podle štítků.")
        sys.exit(1)
    
    print("\nZvol způsob třídění:")
    print("1) Podle struktury")
    print("2) Plochá struktura")
    print("3) Vše do jedné složky")
    sort_choice = input("Zadejte číslo: ").strip()
    
    print("\nZvol zdrojovou složku:")
    print("1) Promo fotky")
    print("2) Promo videa")
    print("3) Vlastní složka")
    print("4) Produktové fotky")
    src_choice = input("Zadejte číslo zdroje: ").strip()

    if src_choice == "1": source_path = SOURCE_PATH_PROMO_FOTO
    elif src_choice == "2": source_path = SOURCE_PATH_PROMO_VIDEA
    elif src_choice == "4": source_path = SOURCE_PATH_ORIGINAL
    else:
        source_path = input("Zadejte cestu ke zdrojové složce: ").strip()

    dest_path = os.path.join(script_dir, "foto_folders")

    if os.path.exists(dest_path):
                try:
                    shutil.rmtree(dest_path)
                    log(f"Odstraněna stará složka: {dest_path}")
                except Exception as e:
                    log(f"Nepodařilo se odstranit složku '{dest_path}': {e}")
                    sys.exit(1)

    os.makedirs(dest_path, exist_ok=True)

    excel_path = None
    for fn in os.listdir(script_dir):
        if fn.lower().endswith((".xlsx", ".xls", ".xlsm")):
            excel_path = os.path.join(script_dir, fn)
            break

    if not excel_path:
        log("❌ Excel nebyl nalezen ve složce se skriptem.")
        sys.exit(1)

    require_structure = (sort_choice == "1")
    flat_structure = (sort_choice == "2")
    root_mode = (sort_choice == "3")

    try:
        mapping = get_mapping_from_excel(excel_path, require_structure=require_structure)
        log(f"Načten Excel: {excel_path}")
    except Exception as e:
        log(f"Chyba při načítání Excelu: {e}")
        sys.exit(1)

    if mode == "3":
        copy_photos_by_excel(source_path, dest_path, mapping, flat_structure, root_mode, log=log, tag_to_find=tag_to_find)
    else:
        copy_mode = "all" if mode == "1" else "first"
        unfound = copy_folders_with_mapping(source_path, dest_path, mapping, copy_mode, flat_structure, root_mode, log=log, tag_to_find=tag_to_find)
        if unfound:
            uf = os.path.join(script_dir, "unfound_folders.txt")
            with open(uf, "w", encoding="utf-8") as f:
                for k in unfound:
                    f.write(k + "\n")
            log(f"Některé složky nebyly nalezeny. Seznam uložen: {uf}")

    log("✅ Kopírování dokončeno.")
