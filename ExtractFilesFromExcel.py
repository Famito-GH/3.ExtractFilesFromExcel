# -- Verze 2.0 Filtr Tagů a Počtu produktů -- GUI Edition --

import os
import sys
import shutil
import threading
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from tkinter.scrolledtext import ScrolledText
import pandas as pd
from datetime import datetime
import re
from enum import Enum
from PIL import Image
from PIL.IptcImagePlugin import getiptcinfo

# ---------------------------------------------------------------------------
# Konstanty a výchozí cesty
# ---------------------------------------------------------------------------
SOURCE_PATH_ORIGINAL    = r"\\NAS\spolecne\1. PRODUKTOVÉ FOTKY\AKTUÁLNÍ"
SOURCE_PATH_PROMO_FOTO  = r"\\NAS\spolecne\00 - PROMO FOTOGRAFIE A VIDEA\fotky"
SOURCE_PATH_PROMO_VIDEA = r"\\NAS\spolecne\00 - PROMO FOTOGRAFIE A VIDEA\videa"

MEDIA_EXTENSIONS = frozenset([
    ".jpg", ".jpeg", ".png", ".gif", ".bmp", ".tif", ".tiff",
    ".mp4", ".avi", ".mov", ".mkv", ".wmv", ".flv",
])


class CopyMode(Enum):
    ALL_FOLDERS = "1"
    FIRST_FILE  = "2"
    BY_EXCEL    = "3"


class MultiChoice(Enum):
    SINGLE   = "1"   # Jeden produkt na fotce (nazev bez carky)
    MULTIPLE = "2"   # Vice produktu na fotce (nazev s carkou)
    ANY      = "3"   # Nezalezi


class SortMode(Enum):
    STRUCTURE = "1"  # Znacka / Kategorie / Produkt
    FLAT      = "2"  # Jen Produkt
    ROOT      = "3"  # Vse do jedne slozky


# ---------------------------------------------------------------------------
# Logger
# ---------------------------------------------------------------------------
def setup_logger(script_dir: str):
    log_file = os.path.join(script_dir, "vypis_konzole.txt")
    if os.path.exists(log_file):
        os.remove(log_file)
    open(log_file, "w", encoding="utf-8").close()

    def log(msg: str):
        ts   = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        line = f"[{ts}] {msg}"
        print(line)
        with open(log_file, "a", encoding="utf-8") as f:
            f.write(line + "\n")

    return log


# ---------------------------------------------------------------------------
# Pomocne funkce
# ---------------------------------------------------------------------------
def clean_cell(val):
    """Vrati cisty string nebo None pro prazdne / neplatne hodnoty."""
    if pd.isna(val):
        return None
    s = str(val).strip()
    if not s or s.lower() in ("nan", "none", "n/a", "na", "-", "null"):
        return None
    return s


def is_media_file(filename: str) -> bool:
    return os.path.splitext(filename)[1].lower() in MEDIA_EXTENSIONS


def passes_multi_filter(name_no_ext: str, multi_choice) -> bool:
    """Vrati True pokud nazev souboru splnuje filtr poctu produktu."""
    if multi_choice is None or multi_choice == MultiChoice.ANY:
        return True
    has_comma = "," in name_no_ext
    if multi_choice == MultiChoice.SINGLE and has_comma:
        return False
    if multi_choice == MultiChoice.MULTIPLE and not has_comma:
        return False
    return True


def extract_product_codes(name_no_ext: str):
    """Extrahuje kody produktu z nazvu souboru (oddelene carkami, bez zavorek)."""
    parts = [p.strip() for p in re.split(r",", name_no_ext) if p.strip()]
    return [re.sub(r"\([^)]*\)", "", p).strip() for p in parts if p.strip()]


def get_output_dir(dest_dir, znacka, kategorie, original_product, root_mode=False, flat_structure=False):
    """Vrati vystupni cestu podle zvoleného zpusobu trideni."""
    if root_mode:
        return dest_dir
    if flat_structure:
        return os.path.join(dest_dir, original_product)
    z = znacka    or "Nezařazeno"
    k = kategorie or "Nezařazeno"
    return os.path.join(dest_dir, z, k, original_product)


# ---------------------------------------------------------------------------
# Nacteni Excelu
# ---------------------------------------------------------------------------
def get_mapping_from_excel(excel_path: str, require_structure: bool = True):
    """
    Vrati slovnik {kod_produktu: (znacka, kategorie)}.
    Pri require_structure=True vyzaduje sloupce Znacka a Kategorie.
    """
    df = pd.read_excel(excel_path)
    col_kod = col_znacka = col_kategorie = None

    for col in df.columns:
        n = str(col).strip().lower()
        if "kód" in n and col_kod is None:
            col_kod = col
        elif "značka" in n and col_znacka is None:
            col_znacka = col
        elif "kategorie" in n and col_kategorie is None:
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
        znacka    = clean_cell(row[col_znacka])    if col_znacka    else None
        kategorie = clean_cell(row[col_kategorie]) if col_kategorie else None
        mapping[kod] = (znacka, kategorie)

    return mapping


# ---------------------------------------------------------------------------
# Kontrola IPTC / EXIF tagu (podporuje vice stitku a AND/OR logiku)
# ---------------------------------------------------------------------------
def _read_file_tags(filepath: str, log) -> set:
    """Precte vsechny IPTC/EXIF stitky souboru a vrati je jako set malych stringu."""
    image_exts = frozenset([".jpg", ".jpeg", ".tif", ".tiff"])
    if os.path.splitext(filepath)[1].lower() not in image_exts:
        return set()
    file_tags = set()
    try:
        with Image.open(filepath) as img:
            iptc = getiptcinfo(img)
            if iptc and (2, 25) in iptc:
                keywords = iptc[(2, 25)]
                if not isinstance(keywords, list):
                    keywords = [keywords]
                file_tags.update(kw.decode("utf-8", errors="ignore").lower().strip() for kw in keywords)
            exif_data = img.getexif()
            if exif_data and 40094 in exif_data:
                raw = exif_data[40094]
                kw_str = raw.decode("utf-16-le").rstrip("\x00")
                file_tags.update(t.strip().lower() for t in kw_str.split(";") if t.strip())
    except Exception as e:
        log(f"Varování: Nelze přečíst metadata ze souboru {os.path.basename(filepath)}. Chyba: {e}")
    return file_tags


def has_tags(filepath: str, tags: list, tag_mode: str, log) -> bool:
    """
    Vrati True pokud soubor splnuje podminku stitku.
    tags     – seznam retezcu (malymi pismeny)
    tag_mode – 'and' (vsechny stitky museji byt pritomny)
               'or'  (staci alespon jeden)
    """
    if not tags:
        return True
    file_tags = _read_file_tags(filepath, log)
    if not file_tags:
        return False
    if tag_mode == "and":
        return all(t in file_tags for t in tags)
    else:  # or
        return any(t in file_tags for t in tags)




# ---------------------------------------------------------------------------
# Kopirovani fotek podle Excelu (mode BY_EXCEL)
# ---------------------------------------------------------------------------
def copy_photos_by_excel(source_dir, dest_dir, mapping, flat_structure=False, root_mode=False,
                          log=None, tags=None, tag_mode="or", multi_choice=None):
    os.makedirs(dest_dir, exist_ok=True)
    mapping_lower = {k.strip().lower(): v for k, v in mapping.items()}
    copied_count  = 0
    tags          = tags or []

    for root, _, files in os.walk(source_dir):
        for fn in files:
            if not is_media_file(fn):
                continue

            name_no_ext = os.path.splitext(fn)[0]

            # --- Filtr poctu produktu ---
            if not passes_multi_filter(name_no_ext, multi_choice):
                continue

            filepath = os.path.join(root, fn)

            # --- Filtr stitku ---
            if tags and not has_tags(filepath, tags, tag_mode, log):
                continue

            # --- Extrakce a validace kodu vuci Excelu ---
            products = extract_product_codes(name_no_ext)
            if not products:
                log(f"Přeskočeno (nelze detekovat produkty): {fn}")
                continue

            missing = [p for p in products if p.lower() not in mapping_lower]
            if missing:
                log(f"Přeskočeno (produkty nenalezeny v Excelu: {', '.join(missing)}): {fn}")
                continue

            # --- Kopirovani pro kazdy produkt ---
            for prod in products:
                prod_key          = prod.lower().strip()
                znacka, kategorie = mapping_lower[prod_key]
                original_product  = next(
                    (k for k in mapping if k.strip().lower() == prod_key), prod
                )
                out_dir = get_output_dir(dest_dir, znacka, kategorie, original_product, root_mode, flat_structure)
                os.makedirs(out_dir, exist_ok=True)
                try:
                    shutil.copy2(filepath, os.path.join(out_dir, fn))
                    log(f"Zkopírován: {fn} -> {out_dir}")
                    copied_count += 1
                except Exception as e:
                    log(f"Chyba při kopírování {fn} -> {out_dir}: {e}")

    msg = f"Celkem zkopírováno {copied_count} souborů." if copied_count else "Nenalezeny žádné soubory."
    log(msg)


# ---------------------------------------------------------------------------
# Kopirovani prvniho media ze slozky
# ---------------------------------------------------------------------------
def copy_first_media(src_dir, dest_dir, log, mapping_lower=None,
                     tags=None, tag_mode="or", multi_choice=None):
    tags = tags or []
    try:
        files = sorted(os.listdir(src_dir))
    except Exception as e:
        log(f"Chyba při čtení složky {src_dir}: {e}")
        return

    for fn in files:
        if not is_media_file(fn):
            continue

        name_no_ext = os.path.splitext(fn)[0]

        # --- Filtr poctu produktu ---
        if not passes_multi_filter(name_no_ext, multi_choice):
            continue

        # --- Kontrola produktu v Excelu ---
        if mapping_lower is not None:
            products = extract_product_codes(name_no_ext)
            if products:
                missing = [p for p in products if p.lower() not in mapping_lower]
                if missing:
                    log(f"Přeskočeno (produkty nenalezeny v Excelu: {', '.join(missing)}): {fn}")
                    continue

        filepath = os.path.join(src_dir, fn)

        # --- Filtr stitku ---
        if tags and not has_tags(filepath, tags, tag_mode, log):
            log(f"Přeskočeno: '{fn}' nesplňuje filtr štítků. Hledám dál...")
            continue

        os.makedirs(dest_dir, exist_ok=True)
        shutil.copy2(filepath, os.path.join(dest_dir, fn))
        log(f"Zkopírován první soubor '{fn}' ze složky '{os.path.basename(src_dir)}'.")
        return


# ---------------------------------------------------------------------------
# Kopirovani slozek podle Excelu (mode ALL_FOLDERS / FIRST_FILE)
# ---------------------------------------------------------------------------
def copy_folders_with_mapping(source_path, dest_path, mapping, copy_mode,
                               flat_structure=False, root_mode=False,
                               log=None, tags=None, tag_mode="or", multi_choice=None):
    unfound       = set(mapping.keys())
    mapping_lower = {k.strip().lower(): v for k, v in mapping.items()}
    tags          = tags or []

    for root, dirs, _ in os.walk(source_path):
        for folder in list(dirs):
            if folder not in unfound:
                continue

            znacka, kategorie = mapping[folder]

            src_dir = os.path.join(root, folder)
            out_dir = get_output_dir(dest_path, znacka, kategorie, folder, root_mode, flat_structure)

            if copy_mode == "all":
                log(f"Kopíruji obsah složky '{folder}' -> '{out_dir}' (s aktivními filtry).")
                files_copied = 0

                for src_root, _, src_files in os.walk(src_dir):
                    for file in src_files:
                        if not is_media_file(file):
                            continue

                        name_no_ext = os.path.splitext(file)[0]

                        # --- Filtr poctu produktu ---
                        if not passes_multi_filter(name_no_ext, multi_choice):
                            continue

                        # --- Kontrola produktu v Excelu ---
                        products = extract_product_codes(name_no_ext)
                        if products:
                            missing = [p for p in products if p.lower() not in mapping_lower]
                            if missing:
                                log(f"Přeskočeno (produkty nenalezeny v Excelu: {', '.join(missing)}): {file}")
                                continue

                        src_file_path = os.path.join(src_root, file)

                        # --- Filtr stitku ---
                        if tags and not has_tags(src_file_path, tags, tag_mode, log):
                            continue

                        rel_path = os.path.relpath(src_root, src_dir)
                        dest_sub = os.path.join(out_dir, rel_path) if rel_path != "." else out_dir
                        os.makedirs(dest_sub, exist_ok=True)
                        shutil.copy2(src_file_path, os.path.join(dest_sub, file))
                        files_copied += 1

                if files_copied:
                    log(f"Zkopírováno {files_copied} souborů (splňujících filtr) ze složky '{folder}'.")

            else:  # copy_mode == "first"
                copy_first_media(
                    src_dir, out_dir, log,
                    mapping_lower=mapping_lower,
                    tags=tags, tag_mode=tag_mode,
                    multi_choice=multi_choice,
                )

            unfound.discard(folder)
            dirs.remove(folder)

    return unfound


# ===========================================================================
# GUI
# ===========================================================================
class App(tk.Tk):
    # Barvy a styly
    BG          = "#ffffff"
    SURFACE     = "#f2f2f2"
    SURFACE2    = "#e0e0e0"
    ACCENT      = "#2563eb"
    ACCENT2     = "#1e40af"
    TEXT        = "#1a1a1a"
    TEXT_DIM    = "#666666"
    SUCCESS     = "#16a34a"
    WARNING     = "#d97706"
    FONT_FAMILY = "Segoe UI"

    def __init__(self):
        super().__init__()
        self.title("ExtractFilesFromExcel  v2.0")
        self.geometry("900x720")
        self.minsize(820, 640)
        self.configure(bg=self.BG)
        self.resizable(True, True)

        self._setup_styles()
        self._build_ui()

    # -----------------------------------------------------------------------
    def _setup_styles(self):
        style = ttk.Style(self)
        style.theme_use("clam")

        style.configure(".",
            background=self.BG,
            foreground=self.TEXT,
            font=(self.FONT_FAMILY, 10),
        )
        style.configure("TFrame",      background=self.BG)
        style.configure("Card.TFrame", background=self.SURFACE, relief="flat")

        style.configure("TLabel",      background=self.BG,      foreground=self.TEXT,     font=(self.FONT_FAMILY, 10))
        style.configure("Dim.TLabel",  background=self.BG,      foreground=self.TEXT_DIM, font=(self.FONT_FAMILY, 9))
        style.configure("Card.TLabel", background=self.SURFACE, foreground=self.TEXT,     font=(self.FONT_FAMILY, 10))
        style.configure("Head.TLabel", background=self.BG,      foreground=self.ACCENT,   font=(self.FONT_FAMILY, 13, "bold"))
        style.configure("Sec.TLabel",  background=self.SURFACE, foreground=self.ACCENT,   font=(self.FONT_FAMILY, 10, "bold"))

        style.configure("TRadiobutton",
            background=self.SURFACE, foreground=self.TEXT,
            font=(self.FONT_FAMILY, 10), focuscolor=self.SURFACE,
        )
        style.map("TRadiobutton",
            background=[("active", self.SURFACE2)],
            foreground=[("active", self.TEXT)],
        )

        style.configure("TCheckbutton",
            background=self.SURFACE, foreground=self.TEXT,
            font=(self.FONT_FAMILY, 10), focuscolor=self.SURFACE,
        )
        style.map("TCheckbutton", background=[("active", self.SURFACE2)])

        style.configure("Run.TButton",
            background=self.ACCENT, foreground="#ffffff",
            font=(self.FONT_FAMILY, 11, "bold"),
            borderwidth=0, relief="flat", padding=(16, 8),
        )
        style.map("Run.TButton",
            background=[("active", "#c73652"), ("disabled", "#555577")],
        )
        style.configure("Outline.TButton",
            background=self.SURFACE, foreground=self.ACCENT,
            font=(self.FONT_FAMILY, 10),
            borderwidth=1, relief="solid", padding=(8, 4),
        )
        style.map("Outline.TButton",
            background=[("active", self.SURFACE2)],
            foreground=[("active", self.ACCENT)],
        )

        style.configure("TEntry",
            fieldbackground=self.SURFACE2, foreground=self.TEXT,
            insertcolor=self.TEXT, borderwidth=1, relief="solid",
            font=(self.FONT_FAMILY, 10),
        )

    # -----------------------------------------------------------------------
    def _card(self, parent, title=""):
        """Vrati ramecek karty s volitelnym nadpisem."""
        outer = ttk.Frame(parent, style="Card.TFrame", padding=12)
        if title:
            ttk.Label(outer, text=title, style="Sec.TLabel").pack(anchor="w", pady=(0, 8))
        return outer

    # -----------------------------------------------------------------------
    def _build_ui(self):
        # ---- Zahlavi ----
        header = ttk.Frame(self, padding=(20, 12, 20, 8))
        header.pack(fill="x")
        ttk.Label(header, text="Extract Files From Excel", style="Head.TLabel").pack(side="left")

        ttk.Separator(self, orient="horizontal").pack(fill="x", padx=20)

        # ---- Log sekce (dole, pevna vyska) ----
        log_frame = ttk.Frame(self)
        log_frame.pack(side="bottom", fill="x", padx=20, pady=(0, 10))
        ttk.Separator(self, orient="horizontal").pack(side="bottom", fill="x", padx=20)
        ttk.Label(self, text="Výstup / Log", style="Head.TLabel").pack(side="bottom", anchor="w", padx=20, pady=(6, 2))

        self.log_box = ScrolledText(
            log_frame, wrap="word", height=6,
            bg=self.SURFACE, fg=self.TEXT,
            insertbackground=self.TEXT,
            font=(self.FONT_FAMILY, 9),
            relief="flat", bd=0,
            padx=8, pady=6,
        )
        self.log_box.pack(fill="x")
        self.log_box.configure(state="disabled")
        self.log_box.tag_config("info",    foreground=self.TEXT)
        self.log_box.tag_config("success", foreground=self.SUCCESS)
        self.log_box.tag_config("warning", foreground=self.WARNING)
        self.log_box.tag_config("error",   foreground=self.ACCENT)

        # ---- Hlavni oblast (scroll) ----
        scroll_area = ttk.Frame(self)
        scroll_area.pack(side="top", fill="both", expand=True)

        main_canvas = tk.Canvas(scroll_area, bg=self.BG, highlightthickness=0)
        scrollbar   = ttk.Scrollbar(scroll_area, orient="vertical", command=main_canvas.yview)
        self.scroll_frame = ttk.Frame(main_canvas)

        self.scroll_frame.bind("<Configure>",
            lambda e: main_canvas.configure(scrollregion=main_canvas.bbox("all"))
        )
        main_canvas.create_window((0, 0), window=self.scroll_frame, anchor="nw")
        main_canvas.configure(yscrollcommand=scrollbar.set)

        scrollbar.pack(side="right", fill="y")
        main_canvas.pack(side="left", fill="both", expand=True)
        main_canvas.bind_all("<MouseWheel>",
            lambda e: main_canvas.yview_scroll(-1 * (e.delta // 120), "units")
        )

        pad = dict(padx=20, pady=6, fill="x")

        # ---- Sekce: Rezim kopirovani ----
        c1 = self._card(self.scroll_frame, "Režim kopírování")
        c1.pack(**pad)
        self.mode_var = tk.StringVar(value=CopyMode.ALL_FOLDERS.value)
        for label, val in [
            ("Celé složky",  CopyMode.ALL_FOLDERS.value),
            ("První soubor", CopyMode.FIRST_FILE.value),
            ("Podle Excelu", CopyMode.BY_EXCEL.value),
        ]:
            ttk.Radiobutton(c1, text=label, variable=self.mode_var, value=val).pack(anchor="w")

        # ---- Sekce: Mnozstvi produktu ----
        c2 = self._card(self.scroll_frame, "Množství produktů na fotce")
        c2.pack(**pad)
        self.multi_var = tk.StringVar(value=MultiChoice.ANY.value)
        for label, val in [
            ("Jeden produkt",  MultiChoice.SINGLE.value),
            ("Více produktů",   MultiChoice.MULTIPLE.value),
            ("Nezáleží",                            MultiChoice.ANY.value),
        ]:
            ttk.Radiobutton(c2, text=label, variable=self.multi_var, value=val).pack(anchor="w")
        ttk.Label(c2,
            text="ℹ  Produkty jsou vždy ověřeny vůči Excelu bez ohledu na volbu.",
            style="Dim.TLabel",
        ).pack(anchor="w", pady=(6, 0))

        # ---- Sekce: Filtr stitku ----
        c3 = self._card(self.scroll_frame, "Filtr štítků")
        c3.pack(**pad)
        self.tag_enabled = tk.BooleanVar(value=False)
        ttk.Checkbutton(c3, text="Filtrovat podle štítků", variable=self.tag_enabled,
                        command=self._toggle_tag).pack(anchor="w")
        ttk.Label(c3, text="Více štítků oddělte čárkou např: pánské, černá",
                  style="Dim.TLabel").pack(anchor="w", pady=(6, 2))
        self.tag_entry = ttk.Entry(c3, width=50)
        self.tag_entry.pack(anchor="w", ipady=4)
        self.tag_entry.configure(state="disabled")
        ttk.Label(c3, text="Logika filtrování:", style="Dim.TLabel").pack(anchor="w", pady=(8, 2))
        self.tag_mode_var = tk.StringVar(value="or")
        tag_mode_row = ttk.Frame(c3, style="Card.TFrame")
        tag_mode_row.pack(anchor="w")
        ttk.Radiobutton(tag_mode_row, text="Alespoň jeden štítek",
                        variable=self.tag_mode_var, value="or").pack(side="left", padx=(0, 12))
        ttk.Radiobutton(tag_mode_row, text="Všechny štítky",
                        variable=self.tag_mode_var, value="and").pack(side="left")

        # ---- Sekce: Zpusob trideni ----
        c4 = self._card(self.scroll_frame, "Způsob třídění výstupu")
        c4.pack(**pad)
        self.sort_var = tk.StringVar(value=SortMode.STRUCTURE.value)
        for label, val in [
            ("Podle struktury", SortMode.STRUCTURE.value),
            ("Plochá struktura", SortMode.FLAT.value),
            ("Vše do jedné složky", SortMode.ROOT.value),
        ]:
            ttk.Radiobutton(c4, text=label, variable=self.sort_var, value=val).pack(anchor="w")

        # ---- Sekce: Zdrojova slozka ----
        c5 = self._card(self.scroll_frame, "Zdrojová složka")
        c5.pack(**pad)
        self.src_var = tk.StringVar(value="4")
        for label, val in [
            ("Promo fotky","1"),
            ("Promo videa","2"),
            ("Vlastní cesta…","3"),
            ("Produktové fotky","4"),
        ]:
            ttk.Radiobutton(c5, text=label, variable=self.src_var, value=val,
                            command=self._toggle_custom_path).pack(anchor="w")

        custom_row = ttk.Frame(c5, style="Card.TFrame")
        custom_row.pack(fill="x", pady=(4, 0))
        self.custom_path_entry = ttk.Entry(custom_row, width=50)
        self.custom_path_entry.pack(side="left", ipady=4, padx=(0, 6))
        ttk.Button(custom_row, text="Procházet…", style="Outline.TButton",
                   command=self._browse_source).pack(side="left")
        self.custom_path_entry.configure(state="disabled")

        # ---- Excel status ----
        excel_frame = ttk.Frame(self.scroll_frame)
        excel_frame.pack(padx=20, pady=(2, 6), fill="x")
        self.excel_label = ttk.Label(excel_frame, text="🔍 Hledám Excel soubor…", style="Dim.TLabel")
        self.excel_label.pack(side="left")
        ttk.Button(excel_frame, text="Vybrat Excel…", style="Outline.TButton",
                   command=self._browse_excel).pack(side="right")

        self.excel_path_var = tk.StringVar()
        self._auto_detect_excel()

        # ---- Tlacitko Spustit ----
        btn_frame = ttk.Frame(self.scroll_frame)
        btn_frame.pack(padx=20, pady=(8, 14), fill="x")
        self.run_btn = ttk.Button(btn_frame, text="▶  Spustit kopírování", style="Run.TButton",
                                  command=self._on_run)
        self.run_btn.pack(side="right")

        # (log sekce je definovana vyse, pred scroll oblasti)

    # -----------------------------------------------------------------------
    # Pomocne metody UI
    # -----------------------------------------------------------------------
    def _toggle_tag(self):
        state = "normal" if self.tag_enabled.get() else "disabled"
        self.tag_entry.configure(state=state)

    def _toggle_custom_path(self):
        state = "normal" if self.src_var.get() == "3" else "disabled"
        self.custom_path_entry.configure(state=state)

    def _browse_source(self):
        path = filedialog.askdirectory(title="Vyberte zdrojovou složku")
        if path:
            self.custom_path_entry.configure(state="normal")
            self.custom_path_entry.delete(0, "end")
            self.custom_path_entry.insert(0, path)

    def _browse_excel(self):
        path = filedialog.askopenfilename(
            title="Vyberte Excel soubor",
            filetypes=[("Excel soubory", "*.xlsx *.xls *.xlsm"), ("Vše", "*.*")],
        )
        if path:
            self.excel_path_var.set(path)
            self.excel_label.configure(text=f"✅ Excel: {os.path.basename(path)}", foreground=self.SUCCESS)

    def _auto_detect_excel(self):
        script_dir = os.path.dirname(os.path.abspath(__file__))
        for fn in os.listdir(script_dir):
            if fn.lower().endswith((".xlsx", ".xls", ".xlsm")):
                path = os.path.join(script_dir, fn)
                self.excel_path_var.set(path)
                self.excel_label.configure(
                    text=f"✅ Nalezen Excel: {fn}",
                    foreground=self.SUCCESS,
                )
                return
        self.excel_label.configure(
            text="❌ Excel nebyl nalezen – vyberte soubor ručně.",
            foreground=self.ACCENT,
        )

    # -----------------------------------------------------------------------
    def _ui_log(self, msg: str):
        """Zapise zpravu do log boxu s barevnym zvyraznenim (thread-safe)."""
        self.after(0, self._append_log, msg)

    def _append_log(self, msg: str):
        ts   = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        line = f"[{ts}] {msg}\n"

        tag = "info"
        ml  = msg.lower()
        if "✅" in msg or "zkopírován" in ml:
            tag = "success"
        elif "❌" in msg or "chyba" in ml:
            tag = "error"
        elif "varování" in ml or "přeskočeno" in ml or "nenalezen" in ml:
            tag = "warning"

        self.log_box.configure(state="normal")
        self.log_box.insert("end", line, tag)
        self.log_box.see("end")
        self.log_box.configure(state="disabled")

        # Zapis do souboru
        script_dir = os.path.dirname(os.path.abspath(__file__))
        log_file   = os.path.join(script_dir, "vypis_konzole.txt")
        with open(log_file, "a", encoding="utf-8") as f:
            f.write(line)

    # -----------------------------------------------------------------------
    def _on_run(self):
        """Spusti kopirovani ve vedlejsim vlakne, aby GUI nezmrzlo."""
        self.run_btn.configure(state="disabled")
        self.log_box.configure(state="normal")
        self.log_box.delete("1.0", "end")
        self.log_box.configure(state="disabled")

        # Priprav log soubor
        script_dir = os.path.dirname(os.path.abspath(__file__))
        log_file   = os.path.join(script_dir, "vypis_konzole.txt")
        if os.path.exists(log_file):
            os.remove(log_file)
        open(log_file, "w", encoding="utf-8").close()

        thread = threading.Thread(target=self._run_copy, daemon=True)
        thread.start()

    # -----------------------------------------------------------------------
    def _run_copy(self):
        log = self._ui_log
        script_dir = os.path.dirname(os.path.abspath(__file__))

        # --- Validace vstupu ---
        excel_path = self.excel_path_var.get()
        if not excel_path or not os.path.isfile(excel_path):
            log("❌ Excel soubor nebyl nalezen nebo zvolen.")
            self.after(0, lambda: self.run_btn.configure(state="normal"))
            return

        src_choice = self.src_var.get()
        if src_choice == "1":
            source_path = SOURCE_PATH_PROMO_FOTO
        elif src_choice == "2":
            source_path = SOURCE_PATH_PROMO_VIDEA
        elif src_choice == "4":
            source_path = SOURCE_PATH_ORIGINAL
        else:
            source_path = self.custom_path_entry.get().strip()

        if not source_path or not os.path.isdir(source_path):
            log(f"❌ Zdrojová složka neexistuje: '{source_path}'")
            self.after(0, lambda: self.run_btn.configure(state="normal"))
            return

        # --- Filtr stitku ---
        tags     = []
        tag_mode = self.tag_mode_var.get()
        if self.tag_enabled.get():
            raw_tags = self.tag_entry.get().strip()
            if not raw_tags:
                log("❌ Filtr štítků je zapnut, ale žádný štítek nebyl zadán.")
                self.after(0, lambda: self.run_btn.configure(state="normal"))
                return
            tags = [t.strip().lower() for t in raw_tags.split(",") if t.strip()]
            log(f"Filtr štítků ({tag_mode.upper()}): {', '.join(tags)}")

        mode_val  = self.mode_var.get()
        multi_val = self.multi_var.get()
        sort_val  = self.sort_var.get()

        mode             = CopyMode(mode_val)
        multi_choice     = MultiChoice(multi_val)
        flat_structure   = (sort_val == SortMode.FLAT.value)
        root_mode        = (sort_val == SortMode.ROOT.value)
        require_structure = (sort_val == SortMode.STRUCTURE.value)

        dest_path = os.path.join(script_dir, "foto_folders")

        # --- Smazani stare vystupni slozky ---
        if os.path.exists(dest_path):
            try:
                shutil.rmtree(dest_path)
                log(f"Odstraněna stará složka: {dest_path}")
            except Exception as e:
                log(f"❌ Nepodařilo se odstranit složku '{dest_path}': {e}")
                self.after(0, lambda: self.run_btn.configure(state="normal"))
                return
        os.makedirs(dest_path, exist_ok=True)

        # --- Nacteni Excelu ---
        try:
            mapping = get_mapping_from_excel(excel_path, require_structure=require_structure)
            log(f"Načten Excel: {os.path.basename(excel_path)}  ({len(mapping)} produktů)")
        except Exception as e:
            log(f"❌ Chyba při načítání Excelu: {e}")
            self.after(0, lambda: self.run_btn.configure(state="normal"))
            return

        log(f"Zdrojová složka: {source_path}")
        log(f"Výstupní složka: {dest_path}")

        try:
            if mode == CopyMode.BY_EXCEL:
                copy_photos_by_excel(
                    source_path, dest_path, mapping,
                    flat_structure=flat_structure,
                    root_mode=root_mode,
                    log=log,
                    tags=tags, tag_mode=tag_mode,
                    multi_choice=multi_choice,
                )
            else:
                copy_mode_str = "all" if mode == CopyMode.ALL_FOLDERS else "first"
                unfound = copy_folders_with_mapping(
                    source_path, dest_path, mapping, copy_mode_str,
                    flat_structure=flat_structure,
                    root_mode=root_mode,
                    log=log,
                    tags=tags, tag_mode=tag_mode,
                    multi_choice=multi_choice,
                )
                if unfound:
                    uf_path = os.path.join(script_dir, "unfound_folders.txt")
                    with open(uf_path, "w", encoding="utf-8") as f:
                        for k in sorted(unfound):
                            f.write(k + "\n")
                    log(f"⚠  {len(unfound)} složek nebylo nalezeno. Seznam: unfound_folders.txt")

            log("✅ Kopírování dokončeno.")
        except Exception as e:
            log(f"❌ Neočekávaná chyba: {e}")
        finally:
            self.after(0, lambda: self.run_btn.configure(state="normal"))


# ===========================================================================
# Vstupni bod
# ===========================================================================
if __name__ == "__main__":
    try:
        app = App()
        app.mainloop()
    except Exception as exc:
        import traceback
        import tkinter.messagebox as mb
        mb.showerror("Chyba při spuštění", traceback.format_exc())
