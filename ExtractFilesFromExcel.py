# -- Verze 2.2 Filtr Tagů a Počtu produktů -- GUI Edition --

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


def get_app_dir() -> str:
    """Vrati slozku, ve ktere se nachazi spustitelny soubor (.exe) nebo skript (.py)."""
    if getattr(sys, 'frozen', False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


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

IGNORED_SYSTEM_FILES = frozenset([
    "thumbs.db", "desktop.ini", ".ds_store",
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


def is_ignored_system_file(filename: str) -> bool:
    return filename.lower() in IGNORED_SYSTEM_FILES or filename.startswith("._")


def extract_product_codes(name_no_ext: str):
    """Extrahuje kody produktu z nazvu souboru (oddelene carkami, bez zavorek)."""
    parts = [p.strip() for p in re.split(r",", name_no_ext) if p.strip()]
    cleaned = [re.sub(r"\([^)]*\)", "", p).strip() for p in parts]
    return [c for c in cleaned if c]


def resolve_product_code(raw_code: str, mapping_lower: dict) -> str:
    """
    Rozpozna kod produktu vuci Excelu:
    1) Nejprve zkusi presnou shodu (case-insensitive).
    2) Jinak postupne odrezava casti od '_' zprava doleva (napr. 'SPEC_01_2' -> 'SPEC_01',
       '6610 C_1' -> '6610 C', '7705 TS_5 a' -> '7705 TS', '1910 TI_5_6' -> '1910 TI').
    """
    if raw_code.lower() in mapping_lower:
        return raw_code
    if "_" in raw_code:
        candidate = raw_code
        while "_" in candidate:
            candidate = candidate.rsplit("_", 1)[0].strip()
            if candidate and candidate.lower() in mapping_lower:
                return candidate
        base = raw_code.split("_", 1)[0].strip()
        if base:
            return base
    return raw_code


def get_preserved_subfolder(root: str, source_dir: str, prod_key: str) -> str:
    """
    Urci relativni podslozku, ktera se ma zachovat pod vystupni slozkou v rezimu BY_EXCEL:
    1) Pokud zdrojova slozka nebo nadrazena cesta odpovida kodu produktu (case-insensitive),
       zachova celou podcestu pod touto produktovou slozkou (stejne jako rezim Cele slozky).
    2) Jinak (napr. Promo fotky) nadrazene slozky nepridava, pouze zachova pripadnou podslozku OLD.
    """
    rel = os.path.relpath(root, source_dir)
    if rel == ".":
        return ""

    src_name = os.path.basename(os.path.normpath(source_dir)).strip().lower()
    if src_name == prod_key:
        return rel

    parts = [p for p in rel.split(os.sep) if p and p != "."]

    for idx in range(len(parts) - 1, -1, -1):
        if parts[idx].strip().lower() == prod_key:
            sub_parts = parts[idx + 1:]
            return os.path.join(*sub_parts) if sub_parts else ""

    for idx in range(len(parts) - 1, -1, -1):
        if parts[idx].strip().lower() == "old":
            return os.path.join(*parts[idx:])

    return ""


def passes_multi_filter(name_no_ext: str, multi_choice) -> bool:
    """Vrati True pokud nazev souboru splnuje filtr poctu produktu."""
    if multi_choice is None or multi_choice == MultiChoice.ANY:
        return True
    products = extract_product_codes(name_no_ext)
    is_multiple = len(products) > 1
    if multi_choice == MultiChoice.SINGLE and is_multiple:
        return False
    if multi_choice == MultiChoice.MULTIPLE and not is_multiple:
        return False
    return True


def get_output_dir(dest_dir, znacka, kategorie, original_product, root_mode=False, flat_structure=False):
    """Vrati vystupni cestu podle zvoleného zpusobu trideni."""
    if root_mode:
        return dest_dir
    if flat_structure:
        return os.path.join(dest_dir, original_product)
    if znacka:
        k = kategorie or "Nezařazeno"
        return os.path.join(dest_dir, znacka, k, original_product)
    # Bez znacky – kategorie se pouzije jako prvni uroven (Kategorie / Produkt)
    if kategorie:
        return os.path.join(dest_dir, kategorie, original_product)
    return os.path.join(dest_dir, original_product)


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
    orig_by_lower = {k.strip().lower(): k for k in mapping}
    copied_count  = 0
    tags          = tags or []

    for root, _, files in os.walk(source_dir):
        for fn in files:
            if is_ignored_system_file(fn):
                continue
            if not is_media_file(fn):
                continue

            name_no_ext = os.path.splitext(fn)[0]

            # --- Filtr poctu produktu ---
            if not passes_multi_filter(name_no_ext, multi_choice):
                continue

            # --- Extrakce a validace kodu vuci Excelu (vcetne pripon za '_') ---
            raw_products = extract_product_codes(name_no_ext)
            if not raw_products:
                log(f"Přeskočeno (nelze detekovat produkty): {fn}")
                continue

            products = [resolve_product_code(p, mapping_lower) for p in raw_products]
            missing  = [p for p in products if p.lower() not in mapping_lower]
            if missing:
                log(f"Přeskočeno (produkty nenalezeny v Excelu: {', '.join(missing)}): {fn}")
                continue

            filepath = os.path.join(root, fn)

            # --- Filtr stitku (az po overeni v Excelu kvuli rychlosti na NASu) ---
            if tags and not has_tags(filepath, tags, tag_mode, log):
                continue

            # --- Kopirovani (se zachovanim podslozek uvnitr produktove slozky / OLD) ---
            out_dirs  = []
            seen_dirs = set()
            for prod in products:
                prod_key          = prod.lower().strip()
                znacka, kategorie = mapping_lower[prod_key]
                original_product  = orig_by_lower.get(prod_key, prod)
                base_dir = get_output_dir(
                    dest_dir, znacka, kategorie, original_product,
                    root_mode=root_mode, flat_structure=flat_structure,
                )
                rel_sub = get_preserved_subfolder(root, source_dir, prod_key)
                out_dir = os.path.join(base_dir, rel_sub) if rel_sub else base_dir
                if out_dir not in seen_dirs:
                    seen_dirs.add(out_dir)
                    out_dirs.append(out_dir)

            for out_dir in out_dirs:
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
# Kopirovani prvniho media ze slozky (mode FIRST_FILE)
# ---------------------------------------------------------------------------
def copy_first_media(src_dir, dest_dir, log, tags=None, tag_mode="or"):
    tags = tags or []
    try:
        files = sorted(os.listdir(src_dir))
    except Exception as e:
        log(f"Chyba při čtení složky {src_dir}: {e}")
        return False

    for fn in files:
        if is_ignored_system_file(fn):
            continue
        filepath = os.path.join(src_dir, fn)
        if not os.path.isfile(filepath):
            continue
        if not is_media_file(fn):
            continue

        # --- Filtr stitku ---
        if tags and not has_tags(filepath, tags, tag_mode, log):
            log(f"Přeskočeno: '{fn}' nesplňuje filtr štítků. Hledám dál...")
            continue

        try:
            os.makedirs(dest_dir, exist_ok=True)
            shutil.copy2(filepath, os.path.join(dest_dir, fn))
            log(f"Zkopírován první soubor '{fn}' ze složky '{os.path.basename(src_dir)}'.")
            return True
        except Exception as e:
            log(f"Chyba při kopírování '{fn}' -> '{dest_dir}': {e}")
            return False

    log(f"Varování: Ve složce '{os.path.basename(src_dir)}' nebyl nalezen žádný vyhovující mediální soubor.")
    return False


# ---------------------------------------------------------------------------
# Kopirovani slozek podle Excelu (mode ALL_FOLDERS / FIRST_FILE)
# ---------------------------------------------------------------------------
def copy_folders_with_mapping(source_path, dest_path, mapping, copy_mode,
                               flat_structure=False, root_mode=False,
                               log=None, tags=None, tag_mode="or"):
    orig_by_lower = {k.strip().lower(): k for k in mapping}
    unfound_lower = set(orig_by_lower.keys())
    total_copied  = 0
    tags          = tags or []

    for root, dirs, _ in os.walk(source_path):
        for folder in list(dirs):
            folder_key = folder.strip().lower()
            if folder_key not in unfound_lower:
                continue

            original_code     = orig_by_lower[folder_key]
            znacka, kategorie = mapping[original_code]

            src_dir = os.path.join(root, folder)
            out_dir = get_output_dir(dest_path, znacka, kategorie, original_code, root_mode, flat_structure)

            if copy_mode == "all":
                log(f"Kopíruji celou složku '{folder}' -> '{out_dir}'.")
                files_copied = 0

                for src_root, _, src_files in os.walk(src_dir):
                    for file in src_files:
                        if is_ignored_system_file(file):
                            continue

                        src_file_path = os.path.join(src_root, file)
                        rel_path      = os.path.relpath(src_root, src_dir)
                        dest_sub      = os.path.join(out_dir, rel_path) if rel_path != "." else out_dir

                        try:
                            os.makedirs(dest_sub, exist_ok=True)
                            shutil.copy2(src_file_path, os.path.join(dest_sub, file))
                            files_copied += 1
                            total_copied += 1
                        except Exception as e:
                            log(f"Chyba při kopírování '{file}' -> '{dest_sub}': {e}")

                if files_copied:
                    log(f"Zkopírováno {files_copied} souborů ze složky '{folder}'.")
                else:
                    log(f"Varování: Složka '{folder}' neobsahovala žádné soubory ke kopírování.")

            else:  # copy_mode == "first"
                if copy_first_media(
                    src_dir, out_dir, log,
                    tags=tags, tag_mode=tag_mode,
                ):
                    total_copied += 1

            unfound_lower.discard(folder_key)
            dirs.remove(folder)

    unfound = {orig_by_lower[k] for k in unfound_lower}
    return unfound, total_copied


# ===========================================================================
# GUI
# ===========================================================================
class App(tk.Tk):
    # Barvy a styly
    BG            = "#ffffff"
    SURFACE       = "#f2f2f2"
    SURFACE2      = "#e0e0e0"
    ACCENT        = "#2563eb"
    ACCENT2       = "#1e40af"
    TEXT          = "#1a1a1a"
    TEXT_DIM      = "#666666"
    TEXT_DISABLED = "#9ca3af"
    SUCCESS       = "#16a34a"
    WARNING       = "#d97706"
    FONT_FAMILY   = "Segoe UI"

    def __init__(self):
        super().__init__()
        self.title("ExtractFilesFromExcel  v2.2")
        self.geometry("900x720")
        self.minsize(820, 640)
        self.configure(bg=self.BG)
        self.resizable(True, True)

        self._is_running = False

        self._setup_styles()
        self._build_ui()
        self._update_ui_state()

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

        style.configure("TLabel",              background=self.BG,      foreground=self.TEXT,          font=(self.FONT_FAMILY, 10))
        style.configure("Dim.TLabel",          background=self.BG,      foreground=self.TEXT_DIM,      font=(self.FONT_FAMILY, 9))
        style.configure("Card.TLabel",         background=self.SURFACE, foreground=self.TEXT,          font=(self.FONT_FAMILY, 10))
        style.configure("CardDim.TLabel",      background=self.SURFACE, foreground=self.TEXT_DIM,      font=(self.FONT_FAMILY, 9))
        style.configure("CardDimDis.TLabel",   background=self.SURFACE, foreground=self.TEXT_DISABLED, font=(self.FONT_FAMILY, 9))
        style.configure("Head.TLabel",         background=self.BG,      foreground=self.ACCENT,        font=(self.FONT_FAMILY, 13, "bold"))
        style.configure("Sec.TLabel",          background=self.SURFACE, foreground=self.ACCENT,        font=(self.FONT_FAMILY, 10, "bold"))
        style.configure("SecDis.TLabel",       background=self.SURFACE, foreground=self.TEXT_DISABLED, font=(self.FONT_FAMILY, 10, "bold"))

        style.configure("TRadiobutton",
            background=self.SURFACE, foreground=self.TEXT,
            font=(self.FONT_FAMILY, 10), focuscolor=self.SURFACE,
        )
        style.map("TRadiobutton",
            background=[("disabled", self.SURFACE), ("active", self.SURFACE2)],
            foreground=[("disabled", self.TEXT_DISABLED), ("active", self.TEXT)],
        )

        style.configure("TCheckbutton",
            background=self.SURFACE, foreground=self.TEXT,
            font=(self.FONT_FAMILY, 10), focuscolor=self.SURFACE,
        )
        style.map("TCheckbutton",
            background=[("disabled", self.SURFACE), ("active", self.SURFACE2)],
            foreground=[("disabled", self.TEXT_DISABLED), ("active", self.TEXT)],
        )

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
            background=[("disabled", self.SURFACE), ("active", self.SURFACE2)],
            foreground=[("disabled", self.TEXT_DISABLED), ("active", self.ACCENT)],
        )

        style.configure("TEntry",
            fieldbackground=self.SURFACE2, foreground=self.TEXT,
            insertcolor=self.TEXT, borderwidth=1, relief="solid",
            font=(self.FONT_FAMILY, 10),
        )
        style.map("TEntry",
            fieldbackground=[("disabled", "#e8e8e8")],
            foreground=[("disabled", self.TEXT_DISABLED)],
        )

    # -----------------------------------------------------------------------
    def _card(self, parent, title=""):
        """Vrati (ramecek karty, popisek nadpisu nebo None)."""
        outer = ttk.Frame(parent, style="Card.TFrame", padding=12)
        title_lbl = None
        if title:
            title_lbl = ttk.Label(outer, text=title, style="Sec.TLabel")
            title_lbl.pack(anchor="w", pady=(0, 8))
        return outer, title_lbl

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
        c1, self._mode_title = self._card(self.scroll_frame, "Režim kopírování")
        c1.pack(**pad)
        self.mode_var = tk.StringVar(value=CopyMode.ALL_FOLDERS.value)
        self._mode_radios = []
        for label, val in [
            ("Celé složky",  CopyMode.ALL_FOLDERS.value),
            ("První soubor", CopyMode.FIRST_FILE.value),
            ("Podle Excelu", CopyMode.BY_EXCEL.value),
        ]:
            rb = ttk.Radiobutton(c1, text=label, variable=self.mode_var, value=val,
                                 command=self._update_ui_state)
            rb.pack(anchor="w")
            self._mode_radios.append(rb)

        # ---- Sekce: Mnozstvi produktu ----
        c2, self._multi_title = self._card(self.scroll_frame, "Množství produktů na fotce")
        c2.pack(**pad)
        self.multi_var = tk.StringVar(value=MultiChoice.ANY.value)
        self._multi_radios = []
        for label, val in [
            ("Jeden produkt", MultiChoice.SINGLE.value),
            ("Více produktů", MultiChoice.MULTIPLE.value),
            ("Nezáleží",      MultiChoice.ANY.value),
        ]:
            rb = ttk.Radiobutton(c2, text=label, variable=self.multi_var, value=val)
            rb.pack(anchor="w")
            self._multi_radios.append(rb)
        self._multi_info_label = ttk.Label(c2,
            text="ℹ  Produkty jsou vždy ověřeny vůči Excelu bez ohledu na volbu.",
            style="CardDim.TLabel",
        )
        self._multi_info_label.pack(anchor="w", pady=(6, 0))

        # ---- Sekce: Filtr stitku ----
        c3, self._tag_title = self._card(self.scroll_frame, "Filtr štítků")
        c3.pack(**pad)
        self.tag_enabled = tk.BooleanVar(value=False)
        self._tag_check = ttk.Checkbutton(c3, text="Filtrovat podle štítků", variable=self.tag_enabled,
                                          command=self._toggle_tag)
        self._tag_check.pack(anchor="w")

        # Widgety skryte ve vychozim stavu
        self._tag_widgets_frame = ttk.Frame(c3, style="Card.TFrame")
        self._tag_inner_labels = []

        lbl_hint = ttk.Label(self._tag_widgets_frame, text="Více štítků oddělte čárkou např: pánské, černá",
                             style="CardDim.TLabel")
        lbl_hint.pack(anchor="w", pady=(6, 2))
        self._tag_inner_labels.append(lbl_hint)

        self.tag_entry = ttk.Entry(self._tag_widgets_frame, width=50)
        self.tag_entry.pack(anchor="w", ipady=4)

        lbl_logic = ttk.Label(self._tag_widgets_frame, text="Logika filtrování:", style="CardDim.TLabel")
        lbl_logic.pack(anchor="w", pady=(8, 2))
        self._tag_inner_labels.append(lbl_logic)

        self.tag_mode_var = tk.StringVar(value="or")
        tag_mode_row = ttk.Frame(self._tag_widgets_frame, style="Card.TFrame")
        tag_mode_row.pack(anchor="w")
        self._tag_mode_radios = []
        rb_or = ttk.Radiobutton(tag_mode_row, text="Alespoň jeden štítek",
                                variable=self.tag_mode_var, value="or")
        rb_or.pack(side="left", padx=(0, 12))
        self._tag_mode_radios.append(rb_or)
        rb_and = ttk.Radiobutton(tag_mode_row, text="Všechny štítky",
                                 variable=self.tag_mode_var, value="and")
        rb_and.pack(side="left")
        self._tag_mode_radios.append(rb_and)

        # ---- Sekce: Zpusob trideni ----
        c4, self._sort_title = self._card(self.scroll_frame, "Způsob třídění výstupu")
        c4.pack(**pad)
        self.sort_var = tk.StringVar(value=SortMode.STRUCTURE.value)
        self._sort_radios = []
        for label, val in [
            ("Podle struktury",     SortMode.STRUCTURE.value),
            ("Plochá struktura",    SortMode.FLAT.value),
            ("Vše do jedné složky", SortMode.ROOT.value),
        ]:
            rb = ttk.Radiobutton(c4, text=label, variable=self.sort_var, value=val)
            rb.pack(anchor="w")
            self._sort_radios.append(rb)

        # ---- Sekce: Zdrojova slozka ----
        c5, self._src_title = self._card(self.scroll_frame, "Zdrojová složka")
        c5.pack(**pad)
        self.src_var = tk.StringVar(value="4")
        self._src_radios = []
        for label, val in [
            ("Promo fotky",      "1"),
            ("Promo videa",      "2"),
            ("Vlastní cesta…",   "3"),
            ("Produktové fotky", "4"),
        ]:
            rb = ttk.Radiobutton(c5, text=label, variable=self.src_var, value=val,
                                 command=self._update_ui_state)
            rb.pack(anchor="w")
            self._src_radios.append(rb)

        custom_row = ttk.Frame(c5, style="Card.TFrame")
        custom_row.pack(fill="x", pady=(4, 0))
        self.custom_path_entry = ttk.Entry(custom_row, width=50)
        self.custom_path_entry.pack(side="left", ipady=4, padx=(0, 6))
        self._browse_src_btn = ttk.Button(custom_row, text="Procházet…", style="Outline.TButton",
                                          command=self._browse_source)
        self._browse_src_btn.pack(side="left")

        # ---- Excel status ----
        excel_frame = ttk.Frame(self.scroll_frame)
        excel_frame.pack(padx=20, pady=(2, 6), fill="x")
        self.excel_label = ttk.Label(excel_frame, text="🔍 Hledám Excel soubor…", style="Dim.TLabel")
        self.excel_label.pack(side="left")
        self._browse_excel_btn = ttk.Button(excel_frame, text="Vybrat Excel…", style="Outline.TButton",
                                            command=self._browse_excel)
        self._browse_excel_btn.pack(side="right")

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
    # Stavovy automat UI a pomocne metody
    # -----------------------------------------------------------------------
    @staticmethod
    def _mode_supports_multi(mode: CopyMode) -> bool:
        return mode == CopyMode.BY_EXCEL

    @staticmethod
    def _mode_supports_tags(mode: CopyMode) -> bool:
        return mode in (CopyMode.FIRST_FILE, CopyMode.BY_EXCEL)

    def _update_ui_state(self):
        """Synchronizuje povoleni/zakazani vsech ovladacich prvku podle aktualniho stavu."""
        running = self._is_running
        mode    = CopyMode(self.mode_var.get())

        general_state = "disabled" if running else "normal"
        multi_enabled = (not running) and self._mode_supports_multi(mode)
        tag_enabled   = (not running) and self._mode_supports_tags(mode)
        custom_active = (not running) and (self.src_var.get() == "3")

        # Rezim kopirovani, trideni, zdrojova slozka, tlacitka
        for rb in self._mode_radios:
            rb.configure(state=general_state)
        self._mode_title.configure(style="SecDis.TLabel" if running else "Sec.TLabel")

        for rb in self._sort_radios:
            rb.configure(state=general_state)
        self._sort_title.configure(style="SecDis.TLabel" if running else "Sec.TLabel")

        for rb in self._src_radios:
            rb.configure(state=general_state)
        self._src_title.configure(style="SecDis.TLabel" if running else "Sec.TLabel")
        self.custom_path_entry.configure(state="normal" if custom_active else "disabled")
        self._browse_src_btn.configure(state=general_state)
        self._browse_excel_btn.configure(state=general_state)
        self.run_btn.configure(state=general_state)

        # Sekce: Mnozstvi produktu na fotce (jen pro Podle Excelu)
        multi_state = "normal" if multi_enabled else "disabled"
        self._multi_title.configure(style="Sec.TLabel" if multi_enabled else "SecDis.TLabel")
        self._multi_info_label.configure(style="CardDim.TLabel" if multi_enabled else "CardDimDis.TLabel")
        for rb in self._multi_radios:
            rb.configure(state=multi_state)

        # Sekce: Filtr stitku (pro Prvni soubor a Podle Excelu; hodnoty i rozbaleni zustavaji zachovany)
        tag_state = "normal" if tag_enabled else "disabled"
        self._tag_title.configure(style="Sec.TLabel" if tag_enabled else "SecDis.TLabel")
        self._tag_check.configure(state=tag_state)
        self.tag_entry.configure(state=tag_state)
        for lbl in self._tag_inner_labels:
            lbl.configure(style="CardDim.TLabel" if tag_enabled else "CardDimDis.TLabel")
        for rb in self._tag_mode_radios:
            rb.configure(state=tag_state)

    def _toggle_tag(self):
        if self.tag_enabled.get():
            self._tag_widgets_frame.pack(anchor="w", fill="x")
        else:
            self._tag_widgets_frame.pack_forget()
        self._update_ui_state()

    def _browse_source(self):
        path = filedialog.askdirectory(title="Vyberte zdrojovou složku")
        if path:
            self.src_var.set("3")
            self._update_ui_state()
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
        script_dir = get_app_dir()
        for fn in sorted(os.listdir(script_dir)):
            if fn.startswith("~$"):
                continue
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
        script_dir = get_app_dir()
        log_file   = os.path.join(script_dir, "vypis_konzole.txt")
        with open(log_file, "a", encoding="utf-8") as f:
            f.write(line)

    # -----------------------------------------------------------------------
    def _on_run(self):
        """Nacte konfiguraci z GUI na hlavnim vlakne a spusti kopirovani na pozadi."""
        mode      = CopyMode(self.mode_var.get())
        sort_mode = SortMode(self.sort_var.get())

        multi_active = self._mode_supports_multi(mode)
        tag_active   = self._mode_supports_tags(mode)
        tag_enabled  = bool(self.tag_enabled.get()) if tag_active else False

        cfg = {
            "excel_path":   self.excel_path_var.get(),
            "src_choice":   self.src_var.get(),
            "custom_path":  self.custom_path_entry.get().strip(),
            "mode":         mode,
            "sort_mode":    sort_mode,
            "multi_choice": MultiChoice(self.multi_var.get()) if multi_active else MultiChoice.ANY,
            "tag_enabled":  tag_enabled,
            "raw_tags":     self.tag_entry.get().strip() if tag_enabled else "",
            "tag_mode":     self.tag_mode_var.get(),
        }

        self._is_running = True
        self._update_ui_state()

        self.log_box.configure(state="normal")
        self.log_box.delete("1.0", "end")
        self.log_box.configure(state="disabled")

        # Priprav log soubor
        script_dir = get_app_dir()
        log_file   = os.path.join(script_dir, "vypis_konzole.txt")
        if os.path.exists(log_file):
            os.remove(log_file)
        open(log_file, "w", encoding="utf-8").close()

        thread = threading.Thread(target=self._run_copy, args=(cfg,), daemon=True)
        thread.start()

    def _on_run_finished(self):
        self._is_running = False
        self._update_ui_state()

    # -----------------------------------------------------------------------
    def _run_copy(self, cfg: dict):
        log = self._ui_log
        script_dir = get_app_dir()

        try:
            # --- Validace vstupu ---
            excel_path = cfg["excel_path"]
            if not excel_path or not os.path.isfile(excel_path):
                log("❌ Excel soubor nebyl nalezen nebo zvolen.")
                return

            src_choice = cfg["src_choice"]
            if src_choice == "1":
                source_path = SOURCE_PATH_PROMO_FOTO
            elif src_choice == "2":
                source_path = SOURCE_PATH_PROMO_VIDEA
            elif src_choice == "4":
                source_path = SOURCE_PATH_ORIGINAL
            else:
                source_path = cfg["custom_path"]

            if not source_path or not os.path.isdir(source_path):
                log(f"❌ Zdrojová složka neexistuje: '{source_path}'")
                return

            mode         = cfg["mode"]
            sort_mode    = cfg["sort_mode"]
            multi_choice = cfg["multi_choice"]

            # --- Filtr stitku (pouze pokud je pro zvoleny rezim povolen a zapnut) ---
            tags     = []
            tag_mode = cfg["tag_mode"]
            if cfg["tag_enabled"]:
                raw_tags = cfg["raw_tags"]
                if not raw_tags:
                    log("❌ Filtr štítků je zapnut, ale žádný štítek nebyl zadán.")
                    return
                tags = [t.strip().lower() for t in raw_tags.split(",") if t.strip()]
                if not tags:
                    log("❌ Filtr štítků je zapnut, ale žádný platný štítek nebyl zadán.")
                    return
                log(f"Filtr štítků ({tag_mode.upper()}): {', '.join(tags)}")

            flat_structure    = (sort_mode == SortMode.FLAT)
            root_mode         = (sort_mode == SortMode.ROOT)
            require_structure = (sort_mode == SortMode.STRUCTURE)

            dest_path = os.path.join(script_dir, "foto_folders")
            uf_path   = os.path.join(script_dir, "unfound_folders.txt")

            # --- Smazani stare vystupni slozky ---
            if os.path.exists(dest_path):
                try:
                    shutil.rmtree(dest_path)
                    log(f"Odstraněna stará složka: {dest_path}")
                except Exception as e:
                    log(f"❌ Nepodařilo se odstranit složku '{dest_path}': {e}")
                    return
            os.makedirs(dest_path, exist_ok=True)

            # --- Nacteni Excelu ---
            try:
                mapping = get_mapping_from_excel(excel_path, require_structure=require_structure)
                log(f"Načten Excel: {os.path.basename(excel_path)}  ({len(mapping)} produktů)")
            except Exception as e:
                log(f"❌ Chyba při načítání Excelu: {e}")
                return

            log(f"Zdrojová složka: {source_path}")
            log(f"Výstupní složka: {dest_path}")

            if mode == CopyMode.BY_EXCEL:
                if os.path.exists(uf_path):
                    os.remove(uf_path)
                copy_photos_by_excel(
                    source_path, dest_path, mapping,
                    flat_structure=flat_structure,
                    root_mode=root_mode,
                    log=log,
                    tags=tags, tag_mode=tag_mode,
                    multi_choice=multi_choice,
                )
                log("✅ Kopírování dokončeno.")
            else:
                copy_mode_str = "all" if mode == CopyMode.ALL_FOLDERS else "first"
                unfound, copied_count = copy_folders_with_mapping(
                    source_path, dest_path, mapping, copy_mode_str,
                    flat_structure=flat_structure,
                    root_mode=root_mode,
                    log=log,
                    tags=tags, tag_mode=tag_mode,
                )
                if unfound:
                    with open(uf_path, "w", encoding="utf-8") as f:
                        for k in sorted(unfound):
                            f.write(k + "\n")
                    log(f"⚠  {len(unfound)} složek nebylo nalezeno. Seznam: unfound_folders.txt")
                elif os.path.exists(uf_path):
                    os.remove(uf_path)
                log(f"✅ Kopírování dokončeno. Celkem zkopírováno {copied_count} souborů.")
        except Exception as e:
            log(f"❌ Neočekávaná chyba: {e}")
        finally:
            self.after(0, self._on_run_finished)


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
