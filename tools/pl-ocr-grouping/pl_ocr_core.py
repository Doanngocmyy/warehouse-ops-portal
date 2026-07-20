#!/usr/bin/env python3
"""
Topologie Packing List Extractor — v7
Output: 2 sheets — All_Items (merged carton cells) + Match_Status
Updates: No/Product Name split, robust HS Code lookup, revised Origin logic, Origin/HS Code repeated per item, no blue fill on content cells
"""
from __future__ import annotations
import re, sys, logging, unicodedata
from collections import defaultdict
from dataclasses import dataclass, field
from pathlib import Path
from typing import Dict, List, Optional, Tuple
import pandas as pd
import pdfplumber
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.styles.borders import Border, Side
from openpyxl.utils import get_column_letter

logging.basicConfig(level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[logging.StreamHandler(sys.stdout)])
log = logging.getLogger(__name__)

# =========================================================
# 1) CONFIG — filled in by app.html from browser uploads (Pyodide virtual FS).
#    Only this block differs from the original notebook; everything below is
#    byte-for-byte identical to the tested OCR_Packing List.ipynb / v7 logic.
# =========================================================
PL_FOLDER = Path("/work/pdfs")
OUTPUT_XLSX = Path("/work/PL_Total.xlsx")
DIM_WEIGHT_FILE = Path("/work/dim.xlsx")
DIM_WEIGHT_SHEET = __DIM_WEIGHT_SHEET__
MASTER_DATA_FILE = Path("/work/master.xlsx")
MASTER_DATA_SHEET = __MASTER_DATA_SHEET__
RECURSIVE = __RECURSIVE__


# ── Constants ──────────────────────────────────────────────────────────────
VALID_UNITS  = {"PCS", "SET", "CARTON", "CTN", "BOX", "PACK"}
UNIT_PAT     = "|".join(VALID_UNITS)
STATUS_WORDS = {"MOI", "NEW", "USED", "CU"}
VN_KEYWORDS  = {"POP", "JION", "QIFENG", "SBGEAR", "SB_GEAR"}
TABLE_HDR_KW = {"stt","barcode","ma vach","ma hang","ten hang",
                "don vi","so luong","tinh trang","condition","quantity"}

RE_PKG_HEADER = re.compile(
    r'(?:M[aã]\s*ki[eệ]n\s*h[aà]ng\s*[:\-]?\s*)?'
    r'(PGKEC[A-Z0-9]{5,})'
    r'(?:\s+(\d+\s*/\s*\d+))?',
    re.IGNORECASE | re.UNICODE)

RE_TOTAL = re.compile(
    r'T[oôồốổỗộòóỏõọ]ng\s+c[oôồốổỗộòóỏõọ]ng\s*:?\s*([\d,\.]+)',
    re.IGNORECASE | re.UNICODE)

RE_BARCODE   = re.compile(r'(?<!\d)(\d{8,14})(?!\d)')
RE_PROD_CODE = re.compile(r'(TP-[A-Z0-9]{2,}(?:-[A-Z0-9]+)*-?)', re.IGNORECASE)
RE_TERMINAL  = re.compile(rf'({UNIT_PAT})\s+[^\d\n]+?\s+([\d,]+)\s*$',
                          re.IGNORECASE | re.UNICODE)
RE_NOISE     = re.compile(
    r'^(STT|No\.\s*$|PACKING\s*LIST|DANH\s*S[AÁ]CH|Page\s*\d)',
    re.IGNORECASE | re.UNICODE)

# ── Helpers ────────────────────────────────────────────────────────────────
def strip_accents(s: str) -> str:
    nfd = unicodedata.normalize('NFD', str(s).strip())
    return ''.join(c for c in nfd if unicodedata.category(c) != 'Mn').upper()

def normalize(s: str) -> str:
    s = strip_accents(s)
    s = re.sub(r'[^\x20-\x7E]', '', s)
    s = re.sub(r'\s*[\(\[]\d+[\)\]]', '', s)
    return s.strip()

def parse_qty(s: str) -> int:
    return int(re.sub(r'[,\.]', '', s.strip()))

def get_origin(source_file: str, reference_code: str) -> str:
    """
    Origin logic:
    - If reference/source explicitly contains TOPOLOGIE or ends with _CN / -CN / space CN => CN.
      Example: CN-2659_SH_CN => CN.
    - Otherwise JION, POP, SBGEAR, QIFENG, VN => VN.
    - Default remains CN for safety.
    """
    raw = f"{source_file} {reference_code}"
    text = strip_accents(raw)
    ref = strip_accents(reference_code)

    # Strong CN markers
    if "TOPOLOGIE" in text:
        return "CN"
    if re.search(r'(?:^|[_\-\s])CN(?:$|[_\-\s.])', ref):
        # The final business code is CN, e.g. CN-2659_SH_CN
        if ref.endswith("_CN") or ref.endswith("-CN") or ref.endswith(" CN") or ref == "CN":
            return "CN"

    # VN markers
    for kw in {"JION", "POP", "SBGEAR", "SB_GEAR", "QI FENG", "QIFENG", "VN"}:
        if kw in text:
            return "VN"

    return "CN"

def join_split_product_code(text: str) -> str:
    return re.sub(r'(TP-[A-Z0-9-]+?-)\s+([A-Z0-9])', r'\1\2',
                  text, flags=re.IGNORECASE)

def is_table_hdr(cells: List[str]) -> bool:
    joined = strip_accents(" ".join(cells)).lower()
    return sum(1 for kw in TABLE_HDR_KW if kw in joined) >= 2

def is_noise(line: str) -> bool:
    return bool(RE_NOISE.match(line.strip()))

def safe_float(v) -> Optional[float]:
    try:
        f = float(v)
        return None if pd.isna(f) else f
    except Exception:
        return None

def normalize_sku_key(text: str) -> str:
    """Robust key for SKU/EAN lookup: remove hidden chars, spaces and punctuation."""
    text = clean_excel_key(text) if 'clean_excel_key' in globals() else str(text or '').strip()
    text = unicodedata.normalize("NFKC", text)
    text = text.replace("\ufffe", "-").replace("￾", "-").replace("­", "-")
    text = re.sub(r"-\s+", "-", text)
    text = re.sub(r"\s+-", "-", text)
    return re.sub(r"[^A-Z0-9]", "", text.upper())

def split_leading_no(product_name: str, current_no: str = "") -> Tuple[str, str]:
    """Split cases like '1 10mm Rope Loop' into No='1', Product Name='10mm Rope Loop'."""
    name = re.sub(r"\s+", " ", str(product_name or "")).strip()
    no = str(current_no or "").strip()
    m = re.match(r"^(\d{1,4})\s+(.+)$", name)
    if m and not no:
        no = m.group(1)
        name = m.group(2).strip()
    return no, name

# ── Data models ────────────────────────────────────────────────────────────
@dataclass
class Item:
    no: str
    product_name: str
    product_code: str
    barcode: str
    unit: str
    quantity: int
    hs_code: str = ""
    parse_method: str = "table"

@dataclass
class Package:
    package_code: str
    source_file: str
    reference_code: str
    pdf_package_seq: str
    items: List[Item] = field(default_factory=list)
    declared_total_qty: Optional[int] = None
    header_count: int = 1
    global_carton_num: str = ""
    length:  Optional[float] = None
    width:   Optional[float] = None
    height:  Optional[float] = None
    weight:  Optional[float] = None
    cbm:     Optional[float] = None
    dim_matched: bool = False
    hs_code: str = ""

    @property
    def calc_qty(self) -> int:
        return sum(i.quantity for i in self.items)
    @property
    def item_count(self) -> int:
        return len(self.items)
    @property
    def origin(self) -> str:
        return get_origin(self.source_file, self.reference_code)

# ── Item parsers ───────────────────────────────────────────────────────────
def parse_item_cells(cells: List[str]) -> Optional[Item]:
    line_no = ""
    barcode = prod_code = unit = ""
    quantity = 0
    name_parts: List[str] = []
    unit_idx = -1
    barcode_idx = -1
    prod_idx = -1

    merged: List[str] = []
    i = 0
    while i < len(cells):
        cell = str(cells[i]).strip()
        if cell.endswith('-') and RE_PROD_CODE.fullmatch(cell) and i + 1 < len(cells):
            merged.append(cell + str(cells[i + 1]).strip())
            i += 2
        else:
            merged.append(cell)
            i += 1

    for idx, raw in enumerate(merged):
        cell = str(raw).strip()
        if not cell:
            continue

        # STT / No column usually appears before barcode or SKU.
        if not line_no and re.fullmatch(r'\d{1,4}', cell) and not barcode and not prod_code and idx <= 2:
            line_no = cell
            continue

        if RE_BARCODE.fullmatch(cell) and len(cell) in (8, 12, 13, 14):
            barcode = cell
            barcode_idx = idx
            continue

        candidate = join_split_product_code(cell)
        m = RE_PROD_CODE.search(candidate)
        if m:
            pc = m.group(1).rstrip('-')
            if pc.count('-') >= 2:
                prod_code = pc.upper()
                prod_idx = idx
                rest = RE_PROD_CODE.sub('', candidate).strip()
                if rest and strip_accents(rest) not in STATUS_WORDS:
                    no_from_name, clean_name = split_leading_no(rest, line_no)
                    line_no = no_from_name
                    if clean_name:
                        name_parts.append(clean_name)
                continue

        if cell.upper() in VALID_UNITS:
            unit = cell.upper()
            unit_idx = idx
            continue

        if re.fullmatch(r'\d{1,5}', cell) and idx > unit_idx >= 0:
            quantity = int(cell)
            continue

        if RE_PKG_HEADER.search(cell):
            continue
        if strip_accents(cell) in STATUS_WORDS:
            continue

        # Avoid putting STT in product name if it appears as a separate numeric cell.
        if re.fullmatch(r'\d{1,4}', cell) and not name_parts and (idx < barcode_idx or idx < prod_idx or idx <= 2):
            if not line_no:
                line_no = cell
            continue

        name_parts.append(cell)

    if not (barcode or prod_code) or quantity == 0:
        return None

    product_name = re.sub(r'\s+', ' ', ' '.join(name_parts)).strip()
    line_no, product_name = split_leading_no(product_name, line_no)
    return Item(no=line_no, product_name=product_name, product_code=prod_code,
                barcode=barcode, unit=unit or "PCS", quantity=quantity,
                parse_method="table")

def parse_item_text(accumulated: str) -> Optional[Item]:
    text = join_split_product_code(accumulated)
    m_term = RE_TERMINAL.search(text)
    if not m_term:
        return None
    unit     = m_term.group(1).upper()
    quantity = parse_qty(m_term.group(2))

    line_no = ""
    m_no = re.match(r"^\s*(\d{1,4})\s+", text)
    if m_no:
        line_no = m_no.group(1)

    barcode = ""
    for m in RE_BARCODE.finditer(text):
        if len(m.group(1)) in (8, 12, 13, 14):
            barcode = m.group(1)
            break
    prod_code = ""
    pc_end = 0
    for m in RE_PROD_CODE.finditer(text):
        pc = m.group(1).rstrip('-')
        if pc.count('-') >= 2:
            prod_code = pc.upper()
            pc_end = m.end()
            break
    product_name = ""
    if prod_code and pc_end < m_term.start():
        region = text[pc_end : m_term.start()].strip()
        region = RE_BARCODE.sub('', region).strip()
        parts = [w for w in region.split() if strip_accents(w) not in STATUS_WORDS]
        product_name = re.sub(r'\s+', ' ', ' '.join(parts)).strip()
    if not (barcode or prod_code) or quantity == 0:
        return None
    line_no, product_name = split_leading_no(product_name, line_no)
    return Item(no=line_no, product_name=product_name, product_code=prod_code,
                barcode=barcode, unit=unit, quantity=quantity,
                parse_method="text")

# ── Parser state machine ───────────────────────────────────────────────────
class Parser:
    def __init__(self):
        self.packages: List[Package] = []
        self._cur: Optional[Package] = None
        self._buf: List[str] = []
        self._source_file    = ""
        self._reference_code = ""

    def set_file(self, pdf_path: Path):
        self._source_file    = pdf_path.name
        self._reference_code = pdf_path.stem

    def feed_table_row(self, cells: List[str]):
        joined = " ".join(cells)
        if not joined.strip() or is_table_hdr(cells):
            return
        m = RE_PKG_HEADER.search(joined)
        if m:
            self._on_pkg_header(m.group(1).upper(), (m.group(2) or "").replace(' ',''))
            return
        m = RE_TOTAL.search(joined)
        if m:
            self._on_total(parse_qty(m.group(1)))
            return
        if self._cur is None:
            return
        item = parse_item_cells(cells)
        if item:
            self._cur.items.append(item)

    def feed_text_line(self, line: str):
        line = line.strip()
        if not line or is_noise(line):
            return
        m = RE_PKG_HEADER.search(line)
        if m:
            self._flush_buf()
            self._on_pkg_header(m.group(1).upper(), (m.group(2) or "").replace(' ',''))
            return
        m = RE_TOTAL.search(line)
        if m:
            self._flush_buf()
            self._on_total(parse_qty(m.group(1)))
            return
        if self._cur is None:
            return
        if RE_BARCODE.search(line):
            if self._buf:
                item = parse_item_text(' '.join(self._buf))
                if item:
                    self._cur.items.append(item)
                self._buf = []
            self._buf.append(line)
            item = parse_item_text(' '.join(self._buf))
            if item:
                self._cur.items.append(item)
                self._buf = []
            return
        if self._buf:
            self._buf.append(line)
            item = parse_item_text(' '.join(self._buf))
            if item:
                self._cur.items.append(item)
                self._buf = []

    def end_of_pdf(self):
        self._flush_buf()

    def finalise(self):
        self._flush_buf()
        if self._cur is not None:
            log.warning(f"EOF: {self._cur.package_code} never saw Tong cong")
            self._force_close()

    def _on_pkg_header(self, pkg_code: str, seq: str):
        if self._cur is None:
            self._cur = Package(package_code=pkg_code,
                                source_file=self._source_file,
                                reference_code=self._reference_code,
                                pdf_package_seq=seq)
        elif self._cur.package_code == pkg_code:
            self._cur.header_count += 1
        else:
            log.warning(f"INTERRUPTED: {self._cur.package_code} -> {pkg_code}")
            self._force_close()
            self._cur = Package(package_code=pkg_code,
                                source_file=self._source_file,
                                reference_code=self._reference_code,
                                pdf_package_seq=seq)

    def _on_total(self, declared: int):
        if self._cur is None:
            log.warning(f"Orphan Tong cong={declared}")
            return
        self._cur.declared_total_qty = declared
        self.packages.append(self._cur)
        self._cur = None

    def _force_close(self):
        if self._cur is not None:
            self.packages.append(self._cur)
            self._cur = None

    def _flush_buf(self):
        if self._buf and self._cur is not None:
            item = parse_item_text(' '.join(self._buf))
            if item:
                self._cur.items.append(item)
        self._buf = []

# ── DIM mapper ─────────────────────────────────────────────────────────────
class DimMapper:
    _ALIASES: Dict[str, List[str]] = {
        "ref":    ["lo","lot","reference_code","reference","ref","job","shipment"],
        "pkg":    ["tracking","package_code","package","pkg","carton_code"],
        "length": ["dai","length","l","len"],
        "width":  ["rong","width","w","wid"],
        "height": ["cao","height","h","hei","high"],
        "weight": ["kg","weight","wt","gross","gw"],
        "cbm":    ["cbm","volume","vol","cubic","m3"],
    }

    def __init__(self, xlsx_path: Path, sheet_name: Optional[str] = None):
        self._data: Dict[str, dict] = {}
        self._load(xlsx_path, sheet_name)

    def _load(self, path: Path, sheet_name: Optional[str] = None):
        log.info(f"Loading DIM <- {path.name}")
        try:
            xl = pd.ExcelFile(str(path))
        except Exception as e:
            log.error(f"Cannot open DIM: {e}")
            return

        if sheet_name:
            if sheet_name in xl.sheet_names:
                sheets_to_try = [sheet_name]
            else:
                log.warning(f"DIM sheet '{sheet_name}' not found. Available sheets: {xl.sheet_names}. Auto-detect instead.")
                sheets_to_try = xl.sheet_names
        else:
            sheets_to_try = xl.sheet_names

        for sheet in sheets_to_try:
            try:
                df = pd.read_excel(path, sheet_name=sheet, dtype=str)
                cols = [str(c).strip() for c in df.columns]
                cm = self._detect(cols)
                if cm is None:
                    continue
                loaded = 0
                for _, row in df.iterrows():
                    ref = normalize(str(row.get(cm["ref"], "")))
                    pkg = normalize(str(row.get(cm["pkg"], "")))
                    if not ref or not pkg or ref == "NAN" or pkg == "NAN":
                        continue
                    self._data[f"{ref}|{pkg}"] = {
                        "length": safe_float(row.get(cm.get("length"))),
                        "width":  safe_float(row.get(cm.get("width"))),
                        "height": safe_float(row.get(cm.get("height"))),
                        "weight": safe_float(row.get(cm.get("weight"))),
                        "cbm":    safe_float(row.get(cm.get("cbm"))),
                    }
                    loaded += 1
                log.info(f"  Sheet '{sheet}': {loaded} rows")
                if loaded:
                    break
            except Exception as e:
                log.warning(f"  Sheet '{sheet}' error: {e}")
        log.info(f"  Total DIM records: {len(self._data)}")

    def _detect(self, cols: List[str]) -> Optional[Dict[str, str]]:
        def norm(s):
            return re.sub(r'[^a-z0-9]', '', strip_accents(s).lower())
        nc = {norm(c): c for c in cols}
        mapping: Dict[str, str] = {}
        for field_name, aliases in self._ALIASES.items():
            for alias in aliases:
                if norm(alias) in nc:
                    mapping[field_name] = nc[norm(alias)]
                    break
        if "ref" not in mapping or "pkg" not in mapping:
            return None
        return mapping

    def lookup(self, ref: str, pkg: str) -> Optional[dict]:
        key = f"{normalize(ref)}|{normalize(pkg)}"
        return self._data.get(key)

# ── HS Code mapper ─────────────────────────────────────────────────────────
def clean_excel_key(text: str) -> str:
    """Clean key for exact Excel-like SKU matching while preserving hyphens."""
    if text is None:
        return ""
    try:
        if pd.isna(text):
            return ""
    except Exception:
        pass
    return re.sub(r"\s+", " ", str(text)).strip()

def norm_col_name(text: str) -> str:
    text = clean_excel_key(text)
    text = unicodedata.normalize("NFD", text)
    text = "".join(ch for ch in text if unicodedata.category(ch) != "Mn")
    text = text.upper()
    return re.sub(r"[^A-Z0-9]+", "", text)

class HsCodeMapper:
    def __init__(self, master_path: Optional[Path] = None, sheet_name: Optional[str] = None):
        self._data_exact: Dict[str, str] = {}
        self._data_norm: Dict[str, str] = {}
        self._data_barcode: Dict[str, str] = {}
        if master_path:
            self._load(master_path, sheet_name)

    def _load(self, path: Path, sheet_name: Optional[str] = None):
        if not path or not path.exists():
            log.warning(f"Master data file not found: {path}")
            return
        log.info(f"Loading HS Code <- {path.name}")
        try:
            xl = pd.ExcelFile(str(path))
        except Exception as e:
            log.error(f"Cannot open Master Data: {e}")
            return

        sheets_to_try = [sheet_name] if sheet_name and sheet_name in xl.sheet_names else xl.sheet_names
        if sheet_name and sheet_name not in xl.sheet_names:
            log.warning(f"Master data sheet '{sheet_name}' not found. Auto-detect instead.")

        for sheet in sheets_to_try:
            try:
                raw = pd.read_excel(path, sheet_name=sheet, header=None, dtype=str)
            except Exception as e:
                log.warning(f"  Sheet '{sheet}' error: {e}")
                continue
            if raw.empty:
                continue

            header_row_idx = None
            for i in range(min(40, len(raw))):
                row_join = " | ".join(clean_excel_key(x) for x in raw.iloc[i].tolist()).upper()
                row_norm = norm_col_name(row_join)
                if ("SKU" in row_norm or "PRODUCTCODE" in row_norm) and "HSCODE" in row_norm:
                    header_row_idx = i
                    break
            if header_row_idx is None:
                continue

            df = pd.read_excel(path, sheet_name=sheet, header=header_row_idx, dtype=str)
            df.columns = [clean_excel_key(c) for c in df.columns]

            sku_col = hs_col = None
            barcode_cols: List[str] = []
            for c in df.columns:
                n = norm_col_name(c)
                if n in {"SKUPRODUCTCODE", "SKU", "PRODUCTCODE"} or ("SKU" in n and "PRODUCT" in n and "CODE" in n):
                    sku_col = c
                if n == "HSCODE" or ("HS" in n and "CODE" in n):
                    hs_col = c
                if n in {"EAN", "BARCODE", "UPC"} or "EAN" in n or "BARCODE" in n:
                    barcode_cols.append(c)

            if not sku_col or not hs_col:
                log.warning(f"  Sheet '{sheet}': cannot find SKU / HS Code columns")
                continue

            loaded = 0
            for _, row in df.iterrows():
                sku = clean_excel_key(row.get(sku_col))
                hs = clean_excel_key(row.get(hs_col))
                if not sku or not hs:
                    continue
                self._data_exact[sku] = hs
                self._data_exact[sku.upper()] = hs
                sku_norm = normalize_sku_key(sku)
                if sku_norm:
                    self._data_norm[sku_norm] = hs
                for bc in barcode_cols:
                    barcode = re.sub(r"\D", "", clean_excel_key(row.get(bc)))
                    if barcode:
                        self._data_barcode[barcode] = hs
                loaded += 1
            log.info(f"  Sheet '{sheet}': loaded {loaded} HS Code records")

        log.info(
            f"  Total HS Code records: exact={len(self._data_exact)}, "
            f"normalized={len(self._data_norm)}, barcode={len(self._data_barcode)}"
        )

    def lookup(self, sku: str, barcode: str = "") -> str:
        sku_clean = clean_excel_key(sku)
        if sku_clean in self._data_exact:
            return self._data_exact[sku_clean]
        if sku_clean.upper() in self._data_exact:
            return self._data_exact[sku_clean.upper()]
        sku_norm = normalize_sku_key(sku_clean)
        if sku_norm in self._data_norm:
            return self._data_norm[sku_norm]
        barcode_clean = re.sub(r"\D", "", clean_excel_key(barcode))
        if barcode_clean in self._data_barcode:
            return self._data_barcode[barcode_clean]
        return ""

# ── Global carton numbers ──────────────────────────────────────────────────
def assign_global_numbers(packages: List[Package]):
    total = len(packages)
    for i, pkg in enumerate(packages, start=1):
        pkg.global_carton_num = f"{i}/{total}"
    log.info(f"Carton numbers: 1/{total} ... {total}/{total}")

# ── Status ─────────────────────────────────────────────────────────────────
def overall_status(pkg: Package) -> Tuple[str, str]:
    decl = pkg.declared_total_qty
    calc = pkg.calc_qty
    if decl is not None and decl > 0 and pkg.item_count == 0:
        return "CRITICAL_ZERO_ITEMS", f"declared={decl} but 0 items parsed"
    if decl is None:
        return "MISMATCH_NO_TOTAL", "Tong cong not found"
    if decl != calc:
        return "MISMATCH_QTY", f"declared={decl} calc={calc} diff={decl - calc}"
    if not pkg.dim_matched:
        return "MISMATCH_DIM", "No row in Final_dim weight"
    return "MATCHED", ""

# ── Excel styles ───────────────────────────────────────────────────────────
HDR_FILL   = PatternFill("solid", fgColor="1F4E79")
HDR_FONT   = Font(bold=True, color="FFFFFF", size=10)
ALT_FILL   = PatternFill()
PKG_FILL   = PatternFill()
PKG_FONT   = Font(bold=True, size=9)

STATUS_FILL = {
    "MATCHED":             PatternFill("solid", fgColor="C6EFCE"),
    "MISMATCH_QTY":        PatternFill("solid", fgColor="FFCCCC"),
    "MISMATCH_DIM":        PatternFill("solid", fgColor="FFF2CC"),
    "MISMATCH_NO_TOTAL":   PatternFill("solid", fgColor="FFD966"),
    "CRITICAL_ZERO_ITEMS": PatternFill("solid", fgColor="FF4444"),
}
CRIT_FONT = Font(bold=True, color="FFFFFF", size=10)

_THICK = Side(style="medium", color="1F4E79")
_THIN  = Side(style="thin",   color="B0BEC5")

# Carton-level columns (1-based): 7=Carton#  8=seq  9=pkg_code
# 10=L 11=W 12=H 13=Wt 14=CBM  17=ref  18=src
# 15=Origin and 16=HS Code are item-level, so they are NOT merged.
_MERGE_COLS = list(range(7, 15)) + [17, 18]

def _hdr(ws, headers):
    ws.append(headers)
    for cell in ws[1]:
        cell.fill = HDR_FILL
        cell.font = HDR_FONT
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws.row_dimensions[1].height = 28

def _auto_w(ws, cap=55):
    for col in ws.columns:
        letter = get_column_letter(col[0].column)
        w = max((len(str(c.value)) if c.value else 0) for c in col)
        ws.column_dimensions[letter].width = min(w + 4, cap)

def _apply_pkg_merge(ws, start_row: int, end_row: int, dim_matched: bool):
    """Merge carton-level columns across all item rows of one package."""
    if end_row <= start_row:
        # Single item — just style, no merge needed
        for col in _MERGE_COLS:
            cell = ws.cell(row=start_row, column=col)
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            if not dim_matched and col in range(10, 15):
                cell.fill = STATUS_FILL["MISMATCH_DIM"]
            else:
                cell.fill = PKG_FILL
            cell.font = PKG_FONT
        return
    for col in _MERGE_COLS:
        letter = get_column_letter(col)
        ws.merge_cells(f"{letter}{start_row}:{letter}{end_row}")
        cell = ws.cell(row=start_row, column=col)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        if not dim_matched and col in range(10, 15):
            cell.fill = STATUS_FILL["MISMATCH_DIM"]
        else:
            cell.fill = PKG_FILL
        cell.font = PKG_FONT

def _border_bottom(ws, row: int, ncols: int):
    """Thick bottom border = visual separator between packages."""
    for c in range(1, ncols + 1):
        cell = ws.cell(row=row, column=c)
        existing = cell.border
        cell.border = Border(
            left=existing.left, right=existing.right, top=existing.top,
            bottom=_THICK)

# ── Workbook writer ────────────────────────────────────────────────────────
def write_workbook(output_path: Path, packages: List[Package]):
    wb = Workbook()
    wb.remove(wb.active)
    NCOLS = 18

    # ── All_Items ──────────────────────────────────────────────────────────
    ws1 = wb.create_sheet("All_Items")
    _hdr(ws1, [
        "No", "Product Name", "product_code", "barcode", "unit", "quantity",
        "Carton number", "pdf_package_seq", "package_code",
        "Length", "Width", "Height", "Weight", "CBM",
        "Origin", "HS Code", "reference_code", "source_file",
    ])

    row_idx = 2
    for pkg in packages:
        origin    = pkg.origin
        pkg_start = row_idx

        if not pkg.items:
            ws1.append([
                "", "", "", "", "", "",
                pkg.global_carton_num, pkg.pdf_package_seq, pkg.package_code,
                pkg.length, pkg.width, pkg.height, pkg.weight, pkg.cbm,
                origin, pkg.hs_code, pkg.reference_code, pkg.source_file,
            ])
            row_idx += 1
        else:
            for i, item in enumerate(pkg.items):
                first = (i == 0)
                ws1.append([
                    item.no, item.product_name, item.product_code, item.barcode,
                    item.unit, item.quantity,
                    pkg.global_carton_num if first else "",
                    pkg.pdf_package_seq   if first else "",
                    pkg.package_code      if first else "",
                    pkg.length  if first else "",
                    pkg.width   if first else "",
                    pkg.height  if first else "",
                    pkg.weight  if first else "",
                    pkg.cbm     if first else "",
                    origin,
                    item.hs_code,
                    pkg.reference_code if first else "",
                    pkg.source_file    if first else "",
                ])
                # Keep content cells unfilled as requested.
                row_idx += 1

        end_row = row_idx - 1
        _apply_pkg_merge(ws1, pkg_start, end_row, pkg.dim_matched)
        _border_bottom(ws1, end_row, NCOLS)

    for r in range(2, row_idx):
        ws1.row_dimensions[r].height = 18
    _auto_w(ws1)
    ws1.freeze_panes = "A2"

    # ── Match_Status ───────────────────────────────────────────────────────
    ws2 = wb.create_sheet("Match_Status")
    _hdr(ws2, [
        "source_file", "reference_code", "package_code",
        "pdf_package_seq", "Carton number",
        "item_count", "calculated_total_qty", "declared_total_qty",
        "qty_match", "dim_match", "overall_status", "remark",
    ])
    for i, pkg in enumerate(packages, start=2):
        decl = pkg.declared_total_qty
        calc = pkg.calc_qty
        q_ok = "YES" if (decl is not None and decl == calc) else "NO"
        d_ok = "YES" if pkg.dim_matched else "NO"
        status, remark = overall_status(pkg)
        ws2.append([
            pkg.source_file, pkg.reference_code, pkg.package_code,
            pkg.pdf_package_seq, pkg.global_carton_num,
            pkg.item_count, calc, decl,
            q_ok, d_ok, status, remark,
        ])
        fill = STATUS_FILL.get(status, STATUS_FILL["MISMATCH_DIM"])
        fnt  = CRIT_FONT if status == "CRITICAL_ZERO_ITEMS" else None
        for c in range(1, 13):
            ws2.cell(row=i, column=c).fill = fill
            if fnt:
                ws2.cell(row=i, column=c).font = fnt
    _auto_w(ws2)
    ws2.freeze_panes = "A2"

    output_path = Path(output_path)
    tmp_path = output_path.with_suffix(output_path.suffix + ".tmp")
    try:
        wb.save(str(tmp_path))
    except PermissionError as e:
        raise PermissionError(
            f"Cannot create temp file for '{output_path.name}': {e}. "
            f"Close any program locking that folder and re-run."
        ) from e
    try:
        if output_path.exists():
            output_path.unlink()
        tmp_path.replace(output_path)
    except PermissionError as e:
        try:
            tmp_path.unlink(missing_ok=True)
        except Exception:
            pass
        raise PermissionError(
            f"Cannot overwrite '{output_path}': the file appears to be open in Excel. "
            f"Close it and re-run."
        ) from e
    log.info(f"Saved -> {output_path}")

# ── Pipeline ───────────────────────────────────────────────────────────────
TABLE_CFG = {
    "vertical_strategy": "lines", "horizontal_strategy": "lines",
    "snap_tolerance": 5, "join_tolerance": 3,
    "edge_min_length": 3, "min_words_vertical": 1,
}

def run_pipeline(pl_folder: Path, dim_xlsx: Path,
                 output_path: Optional[Path] = None,
                 dim_sheet: Optional[str] = None,
                 master_data_file: Optional[Path] = None,
                 master_data_sheet: Optional[str] = None,
                 recursive: bool = False):
    if output_path is None:
        output_path = pl_folder / "PL_Output_v6_HS_DIM.xlsx"

    dim = DimMapper(dim_xlsx, sheet_name=dim_sheet)
    hs_mapper = HsCodeMapper(master_data_file, sheet_name=master_data_sheet)
    pdf_iter = pl_folder.rglob("*.pdf") if recursive else pl_folder.glob("*.pdf")
    pdf_files = sorted(pdf_iter, key=lambda p: p.name.upper())
    if not pdf_files:
        log.error(f"No PDFs in {pl_folder}")
        return []
    log.info(f"PDFs ({len(pdf_files)}): {[f.name for f in pdf_files]}")

    parser = Parser()
    for pdf_path in pdf_files:
        log.info(f"Parsing  {pdf_path.name}")
        parser.set_file(pdf_path)
        with pdfplumber.open(str(pdf_path)) as pdf:
            for page in pdf.pages:
                tables = page.extract_tables(TABLE_CFG)
                used_table = False
                if tables:
                    for table in tables:
                        for row in table:
                            if row is None or all(c is None for c in row):
                                continue
                            cells = [str(c).strip() if c else "" for c in row]
                            if not any(cells):
                                continue
                            parser.feed_table_row(cells)
                    used_table = True
                if not used_table:
                    for line in (page.extract_text() or "").splitlines():
                        parser.feed_text_line(line)
        parser.end_of_pdf()

    parser.finalise()
    packages = parser.packages
    log.info(f"Total packages: {len(packages)}")

    assign_global_numbers(packages)

    matched = 0
    for pkg in packages:
        d = dim.lookup(pkg.reference_code, pkg.package_code)
        if d:
            pkg.length, pkg.width, pkg.height = d["length"], d["width"], d["height"]
            pkg.weight, pkg.cbm = d["weight"], d["cbm"]
            pkg.dim_matched = True
            matched += 1
    log.info(f"DIM matched: {matched}/{len(packages)}")

    hs_matched = 0
    hs_total = 0
    for pkg in packages:
        pkg_hs_codes = []
        for item in pkg.items:
            hs_total += 1
            item.hs_code = hs_mapper.lookup(item.product_code, item.barcode)
            if item.hs_code:
                hs_matched += 1
                pkg_hs_codes.append(item.hs_code)
        # For zero-item packages only, keep a carton-level fallback blank.
        pkg.hs_code = ", ".join(sorted(set(pkg_hs_codes))) if pkg_hs_codes else ""
    log.info(f"HS Code matched: {hs_matched}/{hs_total}")

    counts: Dict[str, int] = defaultdict(int)
    for pkg in packages:
        counts[overall_status(pkg)[0]] += 1
    for st, n in sorted(counts.items()):
        log.info(f"  {st:<30} {n}")

    write_workbook(output_path, packages)
    return packages

# ── Entry point ────────────────────────────────────────────────────────────
# ── Entry point (called directly instead of __main__ guard, since this
#    module is exec'd inside Pyodide rather than run as a script) ─────────
packages = run_pipeline(
    pl_folder=PL_FOLDER,
    dim_xlsx=DIM_WEIGHT_FILE,
    output_path=OUTPUT_XLSX,
    dim_sheet=DIM_WEIGHT_SHEET,
    master_data_file=MASTER_DATA_FILE,
    master_data_sheet=MASTER_DATA_SHEET,
    recursive=RECURSIVE,
)

# =========================================================
# AUTO SPLIT: TOTAL -> FACTORY -> CN PORT -> CN STORE
# Requires: pl_group_export.py in the same folder as this notebook.
# Run this cell AFTER the cell above has produced `packages` via run_pipeline().
# =========================================================
import importlib
import pl_group_export
importlib.reload(pl_group_export)  # pick up edits to pl_group_export.py without restarting the kernel
from pl_group_export import export_grouped_pl

if not packages:
    raise RuntimeError(
        "`packages` is empty \u2014 the OCR pipeline above found no PDFs or produced "
        "no packages. Fix PL_FOLDER / the PDFs first, then re-run the cell above "
        "before running this split step."
    )

SPLIT_OUTPUT_DIR = PL_FOLDER / 'PL_SPLIT_OUTPUT'

try:
    control_file = export_grouped_pl(
        packages=packages,
        output_dir=SPLIT_OUTPUT_DIR,
        write_workbook=write_workbook,
        total_workbook=OUTPUT_XLSX,
        pdf_folder=PL_FOLDER,
        recursive=RECURSIVE,
    )
except RuntimeError as e:
    print("XXXX SPLIT FAILED \u2014 reconciliation mismatch, nothing was silently swallowed XXXX")
    print(e)
    raise
except PermissionError as e:
    print("XXXX SPLIT FAILED \u2014 a target .xlsx is locked/open in Excel XXXX")
    print(e)
    raise
else:
    print(f'Completed: {SPLIT_OUTPUT_DIR}')
    print(f'Control file: {control_file}')
