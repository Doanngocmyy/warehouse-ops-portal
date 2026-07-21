#!/usr/bin/env python3
"""
Topologie Packing List Extractor — v8
Output: 2 sheets — Packing List (matches the official PL template layout,
plain black grid, no bold/no color fill) + Match_Status (internal QC log,
color-coded on purpose so mismatches stay easy to spot).
Updates: No/Product Name split, robust HS Code lookup, revised Origin logic,
Origin/HS Code repeated per item, Packing List sheet now mirrors the real
PL template columns 1:1 (Item#, PO No., Invoice No., Product Name in
English, SKU#, BarCode/UPC, UOM, Quantity, Carton#, Packaging code, Carton
Dimensions L/W/H, Weight, CBM, Origin Country, Origin Country's HTSCODE,
Shipping Mark, PORT, 中国标签名称) with a plain black-bordered grid
(no bold, no blue/red/green fills) instead of the old merged-diagnostic
style. PORT is auto-filled for CN-factory cartons using the same
store/port rule already used for the CN split (pl_group_export.py);
PO No. / Invoice No. / Shipping Mark / 中国标签名称 are intentionally left
blank — fill them in manually, same as before.
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
#    byte-for-byte identical to the tested OCR_Packing List.ipynb / v7 logic,
#    except the Package dataclass (2 new optional fields) and the Excel
#    writer section (Packing List sheet layout), which are the parts this
#    revision (v8) intentionally changes.
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
    text = text.replace("￾", "-").replace("￾", "-").replace("­", "-")
    text = re.sub(r"-\s+", "-", text)
    text = re.sub(r"\s+-", "-", text)
    return re.sub(r"[^A-Z0-9]", "", text.upper())

_INVISIBLE_CHARS_RE = re.compile('[ ​‌‍⁠﻿\xad]')

def sanitize_ocr_cell(s: str) -> str:
    """Strip invisible/zero-width unicode artifacts (NBSP, zero-width space,
    soft hyphen, BOM, ...) that PDF text extraction sometimes inserts inside
    otherwise-contiguous codes. Left uncleaned, these can silently break the
    barcode/SKU regexes or leave a stray character inside a code that should
    read as one unbroken token."""
    if s is None:
        return s
    return _INVISIBLE_CHARS_RE.sub('', str(s))

def dequarantine_code(value: str, label: str, context: str = "") -> str:
    """SKU/product codes and EAN/barcodes never legitimately contain blanks —
    if OCR produced one anyway (stray space from a misread character), strip
    it and log a warning so the anomaly is visible for a quick sanity check,
    rather than silently mismatching or silently 'fixing' with no trace."""
    if not value:
        return value
    cleaned = re.sub(r'\s+', '', value)
    if cleaned != value:
        log.warning(f"{label} contained unexpected blank(s), auto-fixed: "
                     f"{value!r} -> {cleaned!r}" + (f" ({context})" if context else ""))
    return cleaned

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
    # v8: CN store/port classification, filled in by classify_packages_for_port()
    # (same rule as pl_group_export.py's CN split — "quy luật chia port/store").
    # Stays "" for non-CN-factory packages; fill manually if needed.
    port: str = ""
    store: str = ""

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
        cell = sanitize_ocr_cell(str(cells[i]).strip())
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
    prod_code = dequarantine_code(prod_code, "SKU/product_code", "table row")
    barcode   = dequarantine_code(barcode, "EAN/barcode", "table row")
    return Item(no=line_no, product_name=product_name, product_code=prod_code,
                barcode=barcode, unit=unit or "PCS", quantity=quantity,
                parse_method="table")

def parse_item_text(accumulated: str) -> Optional[Item]:
    text = join_split_product_code(sanitize_ocr_cell(accumulated))
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
    prod_code = dequarantine_code(prod_code, "SKU/product_code", "text line")
    barcode   = dequarantine_code(barcode, "EAN/barcode", "text line")
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
        "ref":    ["lo","lot","lohang","reference_code","reference","ref","job","shipment"],
        "pkg":    ["tracking","package_code","package","pkg","carton_code","carton","kien","makien"],
        "length": ["dai","length","l","len","d","chieudai"],
        "width":  ["rong","width","w","wid","r","chieurong"],
        "height": ["cao","height","h","hei","high","c","chieucao"],
        "weight": ["kg","weight","wt","gross","gw","nang"],
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
        claimed: set = set()
        for field_name, aliases in self._ALIASES.items():
            found = None
            # 1) exact normalized match first (original, safest behaviour)
            for alias in aliases:
                a = norm(alias)
                if a in nc and nc[a] not in claimed:
                    found = nc[a]
                    break
            # 2) fallback: alias appears anywhere inside the column name — needed
            #    for real-world templates like "Lô Hàng (QRcode ở giữa)" or
            #    "Mã Kiện (Barcode ở trên hoặc dưới)" where headers are full
            #    descriptive phrases, not the bare alias token. Skip 1-char
            #    aliases here (l/w/h) since substring matching on a single
            #    letter is too prone to false positives.
            if found is None:
                for alias in aliases:
                    a = norm(alias)
                    if len(a) < 2:
                        continue
                    for key, orig_col in nc.items():
                        if orig_col in claimed:
                            continue
                        if a in key:
                            found = orig_col
                            break
                    if found:
                        break
            if found:
                mapping[field_name] = found
                claimed.add(found)
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

# ── v8: CN store/port classification (same rule as pl_group_export.py) ─────
def classify_packages_for_port(packages: List[Package], pdf_folder: Path, recursive: bool):
    """Fill pkg.port / pkg.store for every CN-factory package, using the exact
    same detect_factory() + match_store() rule already used to split CN
    shipments by store/port (pl_group_export.py) — "quy luật chia port và
    store" the warehouse team already uses. Non-CN-factory packages (POP,
    SBGEAR, QIFENG, JION, or unclassifiable) are left with port="" / store=""
    (blank) — fill in manually if needed, same as PO No. / Invoice No."""
    try:
        import pl_group_export as pge
    except ImportError:
        log.warning("pl_group_export not importable — PORT column will stay blank for all packages.")
        return
    cache: Dict[str, str] = {}
    n_cn = n_matched = 0
    for pkg in packages:
        factory = pge.detect_factory(pkg.reference_code, pkg.source_file)
        if factory != "CN":
            continue
        n_cn += 1
        signal = pge._collect_cn_signal(pkg, pdf_folder, recursive, cache)
        store, confidence, suggestion = pge.match_store(signal)
        if store in pge.STORE_MASTER:
            pkg.store = store
            pkg.port = str(pge.STORE_MASTER[store]["port"])
            n_matched += 1
        else:
            pkg.store = "REVIEW"
            pkg.port = ""
    log.info(f"CN store/port classification: {n_matched}/{n_cn} CN package(s) matched to a store+port.")

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
# v8: the "Packing List" sheet (the customer-facing deliverable) is now a
# plain black-grid table — no bold, no header fill, no red/green/yellow
# status colors — matching the real PL template exactly. The Match_Status
# sheet is an internal QC log only (never sent to the customer), so it keeps
# its color-coding on purpose: that's what makes mismatches jump out.
THIN_BLACK = Side(style="thin", color="000000")
PLAIN_BORDER = Border(left=THIN_BLACK, right=THIN_BLACK, top=THIN_BLACK, bottom=THIN_BLACK)
PLAIN_FONT = Font(bold=False, size=10)

STATUS_FILL = {
    "MATCHED":             PatternFill("solid", fgColor="C6EFCE"),
    "MISMATCH_QTY":        PatternFill("solid", fgColor="FFCCCC"),
    "MISMATCH_DIM":        PatternFill("solid", fgColor="FFF2CC"),
    "MISMATCH_NO_TOTAL":   PatternFill("solid", fgColor="FFD966"),
    "CRITICAL_ZERO_ITEMS": PatternFill("solid", fgColor="FF4444"),
}
CRIT_FONT = Font(bold=True, color="FFFFFF", size=10)
MS_HDR_FILL = PatternFill("solid", fgColor="1F4E79")
MS_HDR_FONT = Font(bold=True, color="FFFFFF", size=10)

# Packing List column layout (1-based), matching the real PL template. The
# item table itself starts at row 14 (rows 1-11 = document header block,
# rows 12-13 = the bilingual table header) — same row numbers as the real
# template, so this sheet lines up with it exactly:
#  A Item#            B PO No.           C Invoice No.
#  D Product Name in English             E SKU#            F BarCode/UPC
#  G UOM              H Quantity         I Carton#         J Packaging code
#  K/L/M Carton Dimensions (Length/Width/Height, cm)
#  N Weight (KG)      O CBM
#  P Origin Country   Q Origin Country's HTSCODE            R Shipping Mark
#  S PORT             T 中国标签名称
PL_HEADERS_EN = [
    "Item#", "PO No.", "Invoice No.", "Product Name\nin English", "SKU#",
    "BarCode/UPC", "UOM", "Quantity", "Carton#", "Packaging code",
    "Carton Dimensions (cm)\n(Length*Weight*Height)", "", "",
    "Weight (KG)", "CBM", "Origin Country", "Origin Country's HTSCODE",
    "Shipping Mark", "PORT", "中国标签名称",
]
PL_HEADERS_CN = [
    "项目", "PO 编码", "Invoice 编码", "货品名称", "SKU编码",
    "条形码", "单位", "数量", "箱号", "包装条形码",
    "箱子尺寸", "", "",
    "", "", "原产国", "原产国",
    "", "", "",
]
NCOLS = 20
TABLE_HDR_ROW1 = 12  # English header row (matches the real template exactly)
TABLE_HDR_ROW2 = 13  # Chinese header row
FIRST_ITEM_ROW = 14
# Package/carton-level columns — same physical carton, so merged across all
# item rows of that carton (matches the real template exactly): Carton#,
# Packaging code, L/W/H, Weight, CBM. Everything else (incl. Origin Country,
# HTS Code, Shipping Mark, PORT) is left un-merged / repeated per row, same
# as the template.
_MERGE_COLS = [9, 10, 11, 12, 13, 14, 15]  # I,J,K,L,M,N,O

# v8: document header block (rows 1-11) — SHIPPER / CONSIGNEE are the same
# entity on every CN shipment (confirmed against 2 real PL samples), so they
# are filled in automatically. WPIC Purchase Order#, Seller's EIN#, Date,
# Invoice#, Remark (SO#) and Trade term vary per shipment and are NOT in the
# OCR/DIM/Master data, so they stay blank — fill in manually, same as before.
SHIPPER_BLOCK = (
    "SHIPPER:\n"
    "TOPOLOGIE GLOBAL LIMITED\n"
    "RM G, 9/F, King Palace Plaza\n"
    "55 King Yip Street, Kwun Tong, Hong Kong\n"
    "EMAIL: supplychainhk@topologie.com\n"
    "TEL: 852 3955 9963"
)
CONSIGNEE_BLOCK = (
    "CONSIGNEE:\n"
    "WORKING UNIT SHANGHAI TRADING CO LTD\n"
    "Room 301, No. 47, Branch Lane 51, Lane 2000, Beizhai Road, Minhang District,\n"
    "Shanghai, China\n"
    "13817762730"
)

# Column widths / row heights copied from the real PL template so this sheet
# looks identical when opened in Excel.
PL_COL_WIDTHS = {
    "A": 16.875, "B": 14.25, "C": 13, "D": 27.8125, "E": 25.8125, "F": 17.5,
    "G": 10, "H": 14.25, "I": 11.8125, "J": 23.1875, "K": 8.3125, "L": 8.0625,
    "M": 6.9375, "N": 9.6875, "O": 13, "P": 15.1875, "Q": 17.0625, "R": 27.125,
    "S": 9, "T": 25.75,
}
PL_ROW_HEIGHTS = {1: 33.75, 2: 31.9, 3: 21, 4: 100.9, 5: 158.65, 6: 21.4,
                   7: 21.4, 8: 15, 9: 15, 10: 15, 11: 15, 12: 37.15, 13: 15.75}


def _style_cell(cell, *, bold=False, align="center", wrap=True, size=10):
    cell.font = Font(bold=bold, size=size)
    cell.border = PLAIN_BORDER
    cell.alignment = Alignment(horizontal=align, vertical="center", wrap_text=wrap)


_NO_SIDE = Side(style=None)


def _hdr_border(*, top=False, bottom=False, left=False, right=False):
    return Border(top=THIN_BLACK if top else _NO_SIDE,
                   bottom=THIN_BLACK if bottom else _NO_SIDE,
                   left=THIN_BLACK if left else _NO_SIDE,
                   right=THIN_BLACK if right else _NO_SIDE)


def _style_text(cell, *, bold=False, align="center", wrap=True, size=10):
    """Font + alignment only — border is handled separately by
    _write_pl_doc_header's border pass (see below), so header borders match
    the real template's box exactly instead of a blanket grid."""
    cell.font = Font(bold=bold, size=size)
    cell.alignment = Alignment(horizontal=align, vertical="center", wrap_text=wrap)


def _write_pl_doc_header(ws, notify_party_text: str, is_cn: bool = True):
    """Rows 1-11: PACKING LIST title, WPIC PO# / Date, Seller's EIN# /
    Invoice#, Shipper / Remark(SO#), Consignee / Notify Party, Trade term,
    then the Package/Quantity/Weight/CBM total block.

    Border layout copied cell-by-cell from the real template (checked
    against 3 real PL files): a single box around columns A:Q for rows 2-6
    (a horizontal divider between every field row, one vertical divider
    between the K/L columns splitting left field / right field), the
    "Trade term" row (6) only boxed on the left side (A:K) — matches the
    template exactly. Row 1 (title) and rows 7-11 (blank spacer + the
    Package/Qty/Weight/CBM total labels) only keep the K/L divider line
    continuing down, nothing else. Columns R/S/T are outside the header
    block in the real template and are left completely untouched (no
    border, no fill) — that's what was over-applied last time."""
    for letter, width in PL_COL_WIDTHS.items():
        ws.column_dimensions[letter].width = width
    for r, h in PL_ROW_HEIGHTS.items():
        ws.row_dimensions[r].height = h

    ws.cell(row=1, column=1, value="PACKING LIST（装箱单）")
    ws.merge_cells("A1:Q1")
    _style_text(ws.cell(row=1, column=1), bold=True, align="center", size=16)

    ws.cell(row=2, column=1, value="WPIC Purchase Order#/箱单编号：")
    ws.merge_cells("A2:D2")
    ws.cell(row=2, column=12, value=" 日期/Date：")
    ws.merge_cells("L2:M2")
    ws.merge_cells("N2:Q2")

    ws.cell(row=3, column=1, value="Seller's EIN#：")
    ws.merge_cells("A3:B3")
    ws.cell(row=3, column=12, value="Invoice#:")
    ws.merge_cells("L3:M3")
    ws.merge_cells("N3:Q3")

    ws.cell(row=4, column=1, value=SHIPPER_BLOCK)
    ws.merge_cells("A4:E4")
    ws.cell(row=4, column=12, value="Remark (SO#):")
    ws.merge_cells("L4:M4")
    ws.merge_cells("N4:Q4")

    ws.cell(row=5, column=1, value=_resolve_consignee(is_cn))
    ws.merge_cells("A5:E5")
    ws.cell(row=5, column=12, value=notify_party_text or "NOTIFY PARTY:\nDELIVERY ADDRESS:")
    ws.merge_cells("L5:Q5")

    ws.cell(row=6, column=1, value="成交方式/Trade term：")
    ws.merge_cells("A6:E6")

    ws.cell(row=8, column=1, value="Package Total: ")
    ws.cell(row=9, column=1, value="Quantity Total:")
    ws.cell(row=10, column=1, value="Gross Weight (KG):")
    ws.cell(row=11, column=1, value="CBM")
    # column B values (formulas referencing the TOTAL row) are filled in by
    # write_workbook() once the item table's TOTAL row number is known.

    # ── Font / alignment (all of A:Q, rows 1-11 — text only, no border) ────
    bold_label_cells = {"A1", "A2", "L2", "A3", "L3", "L4", "A6",
                         "A8", "A9", "A10", "A11"}
    for r in range(1, 7):
        for c in range(1, 18):  # A..Q only — R:T are outside the header box
            cell = ws.cell(row=r, column=c)
            align = "left" if r in (4, 5) and c in (1, 12) else "center"
            _style_text(cell, bold=(cell.coordinate in bold_label_cells or r == 1),
                        align=align, wrap=True, size=16 if r == 1 else 10)
    for r in (8, 9, 10, 11):
        _style_text(ws.cell(row=r, column=1), bold=True, align="left", wrap=False)

    # ── Border (matches the real template's box exactly) ───────────────────
    # Row 1: no border at all (floating title).
    ws.cell(row=1, column=1).border = Border()
    # Rows 2-5: full A:Q box, horizontal divider under every row, one
    # vertical divider between K (col 11) and L (col 12).
    for r in (2, 3, 4, 5):
        for c in range(1, 18):
            ws.cell(row=r, column=c).border = _hdr_border(
                top=True, bottom=True, left=(c == 1), right=(c == 17 or c == 11))
        ws.cell(row=r, column=12).border = _hdr_border(top=True, bottom=True, left=True)
    # Row 6 (Trade term): only the left half (A:K) is boxed — matches the
    # template, where the right/Notify-Party box stops at row 5.
    for c in range(1, 12):  # A..K
        ws.cell(row=6, column=c).border = _hdr_border(top=True, left=(c == 1), right=(c == 11))
    # Rows 7-11: blank spacer + totals — only the K/L divider continues.
    for r in range(7, 12):
        ws.cell(row=r, column=11).border = _hdr_border(right=True)
        ws.cell(row=r, column=12).border = _hdr_border(left=True)


def _write_pl_table_header(ws):
    for c, val in enumerate(PL_HEADERS_EN, start=1):
        ws.cell(row=TABLE_HDR_ROW1, column=c, value=val or None)
    for c, val in enumerate(PL_HEADERS_CN, start=1):
        ws.cell(row=TABLE_HDR_ROW2, column=c, value=val or None)
    ws.merge_cells(start_row=TABLE_HDR_ROW1, start_column=11, end_row=TABLE_HDR_ROW1, end_column=13)
    ws.merge_cells(start_row=TABLE_HDR_ROW2, start_column=11, end_row=TABLE_HDR_ROW2, end_column=13)
    for r in (TABLE_HDR_ROW1, TABLE_HDR_ROW2):
        for c in range(1, NCOLS + 1):
            _style_cell(ws.cell(row=r, column=c), bold=False, align="center", wrap=True)


def _auto_w(ws, cap=40):
    for col in ws.columns:
        letter = get_column_letter(col[0].column)
        w = max((len(str(c.value)) if c.value else 0) for c in col)
        ws.column_dimensions[letter].width = min(w + 3, cap)


def _apply_pl_merge(ws, start_row: int, end_row: int):
    """Merge the carton-level columns across all item rows of one package —
    plain style, no fill, no bold (matches the requested no-color grid)."""
    if end_row <= start_row:
        return
    for col in _MERGE_COLS:
        letter = get_column_letter(col)
        ws.merge_cells(f"{letter}{start_row}:{letter}{end_row}")


def _is_all_cn_factory(packages: List[Package]) -> bool:
    """True only if every package in this file is factory=CN (the known CN
    retail network this template's hardcoded CONSIGNEE / STORE_MASTER data
    applies to). Non-CN factories (POP/SBGEAR/QIFENG/JION/REVIEW) go to
    different countries/consignees this tool has no data for — CNEE and
    NOTIFY PARTY must stay blank and be filled in manually for those."""
    if not packages:
        return False
    try:
        import pl_group_export as pge
    except ImportError:
        return False
    return all(pge.detect_factory(p.reference_code, p.source_file) == "CN" for p in packages)


_MANUAL_FILL_NOTE = "(Ngoài CN — vui lòng tự điền chính xác / non-CN: fill in manually)"


def _resolve_consignee(is_cn: bool) -> str:
    if is_cn:
        return CONSIGNEE_BLOCK
    return "CONSIGNEE:\n" + _MANUAL_FILL_NOTE


def _resolve_notify_party(packages: List[Package], is_cn: bool) -> str:
    """Auto-fill Notify Party / Delivery Address ONLY when every package in
    this file belongs to the exact same known CN store (e.g. a
    04_CN_BY_STORE split file) — never guessed/blended when a file mixes
    stores (e.g. PL_Total, a factory file, or a CN-by-port file with several
    stores sharing one port). For non-CN factories, leave blank with a note
    reminding to fill it in manually — this tool has no address data for
    other countries. Ambiguous CN cases are left blank without the note
    (still CN, just needs the specific store filled in)."""
    if not is_cn:
        return "NOTIFY PARTY:\nDELIVERY ADDRESS:\n" + _MANUAL_FILL_NOTE
    stores = {p.store for p in packages if p.store and p.store != "REVIEW"}
    if len(stores) != 1:
        return ""
    try:
        import pl_group_export as pge
    except ImportError:
        return ""
    return pge.notify_party_block(next(iter(stores)))


# ── Workbook writer ────────────────────────────────────────────────────────
def write_workbook(output_path: Path, packages: List[Package]):
    wb = Workbook()
    wb.remove(wb.active)

    # ── Packing List (customer-facing, plain black grid) ────────────────────
    ws1 = wb.create_sheet("Packing List")
    is_cn = _is_all_cn_factory(packages)
    notify_party_text = _resolve_notify_party(packages, is_cn)
    _write_pl_doc_header(ws1, notify_party_text, is_cn)
    _write_pl_table_header(ws1)

    row_idx = FIRST_ITEM_ROW
    item_no = 0
    for pkg in packages:
        origin    = pkg.origin
        pkg_start = row_idx

        if not pkg.items:
            item_no += 1
            ws1.append([
                item_no, "", "",                      # Item# / PO No. / Invoice No. (manual)
                "", "", "", "", "",                     # Product/SKU/Barcode/UOM/Qty (no items)
                pkg.global_carton_num, pkg.package_code,
                pkg.length, pkg.width, pkg.height,
                pkg.weight, pkg.cbm,
                origin, "", "",                          # Origin / HTS / Shipping Mark (manual)
                pkg.port, "",                             # PORT (auto for CN) / 中国标签名称 (manual)
            ])
            row_idx += 1
        else:
            for item in pkg.items:
                item_no += 1
                ws1.append([
                    item_no, "", "",                                  # Item# / PO No. / Invoice No.
                    item.product_name, item.product_code, item.barcode,
                    item.unit, item.quantity,
                    pkg.global_carton_num, pkg.package_code,
                    pkg.length, pkg.width, pkg.height,
                    pkg.weight, pkg.cbm,
                    origin, item.hs_code, "",                          # Shipping Mark (manual)
                    pkg.port, "",                                       # PORT / 中国标签名称
                ])
                row_idx += 1

        end_row = row_idx - 1
        for r in range(pkg_start, end_row + 1):
            align_by_col = {4: "left"}  # Product Name left-aligned, rest centered
            for c in range(1, NCOLS + 1):
                _style_cell(ws1.cell(row=r, column=c), bold=False,
                            align=align_by_col.get(c, "center"), wrap=True)
        _apply_pl_merge(ws1, pkg_start, end_row)

    # ── TOTAL row ────────────────────────────────────────────────────────────
    total_qty = sum(p.calc_qty for p in packages)
    total_cartons = len(packages)
    total_weight = sum(p.weight for p in packages if p.weight is not None)
    total_cbm = sum(p.cbm for p in packages if p.cbm is not None)
    total_row = row_idx
    ws1.cell(row=total_row, column=1, value="TOTAL")
    ws1.merge_cells(start_row=total_row, start_column=1, end_row=total_row, end_column=7)
    ws1.cell(row=total_row, column=8, value=total_qty)
    ws1.cell(row=total_row, column=9, value=f"{total_cartons} Cartons")
    ws1.cell(row=total_row, column=14, value=round(total_weight, 3) if total_weight else 0)
    ws1.cell(row=total_row, column=15, value=round(total_cbm, 6) if total_cbm else 0)
    for c in range(1, NCOLS + 1):
        _style_cell(ws1.cell(row=total_row, column=c), bold=False, align="center", wrap=True)

    # Rows 8-11 (Package/Quantity/Weight/CBM totals, above the table) — live
    # formulas referencing the TOTAL row, same as the real template.
    ws1.cell(row=8, column=2, value=f"={get_column_letter(9)}{total_row}")
    ws1.cell(row=9, column=2, value=f"={get_column_letter(8)}{total_row}")
    ws1.cell(row=10, column=2, value=f"={get_column_letter(14)}{total_row}")
    ws1.cell(row=11, column=2, value=f"={get_column_letter(15)}{total_row}")
    for r in (8, 9, 10, 11):
        _style_text(ws1.cell(row=r, column=2), bold=False, align="left", wrap=False)

    for r in range(FIRST_ITEM_ROW, total_row + 1):
        ws1.row_dimensions[r].height = ws1.row_dimensions[r].height or 18
    for letter, width in PL_COL_WIDTHS.items():
        ws1.column_dimensions[letter].width = width
    ws1.freeze_panes = f"A{FIRST_ITEM_ROW}"

    # ── Match_Status (internal QC only — keeps color-coding on purpose) ─────
    ws2 = wb.create_sheet("Match_Status")
    ws2.append([
        "source_file", "reference_code", "package_code",
        "pdf_package_seq", "Carton number",
        "item_count", "calculated_total_qty", "declared_total_qty",
        "qty_match", "dim_match", "overall_status", "remark",
    ])
    for cell in ws2[1]:
        cell.fill = MS_HDR_FILL
        cell.font = MS_HDR_FONT
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws2.row_dimensions[1].height = 28
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

    # v8: CN store/port classification — fills pkg.port for the Packing List
    # sheet's PORT column, using the exact same rule as the CN split step.
    classify_packages_for_port(packages, pl_folder, recursive)

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
        "`packages` is empty — the OCR pipeline above found no PDFs or produced "
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
    print("XXXX SPLIT FAILED — reconciliation mismatch, nothing was silently swallowed XXXX")
    print(e)
    raise
except PermissionError as e:
    print("XXXX SPLIT FAILED — a target .xlsx is locked/open in Excel XXXX")
    print(e)
    raise
else:
    print(f'Completed: {SPLIT_OUTPUT_DIR}')
    print(f'Control file: {control_file}')
