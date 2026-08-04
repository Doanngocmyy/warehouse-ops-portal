#!/usr/bin/env python3
"""
pl_or_list_import.py
=====================
Optional "OR List" Excel import: STORE / OR / SO, with auto-detected sheet
and header row (no hardcoded sheet name, no hardcoded column position) and
alias-based header matching. Self-contained -- no import from
pl_ocr_core.py / pl_group_export.py, same architecture principle already
used by those two modules (each is independently loadable/testable).

Never required: if no file is uploaded, callers simply never call
load_or_list() and the rest of the pipeline runs exactly as before (spec
section 2.2). If a file IS uploaded but its header can't be confidently
detected, this reports a clear, actionable status instead of guessing --
never silently accepts an ambiguous header (spec section 3).
"""
from __future__ import annotations

import logging
import re
import unicodedata
from dataclasses import dataclass, field
from pathlib import Path
from typing import Dict, List, Optional

import pandas as pd

log = logging.getLogger("pl_or_list_import")
if not log.handlers:
    logging.basicConfig(level=logging.INFO, format="%(asctime)s [%(levelname)s] %(message)s")

# =========================================================================
# 1) Header aliases (spec section 3)
# =========================================================================
STORE_ALIASES_RAW = ["STORE", "STORE NAME", "SHOP", "SHOP NAME",
                      "DESTINATION STORE", "CUSTOMER STORE"]
OR_ALIASES_RAW = ["OR", "OR #", "OR NO", "OR NO.", "OR NUMBER", "OR CODE", "OUTBOUND REQUEST"]
SO_ALIASES_RAW = ["SO", "SO #", "SO NO", "SO NO.", "SO NUMBER", "SO ORDER", "SO ORDER #",
                  "SALES ORDER", "SALES ORDER #"]


def _norm_header(s) -> str:
    """Normalize a header cell for alias matching: Unicode-fold, strip
    accents, uppercase, drop every '_ - / . #' and space (spec section 3:
    "normalization must handle" that exact list)."""
    if s is None:
        return ""
    s = unicodedata.normalize("NFKC", str(s))
    nfd = unicodedata.normalize("NFD", s)
    s = "".join(c for c in nfd if unicodedata.category(c) != "Mn")
    s = s.upper()
    s = re.sub(r"[_\-/.#\s]", "", s)
    return s


_STORE_ALIASES_NORM = {_norm_header(a) for a in STORE_ALIASES_RAW}
_OR_ALIASES_NORM = {_norm_header(a) for a in OR_ALIASES_RAW}
_SO_ALIASES_NORM = {_norm_header(a) for a in SO_ALIASES_RAW}


def _clean_excel_value(v) -> str:
    """Preserve original text (leading zeros, capitalization) -- the ONLY
    normalization applied is removing an artificial trailing ".0" that
    openpyxl/pandas adds when a cell that should be text was stored as a
    plain number in Excel (spec section 3: "Remove only artificial .0
    created by Excel numeric conversion")."""
    if v is None:
        return ""
    try:
        if pd.isna(v):
            return ""
    except Exception:
        pass
    s = str(v).strip()
    if re.fullmatch(r"\d+\.0", s):
        s = s[:-2]
    return s


# =========================================================================
# 2) Data model
# =========================================================================
@dataclass
class OrListRow:
    row_number: int  # 1-based Excel row (for traceability / error messages)
    store_raw: str
    store_norm: str
    or_raw: str
    or_norm: str
    so_raw: str = ""
    so_norm: str = ""


@dataclass
class OrListImportResult:
    status: str = "NO_FILE"  # NO_FILE | OK | HEADER_NOT_FOUND | REQUIRED_FIELD_MISSING | LOAD_ERROR
    sheet_used: Optional[str] = None
    header_row: Optional[int] = None
    rows: List[OrListRow] = field(default_factory=list)
    duplicate_rows: List[OrListRow] = field(default_factory=list)  # exact-duplicate (store,or,so) rows, 2nd+ occurrence
    or_under_multiple_stores: Dict[str, List[str]] = field(default_factory=dict)  # or_norm -> [store_raw, ...]
    errors: List[str] = field(default_factory=list)
    warnings: List[str] = field(default_factory=list)
    sheets_tried: List[str] = field(default_factory=list)

    @property
    def ok(self) -> bool:
        return self.status == "OK"


# =========================================================================
# 3) Header + sheet auto-detection (same scanning approach as
#    pl_ocr_core.HsCodeMapper -- scan the first N rows of each sheet for one
#    containing recognisable aliases, never assume a fixed row/sheet)
# =========================================================================
def _find_header_row(raw_df: pd.DataFrame, max_scan_rows: int = 40):
    """Returns (header_row_idx, {field: column_label}) or (None, None)."""
    for i in range(min(max_scan_rows, len(raw_df))):
        cells = [c for c in raw_df.iloc[i].tolist()]
        norm_cells = {_norm_header(c): c for c in cells if _norm_header(c)}
        if not norm_cells:
            continue
        store_col = next((norm_cells[a] for a in _STORE_ALIASES_NORM if a in norm_cells), None)
        or_col = next((norm_cells[a] for a in _OR_ALIASES_NORM if a in norm_cells), None)
        so_col = next((norm_cells[a] for a in _SO_ALIASES_NORM if a in norm_cells), None)
        if store_col is not None and or_col is not None:
            mapping = {"store": store_col, "or": or_col}
            if so_col is not None:
                mapping["so"] = so_col
            return i, mapping
    return None, None


def load_or_list(path, sheet_name: Optional[str] = None) -> OrListImportResult:
    """Load & validate an OR List workbook. Never raises for a
    missing/malformed file -- always returns a result object with a clear
    `status` the caller can branch on (spec: "khong duoc bao loi fatal chi
    vi thieu OR List" -- that's the caller's job to honor by simply not
    treating a bad OrListImportResult as a pipeline-stopping error)."""
    result = OrListImportResult()
    path = Path(path) if path else None
    if not path or not path.exists():
        result.status = "NO_FILE"
        return result

    try:
        xl = pd.ExcelFile(str(path))
    except Exception as e:
        result.status = "LOAD_ERROR"
        result.errors.append(f"Cannot open OR List workbook: {e}")
        return result

    sheets_to_try = [sheet_name] if sheet_name and sheet_name in xl.sheet_names else xl.sheet_names
    if sheet_name and sheet_name not in xl.sheet_names:
        result.warnings.append(f"OR List sheet '{sheet_name}' not found -- auto-detecting instead.")

    for sheet in sheets_to_try:
        result.sheets_tried.append(sheet)
        try:
            raw = pd.read_excel(path, sheet_name=sheet, header=None, dtype=str)
        except Exception as e:
            result.warnings.append(f"Sheet '{sheet}': could not read ({e})")
            continue
        if raw.empty:
            continue
        header_row_idx, mapping = _find_header_row(raw)
        if header_row_idx is None:
            continue

        df = pd.read_excel(path, sheet_name=sheet, header=header_row_idx, dtype=str)
        df.columns = [str(c).strip() for c in df.columns]
        store_col, or_col = mapping["store"], mapping["or"]
        so_col = mapping.get("so")
        if store_col not in df.columns or or_col not in df.columns:
            result.warnings.append(f"Sheet '{sheet}': header detected at row {header_row_idx + 1} "
                                    f"but the mapped columns didn't survive the header re-read -- skipped.")
            continue

        rows: List[OrListRow] = []
        for i, r in df.iterrows():
            store_raw = _clean_excel_value(r.get(store_col))
            or_raw = _clean_excel_value(r.get(or_col))
            so_raw = _clean_excel_value(r.get(so_col)) if so_col else ""
            if not store_raw and not or_raw and not so_raw:
                continue  # blank row -- ignored, not an error
            excel_row_number = header_row_idx + 2 + i  # +1 for 0->1-based, +1 for the header row itself
            if not store_raw or not or_raw:
                result.errors.append(
                    f"Sheet '{sheet}' row {excel_row_number}: STORE and OR are both required when an "
                    f"OR List is uploaded -- got STORE={store_raw!r} OR={or_raw!r}.")
                continue
            rows.append(OrListRow(
                row_number=excel_row_number,
                store_raw=store_raw, store_norm=_norm_header(store_raw),
                or_raw=or_raw, or_norm=_norm_header(or_raw),
                so_raw=so_raw, so_norm=_norm_header(so_raw),
            ))

        if not rows and result.errors:
            result.status = "REQUIRED_FIELD_MISSING"
            result.sheet_used = sheet
            result.header_row = header_row_idx + 1
            return result
        if rows:
            result.status = "OK"
            result.sheet_used = sheet
            result.header_row = header_row_idx + 1
            result.rows = rows
            _validate_rows(result)
            return result

    result.status = "HEADER_NOT_FOUND"
    result.errors.append(
        "Could not find a header row containing both a STORE alias and an OR alias in any sheet "
        f"({', '.join(result.sheets_tried)}). Recognised STORE aliases: {STORE_ALIASES_RAW}; "
        f"OR aliases: {OR_ALIASES_RAW}.")
    return result


def _validate_rows(result: OrListImportResult):
    """Duplicate-row and OR-under-multiple-Stores detection (spec section
    3). Never silently drops anything -- flags for the caller/UI instead."""
    seen: Dict[tuple, OrListRow] = {}
    for row in result.rows:
        key = (row.store_norm, row.or_norm, row.so_norm)
        if key in seen:
            result.duplicate_rows.append(row)
        else:
            seen[key] = row

    or_to_stores: Dict[str, set] = {}
    for row in result.rows:
        or_to_stores.setdefault(row.or_norm, set()).add(row.store_raw)
    for or_norm, stores in or_to_stores.items():
        if len(stores) > 1:
            result.or_under_multiple_stores[or_norm] = sorted(stores)

    if result.duplicate_rows:
        result.warnings.append(f"{len(result.duplicate_rows)} duplicate row(s) in the OR List (same STORE+OR+SO).")
    if result.or_under_multiple_stores:
        result.warnings.append(
            f"{len(result.or_under_multiple_stores)} OR value(s) appear under more than one Store: "
            f"{result.or_under_multiple_stores}")


def build_or_index(result: OrListImportResult) -> Dict[str, List[OrListRow]]:
    """or_norm -> [OrListRow, ...] (a list because the same OR CAN legitimately
    appear once per Store in a clean file; ambiguity is only when it maps to
    MORE THAN ONE DISTINCT store -- see or_under_multiple_stores)."""
    idx: Dict[str, List[OrListRow]] = {}
    for row in result.rows:
        idx.setdefault(row.or_norm, []).append(row)
    return idx
