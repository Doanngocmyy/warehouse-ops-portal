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
    # v13 (FIX1 traceability): which header-detection tier produced this
    # row's column mapping -- "LITERAL_HEADER" (STORE/OR/SO aliases found
    # directly) or "SEMANTIC_FALLBACK_DUPLICATE_OR_HEADER" (duplicate
    # "OR"/"OR"/"SO" header, first OR column reinterpreted as STORE). Lets
    # the mapping report / UI show exactly how each row's Store came to be
    # recognised instead of leaving it implicit.
    detection_source: str = "LITERAL_HEADER"
    source_sheet: str = ""


@dataclass
class OrListImportResult:
    status: str = "NO_FILE"  # NO_FILE | OK | HEADER_NOT_FOUND | REQUIRED_FIELD_MISSING | LOAD_ERROR
    sheet_used: Optional[str] = None
    header_row: Optional[int] = None
    detection_source: str = ""  # LITERAL_HEADER | SEMANTIC_FALLBACK_DUPLICATE_OR_HEADER | "" (not detected)
    rows: List[OrListRow] = field(default_factory=list)
    duplicate_rows: List[OrListRow] = field(default_factory=list)  # exact-duplicate (store,or,so) rows, 2nd+ occurrence
    or_under_multiple_stores: Dict[str, List[str]] = field(default_factory=dict)  # or_norm -> [store_raw, ...]
    errors: List[str] = field(default_factory=list)
    warnings: List[str] = field(default_factory=list)
    sheets_tried: List[str] = field(default_factory=list)
    # Populated only when header detection fails outright (status ==
    # HEADER_NOT_FOUND) -- human-readable diagnostic text per sheet (sheet
    # names, first 10 non-empty rows' raw + normalized cell values, and
    # which STORE/OR/SO alias -- if any -- each cell in that row matched),
    # so a real uploaded file that genuinely isn't recognised can be
    # diagnosed from the run log instead of guessed at. Never used to
    # silently auto-accept a header -- see _find_header_row(), which is
    # unchanged: still requires an exact alias match, no fuzzy matching.
    diagnostics: List[str] = field(default_factory=list)

    @property
    def ok(self) -> bool:
        return self.status == "OK"


# =========================================================================
# 3) Header + sheet auto-detection (same scanning approach as
#    pl_ocr_core.HsCodeMapper -- scan the first N rows of each sheet for one
#    containing recognisable aliases, never assume a fixed row/sheet)
# =========================================================================
def _find_header_row(raw_df: pd.DataFrame, max_scan_rows: int = 40):
    """Returns (header_row_idx, {field: column_position}, detection_source)
    or (None, None, None).

    Column POSITIONS (0-based int), not labels -- a real-world file can have
    a header row like "OR" / "OR" / "SO" (duplicate literal text, no STORE
    column at all -- see SEMANTIC_FALLBACK_DUPLICATE_OR_HEADER below); pandas
    mangles duplicate column names on read ("OR" -> "OR"/"OR.1"), so a
    label-based map can't reliably say which physical column is meant.
    Position-based mapping is unambiguous for both branches.

    Tier 1 (LITERAL_HEADER, unchanged from before): a row with one column
    matching a STORE alias AND one matching an OR alias -- exact, no
    guessing.

    Tier 2 (SEMANTIC_FALLBACK_DUPLICATE_OR_HEADER, v13/FIX1): only tried
    when NO row anywhere in the sheet has a literal STORE alias. A row with
    >=2 columns matching an OR alias (by position) AND >=1 column matching
    an SO alias is a recognisable, deterministic shape (confirmed against a
    real production OR List: header literally reads "OR"/"OR"/"SO" with no
    STORE label, and the first "OR"-labelled column holds free-text Store
    descriptions, the second holds the actual PO/OR code, matching the
    STORE/OR/SO column ORDER convention used everywhere else in this
    workbook). The leftmost OR-aliased column is reinterpreted as STORE,
    the next one as OR, SO stays SO. This is a structural/positional rule,
    not a fuzzy content guess -- it never fires if a literal STORE header
    exists anywhere in the sheet (that always wins), and it never fires for
    a sheet with only ONE "OR"-aliased column (nothing to reinterpret as
    STORE). Lowest priority: only used when the whole sheet has no literal
    match anywhere."""
    semantic_fallback = None  # (row_idx, {field: position}) -- first candidate seen, kept as fallback only
    for i in range(min(max_scan_rows, len(raw_df))):
        cells = [c for c in raw_df.iloc[i].tolist()]
        norm_positions = [(pos, _norm_header(c)) for pos, c in enumerate(cells)]
        norm_map: Dict[str, int] = {}
        for pos, n in norm_positions:
            if n and n not in norm_map:  # keep the FIRST (leftmost) column for each normalized header text
                norm_map[n] = pos
        if not norm_map:
            continue
        store_pos = next((norm_map[a] for a in _STORE_ALIASES_NORM if a in norm_map), None)
        or_pos = next((norm_map[a] for a in _OR_ALIASES_NORM if a in norm_map), None)
        so_pos = next((norm_map[a] for a in _SO_ALIASES_NORM if a in norm_map), None)
        if store_pos is not None and or_pos is not None:
            mapping = {"store": store_pos, "or": or_pos}
            if so_pos is not None:
                mapping["so"] = so_pos
            return i, mapping, "LITERAL_HEADER"
        if store_pos is None and semantic_fallback is None:
            or_positions = [pos for pos, n in norm_positions if n in _OR_ALIASES_NORM]
            so_positions = [pos for pos, n in norm_positions if n in _SO_ALIASES_NORM]
            if len(or_positions) >= 2 and so_positions:
                store_p, or_p = or_positions[0], or_positions[1]
                so_p = so_positions[0]
                if store_p < or_p:
                    semantic_fallback = (i, {"store": store_p, "or": or_p, "so": so_p})
    if semantic_fallback is not None:
        i, mapping = semantic_fallback
        return i, mapping, "SEMANTIC_FALLBACK_DUPLICATE_OR_HEADER"
    return None, None, None


def _diagnose_sheet(raw_df: pd.DataFrame, sheet: str, max_scan_rows: int) -> str:
    """Builds a human-readable diagnostic block for ONE sheet that failed
    header detection: which rows were actually scanned, their raw and
    normalized cell values, and -- for each row -- which STORE/OR/SO alias
    (if any) each cell matched, plus a "score" (how many of the 2 required
    fields, STORE+OR, that row matched -- a row needs BOTH to qualify, so a
    score of 1 is a visible near-miss, not a silent one). This is diagnostic
    output ONLY -- it never changes what _find_header_row() accepts; a near-
    miss is reported, never auto-promoted to a match (spec: "khong duoc
    fuzzy-match mu quang")."""
    lines = [f"--- Sheet {sheet!r} (scanned up to {max_scan_rows} rows, sheet has {len(raw_df)} row(s) total) ---"]
    non_empty_shown = 0
    for i in range(len(raw_df)):
        if non_empty_shown >= 10:
            lines.append(f"    ... ({len(raw_df) - i} more row(s) not shown)")
            break
        if i >= max_scan_rows:
            lines.append(f"    (row {i + 1}+ beyond the {max_scan_rows}-row scan limit -- not scanned)")
            break
        raw_cells = [c for c in raw_df.iloc[i].tolist()]
        norm_cells = {_norm_header(c): c for c in raw_cells if _norm_header(c)}
        if not norm_cells:
            continue  # blank row -- not counted toward the 10-row diagnostic budget
        non_empty_shown += 1
        matched_store = [a for a in _STORE_ALIASES_NORM if a in norm_cells]
        matched_or = [a for a in _OR_ALIASES_NORM if a in norm_cells]
        matched_so = [a for a in _SO_ALIASES_NORM if a in norm_cells]
        score = (1 if matched_store else 0) + (1 if matched_or else 0)
        lines.append(
            f"    row {i + 1}: raw={raw_cells!r}\n"
            f"             normalized={list(norm_cells.keys())!r}\n"
            f"             candidate STORE match={bool(matched_store)} OR match={bool(matched_or)} "
            f"SO match={bool(matched_so)} -- score={score}/2 required (STORE+OR)"
        )
    if non_empty_shown == 0:
        lines.append("    (no non-empty rows found in this sheet)")
    return "\n".join(lines)


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

    _MAX_SCAN_ROWS = 40  # kept in sync with _find_header_row's own default
    _all_sheet_diagnostics: List[str] = []

    for sheet in sheets_to_try:
        result.sheets_tried.append(sheet)
        try:
            raw = pd.read_excel(path, sheet_name=sheet, header=None, dtype=str)
        except Exception as e:
            result.warnings.append(f"Sheet '{sheet}': could not read ({e})")
            _all_sheet_diagnostics.append(f"--- Sheet {sheet!r}: could not read ({e}) ---")
            continue
        if raw.empty:
            _all_sheet_diagnostics.append(f"--- Sheet {sheet!r}: empty, no rows at all ---")
            continue
        header_row_idx, mapping, detection_source = _find_header_row(raw, max_scan_rows=_MAX_SCAN_ROWS)
        if header_row_idx is None:
            # Header not found on THIS sheet -- build the diagnostic block now
            # (sheet names, first 10 non-empty rows raw+normalized, candidate
            # column scores, scan-row limit) in case every other sheet also
            # fails and this ends up surfaced as HEADER_NOT_FOUND below.
            _all_sheet_diagnostics.append(_diagnose_sheet(raw, sheet, _MAX_SCAN_ROWS))
            continue

        # v13 (FIX1): read data rows POSITIONALLY straight off `raw` (already
        # loaded with header=None) instead of re-reading via pandas with
        # header=header_row_idx -- a duplicate-text header ("OR"/"OR"/"SO")
        # would otherwise get silently mangled into "OR"/"OR.1" column
        # labels by pandas, which the SEMANTIC_FALLBACK path's positional
        # mapping was built to avoid depending on in the first place.
        n_cols = raw.shape[1]
        store_pos, or_pos = mapping["store"], mapping["or"]
        so_pos = mapping.get("so")
        if store_pos >= n_cols or or_pos >= n_cols:
            result.warnings.append(f"Sheet '{sheet}': header detected at row {header_row_idx + 1} "
                                    f"but the mapped column position(s) are out of range -- skipped.")
            continue

        rows: List[OrListRow] = []
        for i in range(header_row_idx + 1, len(raw)):
            row_vals = raw.iloc[i]
            store_raw = _clean_excel_value(row_vals.iloc[store_pos])
            or_raw = _clean_excel_value(row_vals.iloc[or_pos])
            so_raw = _clean_excel_value(row_vals.iloc[so_pos]) if so_pos is not None and so_pos < n_cols else ""
            if not store_raw and not or_raw and not so_raw:
                continue  # blank row -- ignored, not an error
            excel_row_number = i + 1  # raw's 0-based row index -> 1-based Excel row number
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
                detection_source=detection_source, source_sheet=sheet,
            ))

        if detection_source == "SEMANTIC_FALLBACK_DUPLICATE_OR_HEADER":
            result.warnings.append(
                f"Sheet '{sheet}' row {header_row_idx + 1}: no STORE column header found, but a "
                f"duplicate OR/OR/SO-shaped header was recognised -- the first OR-labelled column "
                f"was reinterpreted as STORE (semantic fallback, see OrListRow.detection_source).")

        if not rows and result.errors:
            result.status = "REQUIRED_FIELD_MISSING"
            result.sheet_used = sheet
            result.header_row = header_row_idx + 1
            result.detection_source = detection_source
            return result
        if rows:
            result.status = "OK"
            result.sheet_used = sheet
            result.header_row = header_row_idx + 1
            result.detection_source = detection_source
            result.rows = rows
            _validate_rows(result)
            return result

    result.status = "HEADER_NOT_FOUND"
    result.errors.append(
        "Could not find a header row containing both a STORE alias and an OR alias in any sheet "
        f"({', '.join(result.sheets_tried)}). Recognised STORE aliases: {STORE_ALIASES_RAW}; "
        f"OR aliases: {OR_ALIASES_RAW}.")
    result.diagnostics = _all_sheet_diagnostics
    diag_text = "\n".join(_all_sheet_diagnostics)
    log.warning(
        "OR List header not recognized -- diagnostic scan of every sheet follows "
        "(sheet names, first 10 non-empty rows raw+normalized, candidate STORE/OR/SO "
        f"column scores, {_MAX_SCAN_ROWS}-row scan limit):\n{diag_text}"
    )
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
