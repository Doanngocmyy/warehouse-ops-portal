#!/usr/bin/env python3
"""
pl_or_list_import.py
=====================
Optional "OR List" Excel import. v14: fully dynamic business metadata --
the ONLY structural assumption is that the Store column comes first; every
column after it is an arbitrary business field (e.g. "OR No."/"SO No.",
"PO"/"Invoice No.", or "Ref No."/"Buyer"/"Delivery"/"Batch" -- any count,
any label). Field labels are preserved EXACTLY as uploaded; normalization
is internal-only (for matching), never shown to the user.

Self-contained -- no import from pl_ocr_core.py / pl_group_export.py, same
architecture principle already used by those two modules.

Never required: if no file is uploaded, callers simply never call
load_or_list() and the rest of the pipeline runs exactly as before. If a
file IS uploaded but its header can't be confidently detected, this
reports a clear, actionable status instead of guessing.

Header detection tiers (highest priority first -- each tier is only tried
when every higher tier fails across the WHOLE sheet):
  1) LITERAL_HEADER -- a row with a column matching a known STORE alias
     ("STORE", "SHOP", ...). Every OTHER non-empty column in that row
     becomes a business field, in left-to-right order, using its own
     literal header text.
  2) SEMANTIC_FALLBACK_DUPLICATE_OR_HEADER -- no STORE alias anywhere, but
     >=2 columns match an OR alias and >=1 matches an SO alias (a real
     production shape: header literally reads "OR"/"OR"/"SO", first OR
     column holds free-text Store descriptions). Leftmost OR-aliased
     column is reinterpreted as STORE; every other non-empty column
     (including the reinterpreted OR/SO ones) becomes a business field
     under its own literal header text.
  3) POSITIONAL_FALLBACK -- no recognisable alias text at all (e.g. a
     completely custom header like "Ref No./Buyer/Delivery/Batch"). The
     first row whose column A is non-empty, has >=1 other non-empty
     column, and is immediately followed by a data row (column A also
     non-empty) is treated as the header: column A = STORE, every other
     non-empty column = a business field under its own literal text. This
     is a structural rule (position + "data follows"), never a fuzzy
     content guess.
"""
from __future__ import annotations

import logging
import re
import unicodedata
from collections import OrderedDict
from dataclasses import dataclass, field
from pathlib import Path
from typing import Dict, List, Optional, Tuple

import pandas as pd

log = logging.getLogger("pl_or_list_import")
if not log.handlers:
    logging.basicConfig(level=logging.INFO, format="%(asctime)s [%(levelname)s] %(message)s")

# =========================================================================
# 1) Header aliases -- used ONLY to LOCATE the header row / Store column,
#    never to decide what a business field is called (that's always the
#    literal uploaded text, spec section 6: "Preserve display header
#    exactly as uploaded").
# =========================================================================
STORE_ALIASES_RAW = ["STORE", "STORE NAME", "SHOP", "SHOP NAME",
                      "DESTINATION STORE", "CUSTOMER STORE",
                      "MEMO", "REMARK", "REMARKS", "DESCRIPTION", "DETAIL", "DETAILS"]
OR_ALIASES_RAW = ["OR", "OR #", "OR NO", "OR NO.", "OR NUMBER", "OR CODE", "OUTBOUND REQUEST"]
SO_ALIASES_RAW = ["SO", "SO #", "SO NO", "SO NO.", "SO NUMBER", "SO ORDER", "SO ORDER #",
                  "SALES ORDER", "SALES ORDER #"]


def _norm_header(s) -> str:
    """Normalize a header cell for alias matching: Unicode-fold, strip
    accents, uppercase, drop every '_ - / . #' and space."""
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
    plain number in Excel."""
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


def _clean_header_label(v) -> str:
    """Same as _clean_excel_value but for a HEADER cell -- collapse
    internal whitespace/newlines to a single space (a wrapped Excel header
    like "Product Name\\nin English" should still read naturally), but
    otherwise preserve the text exactly as uploaded."""
    s = _clean_excel_value(v)
    return re.sub(r"\s+", " ", s).strip()


# =========================================================================
# 2) Data model
# =========================================================================
@dataclass
class OrListRow:
    row_number: int  # 1-based Excel row (for traceability / error messages)
    store_raw: str
    store_norm: str
    # v14: fully dynamic business metadata -- display_header -> value, in
    # the exact column order they appeared in the uploaded file. ANY
    # number of fields, ANY labels (spec section 6).
    business_fields: "OrderedDict[str, str]" = field(default_factory=OrderedDict)
    detection_source: str = "LITERAL_HEADER"
    source_sheet: str = ""
    # v18 (SG-533-TEST real-fixture fix): the STORE column's own literal
    # header text (e.g. "OR" for a bare "OR | SO" sheet under
    # POSITIONAL_FALLBACK). Never used to relabel Store/OR No./Ref No. on
    # any CN-Store-matched output -- this exists ONLY so
    # match_store_and_or()'s no-Store-dimension fallback (see below) can
    # correctly recover an OR List that has NO real Store column at all
    # (every row's "store" position actually holds an OR value, e.g. this
    # real production shape: header row literally reads "OR"/"SO", no
    # Store/Shop column anywhere) without silently discarding that
    # column's own value. Defaults to "" so every existing call site /
    # test that constructs OrListRow without it is unaffected.
    store_header: str = ""

    # ---- Backward-compatible convenience accessors -------------------
    # A lot of existing matching/output code (and this session's own
    # earlier v13 work) reads "the OR value" / "the SO value" generically
    # as "the 1st business field" / "the 2nd business field" -- kept as
    # properties (not real dataclass fields) so both old callers and the
    # new fully-dynamic ones work against the SAME underlying data.
    @property
    def or_raw(self) -> str:
        vals = list(self.business_fields.values())
        return vals[0] if len(vals) >= 1 else ""

    @property
    def so_raw(self) -> str:
        vals = list(self.business_fields.values())
        return vals[1] if len(vals) >= 2 else ""

    @property
    def or_norm(self) -> str:
        return _norm_header(self.or_raw)

    @property
    def so_norm(self) -> str:
        return _norm_header(self.so_raw)


@dataclass
class OrListImportResult:
    status: str = "NO_FILE"  # NO_FILE | OK | HEADER_NOT_FOUND | REQUIRED_FIELD_MISSING | LOAD_ERROR
    sheet_used: Optional[str] = None
    header_row: Optional[int] = None
    detection_source: str = ""  # LITERAL_HEADER | SEMANTIC_FALLBACK_DUPLICATE_OR_HEADER | POSITIONAL_FALLBACK | ""
    # v14: the business field labels, in column order, EXACTLY as uploaded
    # -- e.g. ["OR No.", "SO No."] or ["PO", "Invoice No."] or
    # ["Ref No.", "Buyer", "Delivery", "Batch"]. Every OrListRow's
    # business_fields dict uses these same keys in the same order.
    business_field_labels: List[str] = field(default_factory=list)
    rows: List[OrListRow] = field(default_factory=list)
    duplicate_rows: List[OrListRow] = field(default_factory=list)  # exact-duplicate (store + all business fields) rows
    or_under_multiple_stores: Dict[str, List[str]] = field(default_factory=dict)  # 1st business field -> [store_raw, ...]
    errors: List[str] = field(default_factory=list)
    warnings: List[str] = field(default_factory=list)
    sheets_tried: List[str] = field(default_factory=list)
    diagnostics: List[str] = field(default_factory=list)

    @property
    def ok(self) -> bool:
        return self.status == "OK"


# =========================================================================
# 3) Header + sheet auto-detection
# =========================================================================
def _find_header_row_tiers_1_2(raw_df: pd.DataFrame, max_scan_rows: int = 40):
    """Alias-based detection ONLY (tiers 1-2). Returns (header_row_idx,
    store_col_pos, business_col_positions, detection_source) or
    (None, None, None, None). Kept separate from the positional fallback
    (tier 3, below) so callers can try tiers 1-2 across EVERY sheet first,
    and only fall back to the much looser positional rule if NO sheet in
    the whole workbook has a recognisable alias anywhere -- otherwise a
    weak positional match on an early, unrelated sheet (e.g. a "Notes"
    tab) would wrongly win over a clean literal-header match on a later
    sheet.

    business_col_positions: list of 0-based column positions (left-to-
    right, excluding the store column) to read as business fields, using
    THEIR OWN header cell text as the field label -- this is what makes
    field extraction fully dynamic regardless of which detection tier
    found the row."""
    semantic_fallback = None  # (row_idx, store_pos, business_positions) -- kept only if nothing better is found

    for i in range(min(max_scan_rows, len(raw_df))):
        cells = [c for c in raw_df.iloc[i].tolist()]
        norm_positions = [(pos, _norm_header(c)) for pos, c in enumerate(cells)]
        norm_map: Dict[str, int] = {}
        for pos, n in norm_positions:
            if n and n not in norm_map:  # keep the FIRST (leftmost) column for each normalized header text
                norm_map[n] = pos

        # ---- Tier 1: literal STORE alias -----------------------------
        store_pos = next((norm_map[a] for a in _STORE_ALIASES_NORM if a in norm_map), None)
        if store_pos is not None:
            business_positions = [pos for pos, c in enumerate(cells)
                                   if pos != store_pos and _clean_header_label(c)]
            return i, store_pos, business_positions, "LITERAL_HEADER"

        # ---- Tier 2 candidate: duplicate OR/OR/SO shape ---------------
        if semantic_fallback is None:
            or_positions = [pos for pos, n in norm_positions if n in _OR_ALIASES_NORM]
            so_positions = [pos for pos, n in norm_positions if n in _SO_ALIASES_NORM]
            if len(or_positions) >= 2 and so_positions and or_positions[0] < or_positions[1]:
                store_p = or_positions[0]
                business_positions = [pos for pos, c in enumerate(cells)
                                       if pos != store_p and _clean_header_label(c)]
                semantic_fallback = (i, store_p, business_positions)

    if semantic_fallback is not None:
        i, store_p, business_positions = semantic_fallback
        return i, store_p, business_positions, "SEMANTIC_FALLBACK_DUPLICATE_OR_HEADER"

    return None, None, None, None


def _find_header_row_tier_3(raw_df: pd.DataFrame, max_scan_rows: int = 40):
    """Pure positional fallback (spec section 5: "Only assumption: First
    column = Store") -- only tried across the WHOLE workbook after tiers
    1-2 have failed on EVERY sheet (see load_or_list()). Requires: column
    A non-empty, at least one other non-empty column in the same row, AND
    the very next row also has a non-empty column A (data continues) -- a
    structural signal (position + "data follows"), not a content guess.
    Returns (header_row_idx, 0, business_col_positions,
    "POSITIONAL_FALLBACK") or (None, None, None, None)."""
    n_cols = raw_df.shape[1]
    if n_cols == 0:
        return None, None, None, None
    for i in range(min(max_scan_rows, len(raw_df) - 1)):
        cells = [c for c in raw_df.iloc[i].tolist()]
        col_a = _clean_header_label(cells[0])
        if not col_a:
            continue
        other_non_empty = [pos for pos in range(1, n_cols) if _clean_header_label(cells[pos])]
        if not other_non_empty:
            continue
        next_row = raw_df.iloc[i + 1].tolist()
        if not _clean_excel_value(next_row[0]):
            continue
        return i, 0, other_non_empty, "POSITIONAL_FALLBACK"
    return None, None, None, None


def _diagnose_sheet(raw_df: pd.DataFrame, sheet: str, max_scan_rows: int) -> str:
    """Human-readable diagnostic block for ONE sheet that failed header
    detection under every tier -- never used to silently auto-accept a
    header, purely for the run log / UI."""
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
            continue
        non_empty_shown += 1
        matched_store = [a for a in _STORE_ALIASES_NORM if a in norm_cells]
        matched_or = [a for a in _OR_ALIASES_NORM if a in norm_cells]
        matched_so = [a for a in _SO_ALIASES_NORM if a in norm_cells]
        score = (1 if matched_store else 0) + (1 if matched_or else 0)
        lines.append(
            f"    row {i + 1}: raw={raw_cells!r}\n"
            f"             normalized={list(norm_cells.keys())!r}\n"
            f"             candidate STORE match={bool(matched_store)} OR match={bool(matched_or)} "
            f"SO match={bool(matched_so)} -- score={score}/2, positional fallback also failed "
            f"(needs col A + >=1 other non-empty column + non-empty col A on the NEXT row)"
        )
    if non_empty_shown == 0:
        lines.append("    (no non-empty rows found in this sheet)")
    return "\n".join(lines)


def load_or_list(path, sheet_name: Optional[str] = None) -> OrListImportResult:
    """Load & validate an OR List workbook. Never raises for a
    missing/malformed file -- always returns a result object with a clear
    `status` the caller can branch on."""
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

    _MAX_SCAN_ROWS = 40
    _all_sheet_diagnostics: List[str] = []
    _sheet_frames: "OrderedDict[str, pd.DataFrame]" = OrderedDict()

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
        _sheet_frames[sheet] = raw

    def _process_match(sheet, raw, header_row_idx, store_pos, business_positions, detection_source):
        n_cols = raw.shape[1]
        if store_pos >= n_cols or not business_positions:
            result.warnings.append(f"Sheet '{sheet}': header detected at row {header_row_idx + 1} "
                                    f"but had no usable Store/business-field columns -- skipped.")
            return False

        header_cells = raw.iloc[header_row_idx].tolist()
        field_labels = [_clean_header_label(header_cells[pos]) for pos in business_positions]
        # v18 (SG-533-TEST fix): preserve the STORE column's own literal
        # header text too -- see OrListRow.store_header's docstring.
        store_header = _clean_header_label(header_cells[store_pos]) if store_pos < len(header_cells) else ""

        rows: List[OrListRow] = []
        for i in range(header_row_idx + 1, len(raw)):
            row_vals = raw.iloc[i]
            store_raw = _clean_excel_value(row_vals.iloc[store_pos])
            business_fields: "OrderedDict[str, str]" = OrderedDict()
            for pos, label in zip(business_positions, field_labels):
                business_fields[label] = _clean_excel_value(row_vals.iloc[pos]) if pos < len(row_vals) else ""
            any_field_value = any(v for v in business_fields.values())
            if not store_raw and not any_field_value:
                continue  # blank row -- ignored, not an error
            excel_row_number = i + 1
            if not store_raw or not any_field_value:
                result.errors.append(
                    f"Sheet '{sheet}' row {excel_row_number}: Store and at least one business field are both "
                    f"required when an OR List is uploaded -- got Store={store_raw!r} fields={dict(business_fields)!r}.")
                continue
            rows.append(OrListRow(
                row_number=excel_row_number,
                store_raw=store_raw, store_norm=_norm_header(store_raw),
                business_fields=business_fields,
                detection_source=detection_source, source_sheet=sheet,
                store_header=store_header,
            ))

        if detection_source == "SEMANTIC_FALLBACK_DUPLICATE_OR_HEADER":
            result.warnings.append(
                f"Sheet '{sheet}' row {header_row_idx + 1}: no STORE column header found, but a "
                f"duplicate OR/OR/SO-shaped header was recognised -- the first OR-labelled column "
                f"was reinterpreted as STORE (semantic fallback, see OrListRow.detection_source).")
        elif detection_source == "POSITIONAL_FALLBACK":
            result.warnings.append(
                f"Sheet '{sheet}' row {header_row_idx + 1}: no recognised STORE/OR/SO alias text at all -- "
                f"used the first column as Store (positional fallback, see OrListRow.detection_source).")

        if not rows and result.errors:
            result.status = "REQUIRED_FIELD_MISSING"
            result.sheet_used = sheet
            result.header_row = header_row_idx + 1
            result.detection_source = detection_source
            result.business_field_labels = field_labels
            return True
        if rows:
            result.status = "OK"
            result.sheet_used = sheet
            result.header_row = header_row_idx + 1
            result.detection_source = detection_source
            result.business_field_labels = field_labels
            result.rows = rows
            _validate_rows(result)
            return True
        return False

    # Pass 1: alias-based tiers 1-2, across EVERY sheet, before ever trying
    # the much looser positional fallback on any sheet (see
    # _find_header_row_tiers_1_2's docstring for why order matters here).
    for sheet, raw in _sheet_frames.items():
        header_row_idx, store_pos, business_positions, detection_source = _find_header_row_tiers_1_2(
            raw, max_scan_rows=_MAX_SCAN_ROWS)
        if header_row_idx is None:
            continue
        if _process_match(sheet, raw, header_row_idx, store_pos, business_positions, detection_source):
            return result

    # Pass 2: positional fallback (tier 3), only reached when NO sheet
    # matched tiers 1-2 at all.
    for sheet, raw in _sheet_frames.items():
        header_row_idx, store_pos, business_positions, detection_source = _find_header_row_tier_3(
            raw, max_scan_rows=_MAX_SCAN_ROWS)
        if header_row_idx is None:
            _all_sheet_diagnostics.append(_diagnose_sheet(raw, sheet, _MAX_SCAN_ROWS))
            continue
        if _process_match(sheet, raw, header_row_idx, store_pos, business_positions, detection_source):
            return result

    result.status = "HEADER_NOT_FOUND"
    result.errors.append(
        "Could not find a Store column in any sheet -- tried a literal STORE alias, a duplicate OR/OR/SO "
        f"shape, and a positional (first-column) fallback, in sheets ({', '.join(result.sheets_tried)}). "
        f"Recognised STORE aliases: {STORE_ALIASES_RAW}; OR aliases: {OR_ALIASES_RAW}.")
    result.diagnostics = _all_sheet_diagnostics
    diag_text = "\n".join(_all_sheet_diagnostics)
    log.warning(
        "OR List header not recognized -- diagnostic scan of every sheet follows "
        "(sheet names, first 10 non-empty rows raw+normalized, candidate STORE/OR/SO "
        f"column scores, {_MAX_SCAN_ROWS}-row scan limit):\n{diag_text}"
    )
    return result


def _validate_rows(result: OrListImportResult):
    """Duplicate-row and (1st-business-field)-under-multiple-Stores
    detection. Never silently drops anything -- flags for the caller/UI
    instead."""
    seen: Dict[tuple, OrListRow] = {}
    for row in result.rows:
        key = (row.store_norm, tuple(_norm_header(v) for v in row.business_fields.values()))
        if key in seen:
            result.duplicate_rows.append(row)
        else:
            seen[key] = row

    or_to_stores: Dict[str, set] = {}
    for row in result.rows:
        if row.or_norm:
            or_to_stores.setdefault(row.or_norm, set()).add(row.store_raw)
    for or_norm, stores in or_to_stores.items():
        if len(stores) > 1:
            result.or_under_multiple_stores[or_norm] = sorted(stores)

    if result.duplicate_rows:
        result.warnings.append(f"{len(result.duplicate_rows)} duplicate row(s) in the OR List (same Store + all business fields).")
    if result.or_under_multiple_stores:
        primary_label = result.business_field_labels[0] if result.business_field_labels else "1st business field"
        result.warnings.append(
            f"{len(result.or_under_multiple_stores)} {primary_label} value(s) appear under more than one Store: "
            f"{result.or_under_multiple_stores}")


def build_or_index(result: OrListImportResult) -> Dict[str, List[OrListRow]]:
    """1st-business-field-normalized -> [OrListRow, ...] (a list because the
    same value CAN legitimately appear once per Store in a clean file;
    ambiguity is only when it maps to MORE THAN ONE DISTINCT store -- see
    or_under_multiple_stores)."""
    idx: Dict[str, List[OrListRow]] = {}
    for row in result.rows:
        idx.setdefault(row.or_norm, []).append(row)
    return idx
