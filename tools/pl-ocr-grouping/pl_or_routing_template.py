#!/usr/bin/env python3
"""
pl_or_routing_template.py — PL OCR V21 "OR List / Routing Template" import.

Loader + validator for the combined downloadable template that merges
routing identity (Country Code | Port | Store) with business references
(OR No. | Ref No.) into a single 5-column xlsx sheet.

FINAL FIELD POLICY (spec "IMPORTANT OR-TEMPLATE VALIDATION CORRECTION"):
    Country Code   REQUIRED, exactly 2 letters
    Port           CONDITIONAL / may be blank
    Store          CONDITIONAL / may be blank when routing is uniquely
                   resolvable
    OR No.         OPTIONAL-BUT-WARN -- blank must NEVER block loading or
                   pipeline execution
    Ref No.        OPTIONAL-BUT-WARN -- blank must NEVER block loading or
                   pipeline execution

Routing identity is determined by Country / Port / Store. Business
references (OR No. / Ref No.) are a SEPARATE dimension -- a row with a
blank OR No. and/or Ref No. is still a fully valid, usable routing row.
This module must NEVER:
    - reject the whole template because a row's OR/Ref is blank
    - fail the upload
    - stop OCR / stop export
    - classify the routing itself as failed merely because OR/Ref is blank

The only thing that can make an individual ROW unusable is a missing or
malformed Country Code (the one structurally required field) -- and even
then, only that row is skipped; the rest of the file loads normally.

Required per-row warning behavior (all four combinations continue the
pipeline; only the diagnostic code differs):
    OR present + Ref present  -> no warning
    OR present + Ref blank    -> WARNING_MISSING_REF
    OR blank   + Ref present  -> WARNING_MISSING_OR
    OR blank   + Ref blank    -> WARNING_MISSING_OR_REF

Self-contained except for reusing pl_routing_rules' own Country/Port
normalization (single source of truth for "exactly 2 letters", spec
section 3) -- pl_routing_rules.py is a leaf dependency with no reverse
import, so depending on it here does not create a cycle.
"""
from __future__ import annotations

import logging
import re
import unicodedata
from dataclasses import dataclass, field
from pathlib import Path
from typing import Dict, List, Optional, Tuple

import pandas as pd

from pl_routing_rules import normalize_country, normalize_port, is_valid_country_code

log = logging.getLogger("pl_or_routing_template")
if not log.handlers:
    logging.basicConfig(level=logging.INFO, format="%(asctime)s [%(levelname)s] %(message)s")

WARNING_MISSING_OR = "WARNING_MISSING_OR"
WARNING_MISSING_REF = "WARNING_MISSING_REF"
WARNING_MISSING_OR_REF = "WARNING_MISSING_OR_REF"

# =========================================================================
# Header aliases -- used ONLY to locate the 5 columns; display labels
# elsewhere are never derived from these (same principle as
# pl_or_list_import.py: alias lists locate structure, they never become
# user-facing text).
# =========================================================================
COUNTRY_ALIASES_RAW = ["COUNTRY CODE", "COUNTRY", "COUNTRY CODE*", "DESTINATION COUNTRY", "MARKET"]
PORT_ALIASES_RAW = ["PORT"]
STORE_ALIASES_RAW = ["STORE", "STORE NAME", "SHOP", "SHOP NAME"]
OR_ALIASES_RAW = ["OR NO.", "OR NO", "OR NO.*", "OR #", "OR NUMBER", "OR CODE", "OR"]
REF_ALIASES_RAW = ["REF NO.", "REF NO", "REF NO.*", "REF #", "REFERENCE NO.", "REFERENCE NUMBER", "REF"]


def _norm_header(s) -> str:
    """Normalize a header cell for alias matching: Unicode-fold, strip
    accents/asterisks, uppercase, drop '_ - / . #' and spaces."""
    if s is None:
        return ""
    s = unicodedata.normalize("NFKC", str(s))
    nfd = unicodedata.normalize("NFD", s)
    s = "".join(c for c in nfd if unicodedata.category(c) != "Mn")
    s = s.upper().replace("*", "")
    s = re.sub(r"[_\-/.#\s]", "", s)
    return s


def _norm_alias_set(raw: List[str]) -> set:
    return {_norm_header(a) for a in raw}


_COUNTRY_ALIASES_NORM = _norm_alias_set(COUNTRY_ALIASES_RAW)
_PORT_ALIASES_NORM = _norm_alias_set(PORT_ALIASES_RAW)
_STORE_ALIASES_NORM = _norm_alias_set(STORE_ALIASES_RAW)
_OR_ALIASES_NORM = _norm_alias_set(OR_ALIASES_RAW)
_REF_ALIASES_NORM = _norm_alias_set(REF_ALIASES_RAW)


def _clean_cell(v) -> str:
    """Preserve original text (leading zeros, capitalization); the only
    normalization applied is dropping an artificial trailing '.0' that
    openpyxl/pandas adds when a text-like cell was stored as a number."""
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
# Data model
# =========================================================================
@dataclass
class OrRoutingTemplateRow:
    row_number: int  # 1-based Excel row
    country_raw: str
    country_norm: str  # normalized 2-letter code (e.g. "CN")
    port_raw: str
    store_raw: str
    or_raw: str
    ref_raw: str
    warning_code: str = ""  # "" | WARNING_MISSING_OR | WARNING_MISSING_REF | WARNING_MISSING_OR_REF
    warning_text: str = ""

    @property
    def has_or(self) -> bool:
        return bool(self.or_raw)

    @property
    def has_ref(self) -> bool:
        return bool(self.ref_raw)

    def to_routing_rule_dict(self) -> dict:
        """Country/Port/Store only -- feeds directly into
        pl_routing_rules.validate_routing_rules() / match_route(). Business
        references are a separate dimension and intentionally excluded."""
        return {"country": self.country_norm, "port": self.port_raw, "store": self.store_raw}


@dataclass
class OrRoutingTemplateResult:
    status: str = "NO_FILE"  # NO_FILE | OK | HEADER_NOT_FOUND | REQUIRED_FIELD_MISSING | LOAD_ERROR
    sheet_used: Optional[str] = None
    header_row: Optional[int] = None
    rows: List[OrRoutingTemplateRow] = field(default_factory=list)
    # rows dropped for lacking the ONE structurally-required field
    # (Country Code missing or not exactly 2 letters) -- (excel_row, reason)
    skipped_rows: List[Tuple[int, str]] = field(default_factory=list)
    errors: List[str] = field(default_factory=list)
    warnings: List[str] = field(default_factory=list)
    sheets_tried: List[str] = field(default_factory=list)

    @property
    def ok(self) -> bool:
        return self.status == "OK"

    @property
    def rows_missing_or(self) -> List[OrRoutingTemplateRow]:
        return [r for r in self.rows if r.warning_code in (WARNING_MISSING_OR, WARNING_MISSING_OR_REF)]

    @property
    def rows_missing_ref(self) -> List[OrRoutingTemplateRow]:
        return [r for r in self.rows if r.warning_code in (WARNING_MISSING_REF, WARNING_MISSING_OR_REF)]

    def to_routing_rules(self) -> List[dict]:
        """Every valid row's Country/Port/Store, in file order -- ready to
        pass straight to pl_routing_rules.validate_routing_rules(). Rows
        with a blank OR/Ref are included -- business refs never gate
        routing eligibility (spec: separate dimensions)."""
        return [r.to_routing_rule_dict() for r in self.rows]


def _find_header_row(raw_df: pd.DataFrame, max_scan_rows: int = 40):
    """Locate the header row via a literal Country Code alias (the one
    structurally-required column -- spec: "Only Country Code is
    structurally required for the new routing-template mode"). Port/
    Store/OR No./Ref No. columns are located independently by their own
    aliases, in ANY order, and are each optional at the HEADER level too
    (a template missing a Port column, for example, simply yields
    port_raw="" for every row rather than failing to load).
    Returns (header_row_idx, {"country": pos, "port": pos|None,
    "store": pos|None, "or": pos|None, "ref": pos|None}) or (None, None).
    """
    for i in range(min(max_scan_rows, len(raw_df))):
        cells = [c for c in raw_df.iloc[i].tolist()]
        norm_positions = [(pos, _norm_header(c)) for pos, c in enumerate(cells)]
        norm_map: Dict[str, int] = {}
        for pos, n in norm_positions:
            if n and n not in norm_map:
                norm_map[n] = pos

        country_pos = next((norm_map[a] for a in _COUNTRY_ALIASES_NORM if a in norm_map), None)
        if country_pos is None:
            continue

        port_pos = next((norm_map[a] for a in _PORT_ALIASES_NORM if a in norm_map), None)
        store_pos = next((norm_map[a] for a in _STORE_ALIASES_NORM if a in norm_map), None)
        or_pos = next((norm_map[a] for a in _OR_ALIASES_NORM if a in norm_map), None)
        ref_pos = next((norm_map[a] for a in _REF_ALIASES_NORM if a in norm_map), None)

        return i, {"country": country_pos, "port": port_pos, "store": store_pos,
                    "or": or_pos, "ref": ref_pos}
    return None, None


def _row_warning(or_raw: str, ref_raw: str) -> Tuple[str, str]:
    has_or, has_ref = bool(or_raw), bool(ref_raw)
    if has_or and has_ref:
        return "", ""
    if has_or and not has_ref:
        return WARNING_MISSING_REF, "Ref No. is blank -- business reference incomplete, routing/output unaffected."
    if not has_or and has_ref:
        return WARNING_MISSING_OR, "OR No. is blank -- business reference incomplete, routing/output unaffected."
    return WARNING_MISSING_OR_REF, "OR No. and Ref No. are both blank -- routing/output unaffected."


def load_or_routing_template(path, sheet_name: Optional[str] = None) -> OrRoutingTemplateResult:
    """Load & validate an OR List / Routing Template workbook. Never raises
    for a missing/malformed file or for blank OR/Ref values -- always
    returns a result object with a clear `status`, per-row warnings, and
    the pipeline-ready routing rule dicts via `.to_routing_rules()`."""
    result = OrRoutingTemplateResult()
    path = Path(path) if path else None
    if not path or not path.exists():
        result.status = "NO_FILE"
        return result

    try:
        xl = pd.ExcelFile(str(path))
    except Exception as e:
        result.status = "LOAD_ERROR"
        result.errors.append(f"Cannot open OR List / Routing Template workbook: {e}")
        return result

    sheets_to_try = [sheet_name] if sheet_name and sheet_name in xl.sheet_names else xl.sheet_names
    if sheet_name and sheet_name not in xl.sheet_names:
        result.warnings.append(f"Sheet '{sheet_name}' not found -- auto-detecting instead.")

    _MAX_SCAN_ROWS = 40

    for sheet in sheets_to_try:
        result.sheets_tried.append(sheet)
        try:
            raw = pd.read_excel(path, sheet_name=sheet, header=None, dtype=str)
        except Exception as e:
            result.warnings.append(f"Sheet '{sheet}': could not read ({e})")
            continue
        if raw.empty:
            continue

        header_row_idx, positions = _find_header_row(raw, max_scan_rows=_MAX_SCAN_ROWS)
        if header_row_idx is None:
            continue

        rows: List[OrRoutingTemplateRow] = []
        n_missing_or = 0
        n_missing_ref = 0
        n_missing_both = 0

        for i in range(header_row_idx + 1, len(raw)):
            row_vals = raw.iloc[i]
            n_cols = len(row_vals)

            def cell(pos):
                return _clean_cell(row_vals.iloc[pos]) if pos is not None and pos < n_cols else ""

            country_raw = cell(positions["country"])
            port_raw = cell(positions["port"])
            store_raw = cell(positions["store"])
            or_raw = cell(positions["or"])
            ref_raw = cell(positions["ref"])
            excel_row_number = i + 1

            if not (country_raw or port_raw or store_raw or or_raw or ref_raw):
                continue  # fully blank row -- ignored, not an error

            country_norm = normalize_country(country_raw)
            if not country_raw:
                result.skipped_rows.append((excel_row_number, "Country Code is blank (the one structurally-required field)."))
                continue
            if not is_valid_country_code(country_norm):
                result.skipped_rows.append(
                    (excel_row_number, f"Country Code {country_raw!r} must be exactly 2 letters (e.g. CN, SG, JP)."))
                continue

            warning_code, warning_text = _row_warning(or_raw, ref_raw)
            if warning_code == WARNING_MISSING_OR:
                n_missing_or += 1
            elif warning_code == WARNING_MISSING_REF:
                n_missing_ref += 1
            elif warning_code == WARNING_MISSING_OR_REF:
                n_missing_both += 1

            rows.append(OrRoutingTemplateRow(
                row_number=excel_row_number,
                country_raw=country_raw, country_norm=country_norm,
                port_raw=normalize_port(port_raw) if port_raw else "",
                store_raw=store_raw,
                or_raw=or_raw, ref_raw=ref_raw,
                warning_code=warning_code, warning_text=warning_text,
            ))

        result.sheet_used = sheet
        result.header_row = header_row_idx + 1

        for excel_row_number, reason in result.skipped_rows:
            result.errors.append(f"Sheet '{sheet}' row {excel_row_number}: skipped -- {reason}")

        if n_missing_or:
            result.warnings.append(f"{n_missing_or} row(s) missing OR No. ({WARNING_MISSING_OR}) -- pipeline continues, output field left blank.")
        if n_missing_ref:
            result.warnings.append(f"{n_missing_ref} row(s) missing Ref No. ({WARNING_MISSING_REF}) -- pipeline continues, output field left blank.")
        if n_missing_both:
            result.warnings.append(f"{n_missing_both} row(s) missing both OR No. and Ref No. ({WARNING_MISSING_OR_REF}) -- pipeline continues, output fields left blank.")

        if rows:
            result.status = "OK"
            result.rows = rows
        elif result.skipped_rows:
            # every row that had ANY data failed on the one required field
            # (Country Code) -- never silently accepted as "OK", but this
            # is explicitly NOT a reason to fail the upload/stop OCR; the
            # caller can still choose to run with zero routing rows (falls
            # back to legacy behavior) rather than aborting.
            result.status = "REQUIRED_FIELD_MISSING"
        else:
            result.status = "OK"  # header found, no data rows at all -- valid empty template
        return result

    result.status = "HEADER_NOT_FOUND"
    result.errors.append(
        "Could not find a 'Country Code' column in any sheet -- the OR List / Routing Template requires a "
        f"Country Code header (tried sheets: {', '.join(result.sheets_tried)}). "
        f"Recognised aliases: {COUNTRY_ALIASES_RAW}."
    )
    return result
