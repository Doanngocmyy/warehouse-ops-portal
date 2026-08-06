#!/usr/bin/env python3
"""
Regression tests for pl_or_list_import.py -- OR List Excel import, header
auto-detection, and (Turn 12 fix) diagnostic logging when header detection
fails on a real uploaded file.

Same no-framework convention as the rest of this tool's tests. Run with:
    python3 tools/pl-ocr-grouping/tests/test_pl_or_list_import.py
"""
from __future__ import annotations
import sys, tempfile
from pathlib import Path

HERE = Path(__file__).resolve().parent
TOOL_DIR = HERE.parent
if str(TOOL_DIR) not in sys.path:
    sys.path.insert(0, str(TOOL_DIR))

import openpyxl
import pl_or_list_import as oli

_passed = 0
_failed = 0


def test(name, fn):
    global _passed, _failed
    try:
        fn()
        _passed += 1
        print(f"  ok  - {name}")
    except AssertionError as e:
        _failed += 1
        print(f"FAIL  - {name}\n        {e}")
    except Exception as e:
        _failed += 1
        print(f"ERROR - {name}\n        {type(e).__name__}: {e}")


def _write_xlsx(path, sheets: dict):
    """sheets: {sheet_name: [[row1 cells], [row2 cells], ...]}"""
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    for name, rows in sheets.items():
        ws = wb.create_sheet(name)
        for row in rows:
            ws.append(row)
    wb.save(str(path))


# =============================================================================
# 1. Happy path -- header found, rows loaded
# =============================================================================
print("== load_or_list: happy path ==")


def t_load_valid_or_list_ok():
    with tempfile.TemporaryDirectory() as td:
        p = Path(td) / "or_list.xlsx"
        _write_xlsx(p, {"Sheet1": [
            ["STORE", "OR", "SO"],
            ["Kerry", "OR1016", "SO4020"],
            ["Hangzhou", "OR2044", "SO4021"],
        ]})
        result = oli.load_or_list(p)
        assert result.status == "OK", result.errors
        assert result.ok
        assert len(result.rows) == 2
        assert result.sheet_used == "Sheet1"
        assert result.header_row == 1


def t_load_no_file_returns_no_file_status():
    result = oli.load_or_list(None)
    assert result.status == "NO_FILE"
    assert not result.ok


def t_header_found_on_second_sheet_when_first_is_unrelated():
    with tempfile.TemporaryDirectory() as td:
        p = Path(td) / "or_list.xlsx"
        _write_xlsx(p, {
            "Notes": [["some", "unrelated", "notes"], ["nothing here", "", ""]],
            "Data": [["STORE", "OR"], ["Kerry", "OR1016"]],
        })
        result = oli.load_or_list(p)
        assert result.status == "OK", result.errors
        assert result.sheet_used == "Data"


def t_header_row_offset_when_preceded_by_title_rows():
    with tempfile.TemporaryDirectory() as td:
        p = Path(td) / "or_list.xlsx"
        _write_xlsx(p, {"Sheet1": [
            ["OR List -- August shipment"],
            [],
            ["STORE", "OR", "SO"],
            ["Kerry", "OR1016", "SO4020"],
        ]})
        result = oli.load_or_list(p)
        assert result.status == "OK", result.errors
        assert result.header_row == 3


test("valid STORE/OR/SO header -> OK, rows loaded", t_load_valid_or_list_ok)
test("no file path -> NO_FILE, never treated as an error", t_load_no_file_returns_no_file_status)
test("header found on a later sheet when an earlier sheet has no header", t_header_found_on_second_sheet_when_first_is_unrelated)
test("header row correctly detected even when preceded by title/blank rows", t_header_row_offset_when_preceded_by_title_rows)


# =============================================================================
# 2. HEADER_NOT_FOUND -- diagnostics (Turn 12 fix)
# =============================================================================
print("\n== load_or_list: HEADER_NOT_FOUND diagnostics ==")


def t_header_not_found_when_no_recognizable_aliases():
    # v14: a positional fallback exists now (spec section 5: "Only
    # assumption: First column = Store"), so a bare 2-column shape with a
    # blank Store cell on the data row is what genuinely still fails every
    # tier -- no alias anywhere AND no "data continues under column A"
    # structural signal either.
    with tempfile.TemporaryDirectory() as td:
        p = Path(td) / "or_list.xlsx"
        _write_xlsx(p, {"Sheet1": [
            ["Warehouse", "Reference"],
            ["", "B"],
        ]})
        result = oli.load_or_list(p)
        assert result.status == "HEADER_NOT_FOUND"
        assert not result.ok


def t_header_not_found_populates_diagnostics_with_sheet_and_row_detail():
    with tempfile.TemporaryDirectory() as td:
        p = Path(td) / "or_list.xlsx"
        _write_xlsx(p, {"Sheet1": [
            ["Warehouse", "Reference"],
            ["", "B"],
        ]})
        result = oli.load_or_list(p)
        assert result.status == "HEADER_NOT_FOUND"
        assert result.diagnostics, "diagnostics must be populated on HEADER_NOT_FOUND"
        diag_text = "\n".join(result.diagnostics)
        assert "Sheet1" in diag_text
        assert "Warehouse" in diag_text and "Reference" in diag_text
        assert "raw=" in diag_text and "normalized=" in diag_text
        assert "score=" in diag_text
        assert "STORE match=" in diag_text and "OR match=" in diag_text and "SO match=" in diag_text


def t_header_not_found_diagnostics_show_near_miss_score():
    # v14: "Store Name" is itself a literal STORE alias, so with a second
    # populated column it's now a valid Tier-1 header (Store alone is
    # enough -- an "OR" alias is no longer required). A genuine near-miss
    # under the new model is a single-column sheet: no alias match AND no
    # second column for the positional fallback to use as a business field.
    with tempfile.TemporaryDirectory() as td:
        p = Path(td) / "or_list.xlsx"
        _write_xlsx(p, {"Sheet1": [
            ["Reference Code"],
            ["REF001"],
        ]})
        result = oli.load_or_list(p)
        assert result.status == "HEADER_NOT_FOUND"
        diag_text = "\n".join(result.diagnostics)
        assert "score=0/2" in diag_text, diag_text
        assert "STORE match=False" in diag_text
        assert "OR match=False" in diag_text


def t_header_not_found_diagnostics_cover_multiple_sheets():
    with tempfile.TemporaryDirectory() as td:
        p = Path(td) / "or_list.xlsx"
        _write_xlsx(p, {
            "Sheet1": [["Foo", "Bar"]],
            "Sheet2": [["Baz", "Qux"]],
        })
        result = oli.load_or_list(p)
        assert result.status == "HEADER_NOT_FOUND"
        diag_text = "\n".join(result.diagnostics)
        assert "Sheet1" in diag_text and "Sheet2" in diag_text


def t_header_not_found_diagnostics_respect_scan_row_limit():
    # v14: single column -- no second column for the positional fallback to
    # use as a business field, so this genuinely stays HEADER_NOT_FOUND
    # regardless of how many data rows follow.
    with tempfile.TemporaryDirectory() as td:
        p = Path(td) / "or_list.xlsx"
        rows = [["Col A"]] + [[f"x{i}"] for i in range(44)]
        _write_xlsx(p, {"Sheet1": rows})
        result = oli.load_or_list(p)
        assert result.status == "HEADER_NOT_FOUND"
        diag_text = "\n".join(result.diagnostics)
        assert "scan limit" in diag_text or "40 row" in diag_text.lower() or "scanned up to 40" in diag_text


test("no recognizable STORE/OR aliases anywhere -> HEADER_NOT_FOUND", t_header_not_found_when_no_recognizable_aliases)
test("HEADER_NOT_FOUND populates diagnostics with sheet name, raw+normalized cells, scores", t_header_not_found_populates_diagnostics_with_sheet_and_row_detail)
test("HEADER_NOT_FOUND diagnostics show a near-miss score (STORE alias present, OR absent) without auto-matching", t_header_not_found_diagnostics_show_near_miss_score)
test("HEADER_NOT_FOUND diagnostics cover every sheet tried, not just the first", t_header_not_found_diagnostics_cover_multiple_sheets)
test("HEADER_NOT_FOUND diagnostics state the scan-row limit", t_header_not_found_diagnostics_respect_scan_row_limit)


# =============================================================================
# 3. REQUIRED_FIELD_MISSING / duplicate / multi-store OR detection
# =============================================================================
print("\n== load_or_list: validation edge cases ==")


def t_required_field_missing_when_or_blank_for_every_row():
    with tempfile.TemporaryDirectory() as td:
        p = Path(td) / "or_list.xlsx"
        _write_xlsx(p, {"Sheet1": [
            ["STORE", "OR"],
            ["Kerry", ""],
        ]})
        result = oli.load_or_list(p)
        assert result.status == "REQUIRED_FIELD_MISSING"
        assert not result.ok


def t_duplicate_rows_flagged_not_dropped():
    with tempfile.TemporaryDirectory() as td:
        p = Path(td) / "or_list.xlsx"
        _write_xlsx(p, {"Sheet1": [
            ["STORE", "OR", "SO"],
            ["Kerry", "OR1016", "SO4020"],
            ["Kerry", "OR1016", "SO4020"],
        ]})
        result = oli.load_or_list(p)
        assert result.status == "OK"
        assert len(result.rows) == 2, "duplicate rows are flagged, never silently dropped"
        assert len(result.duplicate_rows) == 1


def t_or_under_multiple_stores_flagged():
    with tempfile.TemporaryDirectory() as td:
        p = Path(td) / "or_list.xlsx"
        _write_xlsx(p, {"Sheet1": [
            ["STORE", "OR"],
            ["Kerry", "OR1016"],
            ["Hangzhou", "OR1016"],
        ]})
        result = oli.load_or_list(p)
        assert result.status == "OK"
        assert result.or_under_multiple_stores, "same OR under 2 different Stores must be flagged"


test("blank OR for every row -> REQUIRED_FIELD_MISSING", t_required_field_missing_when_or_blank_for_every_row)
test("exact-duplicate rows are flagged in duplicate_rows, never silently dropped", t_duplicate_rows_flagged_not_dropped)
test("same OR value under two different Stores is flagged", t_or_under_multiple_stores_flagged)


# =============================================================================
# v14: fully dynamic business fields (spec section 5-6) -- ANY number of
# columns after Store, ANY labels, preserved exactly as uploaded.
# =============================================================================
print("\n== load_or_list: v14 dynamic business fields ==")


def t_dynamic_fields_arbitrary_4_column_shape():
    """Store, Ref No., Buyer, Delivery, Batch -- exactly the multi-field
    example from the spec. Labels must be preserved verbatim, values
    correctly keyed to each label, in column order."""
    with tempfile.TemporaryDirectory() as td:
        p = Path(td) / "or_list.xlsx"
        _write_xlsx(p, {"Sheet1": [
            ["Store", "Ref No.", "Buyer", "Delivery", "Batch"],
            ["Kerry", "REF-001", "Acme Co", "2026-08-01", "B7"],
        ]})
        result = oli.load_or_list(p)
        assert result.status == "OK", result.errors
        assert result.business_field_labels == ["Ref No.", "Buyer", "Delivery", "Batch"]
        row = result.rows[0]
        assert row.store_raw == "Kerry"
        assert list(row.business_fields.items()) == [
            ("Ref No.", "REF-001"), ("Buyer", "Acme Co"),
            ("Delivery", "2026-08-01"), ("Batch", "B7"),
        ]
        # backward-compat: 1st/2nd business field still readable the old way
        assert row.or_raw == "REF-001"
        assert row.so_raw == "Acme Co"


def t_dynamic_fields_po_invoice_labels_preserved_exactly():
    """A different real-world label set (PO / Invoice No.) must be
    preserved verbatim -- never silently renamed to "OR No."/"SO No."."""
    with tempfile.TemporaryDirectory() as td:
        p = Path(td) / "or_list.xlsx"
        _write_xlsx(p, {"Sheet1": [
            ["Store", "PO", "Invoice No."],
            ["Kerry", "PO38071", "INV628038"],
        ]})
        result = oli.load_or_list(p)
        assert result.status == "OK", result.errors
        assert result.business_field_labels == ["PO", "Invoice No."]
        assert result.rows[0].business_fields == {"PO": "PO38071", "Invoice No.": "INV628038"}


def t_dynamic_fields_single_business_field_ok():
    """A minimal Store + 1-field shape (no 2nd field at all) is valid --
    the old model required BOTH an OR-shaped and SO-shaped column; the new
    model only requires Store + at least one business field."""
    with tempfile.TemporaryDirectory() as td:
        p = Path(td) / "or_list.xlsx"
        _write_xlsx(p, {"Sheet1": [
            ["STORE", "Reference"],
            ["Kerry", "REF001"],
        ]})
        result = oli.load_or_list(p)
        assert result.status == "OK", result.errors
        assert result.business_field_labels == ["Reference"]
        assert result.rows[0].business_fields == {"Reference": "REF001"}


test("dynamic business fields: arbitrary 4-column shape (Ref No./Buyer/Delivery/Batch)", t_dynamic_fields_arbitrary_4_column_shape)
test("dynamic business fields: PO/Invoice No. labels preserved exactly as uploaded", t_dynamic_fields_po_invoice_labels_preserved_exactly)
test("dynamic business fields: Store + single business field is valid (no 2nd field required)", t_dynamic_fields_single_business_field_ok)


# =========================================================================
# v13 (FIX1): semantic header fallback -- a real production OR List whose
# header row literally reads "OR" / "OR" / "SO" (no STORE column at all),
# with the first OR-labelled column holding free-text Store descriptions
# and the second holding the actual PO/OR code.
# =========================================================================
def t_semantic_fallback_duplicate_or_header_loads_ok():
    with tempfile.TemporaryDirectory() as td:
        p = Path(td) / "or_list.xlsx"
        _write_xlsx(p, {"Sheet1": [
            ["OR", "OR", "SO"],
            ["20260609 CN - Guangzhou Parc Central Replen", "po38070", "inv628037"],
            ["20260609 CN - Hangzhou Mixc Replen", "po38068", "inv628036"],
        ]})
        result = oli.load_or_list(p)
        assert result.status == "OK", result.errors
        assert result.detection_source == "SEMANTIC_FALLBACK_DUPLICATE_OR_HEADER"
        assert len(result.rows) == 2
        r0 = result.rows[0]
        assert r0.store_raw == "20260609 CN - Guangzhou Parc Central Replen"
        assert r0.or_raw == "po38070"
        assert r0.so_raw == "inv628037"
        assert r0.detection_source == "SEMANTIC_FALLBACK_DUPLICATE_OR_HEADER"
        assert r0.source_sheet == "Sheet1"


def t_semantic_fallback_matches_real_7_store_or_list_shape():
    """Reproduces the exact real production OR List shape (7 stores,
    duplicate OR/OR/SO header) -- see audit notes for the source file."""
    with tempfile.TemporaryDirectory() as td:
        p = Path(td) / "or_list.xlsx"
        _write_xlsx(p, {"Sheet1": [
            ["OR", "OR", "SO"],
            ["20260609 CN - Guangzhou Parc Central Replen", "po38070", "inv628037"],
            ["20260609 CN - Hangzhou Mixc Replen", "po38068", "inv628036"],
            ["20260609 CN - Iapm Replen", "po38072", "inv628039"],
            ["20260609 CN - Kerry Center flagship Replen", "po38071", "inv628038"],
            ["20260609 CN - Shanghai Hongqiao Airport Replen", "po38074", "inv628042"],
            ["20260609 CN - Shanghai Taikooli (Shop B1-07b) Replen", "po38076", "inv628041"],
            ["20260609 CN - Shenzhen Mixc City (Shop T228) Replen", "po38073", "inv628040"],
        ]})
        result = oli.load_or_list(p)
        assert result.status == "OK", result.errors
        assert len(result.rows) == 7
        by_or = {r.or_raw: r for r in result.rows}
        assert by_or["po38074"].store_raw.startswith("20260609 CN - Shanghai Hongqiao")
        assert by_or["po38074"].so_raw == "inv628042"


def t_literal_store_header_always_wins_over_semantic_fallback():
    """A sheet with a genuine STORE column elsewhere must never fall back to
    the semantic OR/OR/SO reinterpretation, even if an earlier row happens
    to look duplicate-OR-shaped."""
    with tempfile.TemporaryDirectory() as td:
        p = Path(td) / "or_list.xlsx"
        _write_xlsx(p, {"Sheet1": [
            ["junk", "OR", "OR", "SO"],
            ["STORE", "OR", "junk", "SO"],
            ["Kerry", "OR1016", "", "SO4020"],
        ]})
        result = oli.load_or_list(p)
        assert result.status == "OK", result.errors
        assert result.detection_source == "LITERAL_HEADER"
        assert result.rows[0].store_raw == "Kerry"
        assert result.rows[0].or_raw == "OR1016"


def t_semantic_fallback_never_fires_with_only_one_or_column():
    """A single OR column (no duplicate) has nothing for Tier 2 (semantic
    duplicate-OR/OR/SO fallback) to reinterpret as STORE -- Tier 2 must
    never fire here. v14: Tier 3 (pure positional fallback) still applies
    though, since spec section 5's only real assumption is "first column =
    Store" -- so this now resolves via POSITIONAL_FALLBACK, not
    HEADER_NOT_FOUND and not SEMANTIC_FALLBACK_DUPLICATE_OR_HEADER."""
    with tempfile.TemporaryDirectory() as td:
        p = Path(td) / "or_list.xlsx"
        _write_xlsx(p, {"Sheet1": [
            ["OR", "SO"],
            ["po38070", "inv628037"],
        ]})
        result = oli.load_or_list(p)
        assert result.status == "OK", result.errors
        assert result.detection_source == "POSITIONAL_FALLBACK"
        assert result.rows[0].store_raw == "po38070"
        assert list(result.rows[0].business_fields.values()) == ["inv628037"]


def t_semantic_fallback_never_fires_without_an_so_column():
    """Two OR-aliased columns but no SO column at all is NOT the Tier-2
    recognised shape (spec: >=2 OR columns AND >=1 SO column) -- Tier 2
    must never fire here either. v14: Tier 3 still resolves it positionally
    -- first column becomes Store, so "Kerry Center flagship" (which reads
    like a real store name) is correctly picked up, just via a different
    detection tier than the OR/OR/SO-specific one."""
    with tempfile.TemporaryDirectory() as td:
        p = Path(td) / "or_list.xlsx"
        _write_xlsx(p, {"Sheet1": [
            ["OR", "OR"],
            ["Kerry Center flagship", "po38071"],
        ]})
        result = oli.load_or_list(p)
        assert result.status == "OK", result.errors
        assert result.detection_source == "POSITIONAL_FALLBACK"
        assert result.rows[0].store_raw == "Kerry Center flagship"
        assert list(result.rows[0].business_fields.values()) == ["po38071"]


test("semantic fallback: duplicate OR/OR/SO header (no STORE) loads OK, detection_source recorded", t_semantic_fallback_duplicate_or_header_loads_ok)
test("semantic fallback: matches the real 7-store production OR List shape exactly", t_semantic_fallback_matches_real_7_store_or_list_shape)
test("literal STORE header always wins over the semantic OR/OR/SO fallback", t_literal_store_header_always_wins_over_semantic_fallback)
test("semantic fallback never fires with only one OR-aliased column (nothing to reinterpret)", t_semantic_fallback_never_fires_with_only_one_or_column)
test("semantic fallback never fires without an SO column (not the recognised shape)", t_semantic_fallback_never_fires_without_an_so_column)


print(f"\n{_passed} passed, {_failed} failed")
if _failed:
    sys.exit(1)
