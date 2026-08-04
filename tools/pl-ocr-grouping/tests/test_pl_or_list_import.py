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
    with tempfile.TemporaryDirectory() as td:
        p = Path(td) / "or_list.xlsx"
        _write_xlsx(p, {"Sheet1": [
            ["Warehouse", "Reference", "Notes"],
            ["A", "B", "C"],
        ]})
        result = oli.load_or_list(p)
        assert result.status == "HEADER_NOT_FOUND"
        assert not result.ok


def t_header_not_found_populates_diagnostics_with_sheet_and_row_detail():
    with tempfile.TemporaryDirectory() as td:
        p = Path(td) / "or_list.xlsx"
        _write_xlsx(p, {"Sheet1": [
            ["Warehouse", "Reference", "Notes"],
            ["A", "B", "C"],
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
    with tempfile.TemporaryDirectory() as td:
        p = Path(td) / "or_list.xlsx"
        _write_xlsx(p, {"Sheet1": [
            ["Store Name", "Reference Code"],
            ["Kerry", "REF001"],
        ]})
        result = oli.load_or_list(p)
        assert result.status == "HEADER_NOT_FOUND"
        diag_text = "\n".join(result.diagnostics)
        assert "score=1/2" in diag_text, diag_text
        assert "STORE match=True" in diag_text
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
    with tempfile.TemporaryDirectory() as td:
        p = Path(td) / "or_list.xlsx"
        rows = [["Col A", "Col B"]] + [[f"x{i}", f"y{i}"] for i in range(44)]
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


print(f"\n{_passed} passed, {_failed} failed")
if _failed:
    sys.exit(1)
