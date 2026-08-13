#!/usr/bin/env python3
"""
Regression tests for pl_or_routing_template.py -- the "OR List / Routing
Template" import (Country Code | Port | Store | OR No. | Ref No.).

Spec under test ("IMPORTANT OR-TEMPLATE VALIDATION CORRECTION"):
    Country Code   REQUIRED, exactly 2 letters
    Port           CONDITIONAL / may be blank
    Store          CONDITIONAL / may be blank
    OR No.         OPTIONAL-BUT-WARN -- blank must NEVER block loading
    Ref No.        OPTIONAL-BUT-WARN -- blank must NEVER block loading

The core non-negotiable proven here: all FOUR OR/Ref combinations load
successfully and the pipeline continues (status stays OK, the row is
still present in `.rows`, `.to_routing_rules()` still includes it) --
only the diagnostic warning code differs.

Same no-framework convention as the rest of this tool's tests. Run with:
    python3 tools/pl-ocr-grouping/tests/test_or_routing_template.py
"""
from __future__ import annotations
import sys, tempfile
from pathlib import Path

HERE = Path(__file__).resolve().parent
TOOL_DIR = HERE.parent
if str(TOOL_DIR) not in sys.path:
    sys.path.insert(0, str(TOOL_DIR))

import openpyxl
import pl_or_routing_template as ort

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


HEADER = ["Country Code", "Port", "Store", "OR No.", "Ref No."]


# =============================================================================
# 1. The four required OR/Ref combinations -- each must load AND continue
#    the pipeline (never reject/fail/stop on a blank OR or Ref).
# =============================================================================
print("== the 4 OR/Ref combinations all load and run ==")


def t_or_present_ref_present_no_warning():
    with tempfile.TemporaryDirectory() as td:
        p = Path(td) / "t.xlsx"
        _write_xlsx(p, {"Sheet1": [HEADER, ["CN", "PVG", "CNWorld", "OR1172", "po38533"]]})
        r = ort.load_or_routing_template(p)
        assert r.status == "OK", r.errors
        assert len(r.rows) == 1
        row = r.rows[0]
        assert row.warning_code == "", row.warning_code
        assert row.or_raw == "OR1172" and row.ref_raw == "po38533"
        assert r.to_routing_rules() == [{"country": "CN", "port": "PVG", "store": "CNWorld"}]


def t_or_present_ref_blank_continues_with_warning_missing_ref():
    with tempfile.TemporaryDirectory() as td:
        p = Path(td) / "t.xlsx"
        _write_xlsx(p, {"Sheet1": [HEADER, ["CN", "", "Tmall", "OR1075", ""]]})
        r = ort.load_or_routing_template(p)
        assert r.status == "OK", r.errors
        assert len(r.rows) == 1, "OR-present/Ref-blank row must NOT be dropped"
        row = r.rows[0]
        assert row.warning_code == ort.WARNING_MISSING_REF
        assert row.or_raw == "OR1075" and row.ref_raw == ""
        # routing identity (Country/Port/Store) is unaffected by blank Ref
        assert r.to_routing_rules() == [{"country": "CN", "port": "", "store": "Tmall"}]
        assert any("Ref No." in w for w in r.warnings)


def t_or_blank_ref_present_continues_with_warning_missing_or():
    with tempfile.TemporaryDirectory() as td:
        p = Path(td) / "t.xlsx"
        _write_xlsx(p, {"Sheet1": [HEADER, ["SG", "", "", "", "po38515"]]})
        r = ort.load_or_routing_template(p)
        assert r.status == "OK", r.errors
        assert len(r.rows) == 1, "OR-blank/Ref-present row must NOT be dropped"
        row = r.rows[0]
        assert row.warning_code == ort.WARNING_MISSING_OR
        assert row.or_raw == "" and row.ref_raw == "po38515"
        assert r.to_routing_rules() == [{"country": "SG", "port": "", "store": ""}]
        assert any("OR No." in w for w in r.warnings)


def t_or_blank_ref_blank_continues_with_warning_missing_or_ref():
    with tempfile.TemporaryDirectory() as td:
        p = Path(td) / "t.xlsx"
        _write_xlsx(p, {"Sheet1": [HEADER, ["JP", "NRT", "Rakuten", "", ""]]})
        r = ort.load_or_routing_template(p)
        assert r.status == "OK", r.errors
        assert len(r.rows) == 1, "OR-blank/Ref-blank row must NOT be dropped -- routing identity alone is sufficient"
        row = r.rows[0]
        assert row.warning_code == ort.WARNING_MISSING_OR_REF
        assert row.or_raw == "" and row.ref_raw == ""
        assert r.to_routing_rules() == [{"country": "JP", "port": "NRT", "store": "Rakuten"}]
        assert any("OR No." in w and "Ref No." in w for w in r.warnings)


test("OR present + Ref present -> OK, no warning", t_or_present_ref_present_no_warning)
test("OR present + Ref blank -> pipeline continues, WARNING_MISSING_REF", t_or_present_ref_blank_continues_with_warning_missing_ref)
test("OR blank + Ref present -> pipeline continues, WARNING_MISSING_OR", t_or_blank_ref_present_continues_with_warning_missing_or)
test("OR blank + Ref blank -> pipeline continues, WARNING_MISSING_OR_REF", t_or_blank_ref_blank_continues_with_warning_missing_or_ref)


# =============================================================================
# 2. A single mixed sheet exercising all 4 combinations together (closer to
#    a real uploaded file) -- every row must survive, in order.
# =============================================================================
print("\n== mixed sheet: all 4 combinations in one file ==")


def t_mixed_sheet_all_four_combinations_together():
    with tempfile.TemporaryDirectory() as td:
        p = Path(td) / "t.xlsx"
        _write_xlsx(p, {"Sheet1": [
            HEADER,
            ["CN", "PVG", "CNWorld", "OR1172", "po38533"],   # both present
            ["CN", "", "Tmall", "OR1075", ""],                # OR only
            ["SG", "", "", "", "po38515"],                    # Ref only
            ["JP", "NRT", "Rakuten", "", ""],                 # neither
        ]})
        r = ort.load_or_routing_template(p)
        assert r.status == "OK", r.errors
        assert len(r.rows) == 4, f"all 4 rows must load, got {len(r.rows)}"
        codes = [row.warning_code for row in r.rows]
        assert codes == ["", ort.WARNING_MISSING_REF, ort.WARNING_MISSING_OR, ort.WARNING_MISSING_OR_REF]
        # routing rules present for ALL 4 rows regardless of OR/Ref state
        assert len(r.to_routing_rules()) == 4


test("one sheet mixing all 4 OR/Ref combinations: every row loads, warnings distinct", t_mixed_sheet_all_four_combinations_together)


# =============================================================================
# 3. Country Code is the ONLY structurally-required field
# =============================================================================
print("\n== Country Code required, Port/Store conditional ==")


def t_country_code_blank_row_skipped_but_file_still_loads():
    with tempfile.TemporaryDirectory() as td:
        p = Path(td) / "t.xlsx"
        _write_xlsx(p, {"Sheet1": [
            HEADER,
            ["", "PVG", "CNWorld", "OR1172", "po38533"],  # blank country -- skipped
            ["CN", "", "Tmall", "OR1075", "to10880"],       # valid
        ]})
        r = ort.load_or_routing_template(p)
        assert r.status == "OK", r.errors
        assert len(r.rows) == 1, "only the valid-country row should load"
        assert r.rows[0].country_norm == "CN"
        assert len(r.skipped_rows) == 1
        assert "Country Code" in r.skipped_rows[0][1]


def t_country_code_not_exactly_2_letters_row_skipped():
    with tempfile.TemporaryDirectory() as td:
        p = Path(td) / "t.xlsx"
        _write_xlsx(p, {"Sheet1": [
            HEADER,
            ["CHN", "PVG", "CNWorld", "OR1172", "po38533"],  # 3 letters -- invalid
            ["CN", "", "Tmall", "OR1075", "to10880"],
        ]})
        r = ort.load_or_routing_template(p)
        assert r.status == "OK", r.errors
        assert len(r.rows) == 1
        assert r.rows[0].country_norm == "CN"
        assert len(r.skipped_rows) == 1
        assert "2 letters" in r.skipped_rows[0][1]


def t_port_and_store_both_blank_is_still_a_valid_row():
    """Spec: Port CONDITIONAL/may be blank, Store CONDITIONAL/may be blank
    when routing is uniquely resolvable -- neither ever blocks loading."""
    with tempfile.TemporaryDirectory() as td:
        p = Path(td) / "t.xlsx"
        _write_xlsx(p, {"Sheet1": [HEADER, ["TH", "", "", "OR9001", "ref9001"]]})
        r = ort.load_or_routing_template(p)
        assert r.status == "OK", r.errors
        assert len(r.rows) == 1
        assert r.rows[0].port_raw == "" and r.rows[0].store_raw == ""
        assert r.to_routing_rules() == [{"country": "TH", "port": "", "store": ""}]


def t_country_code_case_insensitive_and_normalized_uppercase():
    with tempfile.TemporaryDirectory() as td:
        p = Path(td) / "t.xlsx"
        _write_xlsx(p, {"Sheet1": [HEADER, ["cn", "pvg", "CNWorld", "OR1172", "po38533"]]})
        r = ort.load_or_routing_template(p)
        assert r.status == "OK", r.errors
        assert r.rows[0].country_norm == "CN"


test("blank Country Code: that row is skipped, rest of file still loads (never a whole-file failure)", t_country_code_blank_row_skipped_but_file_still_loads)
test("Country Code not exactly 2 letters: that row is skipped, rest of file still loads", t_country_code_not_exactly_2_letters_row_skipped)
test("Port and Store both blank: still a fully valid, loadable row", t_port_and_store_both_blank_is_still_a_valid_row)
test("Country Code is normalized to uppercase regardless of input case", t_country_code_case_insensitive_and_normalized_uppercase)


# =============================================================================
# 4. Never a fatal failure just because OR/Ref is blank
# =============================================================================
print("\n== never reject/fail the template merely for blank OR/Ref ==")


def t_all_rows_missing_or_ref_still_status_ok_not_error():
    with tempfile.TemporaryDirectory() as td:
        p = Path(td) / "t.xlsx"
        _write_xlsx(p, {"Sheet1": [
            HEADER,
            ["CN", "PVG", "CNWorld", "", ""],
            ["SG", "", "", "", ""],
        ]})
        r = ort.load_or_routing_template(p)
        assert r.status == "OK", f"blank OR/Ref must never fail the whole template, got status={r.status} errors={r.errors}"
        assert len(r.rows) == 2
        assert all(row.warning_code == ort.WARNING_MISSING_OR_REF for row in r.rows)
        assert not r.errors, f"blank OR/Ref must never produce a hard error, got {r.errors}"


def t_real_uploaded_template_shape_loads_ok():
    """Matches the real downloadable template's exact header row + first
    3 data rows (Country Code | Port | Store | OR No. | Ref No.)."""
    with tempfile.TemporaryDirectory() as td:
        p = Path(td) / "t.xlsx"
        _write_xlsx(p, {"OR Template": [
            ["OR LIST / ROUTING TEMPLATE"],
            [],
            ["MUST FILL", "CONDITIONAL / CAN BE BLANK"],
            HEADER,
            ["CN", "PVG", "CNWorld", "OR1172", "po38533"],
            ["CN", "", "Tmall", "OR1075", "to10880"],
            ["SG", "", "", "OR1159", "po38515"],
        ]})
        r = ort.load_or_routing_template(p)
        assert r.status == "OK", r.errors
        assert len(r.rows) == 3
        assert [row.country_norm for row in r.rows] == ["CN", "CN", "SG"]
        assert r.to_routing_rules() == [
            {"country": "CN", "port": "PVG", "store": "CNWorld"},
            {"country": "CN", "port": "", "store": "Tmall"},
            {"country": "SG", "port": "", "store": ""},
        ]


test("template with every row missing OR/Ref entirely: status OK, zero hard errors, pipeline runs", t_all_rows_missing_or_ref_still_status_ok_not_error)
test("real downloadable template shape (title/legend rows above header) loads OK", t_real_uploaded_template_shape_loads_ok)


# =============================================================================
# 5. Structural edge cases (never crash, never a fatal file-level error)
# =============================================================================
print("\n== structural edge cases ==")


def t_no_file_returns_no_file_status_never_raises():
    r = ort.load_or_routing_template(None)
    assert r.status == "NO_FILE"


def t_no_country_code_column_anywhere_is_header_not_found():
    with tempfile.TemporaryDirectory() as td:
        p = Path(td) / "t.xlsx"
        _write_xlsx(p, {"Sheet1": [["Store", "OR No.", "SO No."], ["Tmall", "OR1075", "so1"]]})
        r = ort.load_or_routing_template(p)
        assert r.status == "HEADER_NOT_FOUND"


test("no file uploaded -> NO_FILE, never raises", t_no_file_returns_no_file_status_never_raises)
test("no Country Code column anywhere -> HEADER_NOT_FOUND (distinct from blank-OR/Ref, which is never fatal)", t_no_country_code_column_anywhere_is_header_not_found)


print(f"\n{_passed} passed, {_failed} failed")
if _failed:
    sys.exit(1)
