#!/usr/bin/env python3
"""
Regression tests for pl_ocr_core.py (PL OCR + Grouping tool).

No test framework dependency, by design -- matches this repo's existing
convention (see tests/outbound-availability.test.js: plain assert-based
tests with a tiny custom runner, no package.json/pytest.ini). Run with:

    python3 tools/pl-ocr-grouping/tests/test_pl_ocr_core.py

Requires the same third-party packages pl_ocr_core.py itself needs
(pandas, openpyxl, pdfplumber) -- install with:

    pip install pandas openpyxl pdfplumber --break-system-packages

pl_ocr_core.py is written to be `exec()`'d by app.html inside Pyodide (see
its own header comment), not imported as a normal module -- its bottom half
unconditionally calls run_pipeline() against hard-coded /work/... paths.
_load_core_defs() / _load_core_pipeline() below reproduce exactly the
placeholder-substitution app.html performs, then exec() the (possibly
truncated) source into a throwaway module so these tests exercise the real
file, not a reimplementation of it.
"""
from __future__ import annotations
import sys, types, tempfile, shutil
from pathlib import Path

HERE = Path(__file__).resolve().parent
TOOL_DIR = HERE.parent
CORE_PY = TOOL_DIR / "pl_ocr_core.py"
if str(TOOL_DIR) not in sys.path:
    sys.path.insert(0, str(TOOL_DIR))
FIXTURES = HERE / "fixtures"
# SANITIZED SYNTHETIC fixtures only (see 2026-08-03 incident: the original
# version of this test suite committed 8 REAL customer packing-list PDFs +
# a real DIM.xlsx containing real customer names/phone numbers/addresses to
# this PUBLIC repo. That history has been purged -- see git log / README
# note for the cleanup commit. These synthetic fixtures were built with
# reportlab (tools/pl-ocr-grouping/tests/fixtures/synthetic/generate.py) to
# reproduce the exact same layout defects (merged condition+quantity cell,
# SKU "-BOX" suffix split across a newline, package header living inside
# the item table, package continuation across pages, one page with two
# packages, repeated table header + about:blank/date noise) using entirely
# fabricated codes/addresses. NO real shipment data ships in this repo.
SYN_PDFS = FIXTURES / "synthetic"
SYN_DIM = FIXTURES / "synthetic" / "SYN-DIM.xlsx"

_ENTRY_MARKER = "# ── Entry point ────────────────────────────────────────────────────────────"
_AUTO_SPLIT_MARKER = "# =========================================================\n# AUTO SPLIT:"


def _substitute_placeholders(src: str, *, dim_sheet=None, master_sheet=None,
                              recursive=False, consignee=None, notify=None,
                              generate_sublist=True, or_list_file=None,
                              generate_sublist_pdf=True) -> str:
    def lit(v):
        return "None" if v is None else repr(v)
    return (src
            .replace("__DIM_WEIGHT_SHEET__", lit(dim_sheet))
            .replace("__MASTER_DATA_SHEET__", lit(master_sheet))
            .replace("__RECURSIVE__", "True" if recursive else "False")
            .replace("__MANUAL_CONSIGNEE__", lit(consignee))
            .replace("__MANUAL_NOTIFY_PARTY__", lit(notify))
            .replace("__GENERATE_SUBLIST__", "True" if generate_sublist else "False")
            .replace("__GENERATE_SUBLIST_PDF__", "True" if generate_sublist_pdf else "False")
            .replace("__OR_LIST_FILE__", lit(str(or_list_file)) if or_list_file else "None")
            .replace("__GIT_COMMIT__", lit("test-suite")))


_module_counter = 0


def _exec_module(src: str, label: str):
    global _module_counter
    _module_counter += 1
    modname = f"pl_ocr_core_test_{_module_counter}"
    mod = types.ModuleType(modname)
    mod.__file__ = str(CORE_PY)
    sys.modules[modname] = mod
    exec(compile(src, str(CORE_PY), "exec"), mod.__dict__)
    return mod.__dict__


def load_core_defs():
    """Load every class/function/constant in pl_ocr_core.py WITHOUT running
    the pipeline (truncated before the entry-point block) -- for pure unit
    tests of parsing/normalization functions."""
    src = CORE_PY.read_text(encoding="utf-8")
    idx = src.index(_ENTRY_MARKER)
    src = _substitute_placeholders(src[:idx])
    return _exec_module(src, "defs")


def load_core_pipeline(pdf_dir: Path, dim_xlsx: Path, master_xlsx: Path, out_dir: Path, **kw):
    """Run the real run_pipeline() end-to-end (truncated just before the
    AUTO SPLIT / factory-store-grouping section, which is a separate
    feature) against real files on disk -- for integration tests."""
    src = CORE_PY.read_text(encoding="utf-8")
    idx = src.index(_AUTO_SPLIT_MARKER)
    src = _substitute_placeholders(src[:idx], **kw)
    src = src.replace('PL_FOLDER = Path("/work/pdfs")', f'PL_FOLDER = Path({str(pdf_dir)!r})')
    src = src.replace('OUTPUT_XLSX = Path("/work/PL_Total.xlsx")', f'OUTPUT_XLSX = Path({str(out_dir / "PL_Total.xlsx")!r})')
    src = src.replace('DIM_WEIGHT_FILE = Path("/work/dim.xlsx")', f'DIM_WEIGHT_FILE = Path({str(dim_xlsx)!r})')
    src = src.replace('MASTER_DATA_FILE = Path("/work/master.xlsx")', f'MASTER_DATA_FILE = Path({str(master_xlsx)!r})')
    return _exec_module(src, "pipeline")


def _make_empty_master_xlsx(path: Path):
    import openpyxl
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["SKU/Product Code", "HS Code", "EAN/Barcode"])
    wb.save(str(path))


# ── tiny test runner (mirrors tests/outbound-availability.test.js) ─────────
_passed = 0
_failed = 0
_failures = []


def test(name, fn):
    global _passed, _failed
    try:
        fn()
        _passed += 1
        print(f"  ok  - {name}")
    except AssertionError as e:
        _failed += 1
        _failures.append(name)
        print(f"FAIL  - {name}")
        print(f"        {e}")
    except Exception as e:
        _failed += 1
        _failures.append(name)
        print(f"ERROR - {name}")
        print(f"        {type(e).__name__}: {e}")


# =============================================================================
# 1. Parser unit tests -- pure functions/classes, no PDF needed
# =============================================================================
print("== Parser unit tests ==")
C = load_core_defs()


def t_cond_qty_merge_split():
    row = ['1', '4894961082009', 'TP-WST-RL10-MCC-02', '10mm Rope Loop', 'PCS', 'Moi 12']
    item = C["parse_item_cells"](row)
    assert item is not None, "item should parse"
    assert item.quantity == 12, f"quantity={item.quantity}"
    assert item.condition, "condition should be captured"


def t_sku_box_suffix_newline_in_cell():
    row = ['2', '4894961081941', 'TP-WST-RL10-MWD-02-\nBOX', '10mm Rope Loop', 'PCS', 'Moi 6']
    item = C["parse_item_cells"](row)
    assert item is not None
    assert item.product_code == "TP-WST-RL10-MWD-02-BOX", item.product_code


def t_sku_box_suffix_next_row():
    # Real-evidence shape (spec item #6: "-BOX" continuation on its own next
    # row/line): the item row's own SKU cell is already a complete code (no
    # trailing '-'), and the "-BOX" suffix arrives as an entirely separate
    # table row directly below it -- NOT as an adjacent cell within the same
    # row (that different case -- a cell ending in '-' merging with the very
    # next cell -- is covered by t_sku_box_suffix_newline_in_cell and is a
    # pre-existing, separate code path).
    Parser = C["Parser"]
    p = Parser()
    p.set_file(Path("CN-9999-Test-CN.pdf"))
    p.set_page(1)
    p.feed_table_row(["Ma kien hang: PGKECTEST0000001 1/1", "", "", "", "", ""])
    p.feed_table_row(["1", "4894961082009", "TP-WST-BGS-CMY-56", "Bag", "PCS", "Moi 4"])
    p.feed_table_row(["-BOX"])
    p.feed_table_row(["Tong cong 4"])
    p.finalise()
    assert len(p.packages) == 1
    item = p.packages[0].items[0]
    assert item.product_code == "TP-WST-BGS-CMY-56-BOX", item.product_code


def t_unicode_dash_artifact_normalized():
    fix = C["fix_unicode_artifacts"]
    assert fix("TP-WBA-CSL14-CMY￾51") == "TP-WBA-CSL14-CMY-51"
    assert fix("TP-WBA-CSL14-CMY￿51") == "TP-WBA-CSL14-CMY-51"


def t_description_newline_join_via_code():
    normalize_code = C["normalize_code"]
    assert normalize_code("TP-WST-BGS-CMY-56-\nBOX") == "TP-WST-BGS-CMY-56-BOX"


def t_gtin13_valid():
    is_valid = C["is_valid_gtin13"]
    assert is_valid("4894961082009") is True
    assert is_valid("123") is False
    assert is_valid("48949610820091") is False  # 14 digits, not 13
    assert is_valid("ABCDEFGHIJKLM") is False


def t_gtin_equals_sku_no_crash():
    row = ['1', '4894961082009', 'TP-4894961082009', 'Some Item', 'PCS', 'Moi 3']
    item = C["parse_item_cells"](row)  # must not raise
    assert item is not None


def t_header_footer_row_not_parsed_as_item():
    is_noise = C["is_noise"]
    assert is_noise("about:blank")
    assert is_noise("8/1/26, 4:27 PM about:blank")
    assert is_noise("# Barcode Ma san pham Ten hang hoa DVT Tinh trang So luong")
    assert not is_noise("1 4894961082009 TP-WST-RL10-MCC-02 10mm Rope Loop PCS Moi 12")


def t_repeated_table_header_not_item():
    is_hdr = C["is_table_hdr"]
    assert is_hdr(["#", "Barcode", "Ma san pham", "Ten hang hoa", "DVT", "Tinh trang So luong"])


def t_package_zero_items_status():
    Package = C["Package"]
    audit_status = C["audit_status"]
    pkg = Package(package_code="PGKECX", source_file="f.pdf", reference_code="f",
                  pdf_package_seq="1/1", declared_total_qty=10)
    assert audit_status(pkg) == "ZERO_ITEMS"


def t_package_missing_total_status():
    Package = C["Package"]
    audit_status = C["audit_status"]
    pkg = Package(package_code="PGKECX", source_file="f.pdf", reference_code="f", pdf_package_seq="1/1")
    pkg.items.append(C["Item"](no="1", product_name="x", product_code="TP-A-1", barcode="4894961082009",
                                unit="PCS", quantity=1))
    assert audit_status(pkg) == "MISSING_TOTAL"


def t_dedup_same_page_same_row_not_double_counted():
    """Table pass parses a row, then a redundant text-fallback pass (e.g. if
    the table produced 0 NEW items and the page's text is re-scanned)
    re-encounters the identical row: it must not be counted twice."""
    Parser = C["Parser"]
    p = Parser()
    p.set_file(Path("CN-9999-Test-CN.pdf"))
    p.set_page(1)
    p.feed_table_row(["Ma kien hang: PGKECTEST0000002 1/1", "", "", "", "", ""])
    row = ["1", "4894961082009", "TP-A-1", "Item A", "PCS", "Moi 5"]
    p.feed_table_row(row)
    p.feed_table_row(row)  # simulate the same physical row being fed twice
    p.feed_table_row(["Tong cong 5"])
    p.finalise()
    assert p.packages[0].item_count == 1, f"expected dedup, got {p.packages[0].item_count} items"
    assert p.duplicate_items_skipped == 1


def t_legit_repeated_sku_different_rows_both_kept():
    """The SAME sku/gtin/qty legitimately repeating on two DIFFERENT lines
    (different line_no) must both be kept -- dedup is context-aware, not a
    blind gtin+qty collapse."""
    Parser = C["Parser"]
    p = Parser()
    p.set_file(Path("CN-9999-Test-CN.pdf"))
    p.set_page(1)
    p.feed_table_row(["Ma kien hang: PGKECTEST0000003 1/1", "", "", "", "", ""])
    p.feed_table_row(["1", "4894961082009", "TP-A-1", "Item A", "PCS", "Moi 5"])
    p.feed_table_row(["2", "4894961082009", "TP-A-1", "Item A", "PCS", "Moi 5"])
    p.feed_table_row(["Tong cong 10"])
    p.finalise()
    assert p.packages[0].item_count == 2, "two genuinely different rows must both survive dedup"


def t_multi_package_same_page():
    Parser = C["Parser"]
    p = Parser()
    p.set_file(Path("CN-9999-Test-CN.pdf"))
    p.set_page(1)
    p.feed_table_row(["Ma kien hang: PGKECTEST0000004 1/2", "", "", "", "", ""])
    p.feed_table_row(["1", "4894961082009", "TP-A-1", "Item A", "PCS", "Moi 5"])
    p.feed_table_row(["Tong cong 5"])
    p.feed_table_row(["Ma kien hang: PGKECTEST0000005 2/2", "", "", "", "", ""])
    p.feed_table_row(["1", "4894961082010", "TP-A-2", "Item B", "PCS", "Moi 7"])
    p.feed_table_row(["Tong cong 7"])
    p.finalise()
    assert len(p.packages) == 2
    assert {pk.package_code for pk in p.packages} == {"PGKECTEST0000004", "PGKECTEST0000005"}


def t_package_spans_pages_totals_reconcile():
    Parser = C["Parser"]
    p = Parser()
    p.set_file(Path("CN-9999-Test-CN.pdf"))
    p.set_page(1)
    p.feed_table_row(["Ma kien hang: PGKECTEST0000006 1/1", "", "", "", "", ""])
    p.feed_table_row(["1", "4894961082009", "TP-A-1", "Item A", "PCS", "Moi 5"])
    p.set_page(2)
    p.feed_table_row(["Ma kien hang: PGKECTEST0000006 1/1", "", "", "", "", ""])  # repeated on continuation page
    p.feed_table_row(["2", "4894961082010", "TP-A-2", "Item B", "PCS", "Moi 8"])
    p.feed_table_row(["Tong cong 13"])
    p.finalise()
    assert len(p.packages) == 1, f"continuation must merge into ONE package, got {len(p.packages)}"
    pkg = p.packages[0]
    assert pkg.item_count == 2
    assert pkg.declared_total_qty == pkg.calc_qty == 13


def t_condition_word_separate_cell_8col():
    row = ["5", "4894961081927", "TP-WST-RL14-RGL-02", "14mm Rope Loop", "PCS", "Moi", "8", ""]
    item = C["parse_item_cells"](row)
    assert item is not None
    assert item.quantity == 8
    assert item.condition


def t_seven_column_row_all_fields():
    """Exactly 7 cells, no trailing blank column (distinct from the 8-col
    fixture above, which has a trailing ""). Verifies every field flagged
    in the follow-up review, not just quantity+condition: GTIN, SKU,
    description, UOM, condition, quantity."""
    row = ["3", "4894961082016", "TP-SVN-RL07-BLK-01", "7mm Rope Loop Black", "PCS", "Moi", "12"]
    item = C["parse_item_cells"](row)
    assert item is not None
    assert item.barcode == "4894961082016", item.barcode                # GTIN
    assert item.product_code == "TP-SVN-RL07-BLK-01", item.product_code  # SKU
    assert "Rope Loop" in item.product_name, item.product_name          # description
    assert item.unit == "PCS", item.unit                                # UOM
    assert item.condition, item.condition                               # condition
    assert item.quantity == 12, item.quantity                           # quantity


def t_delta_fallback_table_zero_items_recovers_via_text():
    """Direct regression test for bug #16 (`if tables: used_table = True`
    anti-pattern). Reproduces the exact Parser-level calls and the exact
    boolean condition run_pipeline() uses to decide whether to fall back to
    the text layer -- table pass yields a non-empty `tables` list whose
    every row is noise/header/blank (0 new items, 0 packages closed), so
    the fallback MUST trigger and the text layer's real item must be
    recovered exactly once (no duplication)."""
    Parser = C["Parser"]
    p = Parser()
    p.set_file(Path("CN-9999-Delta-CN.pdf"))
    p.set_page(1)
    p.feed_table_row(["Ma kien hang: PGKECDLT0000001 1/1", "", "", "", "", ""])

    items_before = p.total_item_count()
    closed_before = len(p.packages)
    # A non-empty `tables` list (3 rows) whose every row fails to produce a
    # valid item -- exactly the shape that used to permanently disable the
    # text fallback for the rest of the page under the old `if tables:` rule.
    garbage_table_rows = [
        ["about:blank", "", "", "", "", ""],                             # noise row
        ["STT", "GTIN", "SKU", "Ten hang", "DVT", "Tinh trang SL"],      # repeated header row
        ["", "", "", "", "", ""],                                        # blank row
    ]
    for row in garbage_table_rows:
        p.feed_table_row(row)
    items_after_table = p.total_item_count()
    closed_after_table = len(p.packages)

    assert items_after_table == items_before, "garbage table rows must not add any item"
    assert closed_after_table == closed_before, "garbage table rows must not close a package"

    # This is the literal run_pipeline() delta-fallback condition (bug #16
    # fix) re-evaluated here against the real Parser counters.
    should_fall_back = (items_after_table == items_before and closed_after_table == closed_before)
    assert should_fall_back, "delta-fallback condition must trigger when the table pass yields nothing new"

    item_line = "1 4894961082009 TP-DLT-001-A Rope Loop PCS Moi 9"
    p.feed_text_line(item_line)
    assert p._cur is not None and p._cur.item_count == 1, "text fallback must recover exactly 1 item"

    # Re-feeding the identical recovered line must NOT double-count it.
    dup_before = p.duplicate_items_skipped
    p.feed_text_line(item_line)
    assert p.duplicate_items_skipped == dup_before + 1, "re-feeding the same line must be caught by dedup"
    assert p._cur.item_count == 1, "item count must stay at 1 after the duplicate re-feed"

    p.feed_text_line("Tong cong 9")
    p.finalise()
    assert len(p.packages) == 1
    pkg = p.packages[0]
    assert pkg.item_count == 1, f"expected exactly 1 item after fallback recovery, got {pkg.item_count}"
    assert pkg.items[0].quantity == 9
    assert pkg.declared_total_qty == 9
    assert pkg.calc_qty == 9


test("condition+quantity merged cell split ('Moi 12')", t_cond_qty_merge_split)
test("SKU '-BOX' suffix joined (newline inside cell)", t_sku_box_suffix_newline_in_cell)
test("SKU '-BOX' suffix joined (suffix on its own next row)", t_sku_box_suffix_next_row)
test("U+FFFE / U+FFFF normalized to '-'", t_unicode_dash_artifact_normalized)
test("SKU split across newline joined via normalize_code", t_description_newline_join_via_code)
test("GTIN validated as exactly 13 digits", t_gtin13_valid)
test("SKU == GTIN does not crash the parser", t_gtin_equals_sku_no_crash)
test("header/footer noise (about:blank, date, repeated header) not parsed as item", t_header_footer_row_not_parsed_as_item)
test("repeated table header row detected", t_repeated_table_header_not_item)
test("zero-item package -> ZERO_ITEMS status", t_package_zero_items_status)
test("package with items but no Tong cong -> MISSING_TOTAL status", t_package_missing_total_status)
test("duplicate row (table+fallback re-read) not double-counted", t_dedup_same_page_same_row_not_double_counted)
test("legitimately repeated SKU on different rows both kept", t_legit_repeated_sku_different_rows_both_kept)
test("one page, multiple packages parsed correctly", t_multi_package_same_page)
test("package spanning pages: continuation merges, totals reconcile", t_package_spans_pages_totals_reconcile)
test("condition + quantity in separate cells (8-col layout)", t_condition_word_separate_cell_8col)
test("7-column row (no trailing blank col): GTIN/SKU/description/UOM/condition/quantity all correct", t_seven_column_row_all_fields)
test("delta-fallback: non-empty garbage table -> text layer recovers item, no duplicate (bug #16)", t_delta_fallback_table_zero_items_recovers_via_text)


# =============================================================================
# 2. DIM loader unit tests -- synthetic fixtures for tiers this real DIM
#    file doesn't happen to exercise (its headers are clean -> HEADER_MAPPING
#    only). PACKAGE_CODE_POSITIONAL_FALLBACK etc. are tested against
#    synthetic workbooks built here, per spec section 15's allowance to use
#    sanitized fixtures when a real broken-header file isn't available.
# =============================================================================
print("\n== DIM loader unit tests ==")


def _dim_wb(rows, headers=None, tmp_name="dim_fixture.xlsx"):
    import openpyxl
    wb = openpyxl.Workbook()
    ws = wb.active
    if headers is not None:
        ws.append(headers)
    for r in rows:
        ws.append(r)
    tmp_dir = Path(tempfile.mkdtemp(prefix="dim_fixture_"))
    path = tmp_dir / tmp_name
    wb.save(str(path))
    return path, tmp_dir


def t_dim_header_mapping_synthetic_file():
    DimMapper = C["DimMapper"]
    dim = DimMapper(SYN_DIM)
    assert dim.detection_method == "HEADER_MAPPING", dim.detection_method
    assert dim.valid_rows == 3, dim.valid_rows
    d = dim.lookup("SYN-1002-WarehouseBeta-CN", "PGKECSYN10020001")
    assert d is not None
    assert d["length"] == 45


def t_dim_partial_header_positional_fill():
    """PARTIAL_HEADER_POSITIONAL (tier 2): ref+pkg found by header alias,
    but the 5 dimension columns have meaningless headers ("X1".."X5") that
    match no alias -- found while writing this test, the original code
    (a) never actually filled those 5 fields positionally despite the
    docstring promising it, and (b) counted the row as "loaded" even though
    _commit_row() rejected it as malformed, so tier 3 was never tried and
    the sheet falsely reported success with 0 usable records. Both are
    fixed: PARTIAL_HEADER_POSITIONAL now reads all 5 dimension cells at
    pkg_col+1..+5 (same fixed order as tier 3), and the row-loaded counters
    only count rows _commit_row() actually accepted."""
    path, tmp = _dim_wb(
        headers=["Lo hang", "Tracking", "X1", "X2", "X3", "X4", "X5"],
        rows=[["REF-P", "PGKECPPP0000001", 10, 20, 15, 2.5, 0.003]],
    )
    try:
        DimMapper = C["DimMapper"]
        dim = DimMapper(path)
        assert dim.detection_method == "PARTIAL_HEADER_POSITIONAL", dim.detection_method
        assert dim.valid_rows == 1, dim.valid_rows
        d = dim.lookup("REF-P", "PGKECPPP0000001")
        assert d is not None
        assert d["length"] == 10 and d["width"] == 20 and d["height"] == 15
        assert d["weight"] == 2.5 and d["cbm"] == 0.003
    finally:
        shutil.rmtree(tmp)


def t_dim_header_alias_rename():
    path, tmp = _dim_wb(
        headers=["Reference", "Package", "L", "W", "H", "WT", "CBM"],
        rows=[["REF-A", "PGKECAAA0000001", 10, 20, 30, 1.5, 0.006]],
    )
    try:
        DimMapper = C["DimMapper"]
        dim = DimMapper(path)
        assert dim.detection_method == "HEADER_MAPPING", dim.detection_method
        assert dim.lookup("REF-A", "PGKECAAA0000001") is not None
    finally:
        shutil.rmtree(tmp)


def t_dim_broken_header_positional_fallback():
    path, tmp = _dim_wb(
        headers=["col1", "col2", "col3", "col4", "col5", "col6", "col7"],
        rows=[["REF-B", "PGKECBBB0000001", 15, 25, 35, 2.2, 0.013]],
    )
    try:
        DimMapper = C["DimMapper"]
        dim = DimMapper(path)
        assert dim.detection_method == "PACKAGE_CODE_POSITIONAL_FALLBACK", dim.detection_method
        d = dim.lookup("REF-B", "PGKECBBB0000001")
        assert d is not None
        assert d["length"] == 15 and d["cbm"] == 0.013
    finally:
        shutil.rmtree(tmp)


def t_dim_positional_with_leading_stt_column():
    path, tmp = _dim_wb(
        headers=["junk1", "junk2", "junk3", "junk4", "junk5", "junk6", "junk7", "junk8"],
        rows=[[1, "REF-C", "PGKECCCC0000001", 12, 22, 32, 3.1, 0.0087]],
    )
    try:
        DimMapper = C["DimMapper"]
        dim = DimMapper(path)
        assert dim.detection_method == "PACKAGE_CODE_POSITIONAL_FALLBACK", dim.detection_method
        d = dim.lookup("REF-C", "PGKECCCC0000001")
        assert d is not None, "STT column before package code must not be mistaken for the reference"
        assert d["length"] == 12
    finally:
        shutil.rmtree(tmp)


def t_dim_positional_package_code_not_fixed_column():
    """Row A has the package code in column 2, row B in column 3 (an empty
    spacer cell shifts it right) -- the fallback must scan every cell, not
    assume a fixed column index. The reference is still the nearest non-
    blank cell to the LEFT of the package code on each row."""
    path, tmp = _dim_wb(
        headers=["h1", "h2", "h3", "h4", "h5", "h6", "h7", "h8"],
        rows=[
            ["REF-D", "PGKECDDD0000001", 11, 21, 31, 4.0, 0.0072, ""],
            ["REF-E", "", "PGKECEEE0000001", 13, 23, 33, 5.0, 0.0099],
        ],
    )
    try:
        DimMapper = C["DimMapper"]
        dim = DimMapper(path)
        assert dim.lookup("REF-D", "PGKECDDD0000001") is not None
        assert dim.lookup("REF-E", "PGKECEEE0000001") is not None, \
            "package code in a different column on a different row must still be found"
    finally:
        shutil.rmtree(tmp)


def t_dim_reference_forward_fill_blank_cell():
    path, tmp = _dim_wb(
        headers=["h1", "h2", "h3", "h4", "h5", "h6", "h7"],
        rows=[
            ["REF-F", "PGKECFFF0000001", 10, 20, 30, 1.0, 0.006],
            ["", "PGKECFFF0000002", 10, 20, 30, 1.0, 0.006],  # blank ref -> forward-fill from row above
        ],
    )
    try:
        DimMapper = C["DimMapper"]
        dim = DimMapper(path)
        assert dim.lookup("REF-F", "PGKECFFF0000002") is not None, "blank reference should forward-fill"
    finally:
        shutil.rmtree(tmp)


def t_dim_insufficient_positional_fields():
    path, tmp = _dim_wb(
        headers=["h1", "h2", "h3"],
        rows=[["REF-G", "PGKECGGG0000001", 10]],  # only 1 cell after the package code, need 5
    )
    try:
        DimMapper = C["DimMapper"]
        dim = DimMapper(path)
        assert dim.lookup("REF-G", "PGKECGGG0000001") is None
        reasons = {d["validation_status"] for d in dim.diagnostics}
        assert "INSUFFICIENT_POSITIONAL_FIELDS" in reasons, reasons
    finally:
        shutil.rmtree(tmp)


def t_dim_non_numeric_dimension_fails_row():
    path, tmp = _dim_wb(
        headers=["h1", "h2", "h3", "h4", "h5", "h6", "h7"],
        rows=[["REF-H", "PGKECHHH0000001", "not-a-number", 20, 30, 1.0, 0.006]],
    )
    try:
        DimMapper = C["DimMapper"]
        dim = DimMapper(path)
        assert dim.lookup("REF-H", "PGKECHHH0000001") is None
        assert dim.malformed_rows >= 1
    finally:
        shutil.rmtree(tmp)


def t_dim_cbm_tolerance_review_required_no_overwrite():
    # length*width*height/1e6 = 10*20*30/1e6 = 0.006 expected; declared 0.5 is
    # wildly outside the 5% tolerance -> DIM_REVIEW_REQUIRED, but the ORIGINAL
    # 0.5 must be kept (never silently recalculated/overwritten).
    path, tmp = _dim_wb(
        headers=["h1", "h2", "h3", "h4", "h5", "h6", "h7"],
        rows=[["REF-I", "PGKECIII0000001", 10, 20, 30, 1.0, 0.5]],
    )
    try:
        DimMapper = C["DimMapper"]
        dim = DimMapper(path)
        d = dim.lookup("REF-I", "PGKECIII0000001")
        assert d is not None
        assert d["cbm"] == 0.5, f"original CBM must be preserved, got {d['cbm']}"
        assert dim.review_required == 1
        diag = [x for x in dim.diagnostics if x["normalized_package_code"] == "PGKECIII0000001"][0]
        assert diag["validation_status"] == "DIM_REVIEW_REQUIRED"
    finally:
        shutil.rmtree(tmp)


def t_dim_duplicate_key_flagged_first_wins():
    path, tmp = _dim_wb(
        headers=["h1", "h2", "h3", "h4", "h5", "h6", "h7"],
        rows=[
            ["REF-J", "PGKECJJJ0000001", 10, 20, 30, 1.0, 0.006],
            ["REF-J", "PGKECJJJ0000001", 99, 99, 99, 9.0, 0.9],  # duplicate key, different values
        ],
    )
    try:
        DimMapper = C["DimMapper"]
        dim = DimMapper(path)
        assert dim.duplicate_keys == 1
        d = dim.lookup("REF-J", "PGKECJJJ0000001")
        assert d["length"] == 10, "first occurrence must win, not be silently overwritten"
    finally:
        shutil.rmtree(tmp)


def t_dim_no_fuzzy_auto_match():
    path, tmp = _dim_wb(
        headers=["h1", "h2", "h3", "h4", "h5", "h6", "h7"],
        rows=[["REF-K", "PGKECKKK0000001", 10, 20, 30, 1.0, 0.006]],
    )
    try:
        DimMapper = C["DimMapper"]
        dim = DimMapper(path)
        # a near-miss (one character different / off-by-one) must NOT match.
        assert dim.lookup("REF-K", "PGKECKKK0000002") is None
        assert dim.lookup("REF-K-TYPO", "PGKECKKK0000001") is None
        diag = dim.diagnose_miss("test.pdf", "REF-K-TYPO", "PGKECKKK0000001")
        assert diag["mismatch_reason"] in ("NO_REFERENCE_IN_DIM", "KEY_COMBINATION_NOT_FOUND")
        # nearest candidates are informational only -- still not auto-applied.
        assert dim.lookup("REF-K-TYPO", "PGKECKKK0000001") is None
    finally:
        shutil.rmtree(tmp)


def t_raw_data_sheet_has_dim_per_item_row():
    out_dir = Path(tempfile.mkdtemp(prefix="pipeline_out_"))
    master_path = out_dir / "master.xlsx"
    _make_empty_master_xlsx(master_path)
    try:
        ns = load_core_pipeline(SYN_PDFS, SYN_DIM, master_path, out_dir)
        import openpyxl
        wb = openpyxl.load_workbook(str(out_dir / "PL_Total.xlsx"))
        assert "Raw_Data" in wb.sheetnames
        assert "Audit_Summary" in wb.sheetnames
        ws = wb["Raw_Data"]
        headers = [c.value for c in ws[1]]
        for col in ("length", "width", "height", "weight", "cbm", "dim_match_status"):
            assert col in headers, f"Raw_Data missing column {col}"
        length_idx = headers.index("length") + 1
        row2 = [ws.cell(row=2, column=c).value for c in range(1, len(headers) + 1)]
        assert row2[length_idx - 1] is not None, "Raw_Data row must carry DIM values directly (not merged cells)"
        wb2 = openpyxl.load_workbook(str(out_dir / "PL_Total.xlsx"))  # re-open sanity check (openpyxl compatible)
        assert wb2 is not None
    finally:
        shutil.rmtree(out_dir)


test("HEADER_MAPPING against the sanitized synthetic DIM.xlsx", t_dim_header_mapping_synthetic_file)
test("PARTIAL_HEADER_POSITIONAL: ref+pkg via header, L/W/H/Wt/CBM filled positionally", t_dim_partial_header_positional_fill)
test("HEADER_MAPPING via alias-renamed headers", t_dim_header_alias_rename)
test("broken headers -> PACKAGE_CODE_POSITIONAL_FALLBACK, fixed L/W/H/Wt/CBM order", t_dim_broken_header_positional_fallback)
test("leading STT column before package code still detected", t_dim_positional_with_leading_stt_column)
test("package code not in a fixed column, scanned per-row", t_dim_positional_package_code_not_fixed_column)
test("blank/merged reference forward-filled from row above", t_dim_reference_forward_fill_blank_cell)
test("insufficient cells after package code -> row fails", t_dim_insufficient_positional_fields)
test("non-numeric dimension -> row fails", t_dim_non_numeric_dimension_fails_row)
test("CBM outside tolerance -> DIM_REVIEW_REQUIRED, original CBM not overwritten", t_dim_cbm_tolerance_review_required_no_overwrite)
test("duplicate DIM key flagged, first occurrence wins (no silent overwrite)", t_dim_duplicate_key_flagged_first_wins)
test("no fuzzy auto-match of package code / reference", t_dim_no_fuzzy_auto_match)
test("Raw_Data sheet carries DIM values on every item row (not merge-only)", t_raw_data_sheet_has_dim_per_item_row)


# =============================================================================
# 3. Synthetic-file integration test (spec section 16 originally called for
#    a real-file integration test; those real files were purged from repo
#    history after the 2026-08-03 security incident -- see module docstring)
# =============================================================================
print("\n== Synthetic-file integration test (2 sanitized PDFs + synthetic DIM.xlsx) ==")


def t_synthetic_files_full_pipeline():
    """End-to-end run against the two sanitized synthetic PDFs + synthetic
    DIM.xlsx. This exercises the SAME layout defects the original real-file
    audit reproduced (merged condition+quantity cells, "-BOX" split across a
    newline, package header living inside the item table, continuation
    across pages, one page with two packages, about:blank/date/repeated-
    header noise) -- see tools/pl-ocr-grouping/tests/fixtures/synthetic/generate.py
    for exactly how each fixture maps to which bug.

    NOTE (see 2026-08-03 security incident): this replaces what used to be a
    genuine real-file integration test against 8 real customer PDFs, which
    were purged from this repo's history entirely. This synthetic run is
    NOT a substitute for periodically re-running the tool against real
    files locally (never committed) before trusting a release -- see
    README/limitations."""
    out_dir = Path(tempfile.mkdtemp(prefix="pipeline_syn_"))
    master_path = out_dir / "master.xlsx"
    _make_empty_master_xlsx(master_path)
    try:
        ns = load_core_pipeline(SYN_PDFS, SYN_DIM, master_path, out_dir)
        packages = ns["packages"]
        audit_status = ns["audit_status"]
        assert len(packages) == 3, f"expected 3 packages (1 spanning 2 pages + 2 on one page), got {len(packages)}"
        assert all(p.item_count > 0 for p in packages), "no package should have 0 items on these fixtures"
        assert sum(p.item_count for p in packages) == 9, f"expected 9 items total, got {sum(p.item_count for p in packages)}"
        statuses = {audit_status(p) for p in packages}
        assert statuses == {"OK"}, f"expected all packages OK, got {statuses}"
        assert all(p.declared_total_qty == p.calc_qty for p in packages), \
            "declared vs calculated totals must reconcile for every package"
        assert all(p.dim_matched for p in packages), "every package must DIM-match against the synthetic DIM.xlsx"
        assert ns["RUN_SUMMARY"]["errors"] == 0
        assert ns["RUN_SUMMARY"]["duplicate_items_skipped"] == 0
        seen_codes = [p.package_code for p in packages]
        assert len(seen_codes) == len(set(seen_codes)), "no duplicate package codes"
        by_code = {p.package_code: p for p in packages}
        assert by_code["PGKECSYN10010001"].item_count == 5, "the 2-page continuation package must merge to 5 items"
        assert by_code["PGKECSYN10010001"].calc_qty == 36
        box_item = next(i for i in by_code["PGKECSYN10010001"].items if i.product_code.endswith("-BOX"))
        assert box_item.product_code == "TP-SYN-A-001-BOX", box_item.product_code
    finally:
        shutil.rmtree(out_dir)


test("synthetic fixtures (2 PDFs, 3 packages, 9 items): all OK, totals reconcile, DIM 3/3", t_synthetic_files_full_pipeline)


# =============================================================================
# 9. Per-Store counting scope for carton numbering (spec section 9) --
#    counting_scope_key must group cartons per Store WITHIN a shipment,
#    never as one flat denominator across an entire multi-store upload.
# =============================================================================
print("\n== Per-Store counting scope (carton numbering) unit tests ==")
D9 = load_core_defs()
Package9 = D9["Package"]
assign_counting_scope_keys = D9["assign_counting_scope_keys"]
assign_global_numbers = D9["assign_global_numbers"]
compute_counting_scope_key = D9["compute_counting_scope_key"]


def _pkg9(ref, or_status="", or_num="", or_store="", store="", seq=1):
    p = Package9(package_code=f"PGKEC{ref}", source_file=f"{ref}.pdf",
                 reference_code=ref, pdf_package_seq=seq)
    p.or_list_match_status = or_status
    p.or_number = or_num
    p.or_list_store = or_store
    p.store = store
    return p


def t_scoped_numbering_spec_example_kerry_hangzhou():
    # Exact user-given example: Kerry 6 cartons, Hangzhou 4 cartons ->
    # independent 1/6..6/6 and 1/4..4/4 -- NEVER a combined 1/10..10/10.
    pkgs = [_pkg9(f"kerry_{i}", "OK", "OR1016", "Kerry", seq=i + 1) for i in range(6)]
    pkgs += [_pkg9(f"hz_{i}", "OK", "OR2044", "Hangzhou", seq=i + 1) for i in range(4)]
    assign_counting_scope_keys(pkgs)
    assign_global_numbers(pkgs)
    kerry = [p for p in pkgs if p.or_list_store == "Kerry"]
    hz = [p for p in pkgs if p.or_list_store == "Hangzhou"]
    assert [p.carton_display for p in kerry] == ["1/6", "2/6", "3/6", "4/6", "5/6", "6/6"],         [p.carton_display for p in kerry]
    assert [p.carton_display for p in hz] == ["1/4", "2/4", "3/4", "4/4"], [p.carton_display for p in hz]
    assert kerry[0].counting_scope_key != hz[0].counting_scope_key
    assert all(p.carton_total == 6 for p in kerry)
    assert all(p.carton_total == 4 for p in hz)


def t_scoped_numbering_no_store_signal_stays_flat_backward_compat():
    # No OR List match, no CN store resolved -> every package falls into
    # the SAME implicit "UPLOAD_BATCH|UNRESOLVED" scope -- i.e. today's
    # flat 1/N..N/N numbering, byte-for-byte, for every pre-v12 caller
    # that never touches counting_scope_key at all.
    pkgs = [_pkg9(f"f{i}", seq=i + 1) for i in range(5)]
    assign_counting_scope_keys(pkgs)
    assign_global_numbers(pkgs)
    assert [p.carton_display for p in pkgs] == ["1/5", "2/5", "3/5", "4/5", "5/5"]
    assert len({p.counting_scope_key for p in pkgs}) == 1


def t_scoped_numbering_cn_store_without_or_list_still_splits():
    # or_list_match_status is NOT "OK" (no OR List uploaded at all) but
    # pkg.store WAS resolved (classify_packages_for_port / CN split) --
    # counting scope must still key off that resolved Store.
    pkgs = [_pkg9(f"a{i}", store="SHENZHEN", seq=i + 1) for i in range(3)]
    pkgs += [_pkg9(f"b{i}", store="GUANGZHOU", seq=i + 1) for i in range(2)]
    assign_counting_scope_keys(pkgs)
    assign_global_numbers(pkgs)
    sz = [p for p in pkgs if p.store == "SHENZHEN"]
    gz = [p for p in pkgs if p.store == "GUANGZHOU"]
    assert [p.carton_display for p in sz] == ["1/3", "2/3", "3/3"]
    assert [p.carton_display for p in gz] == ["1/2", "2/2"]


def t_scoped_numbering_or_list_match_takes_priority_over_cn_store():
    # A package can have BOTH or_list_store (OK match) and a stale/blank
    # pkg.store -- compute_counting_scope_key must prefer the OR-List-
    # matched store, since it's the stronger, more specific signal.
    p = _pkg9("x", or_status="OK", or_num="OR9", or_store="Kerry", store="")
    key, source = compute_counting_scope_key(p)
    assert source == "OR_LIST_GROUPING", source
    assert "KERRY" in key.upper()


def t_scoped_numbering_review_status_falls_back_not_or_list_grouped():
    # or_list_match_status == "REVIEW" (ambiguous/no confident match) must
    # NOT be treated as a confirmed OR-List grouping signal.
    p = _pkg9("x", or_status="REVIEW", or_num="", or_store="", store="SHENZHEN")
    key, source = compute_counting_scope_key(p)
    assert source == "UPLOAD_BATCH_DEFAULT", source
    assert "SHENZHEN" in key.upper()


def t_scoped_numbering_case_insensitive_store_identity_merges():
    # "Kerry" (OR List's own casing) and "KERRY" (a hypothetical differently
    # -cased signal) must resolve to the SAME counting scope, not two.
    p1 = _pkg9("a", or_status="OK", or_num="OR1", or_store="Kerry")
    p2 = _pkg9("b", or_status="OK", or_num="OR1", or_store="KERRY")
    k1, _ = compute_counting_scope_key(p1)
    k2, _ = compute_counting_scope_key(p2)
    assert k1 == k2, (k1, k2)


test("per-Store scope: Kerry 6 cartons + Hangzhou 4 cartons number independently (1/6.. and 1/4.., never 1/10..)",
     t_scoped_numbering_spec_example_kerry_hangzhou)
test("per-Store scope: no store signal at all -> flat 1/N..N/N (backward-compat, matches every pre-v12 caller)",
     t_scoped_numbering_no_store_signal_stays_flat_backward_compat)
test("per-Store scope: CN store resolved without any OR List still splits scopes correctly",
     t_scoped_numbering_cn_store_without_or_list_still_splits)
test("per-Store scope: OK OR List match takes priority over a blank/stale pkg.store",
     t_scoped_numbering_or_list_match_takes_priority_over_cn_store)
test("per-Store scope: REVIEW status is never treated as a confirmed OR List grouping signal",
     t_scoped_numbering_review_status_falls_back_not_or_list_grouped)
test("per-Store scope: store identity is case-insensitive (Kerry == KERRY)",
     t_scoped_numbering_case_insensitive_store_identity_merges)



# ── summary ──────────────────────────────────────────────────────────────
print(f"\n{_passed} passed, {_failed} failed")
if _failed:
    print("FAILED:", ", ".join(_failures))
    sys.exit(1)
sys.exit(0)
