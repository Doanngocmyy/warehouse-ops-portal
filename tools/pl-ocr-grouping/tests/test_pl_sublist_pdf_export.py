#!/usr/bin/env python3
"""
Regression tests for pl_sublist_pdf_export.py (A5 carton Sublist PDF).

Same no-framework convention as the rest of this tool's tests. Run with:

    python3 tools/pl-ocr-grouping/tests/test_pl_sublist_pdf_export.py

Requires reportlab + pdfplumber (used here only to read back generated PDF
text for assertions, not by the module under test itself).
"""
from __future__ import annotations
import sys, tempfile
from pathlib import Path
from types import SimpleNamespace

HERE = Path(__file__).resolve().parent
TOOL_DIR = HERE.parent
if str(TOOL_DIR) not in sys.path:
    sys.path.insert(0, str(TOOL_DIR))

import pl_sublist_pdf_export as ppe

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


def _item(no, ean, qty):
    return SimpleNamespace(product_code=no, barcode=ean, quantity=qty)


def _pkg(seq, total, items, shipping_mark="CN-1666-PVG-KERRY-POP",
         or_number="OR1016", so_number="so402064", weight=35.68,
         pl_gross_weight="", package_code="PGKECO377R7J0320001",
         reference_code="Kerry_POP"):
    return SimpleNamespace(
        carton_sequence=seq, carton_total=total, carton_display=f"{seq}/{total}",
        global_carton_num=f"{seq}/{total}",
        shipping_mark=shipping_mark, shipping_mark_source="PL_TEXT",
        or_number=or_number, or_source="PL_TEXT", so_number=so_number, so_source="OR_LIST",
        weight=weight, pl_gross_weight=pl_gross_weight, package_code=package_code, items=items,
        source_file=reference_code + ".pdf", reference_code=reference_code, pdf_package_seq=str(seq),
    )


# =============================================================================
# 1. Page geometry / capacity sanity
# =============================================================================
print("== Page geometry ==")


def t_page_is_a5_portrait():
    assert ppe.PAGE_W < ppe.PAGE_H, "A5 must be portrait (width < height)"
    # A5 is 148x210mm; at 72pt/in that's ~419.5 x 595.3pt.
    assert 400 < ppe.PAGE_W < 440
    assert 570 < ppe.PAGE_H < 610


def t_items_per_page_is_positive_and_reasonable():
    assert ppe.ITEMS_PER_PDF_PAGE > 0
    assert ppe.ITEMS_PER_PDF_PAGE < 200, "sanity bound -- something is very wrong if this is huge"


def t_no_font_files_referenced_only_base14():
    # Calibri is explicitly NOT to be shipped/distributed -- confirm every
    # font constant in this module is one of reportlab's Base-14 standard
    # fonts (no TTF/OTF embedding required).
    base14 = {"Helvetica", "Helvetica-Bold", "Helvetica-Oblique", "Helvetica-BoldOblique",
              "Times-Roman", "Times-Bold", "Times-Italic", "Times-BoldItalic",
              "Courier", "Courier-Bold", "Courier-Oblique", "Courier-BoldOblique", "Symbol", "ZapfDingbats"}
    used = {ppe.FONT_LABEL, ppe.FONT_VALUE, ppe.FONT_ITEM_HEADER, ppe.FONT_ITEM,
            ppe.FONT_TOTAL, ppe.FONT_CONTINUED}
    assert used.issubset(base14), f"non-Base14 font referenced: {used - base14}"
    assert "calibri" not in " ".join(used).lower()


test("A5 page is portrait with correct real-world dimensions", t_page_is_a5_portrait)
test("ITEMS_PER_PDF_PAGE is positive and sane", t_items_per_page_is_positive_and_reasonable)
test("only Base-14 standard fonts used (no Calibri/font files shipped)", t_no_font_files_referenced_only_base14)


# =============================================================================
# 2. Pagination (_paginate_for_pdf) -- one carton per page/continuation
# =============================================================================
print("\n== Pagination unit tests ==")

import pl_sublist_export as pse


def t_paginate_under_capacity_single_block():
    items = [_item(f"S{i}", f"E{i}", 1) for i in range(3)]
    pkg = _pkg(1, 1, items)
    carton = pse.build_sublist_carton_model(pkg)
    blocks = ppe._paginate_for_pdf([carton])
    assert len(blocks) == 1
    assert blocks[0].block_count == 1
    assert blocks[0].is_last_block
    assert blocks[0].block_carton_label == "1/1"
    assert blocks[0].block_total_qty == 3


def t_paginate_exact_capacity_single_block():
    cap = ppe.ITEMS_PER_PDF_PAGE
    items = [_item(f"S{i}", f"E{i}", 1) for i in range(cap)]
    pkg = _pkg(1, 1, items)
    carton = pse.build_sublist_carton_model(pkg)
    blocks = ppe._paginate_for_pdf([carton])
    assert len(blocks) == 1, f"exactly {cap} items must still fit on one page"


def t_paginate_over_capacity_splits_no_item_lost():
    cap = ppe.ITEMS_PER_PDF_PAGE
    n = cap + 15
    items = [_item(f"S{i}", f"E{i}", 2) for i in range(n)]
    pkg = _pkg(2, 6, items)
    carton = pse.build_sublist_carton_model(pkg)
    blocks = ppe._paginate_for_pdf([carton])
    assert len(blocks) == 2, f"expected 2 pages for {n} items at capacity {cap}"
    total_items = sum(len(b.items) for b in blocks)
    assert total_items == n, "zero items may ever be lost across continuation pages"
    assert not blocks[0].is_last_block and blocks[1].is_last_block
    assert blocks[0].block_total_qty == sum(r.qty for r in blocks[0].items), "non-last block shows a SUBTOTAL"
    assert blocks[1].block_total_qty == carton.total_qty, "last block shows the carton's GRAND TOTAL, not just its own rows"
    assert blocks[1].block_carton_label == "2/6 - Continued 1"


def t_paginate_continuation_never_inflates_carton_identity():
    cap = ppe.ITEMS_PER_PDF_PAGE
    items = [_item(f"S{i}", f"E{i}", 1) for i in range(cap * 2 + 5)]
    pkg = _pkg(1, 1, items)
    carton = pse.build_sublist_carton_model(pkg)
    blocks = ppe._paginate_for_pdf([carton])
    assert len(blocks) == 3
    identities = {b.carton.carton_identity for b in blocks}
    assert len(identities) == 1, "3 pages must still be recognised as ONE carton"


test("<=capacity items -> single page, no continuation", t_paginate_under_capacity_single_block)
test("exactly at ITEMS_PER_PDF_PAGE -> still a single page", t_paginate_exact_capacity_single_block)
test(">capacity items -> splits across pages, zero items lost, subtotal then grand total", t_paginate_over_capacity_splits_no_item_lost)
test("continuation across pages never inflates the unique-carton identity", t_paginate_continuation_never_inflates_carton_identity)


# =============================================================================
# 3. Full PDF generation + content verification (via pdfplumber readback)
# =============================================================================
print("\n== generate_sublist_pdf() integration tests ==")

try:
    import pdfplumber
    HAVE_PDFPLUMBER = True
except ImportError:
    HAVE_PDFPLUMBER = False


def t_generate_pdf_success_basic_fields():
    items = [_item("TP-A-1", "4894961069222", 10), _item("TP-A-2", "4895227935312", 15)]
    pkg = _pkg(1, 1, items, pl_gross_weight="35.68 KG")
    with tempfile.TemporaryDirectory() as td:
        out = Path(td) / "SUBLIST_TOTAL.pdf"
        result = ppe.generate_sublist_pdf([pkg], out)
        assert result.status == "SUCCESS", result.error
        assert out.exists()
        assert result.pages_written == 1
        assert result.cartons_written == 1
        assert result.items_written == 2
        if HAVE_PDFPLUMBER:
            with pdfplumber.open(str(out)) as pdf:
                assert len(pdf.pages) == 1
                text = pdf.pages[0].extract_text() or ""
                for expected in ("Carton #", "1/1", "Shipping Mark", "CN-1666-PVG-KERRY-POP",
                                  "OR #", "OR1016", "SO Order #", "so402064", "GW", "35.68 KG",
                                  "Packing Code #", "PGKECO377R7J0320001",
                                  "Item No.", "EAN", "QTY", "TP-A-1", "4894961069222", "TOTAL QTY", "25"):
                    assert expected in text, f"expected {expected!r} in rendered PDF text, got:\n{text}"


def t_generate_pdf_disabled_writes_nothing():
    items = [_item("A", "B", 1)]
    pkg = _pkg(1, 1, items)
    with tempfile.TemporaryDirectory() as td:
        out = Path(td) / "SUBLIST_TOTAL.pdf"
        result = ppe.generate_sublist_pdf([pkg], out, enabled=False)
        assert result.status == "DISABLED"
        assert not out.exists()


def t_generate_pdf_never_raises_when_pl_sublist_export_missing():
    # Simulate the ImportError branch -- must return FAILED, never raise,
    # so a broken/optional dependency can NEVER break the legacy export
    # (spec: non-blocking failure requirement).
    real_module = sys.modules.pop("pl_sublist_export", None)
    blocked_path = [p for p in sys.path]
    try:
        # Make the module unimportable by temporarily shadowing it with a
        # broken stand-in on sys.modules (simplest reliable simulation
        # without touching the real file on disk).
        import types
        broken = types.ModuleType("pl_sublist_export")
        def _raise(*a, **k):
            raise ImportError("simulated missing dependency")
        # Force a fresh import attempt to fail by deleting the cached module
        # and injecting a finder that raises for this name specifically.
        class _Blocker:
            def find_spec(self, name, path, target=None):
                if name == "pl_sublist_export":
                    raise ImportError("simulated missing dependency")
                return None
        blocker = _Blocker()
        sys.meta_path.insert(0, blocker)
        try:
            items = [_item("A", "B", 1)]
            pkg = _pkg(1, 1, items)
            with tempfile.TemporaryDirectory() as td:
                out = Path(td) / "SUBLIST_TOTAL.pdf"
                result = ppe.generate_sublist_pdf([pkg], out)
                assert result.status == "FAILED"
                assert "pl_sublist_export" in result.error
                assert not out.exists()
        finally:
            sys.meta_path.remove(blocker)
    finally:
        if real_module is not None:
            sys.modules["pl_sublist_export"] = real_module


def t_generate_pdf_never_raises_on_bad_output_path():
    # An unwritable output path must produce status=FAILED, not an
    # exception escaping to the caller (legacy ZIP/export must still run).
    items = [_item("A", "B", 1)]
    pkg = _pkg(1, 1, items)
    # A path with a null byte is guaranteed invalid on every platform.
    bad_path = Path("/nonexistent_root_dir_xyz/\x00bad/SUBLIST_TOTAL.pdf")
    result = ppe.generate_sublist_pdf([pkg], bad_path)
    assert result.status == "FAILED"
    assert result.error


def t_validate_sublist_pdf_clean_on_success():
    items = [_item("A", "B", 1), _item("C", "D", 2)]
    pkg = _pkg(1, 1, items)
    with tempfile.TemporaryDirectory() as td:
        out = Path(td) / "SUBLIST_TOTAL.pdf"
        result = ppe.generate_sublist_pdf([pkg], out)
        problems = ppe.validate_sublist_pdf([pkg], result)
        assert problems == [], problems


def t_validate_sublist_pdf_flags_item_count_mismatch():
    result = ppe.SublistPdfBuildResult(status="SUCCESS", cartons_written=1, items_written=999)
    pkg = _pkg(1, 1, [_item("A", "B", 1)])
    problems = ppe.validate_sublist_pdf([pkg], result)
    assert any("items_written" in p for p in problems)


def t_validate_sublist_pdf_skips_check_when_not_success():
    result = ppe.SublistPdfBuildResult(status="FAILED", error="boom")
    problems = ppe.validate_sublist_pdf([_pkg(1, 1, [])], result)
    assert problems == []


def t_multi_carton_multi_page_zero_loss():
    pkgs = []
    for i in range(4):
        n_items = 3 if i % 2 == 0 else ppe.ITEMS_PER_PDF_PAGE + 7
        items = [_item(f"S{i}-{j}", f"E{i}-{j}", j + 1) for j in range(n_items)]
        pkgs.append(_pkg(i + 1, 4, items, package_code=f"PGKEC{i}"))
    with tempfile.TemporaryDirectory() as td:
        out = Path(td) / "SUBLIST_TOTAL.pdf"
        result = ppe.generate_sublist_pdf(pkgs, out)
        assert result.status == "SUCCESS"
        assert result.cartons_written == 4
        expected_items = sum(len(p.items) for p in pkgs)
        assert result.items_written == expected_items
        problems = ppe.validate_sublist_pdf(pkgs, result)
        assert problems == [], problems
        if HAVE_PDFPLUMBER:
            with pdfplumber.open(str(out)) as pdf:
                assert len(pdf.pages) == result.pages_written


test("generate_sublist_pdf: SUCCESS, correct fields rendered into the PDF text", t_generate_pdf_success_basic_fields)
test("generate_sublist_pdf: enabled=False -> DISABLED, no file written", t_generate_pdf_disabled_writes_nothing)
test("generate_sublist_pdf: missing pl_sublist_export dependency -> FAILED, never raises", t_generate_pdf_never_raises_when_pl_sublist_export_missing)
test("generate_sublist_pdf: bad output path -> FAILED, never raises", t_generate_pdf_never_raises_on_bad_output_path)
test("validate_sublist_pdf: no problems reported on a clean successful build", t_validate_sublist_pdf_clean_on_success)
test("validate_sublist_pdf: flags an items_written mismatch", t_validate_sublist_pdf_flags_item_count_mismatch)
test("validate_sublist_pdf: skips reconciliation when status != SUCCESS", t_validate_sublist_pdf_skips_check_when_not_success)
test("multi-carton batch (mixed normal + oversized cartons): zero item loss, page count matches", t_multi_carton_multi_page_zero_loss)


# ── summary ──────────────────────────────────────────────────────────────
print(f"\n{_passed} passed, {_failed} failed")
if _failed:
    print("FAILED:", ", ".join(_failures))
    sys.exit(1)
sys.exit(0)
