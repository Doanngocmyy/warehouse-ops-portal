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
# 1b. Visual-layout geometry regression tests (spec: metadata block moved
#     to the upper-right, Total QTY placed dynamically after the last
#     item row -- both caught during visual review of the first preview).
# =============================================================================
print("\n== Visual-layout geometry regression tests ==")


def t_metadata_block_right_edge_is_in_the_right_half_of_the_page():
    value_x, label_right_x = ppe._compute_metadata_x(
        ["1/6", "CN-1666-PVG-KERRY-POP", "OR1016", "so402064", "35.68 KG", "PGKECO377R7J0320001"])
    widest = "CN-1666-PVG-KERRY-POP"
    from reportlab.pdfbase import pdfmetrics
    value_width = pdfmetrics.stringWidth(widest, ppe.FONT_VALUE, ppe.SIZE_VALUE)
    block_right_edge = value_x + value_width
    assert block_right_edge <= ppe.CONTENT_RIGHT + 0.5, "value text must never overlap the right margin"
    assert block_right_edge > ppe.CONTENT_RIGHT - 5, "block right edge should sit close to the right margin"


def t_metadata_center_x_is_right_of_page_center():
    value_x, label_right_x = ppe._compute_metadata_x(["1/4", "CN-1667-PVG-HANGZHOU", "OR2044", "so402064", "22.10 KG", "PGKECH0"])
    from reportlab.pdfbase import pdfmetrics
    value_width = pdfmetrics.stringWidth("CN-1667-PVG-HANGZHOU", ppe.FONT_VALUE, ppe.SIZE_VALUE)
    label_width = pdfmetrics.stringWidth("Packing Code #", ppe.FONT_LABEL, ppe.SIZE_LABEL)
    block_left_edge = label_right_x - label_width
    block_right_edge = value_x + value_width
    block_center_x = (block_left_edge + block_right_edge) / 2
    page_center_x = ppe.PAGE_W / 2
    assert block_center_x > page_center_x, \
        f"metadata block center ({block_center_x:.1f}) must be right of page center ({page_center_x:.1f})"


def t_label_value_gap_is_within_the_compact_threshold():
    value_x, label_right_x = ppe._compute_metadata_x(["1/6", "short", "OR1", "so1", "1 KG", "PGKEC1"])
    gap = value_x - label_right_x
    # spec: "approximately 2-4mm only" -- allow a little slack either side.
    gap_mm = gap * 25.4 / 72
    assert 1.5 <= gap_mm <= 5.0, f"label/value gap is {gap_mm:.2f}mm, expected roughly 2-4mm"


def t_metadata_values_never_overlap_right_margin_even_for_a_very_long_value():
    # A value longer than VALUE_COL_MAX_WIDTH now WRAPS instead of
    # growing the column past that cap (business owner's follow-up
    # request: "neu dai qua ky tu thi phai tu wrap text") -- so what must
    # never overlap the margin is each individual WRAPPED LINE, not the
    # raw unwrapped string's own full width.
    long_packing_code = "PGKECO377R7J0320001-EXTRA-LONG-SUFFIX-FOR-STRESS-TEST"
    value_x, label_right_x = ppe._compute_metadata_x(
        ["1/6", "CN-1666-PVG-KERRY-POP", "OR1016", "so402064", "35.68 KG", long_packing_code])
    from reportlab.pdfbase import pdfmetrics
    raw_width = pdfmetrics.stringWidth(long_packing_code, ppe.FONT_VALUE, ppe.SIZE_VALUE)
    assert raw_width > ppe.VALUE_COL_MAX_WIDTH, "test setup: this value should actually need wrapping"
    wrapped_lines = ppe._wrap_text_to_width(long_packing_code, ppe.FONT_VALUE, ppe.SIZE_VALUE, ppe.VALUE_COL_MAX_WIDTH)
    assert len(wrapped_lines) > 1, "an over-cap value must actually wrap into multiple lines"
    for line in wrapped_lines:
        line_width = pdfmetrics.stringWidth(line, ppe.FONT_VALUE, ppe.SIZE_VALUE)
        assert value_x + line_width <= ppe.CONTENT_RIGHT + 0.5, \
            f"wrapped line {line!r} must not be pushed past the right margin"
        assert line_width <= ppe.VALUE_COL_MAX_WIDTH + 0.5


def t_metadata_block_moves_as_one_whole_block_not_just_values():
    # Regression for the exact bug flagged in review: label and value must
    # move TOGETHER -- the gap between them must stay constant regardless
    # of which carton's data is being rendered (short vs. long values).
    v1, l1 = ppe._compute_metadata_x(["1", "A", "B", "C", "D", "E"])
    v2, l2 = ppe._compute_metadata_x(["999999", "A-MUCH-LONGER-VALUE-STRING", "B", "C", "D", "E"])
    assert abs((v1 - l1) - (v2 - l2)) < 0.01, "label/value gap must be identical regardless of content length"


test("metadata block right edge sits near the right margin, never overlapping it", t_metadata_block_right_edge_is_in_the_right_half_of_the_page)
test("metadata block center X is right of page center X (upper-right placement)", t_metadata_center_x_is_right_of_page_center)
test("label/value gap stays within the ~2-4mm compact threshold", t_label_value_gap_is_within_the_compact_threshold)
test("even an unusually long value never overlaps the right margin", t_metadata_values_never_overlap_right_margin_even_for_a_very_long_value)
test("metadata block moves as ONE whole block (constant label/value gap regardless of content)", t_metadata_block_moves_as_one_whole_block_not_just_values)


def t_dynamic_total_5_items_immediately_follows_last_row():
    rule_y, text_y = ppe._compute_total_y(ppe.ITEM_TABLE_START_Y, 5)
    expected_last_item_bottom = ppe.ITEM_TABLE_START_Y - 5 * ppe.ROW_HEIGHT_ITEM
    assert abs(rule_y - (expected_last_item_bottom - ppe.GAP_AFTER_ITEMS)) < 0.01
    gap_between_last_item_and_rule = expected_last_item_bottom - rule_y
    assert 0 < gap_between_last_item_and_rule <= ppe.ROW_HEIGHT_ITEM, \
        "gap between last SKU row and the total rule must be small, not a large blank space"


def t_dynamic_total_12_items_immediately_follows_last_row():
    rule_y, text_y = ppe._compute_total_y(ppe.ITEM_TABLE_START_Y, 12)
    expected_last_item_bottom = ppe.ITEM_TABLE_START_Y - 12 * ppe.ROW_HEIGHT_ITEM
    assert 0 < (expected_last_item_bottom - rule_y) <= ppe.ROW_HEIGHT_ITEM


def t_dynamic_total_26_items_immediately_follows_last_row():
    rule_y, text_y = ppe._compute_total_y(ppe.ITEM_TABLE_START_Y, 26)
    expected_last_item_bottom = ppe.ITEM_TABLE_START_Y - 26 * ppe.ROW_HEIGHT_ITEM
    assert 0 < (expected_last_item_bottom - rule_y) <= ppe.ROW_HEIGHT_ITEM
    # A full-capacity page's dynamic total must still land above the
    # bottom margin -- this is the actual overflow-safety guarantee.
    assert text_y - ppe.TOTAL_TEXT_PADDING >= ppe.MARGIN - 5


def t_dynamic_total_gap_is_small_and_fixed_regardless_of_item_count():
    # The gap between the last item row and the rule must be the SAME
    # small fixed distance whether there are 3 items or 26 -- proving the
    # total is not creeping toward some other anchor as item count changes.
    for n in (1, 3, 5, 12, 20, 26):
        rule_y, _ = ppe._compute_total_y(ppe.ITEM_TABLE_START_Y, n)
        last_item_bottom = ppe.ITEM_TABLE_START_Y - n * ppe.ROW_HEIGHT_ITEM
        gap = last_item_bottom - rule_y
        assert abs(gap - ppe.GAP_AFTER_ITEMS) < 0.01, f"n={n}: gap={gap}, expected exactly GAP_AFTER_ITEMS"


def t_total_qty_is_not_anchored_to_bottom_margin():
    # For a lightly-filled carton (5 items), the total must sit FAR above
    # the bottom margin, not pinned near it -- proving it's not anchored
    # to a fixed bottom-page Y coordinate.
    rule_y, text_y = ppe._compute_total_y(ppe.ITEM_TABLE_START_Y, 5)
    distance_from_bottom_margin = text_y - ppe.MARGIN
    distance_from_top = ppe.ITEM_TABLE_START_Y - text_y
    assert distance_from_bottom_margin > 300, \
        "a 5-item carton's total should be nowhere near the bottom margin"
    assert distance_from_top < 150, "a 5-item carton's total should sit close to the top, right after its few items"


test("dynamic total: 5-item carton -> Total QTY immediately follows item 5", t_dynamic_total_5_items_immediately_follows_last_row)
test("dynamic total: 12-item carton -> Total QTY immediately follows item 12", t_dynamic_total_12_items_immediately_follows_last_row)
test("dynamic total: 26-item (full page) carton -> Total QTY immediately follows item 26, still above margin", t_dynamic_total_26_items_immediately_follows_last_row)
test("dynamic total: last-row-to-rule gap is small and identical regardless of item count", t_dynamic_total_gap_is_small_and_fixed_regardless_of_item_count)
test("dynamic total: NOT anchored to the bottom page margin (verified via a lightly-filled carton)", t_total_qty_is_not_anchored_to_bottom_margin)


# =============================================================================
# 1c. Item table GRID (borders + header fill) + brand LOGO regression
#     tests -- per business owner's follow-up review with a real reference
#     screenshot: the table must have visible borders (not a bare
#     underline), a table is only drawn when there is at least 1 SKU row,
#     and a small "topologie" logo sits top-center on every page.
# =============================================================================
print("\n== Item table grid + logo regression tests ==")

import pdfplumber as _pdfplumber


def t_grid_is_drawn_when_items_present():
    items = [_item(f"S{i}", f"E{i}", i + 1) for i in range(5)]
    pkg = _pkg(1, 1, items)
    with tempfile.TemporaryDirectory() as td:
        out = Path(td) / "SUBLIST_TOTAL.pdf"
        result = ppe.generate_sublist_pdf([pkg], out)
        assert result.status == "SUCCESS", result.error
        with _pdfplumber.open(str(out)) as pdf:
            page = pdf.pages[0]
            assert len(page.rects) >= 2, "expected at least the header fill + outer border rects"
            # header/body separator + (n-1) row separators + 2 column
            # separators + 1 total rule = n + 3 lines for n items.
            assert len(page.lines) == 5 + 3, f"expected {5 + 3} grid lines for 5 items, got {len(page.lines)}"


def t_grid_is_skipped_entirely_when_no_items():
    pkg = _pkg(1, 1, [])
    with tempfile.TemporaryDirectory() as td:
        out = Path(td) / "SUBLIST_TOTAL.pdf"
        result = ppe.generate_sublist_pdf([pkg], out)
        assert result.status == "SUCCESS", result.error
        with _pdfplumber.open(str(out)) as pdf:
            page = pdf.pages[0]
            assert len(page.rects) == 0, "no table grid (fill/border rects) should be drawn for zero items"
            assert len(page.lines) == 1, "only the Total QTY rule line should be drawn, no header/row separators"
            text = page.extract_text() or ""
            assert "TOTAL QTY" in text and "Item No." not in text, \
                "empty carton must show Total QTY directly, with no Item No./EAN/QTY header at all"


def t_grid_row_count_matches_item_count_for_various_sizes():
    for n in (1, 3, 12, 18):
        items = [_item(f"S{i}", f"E{i}", 1) for i in range(n)]
        pkg = _pkg(1, 1, items)
        with tempfile.TemporaryDirectory() as td:
            out = Path(td) / "SUBLIST_TOTAL.pdf"
            result = ppe.generate_sublist_pdf([pkg], out)
            assert result.status == "SUCCESS", result.error
            with _pdfplumber.open(str(out)) as pdf:
                page = pdf.pages[0]
                assert len(page.lines) == n + 3, f"n={n}: expected {n + 3} lines, got {len(page.lines)}"


def t_logo_is_drawn_by_default_on_every_page():
    items = [_item(f"S{i}", f"E{i}", 1) for i in range(3)]
    pkgs = [_pkg(1, 2, items), _pkg(2, 2, items)]
    with tempfile.TemporaryDirectory() as td:
        out = Path(td) / "SUBLIST_TOTAL.pdf"
        result = ppe.generate_sublist_pdf(pkgs, out)
        assert result.status == "SUCCESS", result.error
        with _pdfplumber.open(str(out)) as pdf:
            assert len(pdf.pages) == 2
            for page in pdf.pages:
                assert len(page.images) == 1, "logo should be drawn once per page by default"


def t_logo_can_be_explicitly_suppressed():
    items = [_item("S", "E", 1)]
    pkg = _pkg(1, 1, items)
    with tempfile.TemporaryDirectory() as td:
        out = Path(td) / "SUBLIST_TOTAL.pdf"
        result = ppe.generate_sublist_pdf([pkg], out, logo_path=False)
        assert result.status == "SUCCESS", result.error
        with _pdfplumber.open(str(out)) as pdf:
            assert len(pdf.pages[0].images) == 0


def t_missing_logo_file_does_not_break_generation():
    items = [_item("S", "E", 1)]
    pkg = _pkg(1, 1, items)
    with tempfile.TemporaryDirectory() as td:
        out = Path(td) / "SUBLIST_TOTAL.pdf"
        result = ppe.generate_sublist_pdf([pkg], out, logo_path=Path("/nonexistent/logo.png"))
        assert result.status == "SUCCESS", "a missing logo file must never fail PDF generation"
        with _pdfplumber.open(str(out)) as pdf:
            assert len(pdf.pages[0].images) == 0


def t_default_logo_asset_exists_in_the_repo():
    assert ppe.DEFAULT_LOGO_PATH.exists(), \
        f"expected the bundled logo asset at {ppe.DEFAULT_LOGO_PATH}"


def t_long_shipping_mark_wraps_onto_multiple_lines_in_the_rendered_pdf():
    # Business owner's follow-up: "doi voi shipmark neu dai qua ky tu thi
    # phai tu wrap text" -- a Shipping Mark wider than VALUE_COL_MAX_WIDTH
    # must actually wrap in the real rendered PDF, not just in the pure
    # _wrap_text_to_width() unit test above.
    long_mark = "CN-1666-PVG-KERRY-POP-EXTRA-LONG-SHIPPING-MARK-SUFFIX-FOR-WRAP-TEST"
    items = [_item("A", "B", 1)]
    pkg = _pkg(1, 1, items, shipping_mark=long_mark)
    with tempfile.TemporaryDirectory() as td:
        out = Path(td) / "SUBLIST_TOTAL.pdf"
        result = ppe.generate_sublist_pdf([pkg], out)
        assert result.status == "SUCCESS", result.error
        with _pdfplumber.open(str(out)) as pdf:
            text = pdf.pages[0].extract_text() or ""
            # every fragment of the wrapped value must still be present,
            # even though it's now spread across multiple lines
            for fragment in ["CN-1666-PVG-KERRY-POP", "SUFFIX-FOR-WRAP-TEST"]:
                assert fragment in text.replace("\n", ""), f"{fragment!r} missing from wrapped output"


def t_wrapped_shipping_mark_pushes_the_item_table_down_without_overlap():
    # The item table must start LOWER on a page whose Shipping Mark
    # wrapped to 2 lines than on an otherwise-identical page with a short
    # Shipping Mark -- proving the extra line actually reserves real
    # vertical space rather than being drawn on top of the next row.
    short_items = [_item("A", "B", 1)]
    pkg_short = _pkg(1, 1, short_items, shipping_mark="SHORT")
    long_mark = "CN-1666-PVG-KERRY-POP-EXTRA-LONG-SHIPPING-MARK-SUFFIX-FOR-WRAP-TEST"
    pkg_long = _pkg(1, 1, short_items, shipping_mark=long_mark)
    with tempfile.TemporaryDirectory() as td:
        out_short = Path(td) / "short.pdf"
        out_long = Path(td) / "long.pdf"
        r_short = ppe.generate_sublist_pdf([pkg_short], out_short)
        r_long = ppe.generate_sublist_pdf([pkg_long], out_long)
        assert r_short.status == "SUCCESS" and r_long.status == "SUCCESS"
        with _pdfplumber.open(str(out_short)) as pdf_s, _pdfplumber.open(str(out_long)) as pdf_l:
            # the item-table outer rect's "top" (pdfplumber measures from
            # the page's TOP in image-style coordinates) must be LOWER
            # (larger `top` value) on the wrapped page.
            rect_top_short = min(r["top"] for r in pdf_s.pages[0].rects)
            rect_top_long = min(r["top"] for r in pdf_l.pages[0].rects)
            assert rect_top_long > rect_top_short, \
                "item table must start further down the page when Shipping Mark wraps to 2 lines"


def t_continuation_marker_shown_top_left_only_when_carton_spans_multiple_pages():
    cap = ppe.ITEMS_PER_PDF_PAGE
    items = [_item(f"S{i}", f"E{i}", 1) for i in range(cap + 5)]
    pkg = _pkg(2, 6, items)
    with tempfile.TemporaryDirectory() as td:
        out = Path(td) / "SUBLIST_TOTAL.pdf"
        result = ppe.generate_sublist_pdf([pkg], out)
        assert result.status == "SUCCESS", result.error
        assert result.pages_written == 2
        with _pdfplumber.open(str(out)) as pdf:
            page1_text = pdf.pages[0].extract_text() or ""
            page2_text = pdf.pages[1].extract_text() or ""
            assert "1/2" in page1_text, "top-left page marker '1/2' missing on the first continuation page"
            assert "2/2" in page2_text, "top-left page marker '2/2' missing on the second (last) page"


def t_no_continuation_marker_for_a_single_page_carton():
    items = [_item("A", "B", 1)]
    pkg = _pkg(1, 1, items)
    with tempfile.TemporaryDirectory() as td:
        out = Path(td) / "SUBLIST_TOTAL.pdf"
        result = ppe.generate_sublist_pdf([pkg], out)
        assert result.status == "SUCCESS", result.error
        assert result.pages_written == 1
        with _pdfplumber.open(str(out)) as pdf:
            text = pdf.pages[0].extract_text() or ""
            assert "1/1" not in text.replace("Carton # 1/1", "").replace("1/1", "", 1) or True
            # simpler/robust check: the dedicated marker font/size call was
            # never even invoked (block_count == 1) -- verified indirectly
            # via page.chars not containing an isolated bold 12pt "1/1" at
            # the marker's own top-left coordinate band.
            marker_band_chars = [ch for ch in pdf.pages[0].chars
                                  if ch["top"] < ppe.CONTINUATION_MARKER_TOP_GAP + 4
                                  and ch["x0"] < ppe.CONTENT_LEFT + 30]
            assert not marker_band_chars, "no top-left marker text should be drawn for a single-page carton"


test("a long Shipping Mark actually wraps onto multiple lines in the rendered PDF", t_long_shipping_mark_wraps_onto_multiple_lines_in_the_rendered_pdf)
test("a wrapped Shipping Mark pushes the item table down (no overlap between wrapped lines and the table)", t_wrapped_shipping_mark_pushes_the_item_table_down_without_overlap)
test("top-left 'N/M' continuation marker shown on both pages of a 2-page carton", t_continuation_marker_shown_top_left_only_when_carton_spans_multiple_pages)
test("no top-left continuation marker drawn for a normal single-page carton", t_no_continuation_marker_for_a_single_page_carton)


test("item table grid (borders + header fill) is drawn when items are present", t_grid_is_drawn_when_items_present)
test("item table grid is skipped ENTIRELY when a carton has zero items (total shown directly below metadata)", t_grid_is_skipped_entirely_when_no_items)
test("grid line count matches item count for several carton sizes (1/3/12/18 items)", t_grid_row_count_matches_item_count_for_various_sizes)
test("logo is drawn once per page by default (multi-page batch)", t_logo_is_drawn_by_default_on_every_page)
test("logo can be explicitly suppressed via logo_path=False", t_logo_can_be_explicitly_suppressed)
test("a missing/invalid logo file never breaks PDF generation (non-blocking)", t_missing_logo_file_does_not_break_generation)
test("the bundled default logo asset exists in the repo", t_default_logo_asset_exists_in_the_repo)


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
    # Carton # must read IDENTICALLY on every page of the same carton --
    # continuation is conveyed by the top-left marker/bottom note only.
    assert blocks[1].block_carton_label == "2/6"
    assert blocks[0].block_carton_label == blocks[1].block_carton_label == carton.carton_display


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


def t_continuation_marker_smaller_than_carton_hash_label():
    # Turn 9 requirement: the top-left "N/M" page marker must be visually
    # SMALLER than "Carton #" so it can never be mistaken for the Store's
    # own carton number (e.g. "1/6") shown in the metadata block.
    assert ppe.CONTINUATION_MARKER_SIZE < ppe.SIZE_LABEL, (
        f"marker size {ppe.CONTINUATION_MARKER_SIZE} must be < Carton # label size {ppe.SIZE_LABEL}"
    )
    assert ppe.CONTINUATION_MARKER_SIZE < ppe.SIZE_VALUE, (
        f"marker size {ppe.CONTINUATION_MARKER_SIZE} must be < Carton # value size {ppe.SIZE_VALUE}"
    )
    # Distinct font style (oblique) from the bold label / regular value,
    # so it reads as an annotation rather than a data field.
    assert ppe.CONTINUATION_MARKER_FONT != ppe.FONT_LABEL
    assert ppe.CONTINUATION_MARKER_FONT != ppe.FONT_VALUE


def t_carton_hash_label_unchanged_across_all_continuation_pages():
    cap = ppe.ITEMS_PER_PDF_PAGE
    items = [_item(f"S{i}", f"E{i}", 1) for i in range(cap * 2 + 5)]
    pkg = _pkg(1, 1, items)
    carton = pse.build_sublist_carton_model(pkg)
    blocks = ppe._paginate_for_pdf([carton])
    assert len(blocks) == 3
    labels = {b.block_carton_label for b in blocks}
    assert labels == {carton.carton_display}, (
        f"Carton # must be identical on every page, got {labels}"
    )


test("continuation marker is visually smaller than Carton # (size + font style)", t_continuation_marker_smaller_than_carton_hash_label)
test("Carton # label stays unchanged across all continuation pages", t_carton_hash_label_unchanged_across_all_continuation_pages)


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
