#!/usr/bin/env python3
"""
pl_sublist_pdf_export.py
=========================
Generates the mandatory A5 "Sublist" PDF (BUILD REQUEST spec, section 8):
one page per carton block, A5 portrait, no watermark/title/app branding.
Output: PL_SPLIT_OUTPUT/05_SUBLIST/SUBLIST_TOTAL.pdf.

Font
----
Calibri is NOT bundled with this repo (explicit instruction: "Do not add
or distribute font files" -- shipping a .ttf would make this an
undistributable/licensing liability, and Pyodide/browser execution has no
access to the user's locally-installed fonts anyway). Helvetica is used
instead: it is one of reportlab's built-in "Base 14" PDF standard fonts,
requires NO font file to be embedded or shipped (every PDF reader/printer
already has it), and is the closest metrically-compatible sans-serif
available under that constraint (Arial/Calibri and Helvetica share very
similar glyph widths, which is exactly why Helvetica is the traditional
"no-Arial-available" substitute). This is a documented substitution, not a
silent one -- flag to the user if a different fallback is preferred once
Calibri itself is verified unavailable in the actual deployment target.

Field source / never-overwritten rules (spec section 6) are NOT re-derived
here -- this module reuses pl_sublist_export.build_sublist_carton_model()
(dynamic import, matches this repo's existing self-contained-module
pattern) for the Package -> SublistCartonModel conversion, so the GW
priority (PL text over DIM weight), OR != Shipmark separation, and
Packing Code resolution logic all have exactly ONE implementation shared
by both the Excel and PDF outputs -- never two copies that could drift
apart (the lesson from this session's factory-detection and pending-
metadata bugs: one source of truth, not parallel logic in two places).

Pagination (one carton per page, continuation for oversized cartons) IS
reimplemented here rather than reused from pl_sublist_export.
paginate_carton_blocks(), because that function's item-per-block capacity
(18) is derived from the Excel template's own row/border geometry and has
nothing to do with how much vertical space an A5 PDF page actually has --
reusing it would either leave an A5 page mostly empty or overflow it.
_paginate_for_pdf() below uses ITEMS_PER_PDF_PAGE, computed from this
module's own A5 page geometry (see the "Page geometry" section).
"""
from __future__ import annotations

import logging
from dataclasses import dataclass
from pathlib import Path
from typing import List, Optional

from reportlab.lib.pagesizes import A5
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfgen import canvas

log = logging.getLogger("pl_sublist_pdf_export")
if not log.handlers:
    logging.basicConfig(level=logging.INFO, format="%(asctime)s [%(levelname)s] %(message)s")

# =========================================================================
# 1) Font (see module docstring: Helvetica, no font file shipped/embedded)
# =========================================================================
FONT_LABEL = "Helvetica-Bold"
FONT_VALUE = "Helvetica"
FONT_ITEM_HEADER = "Helvetica-Bold"
FONT_ITEM = "Helvetica"
FONT_TOTAL = "Helvetica-Bold"
FONT_CONTINUED = "Helvetica-Oblique"

SIZE_LABEL = 10
SIZE_VALUE = 10
SIZE_ITEM_HEADER = 9
SIZE_ITEM = 9
SIZE_TOTAL = 10
SIZE_CONTINUED = 8

# =========================================================================
# 2) Page geometry -- A5 portrait, no title/watermark/branding anywhere.
# =========================================================================
PAGE_W, PAGE_H = A5  # (419.53, 595.28) pt -- already portrait (H > W)
MARGIN = 28

CONTENT_LEFT = MARGIN
CONTENT_RIGHT = PAGE_W - MARGIN
CONTENT_TOP = PAGE_H - MARGIN
CONTENT_WIDTH = CONTENT_RIGHT - CONTENT_LEFT

# Metadata block: label (right-aligned) + value (left-aligned), TIGHTLY
# adjacent, positioned in the page's UPPER-RIGHT (revision per visual
# review: the first preview had this block left-biased -- corrected here).
# The WHOLE two-column block moves together -- only its horizontal anchor
# (metadata_right_x, near CONTENT_RIGHT) changes; label/value stay exactly
# as tightly adjacent as before (small fixed gap, not stretched apart).
#
# Both column widths are measured from the ACTUAL text at draw time (see
# _compute_metadata_x() below), not hardcoded, so the block's right edge
# always lands near the page's right margin regardless of how long
# "Shipping Mark" / "Packing Code #" values happen to be on a given
# carton, and short values (e.g. "1/6") never get stranded out at the
# margin with a huge gap back to their label.
ROW_HEIGHT_META = 17
VALUE_LABEL_GAP_MM = 3  # spec: "approximately 2-4 mm only"
VALUE_LABEL_GAP = VALUE_LABEL_GAP_MM * 72 / 25.4  # ~8.5pt
MIN_VALUE_COL_WIDTH = 30  # floor so an all-blank metadata row can't collapse the column to 0

METADATA_ROWS = [
    ("carton", "Carton #"),
    ("shipping_mark", "Shipping Mark"),
    ("or", "OR #"),
    ("so", "SO Order #"),
    ("gw", "GW"),
    ("packing_code", "Packing Code #"),
]
META_BLOCK_HEIGHT = len(METADATA_ROWS) * ROW_HEIGHT_META

GAP_AFTER_METADATA = 14


def _compute_metadata_x(value_texts) -> "tuple[float, float]":
    """-> (metadata_value_x, metadata_label_right_x), both measured from
    the actual value text widths so the block's right edge sits near
    CONTENT_RIGHT (the page's right margin) and the label/value gap stays
    a small fixed distance regardless of content length. Pure function --
    no canvas needed (pdfmetrics.stringWidth uses the font's built-in
    AFM/glyph-width table), so this is directly unit-testable.

        metadata_right_x     = CONTENT_RIGHT               (page_width - right_margin)
        metadata_value_x     = metadata_right_x - value_column_width
        metadata_label_right_x = metadata_value_x - VALUE_LABEL_GAP
    """
    widths = [pdfmetrics.stringWidth(str(v or ""), FONT_VALUE, SIZE_VALUE) for v in value_texts]
    value_column_width = max(widths + [MIN_VALUE_COL_WIDTH])
    metadata_value_x = CONTENT_RIGHT - value_column_width
    metadata_label_right_x = metadata_value_x - VALUE_LABEL_GAP
    return metadata_value_x, metadata_label_right_x

# Item table columns: Item No. (left) | EAN (left) | QTY (center).
QTY_COL_WIDTH = 52
_remaining = CONTENT_WIDTH - QTY_COL_WIDTH
ITEM_NO_COL_WIDTH = _remaining * 0.56
EAN_COL_WIDTH = _remaining * 0.44
COL_X_ITEM_NO = CONTENT_LEFT
COL_X_EAN = COL_X_ITEM_NO + ITEM_NO_COL_WIDTH
COL_X_QTY = COL_X_EAN + EAN_COL_WIDTH
COL_X_QTY_CENTER = COL_X_QTY + QTY_COL_WIDTH / 2

ROW_HEIGHT_ITEM_HEADER = 16
ROW_HEIGHT_ITEM = 14

# Total QTY row is now DYNAMIC -- placed immediately after the last
# rendered item row (revision per visual review: the first preview
# anchored it near the bottom margin, leaving a large empty gap on
# lightly-filled cartons). GAP_AFTER_ITEMS is the small fixed gap between
# the last item row and the rule above "TOTAL QTY"; TOTAL_TEXT_PADDING is
# the fixed clearance between that rule and the total text's own baseline
# (13pt -- this is the exact value that fixed the rule-strikes-through-
# text overlap bug found during the first visual-validation pass; kept
# unchanged here since that fix is independent of the total row's
# position now being dynamic instead of fixed).
GAP_AFTER_ITEMS = 8
TOTAL_TEXT_PADDING = 13
GAP_BEFORE_CONTINUED_NOTE = 4
ROW_HEIGHT_CONTINUED_NOTE = 12

# y-position of the FIRST item row (immediately below the header) -- fixed
# per page (metadata block height doesn't vary), items are what varies.
ITEM_TABLE_START_Y = CONTENT_TOP - META_BLOCK_HEIGHT - GAP_AFTER_METADATA - ROW_HEIGHT_ITEM_HEADER

# Fixed vertical budget that must be reserved BELOW the item rows for the
# total row + a possible "Continued" note, even though the total row's
# actual draw position is now dynamic (see _compute_total_y()) -- this is
# what still caps ITEMS_PER_PDF_PAGE so a full page's dynamic total can
# never be pushed below the bottom margin.
_RESERVED_BELOW_ITEMS = GAP_AFTER_ITEMS + TOTAL_TEXT_PADDING + GAP_BEFORE_CONTINUED_NOTE + ROW_HEIGHT_CONTINUED_NOTE
_AVAILABLE_FOR_ITEMS = (ITEM_TABLE_START_Y - MARGIN) - _RESERVED_BELOW_ITEMS
# Item capacity per A5 page, derived from actual page geometry above (not
# copied from the Excel template's 18 -- see module docstring). Documented
# here rather than silently picked: at MARGIN=28/ROW_HEIGHT_ITEM=14 this
# works out to ~24-26 items/page; verified against a rendered sample (see
# tests/test_pl_sublist_pdf_export.py's visual-validation step).
ITEMS_PER_PDF_PAGE = max(1, int(_AVAILABLE_FOR_ITEMS // ROW_HEIGHT_ITEM))


def _compute_total_y(item_table_start_y: float, rendered_item_count: int) -> "tuple[float, float]":
    """-> (total_rule_y, total_text_y). Pure function -- the exact formula
    requested during visual review:

        last_item_bottom_y = item_table_start_y - rendered_item_count * ROW_HEIGHT_ITEM
        total_rule_y        = last_item_bottom_y - GAP_AFTER_ITEMS
        total_text_y         = total_rule_y - TOTAL_TEXT_PADDING

    Deliberately does NOT reference MARGIN/CONTENT_TOP/page bottom at all
    -- the total's position is a function of how many items were actually
    drawn on THIS page, never a fixed bottom-page Y coordinate. Works
    identically for a normal page (5 items) and a continuation page (still
    counts only the items rendered on that specific page)."""
    last_item_bottom_y = item_table_start_y - rendered_item_count * ROW_HEIGHT_ITEM
    total_rule_y = last_item_bottom_y - GAP_AFTER_ITEMS
    total_text_y = total_rule_y - TOTAL_TEXT_PADDING
    return total_rule_y, total_text_y


# =========================================================================
# 3) Pagination (A5-specific -- see module docstring for why this is not
#    shared with pl_sublist_export.paginate_carton_blocks())
# =========================================================================
@dataclass
class PdfPageBlock:
    carton: "object"  # pl_sublist_export.SublistCartonModel (duck-typed here)
    block_index: int
    block_count: int
    items: list
    is_last_block: bool
    block_carton_label: str
    block_total_qty: int  # subtotal for non-last blocks, GRAND TOTAL on the last block


def _paginate_for_pdf(cartons: list) -> List[PdfPageBlock]:
    """One package == one carton, ALWAYS. A carton with more items than
    ITEMS_PER_PDF_PAGE fits is split across multiple pages ("Continued 1",
    "Continued 2", ...) -- zero items ever lost, subtotal shown on every
    non-last page, grand total only on the carton's last page."""
    blocks: List[PdfPageBlock] = []
    cap = ITEMS_PER_PDF_PAGE
    for carton in cartons:
        n = len(carton.items)
        chunks = [carton.items[i:i + cap] for i in range(0, n, cap)] or [[]]
        block_count = len(chunks)
        for idx, chunk in enumerate(chunks):
            is_last = idx == block_count - 1
            if block_count == 1:
                label = carton.carton_display
            elif idx == 0:
                label = carton.carton_display
            else:
                label = f"{carton.carton_display} - Continued {idx}"
            block_total = carton.total_qty if is_last else sum(r.qty for r in chunk)
            blocks.append(PdfPageBlock(
                carton=carton, block_index=idx, block_count=block_count,
                items=chunk, is_last_block=is_last,
                block_carton_label=label, block_total_qty=block_total,
            ))
    return blocks


# =========================================================================
# 4) Single-page drawing
# =========================================================================
def _draw_metadata_row(c: canvas.Canvas, y: float, label: str, value: str,
                        value_x: float, label_right_x: float):
    c.setFont(FONT_LABEL, SIZE_LABEL)
    c.drawRightString(label_right_x, y, label)
    c.setFont(FONT_VALUE, SIZE_VALUE)
    c.drawString(value_x, y, value or "")


def _draw_page(c: canvas.Canvas, block: PdfPageBlock):
    carton = block.carton
    y = CONTENT_TOP

    # -- Metadata block (Carton# / Shipping Mark / OR# / SO Order# / GW /
    #    Packing Code#) -- label right-aligned, value left-aligned, TIGHTLY
    #    adjacent, positioned in the page's UPPER-RIGHT (see
    #    _compute_metadata_x() docstring for the exact formula) -- the
    #    left side of the page stays visually open, matching the approved
    #    reference layout.
    meta_values = {
        "carton": block.block_carton_label,
        "shipping_mark": carton.shipping_mark,
        "or": carton.or_number,
        "so": carton.so_number,
        "gw": carton.gross_weight_display,
        "packing_code": carton.packing_code,
    }
    value_texts = [meta_values.get(key, "") for key, _label in METADATA_ROWS]
    metadata_value_x, metadata_label_right_x = _compute_metadata_x(value_texts)
    for key, label in METADATA_ROWS:
        y -= ROW_HEIGHT_META
        _draw_metadata_row(c, y, label, meta_values.get(key, ""), metadata_value_x, metadata_label_right_x)

    y -= GAP_AFTER_METADATA

    # -- Item table header --
    c.setFont(FONT_ITEM_HEADER, SIZE_ITEM_HEADER)
    c.drawString(COL_X_ITEM_NO, y, "Item No.")
    c.drawString(COL_X_EAN, y, "EAN")
    c.drawCentredString(COL_X_QTY_CENTER, y, "QTY")
    c.setLineWidth(0.75)
    c.line(CONTENT_LEFT, y - 3, CONTENT_RIGHT, y - 3)
    y -= ROW_HEIGHT_ITEM_HEADER
    assert abs(y - ITEM_TABLE_START_Y) < 0.01, "item_table_start_y drifted from the module constant"

    # -- Item rows: Item No. / EAN left-aligned, QTY centered --
    c.setFont(FONT_ITEM, SIZE_ITEM)
    for row in block.items:
        c.drawString(COL_X_ITEM_NO, y, str(row.item_no))
        c.drawString(COL_X_EAN, y, str(row.ean))
        c.drawCentredString(COL_X_QTY_CENTER, y, str(row.qty))
        y -= ROW_HEIGHT_ITEM

    # -- Total QTY: placed DYNAMICALLY immediately after the last rendered
    #    item row (see _compute_total_y()) -- never anchored to the page
    #    bottom. Subtotal on a non-last continuation page, GRAND TOTAL
    #    (the full carton's total_qty) only on the carton's last page. --
    total_rule_y, total_text_y = _compute_total_y(ITEM_TABLE_START_Y, len(block.items))
    c.setLineWidth(0.75)
    c.line(CONTENT_LEFT, total_rule_y, CONTENT_RIGHT, total_rule_y)
    c.setFont(FONT_TOTAL, SIZE_TOTAL)
    total_label = "TOTAL QTY" if block.is_last_block else "SUBTOTAL QTY"
    c.drawRightString(metadata_label_right_x, total_text_y, total_label)
    c.drawString(metadata_value_x, total_text_y, str(block.block_total_qty))

    # -- Continuation note (only on non-last blocks of a multi-page carton),
    #    also placed dynamically right after the total row it follows. --
    if block.block_count > 1:
        note_y = total_text_y - (ROW_HEIGHT_CONTINUED_NOTE + GAP_BEFORE_CONTINUED_NOTE - 4)
        c.setFont(FONT_CONTINUED, SIZE_CONTINUED)
        note = (f"Continued on next page ({block.block_index + 1}/{block.block_count})"
                if not block.is_last_block else
                f"Continued from previous page ({block.block_index + 1}/{block.block_count})")
        c.drawString(CONTENT_LEFT, note_y, note)


# =========================================================================
# 5) Build result + non-blocking status (spec: an optional PDF must NEVER
#    break the legacy ZIP/export -- SUCCESS/FAILED/DISABLED, never a raise
#    that escapes this module).
# =========================================================================
@dataclass
class SublistPdfBuildResult:
    status: str = "DISABLED"  # SUCCESS | FAILED | DISABLED
    output_path: Optional[Path] = None
    cartons_written: int = 0
    pages_written: int = 0
    items_written: int = 0
    error: str = ""
    warnings: List[str] = None

    def __post_init__(self):
        if self.warnings is None:
            self.warnings = []


def generate_sublist_pdf(packages: list, output_path: Path, *,
                          carton_display_mode: str = "current_total",
                          enabled: bool = True) -> SublistPdfBuildResult:
    """packages -> A5 PDF at `output_path`, one page per carton block.
    NEVER raises -- any failure is captured into the returned result's
    status="FAILED" (spec: Sublist PDF is optional, its failure must never
    break the legacy ZIP/export). Mirrors generate_sublist_workbook()'s
    signature intentionally so app.html/run_pipeline call both the same
    way."""
    result = SublistPdfBuildResult()
    if not enabled:
        result.status = "DISABLED"
        return result
    try:
        import pl_sublist_export as pse
    except ImportError as e:
        result.status = "FAILED"
        result.error = f"pl_sublist_export not importable: {e}"
        log.warning(f"Sublist PDF generation skipped -- {result.error}")
        return result

    try:
        output_path = Path(output_path)
        cartons = [pse.build_sublist_carton_model(pkg, carton_display_mode=carton_display_mode)
                   for pkg in packages]
        blocks = _paginate_for_pdf(cartons)

        output_path.parent.mkdir(parents=True, exist_ok=True)
        tmp_path = output_path.with_suffix(output_path.suffix + ".tmp")
        c = canvas.Canvas(str(tmp_path), pagesize=A5)
        # No document metadata that could act as a de facto "title" in a PDF
        # viewer's title bar (spec: no title/watermark/branding anywhere).
        c.setTitle("")
        c.setAuthor("")
        c.setSubject("")
        c.setCreator("")

        for block in blocks:
            _draw_page(c, block)
            c.showPage()
        c.save()

        if tmp_path.exists():
            if output_path.exists():
                output_path.unlink()
            tmp_path.replace(output_path)

        result.status = "SUCCESS"
        result.output_path = output_path
        result.cartons_written = len(cartons)
        result.pages_written = len(blocks)
        result.items_written = sum(len(b.items) for b in blocks)
        log.info(f"Sublist PDF: {result.cartons_written} carton(s) -> "
                 f"{result.pages_written} page(s), {result.items_written} item row(s) -> {output_path}")
        return result
    except Exception as e:
        result.status = "FAILED"
        result.error = f"{type(e).__name__}: {e}"
        log.warning(f"Sublist PDF generation FAILED (non-blocking, legacy export unaffected): {result.error}")
        return result


# =========================================================================
# 6) Reconciliation check (mirrors validate_sublist() in pl_sublist_export)
# =========================================================================
def validate_sublist_pdf(packages: list, result: SublistPdfBuildResult) -> List[str]:
    """Returns a list of human-readable problems (empty = all good). Never
    raises -- callers decide what to do with a non-empty list (spec: must
    never block the legacy export)."""
    problems: List[str] = []
    if result.status != "SUCCESS":
        return problems  # nothing to reconcile against if it didn't build
    expected_cartons = len(packages)
    if result.cartons_written != expected_cartons:
        problems.append(f"cartons_written={result.cartons_written} != expected {expected_cartons}")
    expected_items = sum(len(getattr(pkg, "items", [])) for pkg in packages)
    if result.items_written != expected_items:
        problems.append(f"items_written={result.items_written} != expected {expected_items}")
    return problems
