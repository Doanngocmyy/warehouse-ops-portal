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
import re
from dataclasses import dataclass
from pathlib import Path
from typing import List, Optional

from reportlab.lib.pagesizes import A5
from reportlab.lib.colors import Color
from reportlab.lib.utils import ImageReader
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
# 1b) Brand logo (small, top-center) -- per business owner's explicit
#     request with a reference screenshot showing the real Sublist
#     template's own small "topologie" wordmark at the top of the page.
#     This is the CUSTOMER'S brand mark on their own shipping document,
#     not app/tool branding -- distinct from the earlier "no app
#     branding/watermark" rule, which was about not stamping this tool's
#     own name/logo onto the customer's paperwork.
# =========================================================================
_MODULE_DIR = Path(__file__).resolve().parent
DEFAULT_LOGO_PATH = _MODULE_DIR / "assets" / "topologie_logo.png"
LOGO_MAX_WIDTH = 50       # "nho nho" (small) -- a modest fraction of the 419pt page width
LOGO_TOP_GAP = 10         # distance from the true page top edge to the logo's top edge

# Top-left "page N/M" marker for a carton that spans multiple PDF pages
# (business owner's follow-up request) -- see _draw_continuation_page_marker().
# Deliberately smaller than SIZE_LABEL (10, used for "Carton #") and in the
# oblique style shared with the bottom "Continued on/from" note -- so this
# top-left page marker reads as an annotation, never as a data value, and
# can't be mistaken for the Store's own carton number (e.g. "1/6") shown
# in the metadata block.
CONTINUATION_MARKER_FONT = "Helvetica-Oblique"
CONTINUATION_MARKER_SIZE = 8
CONTINUATION_MARKER_TOP_GAP = 20  # from the true page top edge, roughly level with the logo


def _draw_logo(c: canvas.Canvas, logo_path):
    """Draws a small centered logo near the top of the page. Never raises
    and never blocks page generation -- a missing/unreadable logo file
    just means no logo is drawn (same non-blocking philosophy as the rest
    of this module; the logo is a visual nicety, not load-bearing data)."""
    if not logo_path:
        return
    logo_path = Path(logo_path)
    if not logo_path.exists():
        return
    try:
        reader = ImageReader(str(logo_path))
        iw, ih = reader.getSize()
        width = LOGO_MAX_WIDTH
        height = width * (ih / iw)
        x = (PAGE_W - width) / 2
        y = PAGE_H - LOGO_TOP_GAP - height
        c.drawImage(reader, x, y, width=width, height=height, mask="auto")
    except Exception as e:
        log.warning(f"Logo not drawn (non-blocking): {type(e).__name__}: {e}")


def _logo_reserved_height(logo_path):
    """Best-effort natural height of the logo at LOGO_MAX_WIDTH, used
    only to reserve vertical space below it -- so the metadata block
    never sits flush under the logo (business owner's feedback: the two
    were too close). Falls back to 0 (no reservation) if the logo can't
    be read, matching _draw_logo's own non-blocking failure behaviour."""
    if not logo_path:
        return 0
    logo_path = Path(logo_path)
    if not logo_path.exists():
        return 0
    try:
        reader = ImageReader(str(logo_path))
        iw, ih = reader.getSize()
        return LOGO_MAX_WIDTH * (ih / iw)
    except Exception:
        return 0

# =========================================================================
# 2) Page geometry -- A5 portrait, no title/watermark/branding anywhere.
# =========================================================================
PAGE_W, PAGE_H = A5  # (419.53, 595.28) pt -- already portrait (H > W)
MARGIN = 28

CONTENT_LEFT = MARGIN
CONTENT_RIGHT = PAGE_W - MARGIN
CONTENT_TOP = PAGE_H - MARGIN
CONTENT_WIDTH = CONTENT_RIGHT - CONTENT_LEFT

# Metadata/item-table vertical start. Reserves the logo's own footprint
# (LOGO_TOP_GAP + its natural height at LOGO_MAX_WIDTH) plus an explicit
# breathing-room gap, so "Carton #" never sits close enough to read as
# touching the logo -- computed once from the bundled default logo so
# capacity math (ITEMS_PER_PDF_PAGE etc.) reflects the real, reserved
# space rather than assuming the full CONTENT_TOP is available for text.
GAP_AFTER_LOGO = 12
_RESERVED_LOGO_HEIGHT = _logo_reserved_height(DEFAULT_LOGO_PATH)
METADATA_TOP_Y = (PAGE_H - LOGO_TOP_GAP - _RESERVED_LOGO_HEIGHT - GAP_AFTER_LOGO
                  if _RESERVED_LOGO_HEIGHT else CONTENT_TOP)

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
# Cap the value column's width -- beyond this, a value (typically Shipping
# Mark or Packing Code #) WRAPS onto additional lines instead of growing
# the column arbitrarily wide (business owner's follow-up request: "neu
# dai qua ky tu thi phai tu wrap text"). 200pt comfortably fits within the
# page while keeping the block's left edge well short of the item table's
# own columns, and fits ~35-40 characters of 10pt Helvetica per line --
# generous for the business codes actually seen in Shipping Mark values.
VALUE_COL_MAX_WIDTH = 200
# Reserved for ITEMS_PER_PDF_PAGE's worst-case capacity math (see below):
# how many EXTRA wrapped lines (beyond the normal 1-per-row) to budget
# for across the whole metadata block. Real Shipping Mark / Packing Code
# values wrapping to 2 lines each (i.e. +1 extra line per field) is
# already a generous assumption for the codes this tool actually sees;
# documented rather than silently picked.
MAX_RESERVED_WRAP_EXTRA_LINES = 2

# v17 (spec sections 12/13/16/17 -- supersedes the v15/v16 fixed-6-row
# design): Carton#/Shipping Mark/OR No./Ref No. are FIXED rows (labels
# never vary -- spec: "Do NOT relabel Ref# as SO/SO Order/Invoice"),
# followed by 0+ OPTIONAL business-field rows (verbatim label, only the
# ones the OR List actually has), then GW/Packing Code# (fixed, unchanged
# meaning). Row COUNT now varies by shipment (6 fixed + n_optional), so
# every pagination constant that used to be a bare module-level number
# derived from len(METADATA_ROWS) is now a FUNCTION of n_optional -- see
# _meta_block_height()/_items_per_pdf_page() below. METADATA_ROWS/META_
# BLOCK_HEIGHT/ITEMS_PER_PDF_PAGE are kept as module constants too (the
# n_optional=0 case) for any direct importer/test that still wants the
# historical bare-constant convenience.
#
# v20 (Sublist display-only UI change): Store was removed as a Sublist
# row entirely -- it is still resolved and required internally
# (Package.store/store_display, CN Store matching, PL_Total, grouped
# Packing Lists, Raw_Data, Match_Status, PL_SPLIT_CONTROL, 04_CN_BY_STORE
# all keep it unchanged) but is intentionally never displayed on the
# Sublist PDF. Removing it drops the fixed-row count from 7 back to 6;
# every height/pagination constant below is still derived from
# len(METADATA_ROWS)/n_rows, so this is the ONLY place that needed to
# change.
METADATA_ROWS = [
    ("carton", "Carton #"),
    ("shipping_mark", "Shipping Mark"),
    ("or", "OR No."),
    ("ref", "Ref No."),
    ("gw", "GW"),
    ("packing_code", "Packing Code #"),
]


def _resolve_metadata_rows(optional_labels=None):
    """-> (key, label) pairs: the fixed Carton#/Shipping Mark/OR No./
    Ref No. rows (v20: no Store row, see module comment -- Store is
    still resolved and used elsewhere, just never displayed on the
    Sublist), then one row per entry in `optional_labels` (verbatim,
    spec section 13) keyed "optional_0", "optional_1", ..., then the
    fixed GW/Packing Code# rows. []/None optional_labels reproduces
    METADATA_ROWS exactly (6 rows, no optional rows)."""
    optional_labels = [str(l) for l in (optional_labels or []) if str(l or "").strip()]
    rows = [
        ("carton", "Carton #"), ("shipping_mark", "Shipping Mark"),
        ("or", "OR No."), ("ref", "Ref No."),
    ]
    for i, label in enumerate(optional_labels):
        rows.append((f"optional_{i}", label))
    rows.append(("gw", "GW"))
    rows.append(("packing_code", "Packing Code #"))
    return rows


def _meta_block_height(n_rows: int) -> float:
    """Single-line ("no wrap") reference height for a metadata block with
    `n_rows` total rows (see _resolve_metadata_rows -- 7 fixed + however
    many optional fields this run has)."""
    return n_rows * ROW_HEIGHT_META


def _meta_block_height_reserved(n_rows: int) -> float:
    """Worst-case (wrapped) reserved height for the same block -- see
    MAX_RESERVED_WRAP_EXTRA_LINES's own docstring below."""
    return (n_rows + MAX_RESERVED_WRAP_EXTRA_LINES) * ROW_HEIGHT_META


META_BLOCK_HEIGHT = _meta_block_height(len(METADATA_ROWS))
META_BLOCK_HEIGHT_RESERVED = _meta_block_height_reserved(len(METADATA_ROWS))

GAP_AFTER_METADATA = 14


def _wrap_text_to_width(text: str, font: str, size: float, max_width: float) -> "list[str]":
    """Wraps `text` into lines that each fit within `max_width`, breaking
    on whitespace AND hyphens (Shipping Mark / reference codes are
    typically hyphen-separated with no spaces at all, e.g.
    'CN-1666-PVG-KERRY-POP') -- a plain whitespace-only wrapper would
    never break such a string at all. Falls back to hard character-wrap
    for the rare case where a single unbreakable token is itself still
    wider than max_width, so a line can never overflow past max_width no
    matter what. Pure function, no canvas needed. Always returns at least
    one line (possibly empty)."""
    text = str(text or "")
    if not text:
        return [""]
    if pdfmetrics.stringWidth(text, font, size) <= max_width:
        return [text]

    tokens = re.findall(r'[^\s-]+-?|\s+', text)
    lines: "list[str]" = []
    current = ""
    for tok in tokens:
        candidate = current + tok
        if not current or pdfmetrics.stringWidth(candidate, font, size) <= max_width:
            current = candidate
        else:
            lines.append(current.rstrip())
            current = tok.lstrip() if tok.strip() else ""
    if current.strip():
        lines.append(current.rstrip())

    final: "list[str]" = []
    for line in lines:
        if pdfmetrics.stringWidth(line, font, size) <= max_width:
            final.append(line)
        else:
            # Hard character-wrap for a single token still too wide.
            chunk = ""
            for ch in line:
                candidate = chunk + ch
                if pdfmetrics.stringWidth(candidate, font, size) <= max_width or not chunk:
                    chunk = candidate
                else:
                    final.append(chunk)
                    chunk = ch
            if chunk:
                final.append(chunk)
    return final or [""]


def _wrap_metadata_values(value_texts) -> "list[list[str]]":
    """value_texts (6 raw strings, in METADATA_ROWS order) -> wrapped
    lines per row, each capped at VALUE_COL_MAX_WIDTH. Pure function --
    the single source of truth both _compute_metadata_x() (sizing) and
    _draw_page() (actual drawing) use, so the computed column width and
    what actually gets drawn can never disagree."""
    return [_wrap_text_to_width(v, FONT_VALUE, SIZE_VALUE, VALUE_COL_MAX_WIDTH) for v in value_texts]


def _compute_metadata_x(value_texts) -> "tuple[float, float]":
    """-> (metadata_value_x, metadata_label_right_x), measured from the
    actual (possibly wrapped) value text so the block's right edge sits
    near CONTENT_RIGHT (the page's right margin) and the label/value gap
    stays a small fixed distance regardless of content length -- a value
    wider than VALUE_COL_MAX_WIDTH wraps onto more lines rather than
    growing the column past that cap. Pure function -- no canvas needed
    (pdfmetrics.stringWidth uses the font's built-in AFM/glyph-width
    table), so this is directly unit-testable.

        metadata_right_x     = CONTENT_RIGHT               (page_width - right_margin)
        metadata_value_x     = metadata_right_x - value_column_width
        metadata_label_right_x = metadata_value_x - VALUE_LABEL_GAP
    """
    wrapped = _wrap_metadata_values(value_texts)
    widths = [pdfmetrics.stringWidth(line, FONT_VALUE, SIZE_VALUE) for lines in wrapped for line in lines]
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

# Grid-line cell geometry: each row's text baseline sits ROW_ASCENT above
# the cell's bottom border and ROW_DESCENT below its top border, so
# ROW_ASCENT + ROW_DESCENT == the row's own height and adjacent cells tile
# with no gap/overlap (row i's bottom border == row i+1's top border).
# Same idea for the (taller, bold) header row with its own ASCENT/DESCENT.
ROW_ASCENT = 10
ROW_DESCENT = 4
HEADER_ASCENT = 10
HEADER_DESCENT = 6
assert ROW_ASCENT + ROW_DESCENT == ROW_HEIGHT_ITEM
assert HEADER_ASCENT + HEADER_DESCENT == ROW_HEIGHT_ITEM_HEADER
HEADER_FILL_COLOR = Color(0.90, 0.90, 0.90)
GRID_LINE_WIDTH = 0.5

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

# y-position of the FIRST item row (immediately below the header), for the
# SINGLE-LINE (no metadata wrapping) reference case -- used as the default
# in pure-function tests and as the base for META_BLOCK_HEIGHT_RESERVED
# below. The ACTUAL per-page value used while drawing is dynamic (see
# _draw_page(), which tracks the real `y` after however many metadata
# lines that page's Shipping Mark/Packing Code actually wrapped to).
def _item_table_start_y(n_rows: int) -> float:
    """y-position of the FIRST item row (immediately below the header),
    single-line reference case, for a metadata block with `n_rows` total
    rows. The ACTUAL per-page value used while drawing is dynamic (see
    _draw_page(), which tracks the real `y` after however many metadata
    lines that page's Shipping Mark/Packing Code/optional-field values
    actually wrapped to, AND however many total metadata rows this run
    has)."""
    return METADATA_TOP_Y - _meta_block_height(n_rows) - GAP_AFTER_METADATA - ROW_HEIGHT_ITEM_HEADER


ITEM_TABLE_START_Y = _item_table_start_y(len(METADATA_ROWS))

# Fixed vertical budget that must be reserved BELOW the item rows for the
# total row + a possible "Continued" note, even though the total row's
# actual draw position is now dynamic (see _compute_total_y()) -- this is
# what still caps items-per-page so a full page's dynamic total can never
# be pushed below the bottom margin.
_RESERVED_BELOW_ITEMS = GAP_AFTER_ITEMS + TOTAL_TEXT_PADDING + GAP_BEFORE_CONTINUED_NOTE + ROW_HEIGHT_CONTINUED_NOTE


def _items_per_pdf_page(n_rows: int) -> int:
    """Item capacity per A5 page for a metadata block with `n_rows` total
    rows (v17: n_rows now varies by shipment -- 7 fixed + however many
    optional business fields this run has -- so this is a function, not a
    bare module constant; see module docstring for why it's derived from
    real page geometry rather than copied from the Excel template's 18).
    Capacity is computed against the WORST-CASE (wrapped) metadata block
    height, not the single-line one -- so a page whose Shipping Mark/
    Packing Code/optional field actually wraps still can't overflow past
    the bottom margin. Never returns less than 1."""
    item_table_start_y_reserved = METADATA_TOP_Y - _meta_block_height_reserved(n_rows) - GAP_AFTER_METADATA - ROW_HEIGHT_ITEM_HEADER
    available_for_items = (item_table_start_y_reserved - MARGIN) - _RESERVED_BELOW_ITEMS
    return max(1, int(available_for_items // ROW_HEIGHT_ITEM))


# Documented here rather than silently picked: at MARGIN=28/ROW_HEIGHT_
# ITEM=14 the 6-row (n_optional=0) case works out to the low-to-mid-20s
# items/page (v20: one row taller than this same n_optional=0 case used
# to be before Store's row was removed again -- see module comment);
# verified against a rendered sample (see tests/test_pl_sublist_pdf_
# export.py's visual-validation step). Kept as a module constant (the
# n_optional=0 case) for any direct importer/test that still wants the
# historical bare-constant convenience.
ITEMS_PER_PDF_PAGE = _items_per_pdf_page(len(METADATA_ROWS))


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


def _paginate_for_pdf(cartons: list, items_per_page: int = None) -> List[PdfPageBlock]:
    """One package == one carton, ALWAYS. A carton with more items than
    `items_per_page` fits is split across multiple pages ("Continued 1",
    "Continued 2", ...) -- zero items ever lost, subtotal shown on every
    non-last page, grand total only on the carton's last page.
    items_per_page (v17): defaults to the module's ITEMS_PER_PDF_PAGE
    (the n_optional=0 case) when not given -- generate_sublist_pdf()
    always passes the ACTUAL value for however many optional business
    fields this run has, via _items_per_pdf_page()."""
    blocks: List[PdfPageBlock] = []
    cap = items_per_page if items_per_page else ITEMS_PER_PDF_PAGE
    for carton in cartons:
        n = len(carton.items)
        chunks = [carton.items[i:i + cap] for i in range(0, n, cap)] or [[]]
        block_count = len(chunks)
        for idx, chunk in enumerate(chunks):
            is_last = idx == block_count - 1
            # Carton # must read identically on every page of the same carton
            # (e.g. always "1/6") -- continuation is conveyed solely by the
            # top-left "N/M" marker and the bottom "Continued on/from" note,
            # never by mutating the Carton # metadata value itself.
            label = carton.carton_display
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
def _draw_metadata_row(c: canvas.Canvas, y: float, label: str, value_lines: "list[str]",
                        value_x: float, label_right_x: float) -> float:
    """Draws one metadata row -- label right-aligned on the FIRST line
    only (a wrapped 2nd+ line has no label repeated, matching how a
    label normally applies to its whole field, not each wrapped
    fragment), value lines left-aligned, one per line, top to bottom.
    Returns the y AFTER this row (i.e. y minus however many lines it
    took), so multi-line rows correctly push the next row down."""
    c.setFont(FONT_LABEL, SIZE_LABEL)
    c.drawRightString(label_right_x, y, label)
    c.setFont(FONT_VALUE, SIZE_VALUE)
    for i, line in enumerate(value_lines or [""]):
        c.drawString(value_x, y - i * ROW_HEIGHT_META, line)
    return y - (len(value_lines or [""]) - 1) * ROW_HEIGHT_META


def _draw_item_table_grid(c: canvas.Canvas, header_bottom_y: float, n_rows: int):
    """Draws the bordered grid (outer box + column separators + row
    separators + a light-gray header fill) for the item table -- ONLY
    called when there is at least 1 item row (spec: "co SKU data moi ke
    bang, con khong thi thoi" -- an empty table is never drawn at all).
    header_bottom_y is the y where the header cell ends / row 0 begins
    (== ITEM_TABLE_START_Y + ROW_ASCENT, i.e. the same header/body split
    used by the text-drawing loop -- kept as one source of truth so the
    grid lines and the text they frame can never drift apart)."""
    table_top = header_bottom_y + ROW_HEIGHT_ITEM_HEADER
    table_bottom = header_bottom_y - n_rows * ROW_HEIGHT_ITEM

    # Header fill (light gray), drawn before the border strokes so the
    # border lines stay crisp on top of it.
    c.setFillColor(HEADER_FILL_COLOR)
    c.rect(CONTENT_LEFT, header_bottom_y, CONTENT_WIDTH, ROW_HEIGHT_ITEM_HEADER, stroke=0, fill=1)
    c.setFillColor(Color(0, 0, 0))

    c.setLineWidth(GRID_LINE_WIDTH)
    # Outer border (header + all item rows together).
    c.rect(CONTENT_LEFT, table_bottom, CONTENT_WIDTH, table_top - table_bottom, stroke=1, fill=0)
    # Header / body separator.
    c.line(CONTENT_LEFT, header_bottom_y, CONTENT_RIGHT, header_bottom_y)
    # Row separators between item rows (n_rows - 1 internal lines).
    for i in range(1, n_rows):
        y_line = header_bottom_y - i * ROW_HEIGHT_ITEM
        c.line(CONTENT_LEFT, y_line, CONTENT_RIGHT, y_line)
    # Column separators, spanning the FULL table height (header + rows).
    col_x_ean_line = COL_X_EAN - (VALUE_LABEL_GAP / 2)
    col_x_qty_line = COL_X_QTY - (VALUE_LABEL_GAP / 2)
    c.line(col_x_ean_line, table_bottom, col_x_ean_line, table_top)
    c.line(col_x_qty_line, table_bottom, col_x_qty_line, table_top)


def _draw_continuation_page_marker(c: canvas.Canvas, block: PdfPageBlock):
    """Small 'page N/M' marker in the page's TOP-LEFT corner -- business
    owner's follow-up request: "voi sublist neu bi dai qua thi tach ra
    ghi ben goc trai la 1/2 2/2 (vi du)". Only drawn when a carton actually
    spans more than one page (block_count > 1); a normal single-page
    carton gets no marker at all, matching the "neu bi dai qua" (only
    when it's actually too long) condition. Distinct from the existing
    'Continued on/from...' sentence near the bottom of the page -- this
    is the quick-glance top-left version, the bottom note stays as the
    fuller-context version; both point at the same block_index/block_count."""
    if block.block_count <= 1:
        return
    c.setFont(CONTINUATION_MARKER_FONT, CONTINUATION_MARKER_SIZE)
    c.drawString(CONTENT_LEFT, PAGE_H - CONTINUATION_MARKER_TOP_GAP,
                 f"{block.block_index + 1}/{block.block_count}")


def _draw_page(c: canvas.Canvas, block: PdfPageBlock, logo_path=None, optional_business_field_labels=None):
    carton = block.carton
    y = METADATA_TOP_Y

    _draw_logo(c, logo_path if logo_path is not None else DEFAULT_LOGO_PATH)
    _draw_continuation_page_marker(c, block)

    # -- Metadata block (Carton# / Shipping Mark / Store / OR No. / Ref
    #    No. / [optional fields] / GW / Packing Code#) -- label right-
    #    aligned, value left-aligned, TIGHTLY adjacent, positioned in the
    #    page's UPPER-RIGHT (see _compute_metadata_x() docstring for the
    #    exact formula) -- the left side of the page stays visually open,
    #    matching the approved reference layout. v17 (spec sections 12/
    #    13/16/17): row LABELS/KEYS come from _resolve_metadata_rows
    #    (optional_business_field_labels) -- OR No./Ref No. are FIXED
    #    (never relabeled), plus one row per actual optional OR List
    #    field beyond OR/Ref (verbatim, never invented when absent).
    #    v20: Store intentionally has no row here (display-only omission,
    #    see module comment) -- it is not read into meta_values below.
    optional_labels = [str(l) for l in (optional_business_field_labels or []) if str(l or "").strip()]
    metadata_rows = _resolve_metadata_rows(optional_labels)
    opt_map = dict(getattr(carton, "optional_business_fields", None) or [])
    meta_values = {
        "carton": block.block_carton_label,
        "shipping_mark": carton.shipping_mark,
        "or": carton.or_number,
        "ref": carton.so_number,
        "gw": carton.gross_weight_display,
        "packing_code": carton.packing_code,
    }
    for i, label in enumerate(optional_labels):
        meta_values[f"optional_{i}"] = opt_map.get(label, "")
    value_texts = [meta_values.get(key, "") for key, _label in metadata_rows]
    wrapped_values = _wrap_metadata_values(value_texts)
    metadata_value_x, metadata_label_right_x = _compute_metadata_x(value_texts)
    for (key, label), lines in zip(metadata_rows, wrapped_values):
        y -= ROW_HEIGHT_META
        y = _draw_metadata_row(c, y, label, lines, metadata_value_x, metadata_label_right_x)

    y -= GAP_AFTER_METADATA
    content_start_y = y  # top of where the item table (or nothing) begins -- reflects
                          # the ACTUAL metadata block height on this page, including any
                          # wrapped Shipping Mark/Packing Code lines.

    n_items = len(block.items)
    if n_items > 0:
        # Grid (fill + border lines) is drawn FIRST, text on TOP of it --
        # drawing order matters with reportlab's painter model (whatever
        # is drawn later covers whatever was drawn earlier), and the
        # header's light-gray fill would otherwise paint over its own
        # "Item No."/"EAN"/"QTY" labels if drawn after them (caught during
        # visual review of this exact change).
        # header_bottom_y (== this page's actual item_table_start_y) is now
        # DYNAMIC -- it equals the fixed ITEM_TABLE_START_Y constant only
        # when no metadata value wrapped to extra lines on this page; when
        # Shipping Mark/Packing Code DID wrap, everything below shifts
        # down by exactly that many extra ROW_HEIGHT_META increments,
        # which is exactly what `y` (tracked through the metadata loop
        # above) already reflects.
        header_bottom_y = y - ROW_HEIGHT_ITEM_HEADER
        _draw_item_table_grid(c, header_bottom_y, n_items)

        # -- Item table header text (bordered, light-gray fill -- spec:
        #    "dung de bang trong tron", draw real grid lines, not just a
        #    bare underline) --
        c.setFont(FONT_ITEM_HEADER, SIZE_ITEM_HEADER)
        # baseline sits HEADER_ASCENT below the cell's TOP edge (== y here,
        # before decrementing) so header text is vertically centered-ish
        # inside its shaded/bordered cell.
        header_baseline_y = y - HEADER_ASCENT
        c.drawString(COL_X_ITEM_NO + 4, header_baseline_y, "Item No.")
        c.drawString(COL_X_EAN + 4, header_baseline_y, "EAN")
        c.drawCentredString(COL_X_QTY_CENTER, header_baseline_y, "QTY")
        y -= ROW_HEIGHT_ITEM_HEADER

        # -- Item rows: Item No. / EAN left-aligned, QTY centered --
        c.setFont(FONT_ITEM, SIZE_ITEM)
        for row in block.items:
            # baseline sits ROW_ASCENT below the cell's TOP edge (== y
            # here, before decrementing) -- same convention as the header.
            row_baseline_y = y - ROW_ASCENT
            c.drawString(COL_X_ITEM_NO + 4, row_baseline_y, str(row.item_no))
            c.drawString(COL_X_EAN + 4, row_baseline_y, str(row.ean))
            c.drawCentredString(COL_X_QTY_CENTER, row_baseline_y, str(row.qty))
            y -= ROW_HEIGHT_ITEM

        # -- Total QTY: placed DYNAMICALLY immediately after the last
        #    rendered item row (see _compute_total_y()) -- never anchored
        #    to the page bottom. Subtotal on a non-last continuation page,
        #    GRAND TOTAL (the full carton's total_qty) only on the
        #    carton's last page. --
        total_rule_y, total_text_y = _compute_total_y(header_bottom_y, n_items)
    else:
        # spec: "co SKU data moi ke bang, con khong thi thoi, khong co ke,
        # de so total ngay ben duoi" -- zero items: skip the header AND
        # the grid entirely, place Total QTY directly below the metadata
        # block instead (nothing to reconcile a table against).
        total_rule_y, total_text_y = _compute_total_y(content_start_y, 0)

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
                          enabled: bool = True,
                          logo_path=None,
                          optional_business_field_labels=None) -> SublistPdfBuildResult:
    """packages -> A5 PDF at `output_path`, one page per carton block.
    NEVER raises -- any failure is captured into the returned result's
    status="FAILED" (spec: Sublist PDF is optional, its failure must never
    break the legacy ZIP/export). Mirrors generate_sublist_workbook()'s
    signature intentionally so app.html/run_pipeline call both the same
    way.

    logo_path: small brand logo drawn top-center on every page (business
    owner's request). Defaults to None, which means "use the bundled
    DEFAULT_LOGO_PATH if it exists" (see _draw_page/_draw_logo) -- pass an
    explicit falsy sentinel like False to suppress the logo entirely if
    ever needed (kept flexible, not hardcoded to always-on).

    optional_business_field_labels (v17, spec sections 12/13/16/17): the
    uploaded OR List's business columns BEYOND OR/Ref (verbatim, in
    original order -- e.g. ["SO","PO","Invoice","Fulfillment No.",
    "Buyer"]). OR No./Ref No. are FIXED metadata rows (spec: never
    relabeled) -- so the Packing List sheet and this PDF's metadata block
    can never disagree. []/None (the default) shows exactly Carton #/
    Shipping Mark/OR No./Ref No./GW/Packing Code# (v20: no Store row --
    still resolved/required everywhere else, just never displayed here,
    see module comment), no invented optional rows. The metadata block's
    height (and therefore how many item rows fit per page / where
    pagination breaks) is recomputed from the ACTUAL number of rows this
    call needs -- see _items_per_pdf_page() -- so an extended OR List's
    taller metadata block can never overlap or clip the item table."""
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
        optional_labels = [str(l) for l in (optional_business_field_labels or []) if str(l or "").strip()]
        # 4 fixed head rows (Carton#/Shipping Mark/OR No./Ref No. -- v20:
        # Store's row removed, see module comment) + n optional + 2 fixed
        # tail rows (GW/Packing Code#) -- see _resolve_metadata_rows().
        n_meta_rows = 4 + len(optional_labels) + 2
        items_per_page = _items_per_pdf_page(n_meta_rows)

        output_path = Path(output_path)
        cartons = [pse.build_sublist_carton_model(pkg, carton_display_mode=carton_display_mode)
                   for pkg in packages]
        blocks = _paginate_for_pdf(cartons, items_per_page=items_per_page)

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
            _draw_page(c, block, logo_path=logo_path, optional_business_field_labels=optional_labels)
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
