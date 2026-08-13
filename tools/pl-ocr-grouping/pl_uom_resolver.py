#!/usr/bin/env python3
"""
pl_uom_resolver.py — PL OCR V21 central UOM / PCS-conversion resolver
(spec sections 21-31).

ONE canonical implementation (spec section 29/49) of:
  - raw UOM -> per-unit PCS multiplier parsing (section 21, 24)
  - the "Convert CARTON UOM to PCS" derived-output resolver (sections
    24-28, 31)

Every customer-facing item writer (PL_Total's Packing List sheet,
grouped PL, Sublist Excel, Sublist PDF) must call resolve_output_uom_qty()
here rather than re-implementing the CARTON_<N>PCS parsing/conversion
independently -- see spec section 30 ("customer-facing consistency").

No import of pl_ocr_core / pl_group_export / pl_sublist_export here --
leaf dependency, importable stand-alone by every one of those modules
without a circular import (pl_sublist_export.py in particular is
documented as deliberately never importing pl_ocr_core.py).
"""
from __future__ import annotations
import re
from dataclasses import dataclass
from typing import Optional

# Spec section 21: "Generic structure: CARTON_<positive integer>PCS".
# Strict on purpose -- this is used to derive a real multiplier, never a
# best-effort guess (spec section 32: malformed UOM must never fabricate a
# multiplier).
RE_CARTON_XXPCS = re.compile(r'^CARTON_(\d+)PCS$', re.IGNORECASE)

# Any of these are recognised as "this cell IS a carton-style UOM" even
# when the multiplier itself can't be safely parsed (spec section 32
# examples: CARTON_PCS, CARTON_XXPCS, CARTON_-10PCS, CARTON_0PCS) -- used
# only to decide whether to raise the malformed-UOM review flag versus
# treating the value as an ordinary non-carton unit (e.g. "PCS", "SET").
RE_CARTON_LIKE = re.compile(r'^CARTON[_-]', re.IGNORECASE)

REVIEW_MALFORMED_UOM = "UOM_CONVERSION_REVIEW"


def parse_carton_multiplier(uom_raw: str) -> Optional[int]:
    """Returns the per-carton PCS multiplier for a `CARTON_<N>PCS` UOM, or
    None when `uom_raw` isn't that exact shape (including malformed
    variants -- section 32: never fabricate a multiplier for those)."""
    m = RE_CARTON_XXPCS.fullmatch(str(uom_raw or "").strip())
    if not m:
        return None
    n = int(m.group(1))
    return n if n > 0 else None


@dataclass
class UomResolution:
    output_uom: str
    output_qty: int
    pcs_per_unit: int          # 1 for PCS-native rows, N for CARTON_NPCS
    qty_pcs: int                 # raw_qty * pcs_per_unit, ALWAYS computed
                                    # (independent of convert_to_pcs -- see
                                    # Raw_Data architecture, spec section 37)
    review_flag: str = ""       # REVIEW_MALFORMED_UOM or ""


def resolve_output_uom_qty(uom_raw: str, qty_raw: int, convert_to_pcs: bool) -> UomResolution:
    """The ONE place every exporter asks "what UOM/QTY do I print for this
    item row" (spec sections 24-31). Never mutates the item -- caller
    still owns uom_raw/qty_raw (Raw_Data keeps those untouched regardless
    of what this returns, spec section 23)."""
    uom_raw = str(uom_raw or "PCS").strip().upper() or "PCS"
    qty_raw = int(qty_raw or 0)

    multiplier = parse_carton_multiplier(uom_raw)
    is_carton_like = RE_CARTON_LIKE.match(uom_raw) is not None
    review_flag = REVIEW_MALFORMED_UOM if (is_carton_like and multiplier is None) else ""

    if multiplier is not None:
        pcs_per_unit = multiplier
    else:
        pcs_per_unit = 1

    qty_pcs = qty_raw * pcs_per_unit

    if not convert_to_pcs:
        # Section 25: OFF -- preserve original UOM/QTY semantics exactly.
        return UomResolution(output_uom=uom_raw, output_qty=qty_raw,
                              pcs_per_unit=pcs_per_unit, qty_pcs=qty_pcs, review_flag=review_flag)

    if review_flag:
        # Section 32: malformed CARTON_* UOM -- never fabricate a
        # multiplier, so conversion is a no-op for this row (raw values
        # preserved) even though Convert-to-PCS is ON; the review flag
        # tells the caller to surface it for manual attention.
        return UomResolution(output_uom=uom_raw, output_qty=qty_raw,
                              pcs_per_unit=pcs_per_unit, qty_pcs=qty_pcs, review_flag=review_flag)

    # Section 24/26-28: PCS source -> unchanged; CARTON_<N>PCS -> N x qty,
    # output UOM always "PCS" once conversion is ON.
    return UomResolution(output_uom="PCS", output_qty=qty_pcs,
                          pcs_per_unit=pcs_per_unit, qty_pcs=qty_pcs, review_flag="")
