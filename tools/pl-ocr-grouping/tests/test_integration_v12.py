#!/usr/bin/env python3
"""
Integration tests for the v12 (BUILD REQUEST turn 5) features working
TOGETHER end-to-end: OR List match -> per-Store counting_scope_key ->
carton numbering -> A5 Sublist PDF -- using real pl_ocr_core.Package /
Item objects (not just duck-typed doubles), so a mismatch between the
dataclass's real field names and what the other modules expect would be
caught here even if it slipped through each module's own isolated tests.

Does NOT re-run PDF text-layer OCR (that is already covered by
test_pl_ocr_core.py's synthetic-fixture tests) -- packages are built
directly, exactly like a completed run_pipeline() would have left them.

Same no-framework convention as the rest of this tool's tests.
"""
from __future__ import annotations
import sys, tempfile
from pathlib import Path

HERE = Path(__file__).resolve().parent
TOOL_DIR = HERE.parent
CORE_PY = TOOL_DIR / "pl_ocr_core.py"
if str(TOOL_DIR) not in sys.path:
    sys.path.insert(0, str(TOOL_DIR))

import types

_ENTRY_MARKER = "# ── Entry point ────────────────────────────────────────────────────────────"


def _load_core_defs():
    src = CORE_PY.read_text(encoding="utf-8")
    src = src[:src.index(_ENTRY_MARKER)]
    src = (src.replace("__DIM_WEIGHT_SHEET__", "None").replace("__MASTER_DATA_SHEET__", "None")
              .replace("__RECURSIVE__", "False").replace("__MANUAL_CONSIGNEE__", "None")
              .replace("__MANUAL_NOTIFY_PARTY__", "None").replace("__GENERATE_SUBLIST__", "True")
              .replace("__GENERATE_SUBLIST_PDF__", "True").replace("__OR_LIST_FILE__", "None")
              .replace("__GIT_COMMIT__", repr("test-suite")))
    mod = types.ModuleType("pl_ocr_core_integration")
    mod.__file__ = str(CORE_PY)
    sys.modules[mod.__name__] = mod
    exec(compile(src, str(CORE_PY), "exec"), mod.__dict__)
    return mod.__dict__


C = _load_core_defs()
Package = C["Package"]
Item = C["Item"]
assign_counting_scope_keys = C["assign_counting_scope_keys"]
assign_global_numbers = C["assign_global_numbers"]

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


def _real_pkg(ref, seq, or_store, or_number, n_items=3, weight_kg="12.50 KG"):
    pkg = Package(package_code=f"PGKEC{ref}", source_file=f"{ref}.pdf",
                   reference_code=ref, pdf_package_seq=seq)
    pkg.or_list_match_status = "OK"
    pkg.or_list_store = or_store
    pkg.or_number = or_number
    pkg.so_number = f"SO-{or_number}"
    pkg.shipping_mark = ref
    pkg.pl_gross_weight = weight_kg
    pkg.items = [Item(no=str(i + 1), product_name="", product_code=f"{ref}-SKU{i}",
                       barcode=f"48900000000{i:02d}", unit="PCS", quantity=i + 1)
                 for i in range(n_items)]
    return pkg


def t_kerry_hangzhou_numbering_flows_through_to_the_actual_pdf_text():
    # The user's own worked example, run all the way through to rendered
    # PDF text -- not just the in-memory model (Task 18's tests already
    # cover that layer alone; this closes the gap up to the actual output
    # file a warehouse worker would print and use).
    pkgs = [_real_pkg(f"KERRY_{i}", i + 1, "Kerry", "OR1016") for i in range(6)]
    pkgs += [_real_pkg(f"HZ_{i}", i + 1, "Hangzhou", "OR2044") for i in range(4)]

    assign_counting_scope_keys(pkgs)
    assign_global_numbers(pkgs)

    kerry = [p for p in pkgs if p.or_list_store == "Kerry"]
    hz = [p for p in pkgs if p.or_list_store == "Hangzhou"]
    assert [p.carton_display for p in kerry] == ["1/6", "2/6", "3/6", "4/6", "5/6", "6/6"]
    assert [p.carton_display for p in hz] == ["1/4", "2/4", "3/4", "4/4"]

    with tempfile.TemporaryDirectory() as td:
        out = Path(td) / "SUBLIST_TOTAL.pdf"
        result = ppe.generate_sublist_pdf(pkgs, out)
        assert result.status == "SUCCESS", result.error
        assert result.pages_written == 10, "6 Kerry + 4 Hangzhou cartons, none oversized -> 10 pages, 1 each"
        problems = ppe.validate_sublist_pdf(pkgs, result)
        assert problems == [], problems

        import pdfplumber
        with pdfplumber.open(str(out)) as pdf:
            assert len(pdf.pages) == 10
            page_texts = [(p.extract_text() or "") for p in pdf.pages]
        kerry_pages = [t for t in page_texts if "Kerry" not in t and "OR1016" in t]  # shipping_mark is the ref code here, OR1016 identifies Kerry rows
        hz_pages = [t for t in page_texts if "OR2044" in t]
        assert len(kerry_pages) == 6, f"expected 6 Kerry pages (OR1016), got {len(kerry_pages)}"
        assert len(hz_pages) == 4, f"expected 4 Hangzhou pages (OR2044), got {len(hz_pages)}"
        # The literal combined-denominator bug this whole feature exists to
        # prevent: no page's Carton # line may ever read as a combined /10
        # denominator (e.g. "3/10") -- only Carton # lines are checked here
        # (not the whole page text), since item barcodes/codes elsewhere on
        # the page can coincidentally contain the digits "10".
        import re
        for t in page_texts:
            carton_line = next((ln for ln in t.splitlines() if ln.strip().startswith("Carton #")), "")
            assert not re.search(r"/10\b", carton_line), \
                f"Carton # line must never show a combined /10 denominator, got: {carton_line!r}"
        first_kerry_display = next(t for t in page_texts if "OR1016" in t)
        assert "1/6" in first_kerry_display
        first_hz_display = next(t for t in page_texts if "OR2044" in t)
        assert "1/4" in first_hz_display


def t_zero_item_loss_across_full_v12_flow_multi_store_mixed_sizes():
    # One oversized carton (forces PDF continuation) mixed in with normal
    # ones, across 2 different Store scopes -- nothing may be lost or
    # duplicated anywhere in the chain.
    pkgs = [_real_pkg(f"KERRY_{i}", i + 1, "Kerry", "OR1016", n_items=3) for i in range(3)]
    big = _real_pkg("KERRY_BIG", 4, "Kerry", "OR1016", n_items=ppe.ITEMS_PER_PDF_PAGE + 10)
    pkgs.append(big)
    pkgs += [_real_pkg(f"HZ_{i}", i + 1, "Hangzhou", "OR2044", n_items=2) for i in range(2)]

    assign_counting_scope_keys(pkgs)
    assign_global_numbers(pkgs)

    kerry = [p for p in pkgs if p.or_list_store == "Kerry"]
    hz = [p for p in pkgs if p.or_list_store == "Hangzhou"]
    assert len(kerry) == 4 and all(p.carton_total == 4 for p in kerry)
    assert len(hz) == 2 and all(p.carton_total == 2 for p in hz)

    expected_total_items = sum(len(p.items) for p in pkgs)
    with tempfile.TemporaryDirectory() as td:
        out = Path(td) / "SUBLIST_TOTAL.pdf"
        result = ppe.generate_sublist_pdf(pkgs, out)
        assert result.status == "SUCCESS", result.error
        assert result.cartons_written == 6
        assert result.items_written == expected_total_items
        assert result.pages_written == 7, "5 normal cartons (1 page each) + 1 oversized carton (2 pages) = 7"
        problems = ppe.validate_sublist_pdf(pkgs, result)
        assert problems == [], problems


test("Kerry(6)/Hangzhou(4) numbering flows through to the rendered PDF text (never a combined 1/10..)",
     t_kerry_hangzhou_numbering_flows_through_to_the_actual_pdf_text)
test("zero item loss across the full v12 flow: multi-store + oversized-carton continuation combined",
     t_zero_item_loss_across_full_v12_flow_multi_store_mixed_sizes)


# ── summary ──────────────────────────────────────────────────────────────
print(f"\n{_passed} passed, {_failed} failed")
if _failed:
    print("FAILED:", ", ".join(_failures))
    sys.exit(1)
sys.exit(0)
