#!/usr/bin/env python3
"""
tests/test_v21_routing_and_uom.py
==================================
Permanent regression tests for the "PL OCR V21" pass:

  - pl_routing_rules.py: routing-rule validation/normalization, Shipping
    Mark tokenization, and rule matching (exact / partial-unique / unique-
    fallback / conflict) -- spec sections 3-13, tests per section 42.
  - pl_uom_resolver.py: CARTON_<N>PCS UOM parsing + PCS conversion --
    spec sections 21-32, tests per section 43.
  - pl_sublist_export.py: the 4 checkbox-mode Sublist outputs (Convert x
    Show-UOM) -- spec sections 25-31, tests per section 44.
  - pl_ocr_core.classify_packages_for_port() / pl_group_export.
    match_store_and_or_v21(): routing_rules wired into the real production
    Store/Port + OR/Ref resolution, including the CN|blank|Tmall case
    (spec sections 14/38) and the CN-6557 multi-store production counts
    re-expressed as user routing rules (spec section 40).

No test-framework dependency, by design -- matches this repo's existing
convention. Run with:

    python3 tools/pl-ocr-grouping/tests/test_v21_routing_and_uom.py

NOTE ON DATA: only literal Shipping Mark / SKU / EAN TEXT STRINGS (either
invented for this test or the same real CN-6557 production Shipmark
shapes already used by test_v15_bugfix_regression.py) are used below --
no customer PDFs, DIM files, addresses, or phone numbers.
"""
from __future__ import annotations
import sys, types
from pathlib import Path
from collections import Counter

HERE = Path(__file__).resolve().parent
TOOL_DIR = HERE.parent
CORE_PY = TOOL_DIR / "pl_ocr_core.py"
if str(TOOL_DIR) not in sys.path:
    sys.path.insert(0, str(TOOL_DIR))

import pl_routing_rules as prr
import pl_uom_resolver as uomr
import pl_group_export as pge
import pl_sublist_export as pse

_ENTRY_MARKER = "# ── Entry point ────────────────────────────────────────────────────────────"


def _substitute_placeholders(src: str) -> str:
    def lit(v):
        return "None" if v is None else repr(v)
    return (src
            .replace("__DIM_WEIGHT_SHEET__", lit(None))
            .replace("__MASTER_DATA_SHEET__", lit(None))
            .replace("__RECURSIVE__", "False")
            .replace("__MANUAL_CONSIGNEE__", lit(None))
            .replace("__MANUAL_NOTIFY_PARTY__", lit(None))
            .replace("__GENERATE_SUBLIST__", "True")
            .replace("__GENERATE_SUBLIST_PDF__", "True")
            .replace("__OR_LIST_FILE__", "None")
            .replace("__ROUTING_RULES_JSON__", "[]")
            .replace("__CONVERT_TO_PCS__", "False")
            .replace("__SHOW_UOM_IN_SUBLIST__", "False")
            .replace("__GIT_COMMIT__", lit("test-v21")))


def _exec_module(src: str):
    modname = f"pl_ocr_core_v21test_{id(src)}"
    mod = types.ModuleType(modname)
    mod.__file__ = str(CORE_PY)
    sys.modules[modname] = mod
    exec(compile(src, str(CORE_PY), "exec"), mod.__dict__)
    return mod.__dict__


def load_core_defs():
    src = CORE_PY.read_text(encoding="utf-8")
    idx = src.index(_ENTRY_MARKER)
    src = _substitute_placeholders(src[:idx])
    return _exec_module(src)


D = load_core_defs()
Package = D["Package"]
classify_packages_for_port = D["classify_packages_for_port"]

# ── tiny test runner (mirrors every other test_*.py in this directory) ─────
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


def _pkg(mark: str, seq: int, country: str = "CN") -> "Package":
    p = Package(package_code=f"PKG{seq:03d}", source_file=f"{mark}.pdf",
                reference_code=mark, pdf_package_seq=0)
    p.shipping_mark = mark
    p.shipping_mark_source = "PDF_STANDALONE_CODE"
    p.country = country
    return p


# =========================================================================
# Section 42: routing tests
# =========================================================================
print("== V21 section 42: routing-rule validation ==")

test("Country Code: valid 2-letter codes accepted (incl. unseen ones)",
     lambda: [prr.is_valid_country_code(c) for c in ("CN", "SG", "JP", "ZZ")] == [True] * 4)
test("Country Code: lowercase normalized to upper",
     lambda: prr.normalize_country("sg") == "SG")
test("Country Code: invalid single letter rejected",
     lambda: prr.is_valid_country_code("C") is False)
test("Country Code: invalid 3-letter code rejected",
     lambda: prr.is_valid_country_code("CHN") is False)
test("Country Code: invalid numeric rejected",
     lambda: prr.is_valid_country_code("12") is False)


def t_duplicate_rule_rejected():
    v = prr.validate_routing_rules([
        {"country": "CN", "port": "PVG", "store": "Tmall"},
        {"country": "cn", "port": "pvg", "store": "Tmall"},
    ])
    assert len(v.rules) == 1, v.rules
    assert len(v.errors) == 1, v.errors


test("duplicate rule rejection (case/whitespace-insensitive)", t_duplicate_rule_rejected)


def t_multiple_stores_same_country_valid():
    v = prr.validate_routing_rules([
        {"country": "CN", "port": "PEK", "store": "ChinaWorld"},
        {"country": "CN", "port": "PVG", "store": "Hangzhou"},
        {"country": "CN", "port": "PVG", "store": "Kerry"},
        {"country": "CN", "port": "SZX", "store": "Shenzhen"},
    ])
    assert v.is_valid and len(v.rules) == 4


test("multiple Stores under the same Country: all valid", t_multiple_stores_same_country_valid)
test("blank Port is a valid rule (never an error)",
     lambda: prr.validate_routing_rules([{"country": "CN", "port": "", "store": "Tmall"}]).is_valid)
test("blank Store is a valid rule (never an error)",
     lambda: prr.validate_routing_rules([{"country": "SG", "port": "", "store": ""}]).is_valid)


def t_shipmark_cn_1000_pvg_cnworld_pop():
    tok = prr.tokenize_shipping_mark("CN-1000-PVG-CNWorld-POP")
    assert tok.country == "CN"
    assert tok.total_qty == "1000"
    assert tok.body_tokens == ["PVG", "CNWorld"]
    assert tok.suffix == "POP"


test("Shipping Mark CN-1000-PVG-CNWorld-POP decomposes correctly", t_shipmark_cn_1000_pvg_cnworld_pop)


def t_shipmark_separator_variants():
    rules = prr.validate_routing_rules([{"country": "CN", "port": "PVG", "store": "CNWorld"}]).rules
    for variant in ("CN-1000-PVG-CNWorld-POP", "CN_1000_PVG_CNWorld_POP", "CN 1000 PVG CNWorld POP"):
        tok = prr.tokenize_shipping_mark(variant)
        m = prr.match_route(tok.country, tok.body_tokens, rules)
        assert m.status == "MATCHED" and m.port == "PVG" and m.store == "CNWorld", (variant, m)


test("Shipping Mark separator variants (-, _, space) all resolve identically", t_shipmark_separator_variants)


def t_shipmark_port_before_store_and_store_before_port():
    rules = prr.validate_routing_rules([{"country": "CN", "port": "PVG", "store": "CNWorld"}]).rules
    m1 = prr.match_route("CN", ["PVG", "CNWorld"], rules)
    m2 = prr.match_route("CN", ["CNWorld", "PVG"], rules)
    assert m1.status == "MATCHED" and m2.status == "MATCHED"
    assert m1.port == m2.port == "PVG"
    assert m1.store == m2.store == "CNWorld"


test("Port-before-Store and Store-before-Port both resolve (no fixed BODY position)",
     t_shipmark_port_before_store_and_store_before_port)


def t_suffix_never_treated_as_routing_signal():
    # SG-500-...-CN: destination SG, suffix "CN" is metadata/origin, NEVER
    # reinterpreted as destination country/Store/Port (spec section 9).
    rules = prr.validate_routing_rules([{"country": "SG", "port": "", "store": ""}]).rules
    tok = prr.tokenize_shipping_mark("SG-500-WAREHOUSE-CN")
    assert tok.country == "SG"
    assert tok.suffix == "CN"
    m = prr.match_route(tok.country, tok.body_tokens, rules)
    assert m.status == "MATCHED" and m.country == "SG"


test("suffix (CN/VN/POP/QF/...) never reinterpreted as destination Country/Store/Port",
     t_suffix_never_treated_as_routing_signal)


def t_matching_hierarchy():
    rules = prr.validate_routing_rules([
        {"country": "CN", "port": "PVG", "store": "Hangzhou"},
        {"country": "CN", "port": "PVG", "store": "Kerry"},
        {"country": "CN", "port": "SZX", "store": "Shenzhen"},
    ]).rules
    # Country+Port+Store exact
    m = prr.match_route("CN", ["PVG", "Hangzhou"], rules)
    assert m.status == "MATCHED" and m.method == prr.METHOD_EXACT and m.port == "PVG" and m.store == "Hangzhou"
    # Country+unique Store (Port token missing from OCR)
    m = prr.match_route("CN", ["Shenzhen"], rules)
    assert m.status == "MATCHED" and m.method == prr.METHOD_PARTIAL and m.port == "SZX"
    # ambiguous candidates -> REVIEW (Port shared by 2 Stores, no Store signal)
    m = prr.match_route("CN", ["PVG"], rules)
    assert m.status == "REVIEW"
    # contradictory Port/Store -> REVIEW
    m = prr.match_route("CN", ["PVG", "Shenzhen"], rules)
    assert m.status == "REVIEW" and m.method == prr.METHOD_CONFLICT
    # no route at all (unknown country) -> REVIEW
    m = prr.match_route("JP", ["TOKYO"], rules)
    assert m.status == "REVIEW" and m.method == prr.METHOD_NO_RULE_FOR_COUNTRY


test("matching hierarchy: exact / unique-Store / ambiguous / conflict / no-route",
     t_matching_hierarchy)


def t_tmall_blank_port_no_special_case():
    rules = prr.validate_routing_rules([{"country": "CN", "port": "", "store": "Tmall"}]).rules
    tok = prr.tokenize_shipping_mark("CN_TMALL-6676-CTN_POP")
    m = prr.match_route(tok.country, tok.body_tokens, rules)
    assert m.status == "MATCHED"
    assert m.port == ""
    assert m.store == "Tmall"
    assert m.method == prr.METHOD_EXACT  # blank-by-design, not a "missed OCR field"


test("CN|blank|Tmall -> blank Port is a clean MATCHED result, not an error",
     t_tmall_blank_port_no_special_case)


# =========================================================================
# Section 43: UOM tests
# =========================================================================
print("== V21 section 43: UOM parsing + conversion ==")

test("parser: PCS", lambda: uomr.parse_carton_multiplier("PCS") is None)
test("parser: CARTON_10PCS -> 10", lambda: uomr.parse_carton_multiplier("CARTON_10PCS") == 10)
test("parser: CARTON_20PCS -> 20", lambda: uomr.parse_carton_multiplier("CARTON_20PCS") == 20)
test("parser: CARTON_40PCS -> 40", lambda: uomr.parse_carton_multiplier("CARTON_40PCS") == 40)
test("parser: CARTON_200PCS -> 200", lambda: uomr.parse_carton_multiplier("CARTON_200PCS") == 200)


def t_conversion_table():
    assert uomr.resolve_output_uom_qty("PCS", 15, True).output_qty == 15
    assert uomr.resolve_output_uom_qty("CARTON_10PCS", 2, True).output_qty == 20
    assert uomr.resolve_output_uom_qty("CARTON_40PCS", 3, True).output_qty == 120
    assert uomr.resolve_output_uom_qty("CARTON_200PCS", 4, True).output_qty == 800
    for uom, qty, expected in (("PCS", 15, 15), ("CARTON_10PCS", 2, 20),
                               ("CARTON_40PCS", 3, 120), ("CARTON_200PCS", 4, 800)):
        assert uomr.resolve_output_uom_qty(uom, qty, True).output_uom == "PCS"


test("PCSx15/CARTON_10PCSx2/CARTON_40PCSx3/CARTON_200PCSx4 conversion table", t_conversion_table)


def t_raw_values_unchanged_regardless_of_convert():
    r_off = uomr.resolve_output_uom_qty("CARTON_10PCS", 2, False)
    assert r_off.output_uom == "CARTON_10PCS" and r_off.output_qty == 2
    # Raw values are never destroyed by requesting conversion -- pcs_per_unit/
    # qty_pcs are ALWAYS derived, independent of convert_to_pcs (spec section 23).
    r_on = uomr.resolve_output_uom_qty("CARTON_10PCS", 2, True)
    assert r_on.pcs_per_unit == 10 and r_on.qty_pcs == 20 == r_off.qty_pcs


test("raw UOM/QTY never destroyed; pcs_per_unit/qty_pcs always derived",
     t_raw_values_unchanged_regardless_of_convert)


def t_malformed_uom_review_no_fabrication():
    for bad in ("CARTON_PCS", "CARTON_XXPCS", "CARTON_-10PCS", "CARTON_0PCS"):
        assert uomr.parse_carton_multiplier(bad) is None, bad
        r = uomr.resolve_output_uom_qty(bad, 5, True)
        assert r.review_flag == uomr.REVIEW_MALFORMED_UOM, (bad, r)
        assert r.output_uom == bad.upper() and r.output_qty == 5, "must preserve raw, never fabricate"


test("malformed UOM (CARTON_PCS/XXPCS/-10PCS/0PCS) -> REVIEW, raw preserved, no fabricated multiplier",
     t_malformed_uom_review_no_fabrication)


# =========================================================================
# Section 44: four-mode Sublist tests
# =========================================================================
print("== V21 section 44: four-mode Sublist (Convert x Show-UOM) ==")


class _FakeItem:
    def __init__(self, product_code, barcode, unit, quantity):
        self.product_code = product_code
        self.barcode = barcode
        self.unit = unit
        self.quantity = quantity


class _FakePkg:
    def __init__(self, items):
        self.items = items
        self.pl_gross_weight = ""
        self.weight = None
        self.shipping_mark = "CN-1000-PVG-CNWorld-POP"
        self.reference_code = "CN-1000-PVG-CNWorld-POP"
        self.shipping_mark_source = "PDF_STANDALONE_CODE"
        self.store_display = "CNWorld"
        self.or_number = "OR1"
        self.or_source = "OR_LIST"
        self.so_number = "SO1"
        self.so_source = "OR_LIST"
        self.business_fields = {}
        self.package_code = "PKG0001"
        self.carton_sequence = 1
        self.carton_total = 1
        self.carton_display = "1/1"
        self.global_carton_num = "1/1"
        self.source_file = "test.pdf"
        self.pdf_package_seq = "1/1"
        self.counting_scope_key = ""


def _sample_pkg():
    return _FakePkg([_FakeItem("SKU-A", "EAN-A", "CARTON_10PCS", 2),
                     _FakeItem("SKU-B", "EAN-B", "PCS", 15)])


def t_mode_a_off_off():
    model = pse.build_sublist_carton_model(_sample_pkg(), convert_to_pcs=False)
    assert [i.uom for i in model.items] == ["CARTON_10PCS", "PCS"]
    assert [i.qty for i in model.items] == [2, 15]
    assert model.total_qty == 17
    assert model.store_display == "CNWorld"  # resolved internally, never displayed by the writer


test("Mode A (Convert OFF / Show-UOM OFF): raw UOM/QTY, SKU|EAN|QTY headers, Store absent from output columns",
     t_mode_a_off_off)


def t_mode_b_off_on():
    model = pse.build_sublist_carton_model(_sample_pkg(), convert_to_pcs=False)
    assert [i.uom for i in model.items] == ["CARTON_10PCS", "PCS"]
    assert [i.qty for i in model.items] == [2, 15]
    assert model.total_qty == 17
    assert pse.ITEM_HEADERS_WITH_UOM == ["Item No.", "EAN", "UOM", "QTY"]  # UOM sits between EAN and QTY


test("Mode B (Convert OFF / Show-UOM ON): UOM column between EAN and QTY, raw values", t_mode_b_off_on)


def t_mode_c_on_off():
    model = pse.build_sublist_carton_model(_sample_pkg(), convert_to_pcs=True)
    assert [i.uom for i in model.items] == ["PCS", "PCS"]
    assert [i.qty for i in model.items] == [20, 15]
    assert model.total_qty == 35


test("Mode C (Convert ON / Show-UOM OFF): CARTON_10PCSx2->20 PCS, total 35", t_mode_c_on_off)


def t_mode_d_on_on():
    model = pse.build_sublist_carton_model(_sample_pkg(), convert_to_pcs=True)
    assert [i.uom for i in model.items] == ["PCS", "PCS"]
    assert [i.qty for i in model.items] == [20, 15]
    assert model.total_qty == 35


test("Mode D (Convert ON / Show-UOM ON): same converted values, UOM column shown", t_mode_d_on_on)


def t_sublist_workbook_all_four_modes_generate_and_never_show_store():
    import tempfile, openpyxl
    pkg = _sample_pkg()
    for convert, show_uom in ((False, False), (False, True), (True, False), (True, True)):
        with tempfile.TemporaryDirectory() as td:
            out = Path(td) / "s.xlsx"
            result = pse.generate_sublist_workbook([pkg], out, convert_to_pcs=convert, show_uom=show_uom)
            assert out.exists()
            wb = openpyxl.load_workbook(out)
            ws = wb.active
            all_values = [ws.cell(row=r, column=c).value for r in range(1, 40) for c in range(1, 6)]
            # Store must remain absent from the Sublist in ALL four modes
            # (spec section 44 / v20 display-only omission preserved).
            assert "CNWorld" not in [v for v in all_values if isinstance(v, str)], \
                f"Store leaked into Sublist output (convert={convert}, show_uom={show_uom})"


test("Sublist Excel: all four modes generate; Store never appears in any of them",
     t_sublist_workbook_all_four_modes_generate_and_never_show_store)


def t_default_no_checkbox_selected_backward_compatible():
    # Section 28: default (neither checkbox) must still correctly preserve
    # BOTH PCS and CARTON_<N>PCS internally, and the Sublist stays exactly
    # SKU|EAN|QTY with no UOM column and no conversion.
    model = pse.build_sublist_carton_model(_sample_pkg())  # no kwargs at all
    assert [i.uom for i in model.items] == ["CARTON_10PCS", "PCS"]
    assert [i.qty for i in model.items] == [2, 15]


test("default (no checkbox args at all) == Mode A, fully backward compatible",
     t_default_no_checkbox_selected_backward_compatible)


# =========================================================================
# Integration: routing_rules wired into classify_packages_for_port() /
# match_store_and_or_v21() against the REAL production shapes
# =========================================================================
print("== V21 integration: classify_packages_for_port(routing_rules=...) ==")


def t_cn6557_via_routing_rules_matches_production_counts():
    # Same real CN-6557 Shipmark shapes as test_v15_bugfix_regression.py,
    # but routed via a USER routing-rule table instead of STORE_MASTER
    # (spec section 40) -- must reproduce the exact same PEK=6/PVG=31/
    # SZX=11 (VN 1/6/2) production counts.
    stores = [
        ("CNWORLD", "PEK", 5, 1), ("GUANGZHOU", "SZX", 4, 1), ("HANGZHOU", "PVG", 4, 1),
        ("IAPM", "PVG", 5, 1), ("KERRY", "PVG", 5, 2), ("SHANGHAI_HONGQIAO", "PVG", 7, 1),
        ("SHANGHAI_TAIKOOLI", "PVG", 4, 1), ("SHENZHEN", "SZX", 5, 1),
    ]
    routing_rules = [{"country": "CN", "port": port, "store": store} for store, port, _, _ in stores]

    packages = []
    seq = 0
    for store_token, port, n_cn, n_vn in stores:
        for _ in range(n_cn):
            seq += 1
            packages.append(_pkg(f"CN-6557-{store_token}_{port}_CN", seq))
        for _ in range(n_vn):
            seq += 1
            packages.append(_pkg(f"CN-6557-{store_token}_{port}-VN", seq))
    assert len(packages) == 48

    classify_packages_for_port(packages, None, False, routing_rules=routing_rules)

    port_counts = Counter(p.port for p in packages)
    assert port_counts == {"PEK": 6, "PVG": 31, "SZX": 11}, dict(port_counts)
    assert all(p.route_match_status == "MATCHED" for p in packages)
    assert all(p.store != "REVIEW" for p in packages)

    vn_packages = [p for p in packages if p.reference_code.upper().endswith("VN")]
    vn_port_counts = Counter(p.port for p in vn_packages)
    assert vn_port_counts == {"PEK": 1, "PVG": 6, "SZX": 2}, dict(vn_port_counts)


test("CN-6557 real production counts reproduced via routing_rules (48 cartons, PEK6/PVG31/SZX11, VN 1/6/2)",
     t_cn6557_via_routing_rules_matches_production_counts)


def t_tmall_classify_and_or_match_end_to_end():
    routing_rules = [{"country": "CN", "port": "", "store": "Tmall"}]
    packages = [_pkg("CN_TMALL-6676-CTN_POP", 1), _pkg("CN_TMALL-6676-PCS_QF", 2)]
    classify_packages_for_port(packages, None, False, routing_rules=routing_rules)
    for p in packages:
        assert p.route_match_status == "MATCHED", p.route_match_reason
        assert p.store == "Tmall"
        assert p.port == ""

    # 03_CN_BY_PORT membership: is_cn_port_eligible() is still True (country
    # ==CN), but PORT_FILE_MAP has no "" key, so blank-Port packages are
    # never written to a physical 03_CN_BY_PORT file -- that's the natural
    # mechanism, no special-cased "if store == Tmall" branch anywhere.
    assert pge.is_cn_port_eligible(packages[0]) is True
    assert pge.PORT_FILE_MAP.get(packages[0].port) is None

    # OR/Ref: a single-row OR List (Shop=Tmall/OR=OR1075/Ref=to10880) must
    # resolve via match_store_and_or_v21(), using the ALREADY-resolved
    # V21 Store, not the legacy ~9-store CN alias table.
    class _OrRow:
        def __init__(self):
            self.store_raw = "Tmall"
            self.business_fields = {"OR": "OR1075", "Ref": "to10880"}
    or_index = {"OR1075": [_OrRow()]}
    m = pge.match_store_and_or_v21(packages[0], or_index)
    assert m.status == "OK"
    assert m.matched_or == "OR1075"
    assert m.matched_so == "to10880"


test("CN|blank|Tmall: classify + is_cn_port_eligible/PORT_FILE_MAP exclusion + OR/Ref v21 match",
     t_tmall_classify_and_or_match_end_to_end)


def t_or_v21_unique_fallback_when_store_blank():
    # spec section 13/19: Store blank in the routed package, but the OR
    # List has exactly one unique business record -> still resolves.
    pkg = _pkg("SG-500-WAREHOUSE-A", 1, country="SG")
    pkg.store = ""  # e.g. a `SG | blank | blank` routing rule -- no Store at all

    class _OrRow:
        def __init__(self):
            self.store_raw = ""
            self.business_fields = {"OR": "OR9001", "Ref": "ref-only-one"}
    or_index = {"OR9001": [_OrRow()]}
    m = pge.match_store_and_or_v21(pkg, or_index)
    assert m.status == "OK"
    assert m.matched_or == "OR9001"
    assert m.matched_so == "ref-only-one"


test("V21 OR/Ref: unique single-record fallback applies even when routed Store is blank",
     t_or_v21_unique_fallback_when_store_blank)


def t_or_v21_review_when_multiple_records_and_no_store_match():
    pkg = _pkg("CN_UNKNOWNSTORE-1", 1)
    pkg.store = "SomeStoreNotInOrList"

    class _OrRow:
        def __init__(self, store, or_, ref):
            self.store_raw = store
            self.business_fields = {"OR": or_, "Ref": ref}
    or_index = {"OR1": [_OrRow("StoreA", "OR1", "R1")], "OR2": [_OrRow("StoreB", "OR2", "R2")]}
    m = pge.match_store_and_or_v21(pkg, or_index)
    assert m.status == "REVIEW"


test("V21 OR/Ref: REVIEW (never a guess) when Store doesn't match and multiple distinct records exist",
     t_or_v21_review_when_multiple_records_and_no_store_match)


# =========================================================================
# Static-audit gate fixes (V21 final round):
#   1. Generic 2-letter Country Code must route end-to-end (not just a
#      finite whitelist) -- PH/TH/AU/MY/ID/SG/CN all exercised through the
#      REAL detector (resolve_package_country/detect_shipment_country),
#      not a pre-set pkg.country shortcut.
#   2. NO SILENT SKIP when routing_rules is supplied: every package must
#      get an explicit route diagnostic, never a quiet `continue` that
#      leaves route_match_status/method/reason untouched.
#   5. Legacy STORE_MASTER fallback is allowed ONLY when the whole run has
#      no routing rules at all.
# =========================================================================
print("== V21 static-audit gate: generic country + no-silent-skip ==")

resolve_package_country = D["resolve_package_country"]
detect_shipment_country = D["detect_shipment_country"]


def _pkg_with_real_country_resolution(mark: str, seq: int) -> "Package":
    # Deliberately does NOT pre-set p.country -- goes through the exact
    # same resolve_package_country() the real pipeline calls, so this
    # proves the detector itself, not just classify_packages_for_port's
    # consumption of an already-known-good country.
    p = Package(package_code=f"PKG{seq:03d}", source_file=f"{mark}.pdf",
                reference_code=mark, pdf_package_seq=0)
    p.shipping_mark = mark
    p.shipping_mark_source = "PDF_STANDALONE_CODE"
    p.filename_reference = mark
    resolve_package_country(p)
    return p


_GENERIC_COUNTRY_CODES = ["PH", "TH", "AU", "MY", "ID", "SG", "CN"]


def t_generic_2letter_country_detected_end_to_end():
    for code in _GENERIC_COUNTRY_CODES:
        pkg = _pkg_with_real_country_resolution(f"{code}-1234-GenericHub-POP", 1)
        assert pkg.country == code, f"{code}: detected {pkg.country!r} via real resolve_package_country()"


test("Generic 2-letter Country Code detector resolves ALL of PH/TH/AU/MY/ID/SG/CN end-to-end "
     "(structural detector, not a finite whitelist)",
     t_generic_2letter_country_detected_end_to_end)


def t_generic_2letter_country_routes_end_to_end():
    # Full path: real country resolution -> routing_rules (one rule per
    # country, arbitrary/never-seen-before codes included) ->
    # classify_packages_for_port() -> MATCHED with the correct Store.
    seq = 0
    packages = []
    for code in _GENERIC_COUNTRY_CODES:
        seq += 1
        packages.append(_pkg_with_real_country_resolution(f"{code}-1234-GenericHub-POP", seq))
    routing_rules = [{"country": code, "port": "", "store": "GenericHub"} for code in _GENERIC_COUNTRY_CODES]
    classify_packages_for_port(packages, None, False, routing_rules=routing_rules)
    for pkg, code in zip(packages, _GENERIC_COUNTRY_CODES):
        assert pkg.route_match_status == "MATCHED", (code, pkg.route_match_reason)
        assert pkg.store == "GenericHub", (code, pkg.store)


test("Any valid 2-letter Country Code used in a manual routing rule is routable end-to-end "
     "(PH/TH/AU/MY/ID/SG/CN all matched to their Store via one generic rule table)",
     t_generic_2letter_country_routes_end_to_end)


def t_trailing_factory_suffix_never_read_as_destination_country():
    # CN/VN/POP/QF etc as a TRAILING factory/origin suffix must never be
    # misread as the destination Country -- only the LEADING structural
    # prefix counts (detector anchors on ^, never scans the suffix).
    cases = [
        ("TH-1000-SomeHub-CN", "TH"),
        ("MY-2000-SomeHub-VN", "MY"),
        ("ID-3000-SomeHub-POP", "ID"),
        ("PH-4000-SomeHub-QF", "PH"),
    ]
    for mark, expected in cases:
        pkg = _pkg_with_real_country_resolution(mark, 1)
        assert pkg.country == expected, f"{mark}: got {pkg.country!r}, trailing factory suffix must not win"


test("Trailing factory/origin suffix (CN/VN/POP/QF) is never misread as destination Country "
     "(only the leading structural prefix counts)",
     t_trailing_factory_suffix_never_read_as_destination_country)


def t_no_silent_skip_blank_country_is_review_not_skipped():
    pkg = Package(package_code="PKG900", source_file="blank-country.pdf",
                  reference_code="UNRECOGNIZABLE_MARK_NO_PREFIX", pdf_package_seq=0)
    pkg.shipping_mark = "UNRECOGNIZABLE_MARK_NO_PREFIX"
    pkg.shipping_mark_source = "PDF_STANDALONE_CODE"
    pkg.country = ""  # never resolved
    routing_rules = [{"country": "CN", "port": "", "store": "Tmall"}]
    classify_packages_for_port([pkg], None, False, routing_rules=routing_rules)
    assert pkg.route_match_status == "REVIEW", pkg.route_match_status
    assert pkg.route_match_method == prr.METHOD_COUNTRY_UNRESOLVED, pkg.route_match_method
    assert pkg.route_match_reason, "reason must be explicit, never blank"
    assert pkg.store == "REVIEW"


test("Blank pkg.country with routing_rules supplied -> explicit REVIEW/COUNTRY_UNRESOLVED, "
     "never a silent skip (spec gap-fix #2)",
     t_no_silent_skip_blank_country_is_review_not_skipped)


def t_no_silent_skip_country_with_no_rule_is_review_not_skipped():
    # pkg.country resolves fine (TH), but the routing table only has a
    # rule for CN -- must still get an explicit diagnostic, not a skip.
    pkg = _pkg_with_real_country_resolution("TH-9000-SomeHub-POP", 1)
    routing_rules = [{"country": "CN", "port": "", "store": "Tmall"}]
    classify_packages_for_port([pkg], None, False, routing_rules=routing_rules)
    assert pkg.route_match_status == "REVIEW", pkg.route_match_status
    assert pkg.route_match_method == prr.METHOD_NO_RULE_FOR_COUNTRY, pkg.route_match_method
    assert pkg.route_match_reason, "reason must be explicit, never blank"
    assert pkg.store == "REVIEW"


test("pkg.country resolved but no routing rule exists for it -> explicit REVIEW/NO_RULE_FOR_COUNTRY, "
     "never a silent skip (spec gap-fix #2)",
     t_no_silent_skip_country_with_no_rule_is_review_not_skipped)


def t_mixed_run_some_matched_some_review_none_silently_skipped():
    # One package per outcome, all in the SAME run: MATCHED, blank-country
    # REVIEW, and no-rule-for-country REVIEW -- proves the loop handles all
    # three without ever leaving a package's route fields untouched.
    matched_pkg = _pkg_with_real_country_resolution("CN-1-SomeHub-POP", 1)
    blank_country_pkg = Package(package_code="PKG2", source_file="x.pdf",
                                 reference_code="NOPREFIXATALL", pdf_package_seq=0)
    blank_country_pkg.shipping_mark = "NOPREFIXATALL"
    blank_country_pkg.shipping_mark_source = "PDF_STANDALONE_CODE"
    blank_country_pkg.country = ""
    no_rule_pkg = _pkg_with_real_country_resolution("TH-3-SomeHub-POP", 3)

    packages = [matched_pkg, blank_country_pkg, no_rule_pkg]
    routing_rules = [{"country": "CN", "port": "", "store": "SomeHub"}]
    classify_packages_for_port(packages, None, False, routing_rules=routing_rules)

    assert matched_pkg.route_match_status == "MATCHED"
    assert blank_country_pkg.route_match_status == "REVIEW"
    assert blank_country_pkg.route_match_method == prr.METHOD_COUNTRY_UNRESOLVED
    assert no_rule_pkg.route_match_status == "REVIEW"
    assert no_rule_pkg.route_match_method == prr.METHOD_NO_RULE_FOR_COUNTRY
    # every single package got an explicit non-blank status -- none skipped
    assert all(p.route_match_status in ("MATCHED", "REVIEW") for p in packages)


test("Mixed run (matched + blank-country + no-rule-for-country) -- every package gets an explicit "
     "diagnostic, none silently skipped",
     t_mixed_run_some_matched_some_review_none_silently_skipped)


def t_legacy_fallback_only_when_no_routing_rules_at_all():
    # requirement #5: legacy STORE_MASTER path only runs when routing_rules
    # is None/empty for the WHOLE run -- route_match_status stays "" (never
    # touched) because the legacy path doesn't know about V21 diagnostics
    # at all, which is the existing, unchanged behaviour being locked in.
    pkg = _pkg("CN-1666-PVG-KERRY-POP", 1)
    classify_packages_for_port([pkg], None, False, routing_rules=None)
    assert pkg.route_match_status == "", pkg.route_match_status
    assert pkg.store == "KERRY", pkg.store
    assert pkg.port == "PVG", pkg.port


test("No routing rules at all -> legacy STORE_MASTER fallback runs exactly as before "
     "(route_match_status stays blank, spec gap-fix #5)",
     t_legacy_fallback_only_when_no_routing_rules_at_all)


# ── summary ──────────────────────────────────────────────────────────────
print(f"\n{_passed} passed, {_failed} failed")
if _failed:
    print("FAILED tests present -- see above.")
    sys.exit(1)
sys.exit(0)
