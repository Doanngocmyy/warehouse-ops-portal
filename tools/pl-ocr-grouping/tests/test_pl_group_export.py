#!/usr/bin/env python3
"""
Regression tests for pl_group_export.py -- previously 0% covered (see
Phase-1 audit). Same no-framework convention as test_pl_ocr_core.py: run
with `python3 tools/pl-ocr-grouping/tests/test_pl_group_export.py`.
"""
from __future__ import annotations
import sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
TOOL_DIR = HERE.parent
if str(TOOL_DIR) not in sys.path:
    sys.path.insert(0, str(TOOL_DIR))

import pl_group_export as pge

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


print("== pl_group_export.detect_factory ==")


def t_factory_suffix_last_token_wins():
    assert pge.detect_factory("CN-2569_SH_PVG_POP", "x.pdf") == "POP", \
        "leading CN must NOT be read as factory -- only the final suffix counts"


def t_factory_plain_cn_suffix():
    assert pge.detect_factory("SYN-1001-WarehouseAlpha-CN", "x.pdf") == "CN"


def t_factory_split_spelling_sbgear():
    assert pge.detect_factory("Kerry_SB_GEAR", "x.pdf") == "SBGEAR"


def t_factory_split_spelling_qifeng():
    assert pge.detect_factory("Kerry_QI_FENG", "x.pdf") == "QIFENG"


def t_factory_flat_suffix_fallback():
    assert pge.detect_factory("TW8785SBGEAR", "x.pdf") == "SBGEAR"


def t_factory_falls_back_to_filename():
    # reference_code has no recognisable suffix at all -> try filename stem
    assert pge.detect_factory("random-code-xyz", "Kerry_POP.pdf") == "POP"


def t_factory_unclassifiable_is_review():
    assert pge.detect_factory("totally-unknown-code", "unknown.pdf") == "REVIEW"


test("leading 'CN' token is not misread as factory=CN (only final suffix counts)",
     t_factory_suffix_last_token_wins)
test("plain '-CN' suffix -> factory CN", t_factory_plain_cn_suffix)
test("split spelling 'SB_GEAR' -> SBGEAR", t_factory_split_spelling_sbgear)
test("split spelling 'QI_FENG' -> QIFENG", t_factory_split_spelling_qifeng)
test("flat no-separator suffix 'TW8785SBGEAR' -> SBGEAR", t_factory_flat_suffix_fallback)
test("reference_code with no suffix falls back to filename stem", t_factory_falls_back_to_filename)
test("unclassifiable code -> REVIEW (never guessed)", t_factory_unclassifiable_is_review)

print("== pl_group_export.match_store ==")


def t_store_exact_alias_high_confidence():
    store, conf, sug = pge.match_store("Shop NB1-23B, Jing'an Kerry Centre")
    assert store == "KERRY", store
    assert conf == 1.0, conf


def t_store_no_signal_is_review():
    store, conf, sug = pge.match_store("")
    assert store == "REVIEW"
    assert conf == 0.0


def t_store_weak_signal_is_review_not_guessed():
    store, conf, sug = pge.match_store("some random shipping text with no store info at all")
    assert store == "REVIEW", (store, conf, sug)


test("exact alias substring -> store matched with confidence 1.0", t_store_exact_alias_high_confidence)
test("empty signal -> REVIEW, confidence 0.0", t_store_no_signal_is_review)
test("weak/no signal -> REVIEW rather than a low-confidence guess", t_store_weak_signal_is_review_not_guessed)

print("== pl_group_export.CARTON_FACTORY_ORDER_WITH_CO ==")


def t_carton_factory_order_matches_spec():
    assert pge.CARTON_FACTORY_ORDER_WITH_CO == ["POP", "SBGEAR", "QIFENG", "JION", "CN"], \
        pge.CARTON_FACTORY_ORDER_WITH_CO


def t_carton_factory_order_all_vn_before_cn():
    order = pge.CARTON_FACTORY_ORDER_WITH_CO
    cn_idx = order.index("CN")
    for vn_factory in ("POP", "SBGEAR", "QIFENG", "JION"):
        assert order.index(vn_factory) < cn_idx, f"{vn_factory} must precede CN"


test("CARTON_FACTORY_ORDER_WITH_CO matches spec order exactly", t_carton_factory_order_matches_spec)
test("every VN factory precedes CN (satisfies the 'no CO: VN before CN' rule for free)",
     t_carton_factory_order_all_vn_before_cn)

print(f"\n{_passed} passed, {_failed} failed")
if _failed:
    sys.exit(1)
