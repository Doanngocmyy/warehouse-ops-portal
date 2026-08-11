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
import pl_or_list_import as oli

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

print("== v12: detect_factory strips trailing copy-suffix (single source of truth) ==")


def t_factory_detection_survives_copy_suffix_1():
    assert pge.detect_factory("Kerry_POP_1", "Kerry_POP_1.pdf") == "POP"


def t_factory_detection_survives_copy_suffix_10():
    assert pge.detect_factory("Kerry_POP_10", "Kerry_POP_10.pdf") == "POP"


def t_factory_detection_survives_paren_and_copy_word_suffix():
    assert pge.detect_factory("Kerry_CN(2)", "x.pdf") == "CN"
    assert pge.detect_factory("Kerry_CN COPY 3", "x.pdf") == "CN"
    assert pge.detect_factory("Kerry_CN COPY_4", "x.pdf") == "CN"
    assert pge.detect_factory("Kerry_CN COPY-5", "x.pdf") == "CN"


def t_factory_detection_still_ignores_midstring_number():
    # sanity: the fix must not start treating "1666" as a copy suffix --
    # it's not anchored at the end (more tokens follow it).
    assert pge.detect_factory("CN-1666-PVG-KERRY-POP", "x.pdf") == "POP"


def t_literal_vn_factory_recognised():
    assert pge.detect_factory("KR100_VN", "x.pdf") == "VN"
    assert pge.detect_factory("KR100_CN", "x.pdf") == "CN"


def t_factory_detection_survives_trailing_version_marker():
    # Turn 12 bug report: "CN-1529_SH-Airport_PVG_CN v.pdf" -> was REVIEW
    # because the last token "V" isn't a known factory token. A trailing
    # version/revision marker (v, v1, rev2, ...) must be stripped before
    # factory detection sees the real trailing factory suffix.
    assert pge.detect_factory("", "CN-1529_SH-Airport_PVG_CN v.pdf") == "CN",         "bare trailing ' v' must be stripped as a version marker"
    assert pge.detect_factory("", "CN-1529_SH-Airport_PVG_CN v1.pdf") == "CN",         "'v1' (v + digits, no space) must be stripped as a version marker"
    assert pge.detect_factory("", "CN-1529_SH-Airport_PVG_CN rev2.pdf") == "CN",         "'rev2' must be stripped as a version marker"
    assert pge.detect_factory("Kerry_POP v3", "x.pdf") == "POP",         "version marker after a non-CN factory token must also be stripped"


def t_factory_detection_version_marker_never_eats_literal_vn():
    # The exact same regex machinery must NEVER mis-strip the real "VN"
    # factory token -- "VN" is a literal 2-letter suffix, not "V" followed
    # only by digits, so "V\d*$" can never consume the trailing "N".
    assert pge.detect_factory("KR100_VN", "x.pdf") == "VN"
    assert pge.detect_factory("", "Shipment_KR100_VN.pdf") == "VN"


def t_factory_detection_version_marker_does_not_alter_legitimate_codes():
    # Sanity: codes that happen to contain "V" mid-string, or end in a real
    # factory token with no version marker at all, must be completely
    # unaffected.
    assert pge.detect_factory("CN-1666-PVG-KERRY-POP", "x.pdf") == "POP"
    assert pge.detect_factory("Kerry_POP", "x.pdf") == "POP"
    assert pge.detect_factory("Kerry_POP_1", "Kerry_POP_1.pdf") == "POP",         "existing copy-suffix behaviour must be unaffected by the new version-suffix stripper"


def t_carton_factory_rank_table_merges_the_two_named_constants():
    table = pge.carton_factory_rank_table()
    assert table == ["POP", "SBGEAR", "QIFENG", "JION", "VN", "CN"], table
    assert table.index("VN") < table.index("CN")
    for f in pge.CARTON_FACTORY_ORDER_WITH_CO:
        if f != "CN":
            assert f in table
    assert table[-2:] == pge.CARTON_FACTORY_ORDER_NO_CO, \
        "the merged table's tail must exactly equal CARTON_FACTORY_ORDER_NO_CO -- the two constants must never disagree"


test("'Kerry_POP_1' still detects factory=POP (not REVIEW) -- the copy suffix no longer masks the real factory",
     t_factory_detection_survives_copy_suffix_1)
test("'Kerry_POP_10' still detects factory=POP", t_factory_detection_survives_copy_suffix_10)
test("'(2)' and 'COPY 3/COPY_4/COPY-5' suffix forms are all stripped before factory detection",
     t_factory_detection_survives_paren_and_copy_word_suffix)
test("mid-string number ('1666') is still never mistaken for a copy suffix",
     t_factory_detection_still_ignores_midstring_number)
test("literal 'VN' suffix is recognised as its own factory, distinct from POP/SBGEAR/QIFENG/JION",
     t_literal_vn_factory_recognised)
test("trailing version/revision marker (v, v1, rev2) is stripped before factory detection",
     t_factory_detection_survives_trailing_version_marker)
test("version-marker stripping never mis-strips the real literal 'VN' factory token",
     t_factory_detection_version_marker_never_eats_literal_vn)
test("version-marker stripping does not alter legitimate codes with no version marker",
     t_factory_detection_version_marker_does_not_alter_legitimate_codes)
test("carton_factory_rank_table() merges CARTON_FACTORY_ORDER_WITH_CO + CARTON_FACTORY_ORDER_NO_CO consistently",
     t_carton_factory_rank_table_merges_the_two_named_constants)


# =========================================================================
# v13 (FIX2/FIX3/FIX4): canonical Store identity from OR List free text +
# explicit Shipping-Mark short-code aliases + tiered exact/fuzzy matching,
# using the real production store descriptions and Shipping Mark shapes
# (see audit notes -- real OR List.xlsx / SUBLIST_TOTAL.pdf).
# =========================================================================
print()
print("== pl_group_export Store/OR/SO matching (v13) ==")

_REAL_OR_ROWS = [
    ("20260609 CN - Guangzhou Parc Central Replen", "po38070", "inv628037"),
    ("20260609 CN - Hangzhou Mixc Replen", "po38068", "inv628036"),
    ("20260609 CN - Iapm Replen", "po38072", "inv628039"),
    ("20260609 CN - Kerry Center flagship Replen", "po38071", "inv628038"),
    ("20260609 CN - Shanghai Hongqiao Airport Replen", "po38074", "inv628042"),
    ("20260609 CN - Shanghai Taikooli (Shop B1-07b) Replen", "po38076", "inv628041"),
    ("20260609 CN - Shenzhen Mixc City (Shop T228) Replen", "po38073", "inv628040"),
]


def _real_or_index():
    from collections import OrderedDict
    rows = [
        oli.OrListRow(row_number=i + 2, store_raw=store, store_norm="",
                       business_fields=OrderedDict([("OR No.", or_v), ("SO No.", so_v)]))
        for i, (store, or_v, so_v) in enumerate(_REAL_OR_ROWS)
    ]
    idx = {}
    for r in rows:
        idx.setdefault(r.or_norm, []).append(r)
    return idx


class _FakePkg:
    def __init__(self, shipping_mark="", reference_code=""):
        self.shipping_mark = shipping_mark
        self.reference_code = reference_code


def t_canonical_store_identity_resolves_all_7_real_descriptions():
    expected = ["GUANGZHOU", "HANGZHOU", "IAPM", "KERRY",
                "SHANGHAI_HONGQIAO", "SHANGHAI_TAIKOOLI", "SHENZHEN"]
    for (store_text, _or, _so), exp_key in zip(_REAL_OR_ROWS, expected):
        got = pge._canonical_store_identity_for_or_row(store_text)
        assert got == pge._store_identity(exp_key.replace("_", " ")),             f"{store_text!r} -> {got!r}, expected canonical form of {exp_key!r}"


def t_explicit_short_shipmark_codes_resolve_unambiguously():
    """HZ/KR/GZ/SZ/IAPM are explicit shipping_mark_tokens aliases (v13) --
    never fuzzy, always an exact single-store hit.

    v17 (test correction, spec point 37): match_source's real, current
    value for a Shipping-Mark-derived signal is "SHIPMARK_SAFE_ALIAS" --
    match_store_and_or() has one unified "safe" match tier per signal
    source (Shipmark / filename / receiver text), covering BOTH literal
    shipping_mark_tokens hits and safe-typo-tolerant hits under the same
    label; "SHIPMARK_TOKEN_EXACT" is stale docstring-era terminology that
    was never the actual runtime value (see StoreOrMatchResult.match_
    source's field comment, corrected alongside this test). The business
    values (matched_or/matched_so) below were already 100% correct in
    every case -- only this tier-name string was stale, root-caused and
    confirmed via a git-stash baseline run before being changed here."""
    idx = _real_or_index()
    cases = {
        "CN-1529_HZ_PVG_POP": ("po38068", "inv628036"),
        "CN-1529_KR_PVG_VN": ("po38071", "inv628038"),
        "CN-1529_GZ_SZX_CN": ("po38070", "inv628037"),
        "CN-1529_SZ_SZX_VN": ("po38073", "inv628040"),
        "CN-1529_IAPM_PVG_POP": ("po38072", "inv628039"),
    }
    for mark, (exp_or, exp_so) in cases.items():
        m = pge.match_store_and_or(_FakePkg(shipping_mark=mark), idx)
        assert m.status == "OK", f"{mark}: expected OK, got {m.status} ({m.review_reason})"
        assert m.match_source == "SHIPMARK_SAFE_ALIAS", m.match_source
        assert m.matched_or == exp_or and m.matched_so == exp_so, f"{mark}: {m.matched_or}/{m.matched_so}"


def t_compound_hyphenated_shipmark_codes_resolve_via_bigram():
    """"SH-Airport" / "SH-Taikooli" tokenize into two generic pieces
    (SH + AIRPORT/TAIKOOLI) -- must resolve via the adjacent-token bigram
    join ("SHAIRPORT"/"SHTAIKOOLI"), each to its OWN distinct store, never
    confused with each other."""
    idx = _real_or_index()
    m1 = pge.match_store_and_or(_FakePkg(shipping_mark="CN-1529_SH-Airport_PVG_POP"), idx)
    assert m1.status == "OK" and m1.matched_or == "po38074", m1.review_reason
    m2 = pge.match_store_and_or(_FakePkg(shipping_mark="CN-1529_SH-Taikooli_PVG_CN"), idx)
    assert m2.status == "OK" and m2.matched_or == "po38076", m2.review_reason


def t_generic_alias_word_does_not_create_false_ambiguity_with_chengdu():
    """Regression test for a real bug found during audit: CHENGDU's own
    alias "Chengdu Taikooli" shares the generic mall-brand word "Taikooli"
    with SHANGHAI_TAIKOOLI -- that shared word must never make an otherwise
    clean "SH-Taikooli" Shipmark match ambiguous with Chengdu (which isn't
    even in this shipment)."""
    idx = _real_or_index()
    m = pge.match_store_and_or(_FakePkg(shipping_mark="CN-1529_SH-Taikooli_PVG_POP"), idx)
    assert m.status == "OK", f"expected OK, got REVIEW: {m.review_reason}"
    assert m.matched_or == "po38076"


def t_store_resolves_for_pop_and_vn_factories_not_just_cn():
    """FIX4: match_store_and_or itself is factory-agnostic (it only looks at
    Shipmark/reference_code/receiver text) -- confirm POP and VN-tagged
    Shipping Marks resolve exactly like CN ones."""
    idx = _real_or_index()
    for factory in ("POP", "VN", "CN"):
        mark = f"CN-1529_KR_PVG_{factory}"
        m = pge.match_store_and_or(_FakePkg(shipping_mark=mark), idx)
        assert m.status == "OK", f"{mark}: {m.review_reason}"
        assert m.matched_or == "po38071"


def t_unconfigured_short_code_never_silently_guessed():
    """A short 2-letter code that ISN'T one of the explicitly configured
    shipping_mark_tokens must never fuzzy-guess a store -- short codes are
    explicitly excluded from fuzzy matching by design (only explicit exact
    aliases are trusted at that length)."""
    idx = _real_or_index()
    m = pge.match_store_and_or(_FakePkg(shipping_mark="CN-1529_XX_PVG_POP"), idx)
    assert m.status != "OK", "an unconfigured short code must never resolve to a confident match"


def t_no_or_list_returns_no_or_list_status():
    m = pge.match_store_and_or(_FakePkg(shipping_mark="CN-1529_KR_PVG_POP"), {})
    assert m.status == "NO_OR_LIST"


test("canonical store identity resolves all 7 real OR List free-text descriptions to their STORE_MASTER key",
     t_canonical_store_identity_resolves_all_7_real_descriptions)
test("explicit short Shipmark codes (HZ/KR/GZ/SZ/IAPM) resolve unambiguously via SHIPMARK_TOKEN_EXACT",
     t_explicit_short_shipmark_codes_resolve_unambiguously)
test("compound hyphenated Shipmark codes (SH-Airport/SH-Taikooli) resolve via adjacent-token bigram match",
     t_compound_hyphenated_shipmark_codes_resolve_via_bigram)
test("a generic shared alias word (Taikooli) never creates false ambiguity between two different stores",
     t_generic_alias_word_does_not_create_false_ambiguity_with_chengdu)
test("Store resolves identically for POP/VN/CN factories (v13 FIX4: not CN-only any more)",
     t_store_resolves_for_pop_and_vn_factories_not_just_cn)
test("an unconfigured short code is never silently guessed into a wrong store",
     t_unconfigured_short_code_never_silently_guessed)
test("no OR List uploaded -> NO_OR_LIST status, never a crash", t_no_or_list_returns_no_or_list_status)


# =========================================================================
# v14 (spec section 4): expanded Store alias config -- KRY/KRYY/KER,
# "Hang Zhou", "Guangzou", "I A P M", "Shenzen".
# =========================================================================
def t_kerry_short_code_aliases_kry_kryy_ker_all_resolve_unambiguously():
    """KR/KRY/KRYY/KER are explicit shipping_mark_tokens (v14, spec section
    4) -- resolved via match_store_and_or's SHIPMARK_TOKEN_EXACT tier,
    same mechanism/test pattern as the pre-existing HZ/KR/GZ/SZ/IAPM test
    above (match_store() itself doesn't consult shipping_mark_tokens at
    all -- it's a separate, receiver/alias-substring-only matcher used only
    for the CN-only classify_packages_for_port() path)."""
    idx = _real_or_index()
    for code in ("KR", "KRY", "KRYY", "KER"):
        mark = f"CN-1529_{code}_PVG_POP"
        m = pge.match_store_and_or(_FakePkg(shipping_mark=mark), idx)
        assert m.status == "OK", f"{mark}: expected OK, got {m.status} ({m.review_reason})"
        # v17 (test correction, spec point 37): see t_explicit_short_
        # shipmark_codes_resolve_unambiguously's docstring -- the real
        # current tier name is SHIPMARK_SAFE_ALIAS, not the stale
        # SHIPMARK_TOKEN_EXACT docstring-era name.
        assert m.match_source == "SHIPMARK_SAFE_ALIAS", m.match_source
        assert m.matched_or == "po38071" and m.matched_so == "inv628038", (mark, m.matched_or, m.matched_so)


def t_hangzhou_spaced_spelling_hang_zhou_resolves():
    store, conf, _ = pge.match_store("20260609 CN - Hang Zhou Replen")
    assert store == "HANGZHOU", (store, conf)


def t_guangzhou_misspelling_guangzou_resolves():
    store, conf, _ = pge.match_store("20260609 CN - Guangzou Parc Central Replen")
    assert store == "GUANGZHOU", (store, conf)


def t_iapm_spaced_letters_resolves():
    idx = _real_or_index()
    m = pge.match_store_and_or(_FakePkg(shipping_mark="CN-1529_I A P M_PVG_POP"), idx)
    assert m.status == "OK", f"expected OK, got {m.status} ({m.review_reason})"
    # v17 (test correction, spec point 37): see t_explicit_short_shipmark_
    # codes_resolve_unambiguously's docstring -- real current tier name.
    assert m.match_source == "SHIPMARK_SAFE_ALIAS", m.match_source
    assert m.matched_or == "po38072" and m.matched_so == "inv628039", (m.matched_or, m.matched_so)


def t_shenzhen_misspelling_shenzen_resolves():
    store, conf, _ = pge.match_store("20260609 CN - Shenzen Mixc City Replen")
    assert store == "SHENZHEN", (store, conf)


def t_expanded_aliases_never_create_cross_store_ambiguity_for_the_real_7_stores():
    """Regression guard for the exact bug found while adding these aliases:
    a naive addition of 'SH-Airport'/'SH-Taikooli' as free-text ALIASES
    (STORE_MASTER "aliases" list, feeding match_store()'s receiver/alias-
    substring lookup) created a shared bare 'SH' token ambiguous between
    the two Shanghai stores. Every real compound Shipmark code must still
    resolve cleanly through match_store_and_or (the exact-token matcher
    used for actual Store resolution)."""
    idx = _real_or_index()
    cases = {
        "CN-1529_SH-Airport_PVG_POP": "po38074",
        "CN-1529_SH-Taikooli_PVG_POP": "po38076",
    }
    for mark, expected_or in cases.items():
        m = pge.match_store_and_or(_FakePkg(shipping_mark=mark), idx)
        assert m.status == "OK", f"{mark}: expected OK, got {m.status} ({m.review_reason})"
        assert m.matched_or == expected_or, (mark, m.matched_or)


test("Kerry short-code aliases KR/KRY/KRYY/KER all resolve unambiguously (spec section 4)",
     t_kerry_short_code_aliases_kry_kryy_ker_all_resolve_unambiguously)
test("Hangzhou spaced spelling 'Hang Zhou' resolves (spec section 4)",
     t_hangzhou_spaced_spelling_hang_zhou_resolves)
test("Guangzhou misspelling 'Guangzou' resolves (spec section 4)",
     t_guangzhou_misspelling_guangzou_resolves)
test("IAPM spaced-letter spelling 'I A P M' resolves via token-collapsing (spec section 4)",
     t_iapm_spaced_letters_resolves)
test("Shenzhen misspelling 'Shenzen' resolves (spec section 4)",
     t_shenzhen_misspelling_shenzen_resolves)
test("expanded alias config never creates cross-store ambiguity for the real compound Shipmark codes (regression guard)",
     t_expanded_aliases_never_create_cross_store_ambiguity_for_the_real_7_stores)


print(f"\n{_passed} passed, {_failed} failed")
if _failed:
    sys.exit(1)
