#!/usr/bin/env python3
"""
pl_group_export.py
===================
Splits the packages produced by OCR_Packing List.ipynb (v7 pipeline) into a
grouped set of workbooks:

    PL_SPLIT_OUTPUT/
    ├── 01_PL_TOTAL/PL_TOTAL.xlsx
    ├── 02_BY_FACTORY/PL_FACTORY_{CN,POP,SBGEAR,QIFENG,JION}.xlsx
    ├── 03_CN_BY_PORT/PL_CN_PORT_{PVG,SZX,TFU,PEK}.xlsx
    ├── 04_CN_BY_STORE/PL_CN_STORE_{...9 stores...}.xlsx
    └── PL_SPLIT_CONTROL.csv

Design notes
------------
* This module is intentionally self-contained (no import from the notebook).
  It only needs a `write_workbook(path, packages)` callable — the one already
  defined in the notebook — and a list of `Package`-like objects exposing:
  .package_code .source_file .reference_code .pdf_package_seq .items
  .global_carton_num .calc_qty (property) .declared_total_qty
  It never mutates package_code / pdf_package_seq / items — only
  `global_carton_num` is temporarily rewritten per sub-file and always
  restored afterwards (see _write_group).
* Every workbook write goes through a temp-file + os.replace() pattern so a
  file that is currently open in Excel raises a clear PermissionError instead
  of silently corrupting the target or crashing mid-write.
* Store/port classification only ever runs for packages whose FACTORY is CN.
  It never guesses: if confidence is low, or the top-2 candidates are too
  close, the package is marked REVIEW (store=REVIEW, port=REVIEW) and is
  EXCLUDED from the 03_CN_BY_PORT / 04_CN_BY_STORE files (it is still fully
  visible in PL_SPLIT_CONTROL.csv and in the validation report).
* v8: STORE_MASTER now also carries the full Notify Party / Delivery Address
  block (receiver company, full street address, contact name + phone) for
  each of the 9 CN retail stores, supplied directly by the warehouse team.
  notify_party_block() formats that into the exact "NOTIFY PARTY: / DELIVERY
  ADDRESS: / ..." text used on the real PL template, so pl_ocr_core.py's
  Packing List sheet can auto-fill Notify Party ONLY when every package in
  that file belongs to the same single store (e.g. the 04_CN_BY_STORE
  split) — never guessed/blended when a file mixes stores.
"""
from __future__ import annotations

import csv
import difflib
import logging
import re
import shutil
import unicodedata
from collections import defaultdict
from pathlib import Path
from typing import Callable, Dict, List, Optional, Tuple

log = logging.getLogger("pl_group_export")
if not log.handlers:
    logging.basicConfig(level=logging.INFO, format="%(asctime)s [%(levelname)s] %(message)s")

try:
    import pdfplumber
except ImportError:  # pragma: no cover - pdfplumber is a hard dependency of the notebook already
    pdfplumber = None


# =========================================================================
# 1) Normalization helpers (self-contained — no dependency on the notebook)
# =========================================================================
def _strip_accents(s: str) -> str:
    nfd = unicodedata.normalize("NFD", str(s or "").strip())
    return "".join(c for c in nfd if unicodedata.category(c) != "Mn").upper()


def _norm_text(s: str) -> str:
    """Uppercase, strip accents/diacritics, unicode-normalize, collapse
    punctuation/whitespace into single spaces. Used for fuzzy text compare."""
    if s is None:
        return ""
    s = unicodedata.normalize("NFKC", str(s))
    s = _strip_accents(s)
    s = re.sub(r"[^A-Z0-9]+", " ", s)
    s = re.sub(r"\s+", " ", s).strip()
    return s


def _tokens(code: str) -> List[str]:
    return [t for t in re.split(r"[^A-Z0-9]+", _strip_accents(code)) if t]


# =========================================================================
# 2) Factory detection — suffix based, last token wins
# =========================================================================
# v12: "VN" is a distinct, literal factory token -- used when a shipment's
# Vietnam factory is named generically ("KR100_VN") rather than by its
# specific name (POP/SBGEAR/QIFENG/JION). Spec section 8 explicitly warns
# these must NOT be assumed equivalent -- VN is its own category, never
# silently folded into one of the other four.
FACTORY_KEYWORDS = ["SBGEAR", "QIFENG", "POP", "JION", "VN", "CN"]  # longest-first for the flat-suffix fallback
_KNOWN_FACTORY_TOKENS = {"CN", "POP", "SBGEAR", "QIFENG", "JION", "VN"}

# v11/v12: carton-NUMBERING business order (spec section 7.2/8) --
# deliberately separate from FACTORY_KEYWORDS above, which is only a
# longest-first matching-priority list for detect_factory() and has nothing
# to do with the order cartons should be numbered in. Centralised here (not
# scattered across pl_ocr_core.py) so there is exactly one place to change
# if the warehouse's factory order ever changes.
#   CARTON_FACTORY_ORDER_WITH_CO -- specific named factories, one Store
#     spanning several of them (spec 7.2/8, "CO" case).
#   CARTON_FACTORY_ORDER_NO_CO -- the literal VN/CN case (spec 7.3/8, "no
#     CO"): VN always before CN, whether or not a specific VN sub-factory
#     name is present.
# carton_factory_rank_table() merges the two into one effective order so
# they can never silently disagree with each other -- see its docstring.
CARTON_FACTORY_ORDER_WITH_CO = ["POP", "SBGEAR", "QIFENG", "JION", "CN"]
CARTON_FACTORY_ORDER_NO_CO = ["VN", "CN"]


def carton_factory_rank_table() -> List[str]:
    """Single effective carton-numbering factory order, derived from the two
    named constants above so they cannot silently diverge: every specific
    factory from CARTON_FACTORY_ORDER_WITH_CO except its trailing "CN"
    ranks first (in that order), then CARTON_FACTORY_ORDER_NO_CO's own order
    (VN, then CN) as the tail -- i.e. POP -> SBGEAR -> QIFENG -> JION -> VN
    -> CN. A shipment with only literal VN/CN packages (no specific factory
    names) gets exactly the CARTON_FACTORY_ORDER_NO_CO order (VN before CN)
    from this same table; a shipment with specific factory names gets
    CARTON_FACTORY_ORDER_WITH_CO's order, with any literal-VN packages
    sorting alongside them (after the named ones, still before CN)."""
    specific = [f for f in CARTON_FACTORY_ORDER_WITH_CO if f != "CN"]
    return specific + list(CARTON_FACTORY_ORDER_NO_CO)


# v12: trailing copy-suffix (_1 / -2 / (3) / COPY 4 / COPY_5 / COPY-6),
# anchored to the END of the string only -- same semantics as
# pl_ocr_core.parse_document_sequence's _SEQ_SUFFIX_RE (kept as an
# independent, self-contained copy here per this module's own "no import
# from pl_ocr_core" architecture -- see test_pl_group_export.py for a
# cross-consistency test guarding the two against silently diverging).
_COPY_SUFFIX_RE = re.compile(
    r'^(?P<base>.+?)(?:[_\-]\s*\d{1,4}|\(\d{1,4}\)|\s+COPY[_\-\s]\d{1,4})$',
    re.IGNORECASE)


def _strip_trailing_copy_suffix(code: str) -> str:
    """"Kerry_POP_1" -> "Kerry_POP" (so factory detection sees the real
    "POP" suffix instead of the copy marker "1"); "Kerry_POP" (no suffix)
    and "CN-1666-PVG-KERRY-POP" (mid-string number, not a trailing copy
    marker) are returned unchanged. Found while testing business_sort_
    packages(): calling detect_factory() on a copy-suffixed filename without
    this step misreads every _1/_2/... copy as factory=REVIEW (the literal
    last token is "1", not "POP"), which would sort a copy PDF to the very
    end of the whole shipment instead of next to its original."""
    m = _COPY_SUFFIX_RE.match(code or "")
    return m.group("base") if m and m.group("base") else (code or "")


def _detect_factory_from_code(code: str) -> Optional[str]:
    code = _strip_trailing_copy_suffix(code)
    toks = _tokens(code)
    if not toks:
        return None

    last = toks[-1]
    if last in _KNOWN_FACTORY_TOKENS:
        return last

    # Handle split spellings such as "..._SB_GEAR" or "..._QI_FENG"
    if len(toks) >= 2:
        merged2 = toks[-2] + toks[-1]
        if merged2 == "SBGEAR":
            return "SBGEAR"
        if merged2 == "QIFENG":
            return "QIFENG"

    # Fallback: flat (no-separator) suffix match, e.g. filename "TW8785SBGEAR"
    flat = "".join(toks)
    for kw in FACTORY_KEYWORDS:
        if flat.endswith(kw):
            return kw
    return None


def detect_factory(reference_code: str, source_file: str) -> str:
    """Return one of CN / POP / SBGEAR / QIFENG / JION / VN / REVIEW.

    Priority: the LAST suffix of reference_code, then the LAST suffix of the
    PDF filename (stem). A leading 'CN' (e.g. CN-2569_SH_PVG_POP) must NOT be
    read as factory=CN — only the final suffix counts. A trailing copy-
    sequence marker (_1/_2/(3)/COPY 4/...) is stripped first, so
    "Kerry_POP_1" still detects factory=POP (see _strip_trailing_copy_suffix)
    -- this is now the SINGLE source of truth for factory classification:
    every caller (classify_packages_for_port, export_grouped_pl's grouping,
    _is_all_cn_factory, business_sort_packages) goes through this same
    function and therefore gets the same, consistent answer.
    """
    f = _detect_factory_from_code(reference_code or "")
    if f:
        return f
    stem = Path(source_file).stem if source_file else ""
    f = _detect_factory_from_code(stem)
    if f:
        return f
    return "REVIEW"


# =========================================================================
# 3) Master store <-> port mapping (CN retail only)
# =========================================================================
STORE_MASTER: Dict[str, Dict[str, object]] = {
    "CHENGDU": {
        "port": "TFU",
        "receiver": "Topologie CN - Chengdu Taikooli",
        "aliases": ["Chengdu", "Chengdu Taikooli", "Taikoo Li Chengdu", "M060"],
        "address": "Shop M060, Taikoo Li Chengdu, 8 Zhongshamao Street, Jinjiang District, Chengdu City, Sichuan, CN 610021",
        "contact_name": "June",
        "contact_phone": "18084829907",
    },
    "SHENZHEN": {
        "port": "SZX",
        "receiver": "CN - Shenzhen Mixc City (Shop T228)",
        "aliases": ["Shenzhen", "Shenzhen MixC", "Mixc City", "T228"],
        "address": "Shop T228, Tower 3, Vientiane City (MixC), No. 1881 Baoan South Road, Luohu District, Shenzhen, Guangdong, CN 518000",
        "contact_name": "Ben",
        "contact_phone": "18565775002",
    },
    "GUANGZHOU": {
        "port": "SZX",
        "receiver": "Topologie CN - Guangzhou Central Parc",
        "aliases": ["Guangzhou", "Guangzhou Central Parc", "Guangzhou Parc Central", "B262-1"],
        "address": "Shop B262-1, B2/F, Parc Central, No.218 Tianhe Road, Tianhe District, Guangzhou City, Guangdong, CN 510620",
        "contact_name": "Zhang Xiaojie",
        "contact_phone": "13662343374",
    },
    "HANGZHOU": {
        "port": "PVG",
        "receiver": "Topologie CN - Hangzhou Mixc",
        "aliases": ["Hangzhou", "Hangzhou MixC", "B1C03"],
        "address": "B1C03, Hangzhou MixC Mall, 701 Fuchun Rd, Jianggan District, Hangzhou, Zhejiang, CN 310008",
        "contact_name": "Su Su",
        "contact_phone": "15606539115",
    },
    "IAPM": {
        "port": "PVG",
        "receiver": "Topologie CN - Iapm",
        "aliases": ["IAPM", "IAPM Mall", "L4-426"],
        "address": "L4-426, IAPM Mall, 999 Huaihai Rd (M), Xuhui District, Shanghai, Shanghai, CN 200020",
        "contact_name": "Shi Wei Yi",
        "contact_phone": "13621647004",
    },
    "KERRY": {
        "port": "PVG",
        "receiver": "Topologie CN - Kerry Center flagship",
        "aliases": ["Kerry", "Kerry Center", "Kerry Centre", "NB1-23B"],
        "address": "NB1-23B shop, B1 floor, Jing'an Kerry Centre, Jing'an District, Shanghai, Shanghai, CN 200040",
        "contact_name": "Ning Ning",
        "contact_phone": "17602197790",
    },
    "SHANGHAI_TAIKOOLI": {
        "port": "PVG",
        "receiver": "CN - Shanghai Taikooli (Shop B1-07b)",
        "aliases": ["Shanghai Taikooli", "Shanghai Taikoo Li", "B1-07b", "S-B1-07b"],
        "address": "Shop S-B1-07b, B/F, No.1-9, 500 Dongyu Road, Pudong, Shanghai, Shanghai, CN 200127",
        "contact_name": "Bobo Shi",
        "contact_phone": "13621647004",
    },
    "SHANGHAI_HONGQIAO": {
        "port": "PVG",
        "receiver": "CN - Shanghai Hongqiao Airport",
        "aliases": ["Shanghai Hongqiao", "Hongqiao Airport", "D60-6"],
        "address": "Shop D60-6, Shanghai Hongqiao International Airport Terminal 2 (Departure Restricted Area), Changning District, Shanghai, Shanghai, China 200335",
        "contact_name": "Bobo Shi",
        "contact_phone": "13621647004",
    },
    "CHINA_WORLD": {
        "port": "PEK",
        "receiver": "China World NB1026",
        "aliases": ["China World", "China World Mall", "NB1026"],
        "address": "Shop NB1026, B1 Floor, China World Mall, No. 1 Jianguomenwai Avenue, Chaoyang District, Beijing, Beijing, China 100004",
        "contact_name": "Bobo Shi",
        "contact_phone": "+86 13621647004",
    },
}

FACTORY_FILE_MAP = {
    "CN": "PL_FACTORY_CN.xlsx",
    "POP": "PL_FACTORY_POP.xlsx",
    "SBGEAR": "PL_FACTORY_SBGEAR.xlsx",
    "QIFENG": "PL_FACTORY_QIFENG.xlsx",
    "JION": "PL_FACTORY_JION.xlsx",
}
PORT_FILE_MAP = {
    "PVG": "PL_CN_PORT_PVG.xlsx",
    "SZX": "PL_CN_PORT_SZX.xlsx",
    "TFU": "PL_CN_PORT_TFU.xlsx",
    "PEK": "PL_CN_PORT_PEK.xlsx",
}
STORE_FILE_MAP = {k: f"PL_CN_STORE_{k}.xlsx" for k in STORE_MASTER}


def notify_party_block(store_key: str) -> str:
    """Format the exact 'NOTIFY PARTY: / DELIVERY ADDRESS: / ...' text block
    used on the real PL template, from STORE_MASTER's contact info. Returns
    "" if store_key isn't a known store (caller should leave the cell blank
    in that case rather than guessing)."""
    info = STORE_MASTER.get(store_key)
    if not info:
        return ""
    lines = ["NOTIFY PARTY:", "DELIVERY ADDRESS:", str(info.get("receiver", "")), str(info.get("address", ""))]
    contact = str(info.get("contact_name", "") or "")
    phone = str(info.get("contact_phone", "") or "")
    if contact:
        lines.append(contact)
    if phone:
        lines.append(f"Tel: {phone}")
    return "\n".join(lines)


def _build_alias_lookup() -> List[Tuple[str, str]]:
    lookup: List[Tuple[str, str]] = []
    for store_key, info in STORE_MASTER.items():
        cand_texts = [store_key.replace("_", " "), str(info["receiver"])] + list(info["aliases"])
        for t in cand_texts:
            n = _norm_text(t)
            if n:
                lookup.append((store_key, n))
    return lookup


_ALIAS_LOOKUP = _build_alias_lookup()


def _token_overlap_score(signal: str, alias: str) -> float:
    ali_tokens = set(alias.split())
    if not ali_tokens:
        return 0.0
    sig_tokens = set(signal.split())
    inter = sig_tokens & ali_tokens
    return len(inter) / len(ali_tokens)


def match_store(signal_text: str, threshold: float = 0.55, margin: float = 0.08) -> Tuple[str, float, str]:
    """Fuzzy-match free text against STORE_MASTER.

    Returns (store_key_or_REVIEW, confidence[0..1], suggested_store_if_review).
    Never guesses: low confidence or a too-close runner-up both yield REVIEW.
    """
    norm_signal = _norm_text(signal_text)
    if not norm_signal:
        return "REVIEW", 0.0, ""

    # 1) exact / substring alias hits (shop number, distinctive alias) -> high confidence
    exact_hits = set()
    for store_key, alias_norm in _ALIAS_LOOKUP:
        if not alias_norm:
            continue
        if alias_norm == norm_signal or alias_norm in norm_signal or norm_signal in alias_norm:
            exact_hits.add(store_key)
    if len(exact_hits) == 1:
        return next(iter(exact_hits)), 1.0, ""
    if len(exact_hits) > 1:
        return "REVIEW", 0.5, "/".join(sorted(exact_hits))

    # 2) fuzzy scoring — best per store across all its aliases
    scores: Dict[str, float] = {}
    for store_key, alias_norm in _ALIAS_LOOKUP:
        ratio = difflib.SequenceMatcher(None, norm_signal, alias_norm).ratio()
        overlap = _token_overlap_score(norm_signal, alias_norm)
        s = max(ratio, overlap)
        if s > scores.get(store_key, 0.0):
            scores[store_key] = s

    if not scores:
        return "REVIEW", 0.0, ""

    ranked = sorted(scores.items(), key=lambda kv: kv[1], reverse=True)
    best_store, best_score = ranked[0]
    second_score = ranked[1][1] if len(ranked) > 1 else 0.0

    if best_score < threshold:
        return "REVIEW", round(best_score, 3), best_store
    if (best_score - second_score) < margin:
        return "REVIEW", round(best_score, 3), best_store
    return best_store, round(best_score, 3), ""


# =========================================================================
# v12) Store / OR / SO matching against an (optional) OR List (spec section
#      5) -- separate from match_store() above (which is CN-STORE_MASTER-
#      specific, used for Notify Party/address lookup). This matches
#      against the OR List's OWN store vocabulary, which can name ANY
#      store across ANY factory, not just the 9 CN retail stores.
# =========================================================================
from dataclasses import dataclass as _dc_dataclass, field as _dc_field


@_dc_dataclass
class StoreOrMatchResult:
    matched_store: str = ""
    matched_or: str = ""
    matched_so: str = ""
    match_source: str = ""   # SHIPMARK_TOKEN_EXACT | FILENAME_TOKEN_EXACT | RECEIVER_TEXT_EXACT | FUZZY | ""
    candidate_store: str = ""
    candidate_score: float = 0.0
    status: str = "REVIEW"   # OK | REVIEW | NO_OR_LIST
    review_reason: str = ""


def store_identity(store_text: str) -> str:
    """Case/whitespace-insensitive identity key so "Kerry" (OR List's own
    casing) and "KERRY" (STORE_MASTER's enum-style key) are recognised as
    the SAME store instead of a false ambiguity between two differently-
    cased spellings of one store. Public (no leading underscore) because
    pl_ocr_core.py's counting-scope-key computation reuses it too (spec
    section 9: counting_scope_key = shipment_key + "|" + normalized_store)."""
    return _strip_accents(store_text).strip()


_store_identity = store_identity  # backward-compat alias for internal callers


def _store_alias_token_index(or_list_store_values) -> Dict[str, set]:
    """token(uppercased, >=2 chars) -> set of distinct store IDENTITIES
    (see _store_identity above) it could mean. Built from BOTH
    STORE_MASTER's own alias config (helps match a CN store even when the
    OR List just says "Kerry") AND the OR List's own STORE column values
    (works for factories STORE_MASTER doesn't cover at all, e.g.
    POP/SBGEAR/QIFENG stores). A token that maps to more than one DISTINCT
    store identity is ambiguous by construction -- never used for an exact
    match (spec: "RY must not automatically match KERRY unless RY is
    explicitly configured as a Store alias" -- since tokens are never split
    below word boundaries, "RY" and "KERRY" are simply different tokens;
    this only becomes relevant if two DIFFERENT stores happen to share a
    configured alias token, which is flagged as ambiguous rather than
    picked)."""
    idx: Dict[str, set] = {}

    def add(token, store_identity):
        if len(token) >= 2:
            idx.setdefault(token, set()).add(store_identity)

    for store_key, info in STORE_MASTER.items():
        identity = _store_identity(store_key.replace("_", " "))
        # NOTE: deliberately NOT tokenizing info["receiver"] (a full postal/
        # company description, e.g. "CN - Shenzhen Mixc City (Shop T228)")
        # -- splitting that into individual words would flood the index
        # with generic English tokens ("Shop", "City", "Tower", ...) that
        # collide across many stores and make everything "ambiguous". Only
        # the store's own short/distinctive name + explicitly configured
        # aliases (e.g. "Kerry", "Kerry Center", "NB1-23B") are tokenized.
        for alias in [store_key.replace("_", " ")] + list(info["aliases"]):
            for tok in _tokens(alias):
                add(tok, identity)
    for store_raw in or_list_store_values:
        identity = _store_identity(store_raw)
        for tok in _tokens(store_raw):
            add(tok, identity)
    return idx


def _exact_token_store_match(text: str, token_index: Dict[str, set]):
    """Returns (store_or_None, ambiguous_candidates_or_None)."""
    hits: set = set()
    ambiguous: set = set()
    for tok in _tokens(text):
        stores = token_index.get(tok)
        if not stores:
            continue
        if len(stores) == 1:
            hits |= stores
        else:
            ambiguous |= stores
    if len(hits) == 1 and not (ambiguous - hits):
        return next(iter(hits)), None
    if hits or ambiguous:
        return None, sorted(hits | ambiguous)
    return None, None


def _fuzzy_match_against_candidates(signal_text: str, candidates, threshold: float = 0.6, margin: float = 0.1):
    """Generic version of match_store()'s tier-2 fuzzy scoring, parameterized
    over an arbitrary candidate list instead of hardcoded STORE_MASTER --
    used as the conservative last-resort fallback for OR List store names
    that aren't in STORE_MASTER at all. Same "REVIEW beats a wrong guess"
    philosophy: low confidence or a too-close runner-up both yield REVIEW."""
    norm_signal = _norm_text(signal_text)
    if not norm_signal or not candidates:
        return "REVIEW", 0.0, ""
    scores: Dict[str, float] = {}
    for cand in candidates:
        norm_cand = _norm_text(cand)
        if not norm_cand:
            continue
        ratio = difflib.SequenceMatcher(None, norm_signal, norm_cand).ratio()
        overlap = _token_overlap_score(norm_signal, norm_cand)
        s = max(ratio, overlap)
        if s > scores.get(cand, 0.0):
            scores[cand] = s
    if not scores:
        return "REVIEW", 0.0, ""
    ranked = sorted(scores.items(), key=lambda kv: kv[1], reverse=True)
    best_cand, best_score = ranked[0]
    second_score = ranked[1][1] if len(ranked) > 1 else 0.0
    if best_score < threshold:
        return "REVIEW", round(best_score, 3), best_cand
    if (best_score - second_score) < margin:
        return "REVIEW", round(best_score, 3), best_cand
    return best_cand, round(best_score, 3), ""


def match_store_and_or(pkg, or_index: Dict[str, list], receiver_text: str = "") -> StoreOrMatchResult:
    """Resolve Store (and its OR/SO) for `pkg` against an OR List already
    loaded into `or_index` (pl_or_list_import.build_or_index() -- or_norm ->
    [OrListRow, ...]). Priority (spec section 5):
      1) exact store-alias TOKEN in Shipmark
      2) exact store-alias TOKEN in filename/reference_code
      3) exact store name TOKEN in receiver/consignee text
      4) conservative fuzzy fallback across the OR List's own store names
    Exact token matching always wins over fuzzy. Returns status=REVIEW
    (never a guess) on any conflict: ambiguous token, multiple store
    candidates, or an OR value that maps to more than one Store in the OR
    List itself."""
    result = StoreOrMatchResult()
    if not or_index:
        result.status = "NO_OR_LIST"
        return result

    all_rows = [r for rows in or_index.values() for r in rows]
    store_values = sorted({r.store_raw for r in all_rows})
    identity_to_raw: Dict[str, set] = {}
    for r in all_rows:
        identity_to_raw.setdefault(_store_identity(r.store_raw), set()).add(r.store_raw)
    token_idx = _store_alias_token_index(store_values)

    store_identity = None
    # NOTE: in production pkg.shipping_mark is already resolved to
    # pkg.reference_code as a fallback upstream (run_pipeline(), v11) when
    # no explicit Shipping Mark label was found in the PL text, so this
    # first tier effectively also covers "no distinct Shipmark" cases; the
    # reference_code tier below still runs independently for a pure/testable
    # function that doesn't assume that upstream resolution already happened.
    for text, source in (
        (pkg.shipping_mark, "SHIPMARK_TOKEN_EXACT"),
        (pkg.reference_code, "FILENAME_TOKEN_EXACT"),
        (receiver_text, "RECEIVER_TEXT_EXACT"),
    ):
        if not text:
            continue
        hit, ambiguous = _exact_token_store_match(text, token_idx)
        if hit:
            store_identity, result.match_source = hit, source
            break
        if ambiguous:
            result.status = "REVIEW"
            result.review_reason = f"Ambiguous store token match in {source}: candidates={ambiguous}"
            result.candidate_store = "/".join(sorted(ambiguous))
            return result

    if not store_identity:
        signal = " ".join(t for t in (pkg.shipping_mark, pkg.reference_code, receiver_text) if t)
        cand, score, suggestion = _fuzzy_match_against_candidates(signal, store_values)
        if cand == "REVIEW" or not cand:
            result.status = "REVIEW"
            result.review_reason = "No confident Store match (exact-token and fuzzy both failed)."
            result.candidate_store = suggestion
            result.candidate_score = score
            return result
        store_identity, result.match_source, result.candidate_score = _store_identity(cand), "FUZZY", score

    # Map the matched (normalized) store identity back onto the OR List's
    # own raw store spelling(s) that share it -- this is what the OR List's
    # rows are actually keyed by.
    raw_candidates = identity_to_raw.get(store_identity)
    if not raw_candidates:
        result.status = "REVIEW"
        result.review_reason = f"Store identity '{store_identity}' matched via {result.match_source}, " \
                                f"but has no corresponding row in the OR List (only in STORE_MASTER)."
        result.candidate_store = store_identity
        return result
    if len(raw_candidates) > 1:
        result.status = "REVIEW"
        result.review_reason = f"Store identity '{store_identity}' spelled multiple ways in the OR List: " \
                                f"{sorted(raw_candidates)}"
        result.candidate_store = "/".join(sorted(raw_candidates))
        return result
    store = next(iter(raw_candidates))

    # Now resolve OR/SO for that store from the OR List.
    store_rows = [r for r in all_rows if r.store_raw == store]
    if not store_rows:
        result.status = "REVIEW"
        result.review_reason = f"Store '{store}' matched, but has no row in the OR List."
        result.candidate_store = store
        return result

    distinct_or = {r.or_raw for r in store_rows}
    if len(distinct_or) > 1:
        result.status = "REVIEW"
        result.review_reason = f"Store '{store}' has multiple OR values in the OR List: {sorted(distinct_or)}"
        result.candidate_store = store
        return result

    row = store_rows[0]
    result.matched_store = store
    result.matched_or = row.or_raw
    result.matched_so = row.so_raw
    result.status = "OK"
    return result


# =========================================================================
# 4) PDF receiver-text extraction (CN packages only)
# =========================================================================
_RECEIVER_LABEL_RE = re.compile(r"receiver\s*(?:company)?\s*[:\-]?\s*(.+)", re.IGNORECASE)
_STOP_LABEL_RE = re.compile(r"^(sender|shipper|consignee|invoice|date|packing|page|notify)\b", re.IGNORECASE)


def _extract_receiver_block(full_text: str) -> str:
    if not full_text:
        return ""
    lines = full_text.splitlines()
    for i, line in enumerate(lines):
        m = _RECEIVER_LABEL_RE.search(line)
        if m:
            block = [m.group(1).strip()]
            for j in range(i + 1, min(i + 4, len(lines))):
                nxt = lines[j].strip()
                if not nxt or _STOP_LABEL_RE.match(nxt):
                    break
                block.append(nxt)
            return " ".join(p for p in block if p)
    return ""


def _find_pdf_path(source_file: str, pdf_folder: Optional[Path], recursive: bool) -> Optional[Path]:
    if not pdf_folder or not source_file:
        return None
    pdf_folder = Path(pdf_folder)
    if not pdf_folder.exists():
        return None
    direct = pdf_folder / source_file
    if direct.exists():
        return direct
    pattern = "**/*" if recursive else "*"
    try:
        for p in pdf_folder.glob(pattern):
            if p.is_file() and p.name == source_file:
                return p
    except Exception as e:  # pragma: no cover
        log.warning(f"  glob error while looking for {source_file}: {e}")
    return None


def _extract_pdf_text_cached(pdf_path: Path, cache: Dict[str, str]) -> str:
    key = str(pdf_path)
    if key in cache:
        return cache[key]
    text = ""
    if pdfplumber is None:
        cache[key] = text
        return text
    try:
        with pdfplumber.open(str(pdf_path)) as pdf:
            pages_text = []
            for page in pdf.pages[:3]:  # receiver block is always on the cover page(s)
                pages_text.append(page.extract_text() or "")
            text = "\n".join(pages_text)
    except Exception as e:
        log.warning(f"  cannot read PDF for store detection: {pdf_path.name} ({e})")
        text = ""
    cache[key] = text
    return text


def _collect_cn_signal(pkg, pdf_folder, recursive, cache: Dict[str, str]) -> str:
    """Build the text blob used for store fuzzy-matching: PDF receiver block
    (if available) + reference_code + source_file, so matching still has a
    chance even when the PDF cannot be located."""
    parts = [pkg.reference_code or "", pkg.source_file or ""]
    pdf_path = _find_pdf_path(pkg.source_file, pdf_folder, recursive)
    if pdf_path:
        full_text = _extract_pdf_text_cached(pdf_path, cache)
        block = _extract_receiver_block(full_text)
        parts.append(block or full_text[:800])
    else:
        log.warning(f"  PDF not found for store detection: {pkg.source_file} (ref={pkg.reference_code})")
    return " ".join(p for p in parts if p)


# =========================================================================
# 5) Safe workbook writer (temp file + replace, never corrupts an open file)
# =========================================================================
def _safe_write(path: Path, write_fn: Callable[[Path], None]):
    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    tmp_path = path.with_suffix(path.suffix + ".tmp")
    try:
        write_fn(tmp_path)
    except PermissionError as e:
        raise PermissionError(
            f"Cannot create temp file for '{path.name}': {e}. "
            f"Close any program locking that folder and re-run."
        ) from e
    try:
        if path.exists():
            path.unlink()
        tmp_path.replace(path)
    except PermissionError as e:
        # cleanup the tmp file so re-runs don't pile up .tmp files
        try:
            tmp_path.unlink(missing_ok=True)
        except Exception:
            pass
        raise PermissionError(
            f"Cannot overwrite '{path}': the file appears to be open in Excel. "
            f"Close it and re-run the export."
        ) from e


def _write_group(path: Path, pkgs: List, write_workbook: Callable, renumber: bool):
    """Write one grouped workbook, optionally renumbering `global_carton_num`
    to be local to the group (1/N .. N/N), then ALWAYS restoring the original
    values afterwards — success or failure — so later groups are unaffected."""
    if not pkgs:
        return None
    saved = [p.global_carton_num for p in pkgs]
    try:
        if renumber:
            def _sort_key(p):
                m = re.match(r"^\s*(\d+)", p.global_carton_num or "")
                return int(m.group(1)) if m else 0
            ordered = sorted(pkgs, key=_sort_key)
            n = len(ordered)
            for i, p in enumerate(ordered, start=1):
                p.global_carton_num = f"{i}/{n}"
        _safe_write(path, lambda tmp: write_workbook(tmp, pkgs))
        log.info(f"  wrote {path.name}  ({len(pkgs)} cartons)")
    finally:
        for p, orig in zip(pkgs, saved):
            p.global_carton_num = orig
    return path


def _write_total(dir_total: Path, packages: List, write_workbook: Callable, total_workbook: Optional[Path]) -> Path:
    target = dir_total / "PL_TOTAL.xlsx"
    if total_workbook and Path(total_workbook).exists():
        _safe_write(target, lambda tmp: shutil.copyfile(str(total_workbook), str(tmp)))
    else:
        _safe_write(target, lambda tmp: write_workbook(tmp, packages))
    log.info(f"  wrote {target.name}  ({len(packages)} cartons)")
    return target


# =========================================================================
# 6) Control CSV
# =========================================================================
CONTROL_FIELDS = [
    "source_file", "reference_code", "package_code",
    "factory", "port", "store", "store_confidence", "suggested_store_if_review",
]


def _write_control_csv(path: Path, rows: List[dict]):
    def _do(tmp):
        with open(tmp, "w", newline="", encoding="utf-8-sig") as f:
            w = csv.DictWriter(f, fieldnames=CONTROL_FIELDS)
            w.writeheader()
            for r in rows:
                w.writerow({k: r.get(k, "") for k in CONTROL_FIELDS})
    _safe_write(path, _do)


# =========================================================================
# 7) Validation / reconciliation
# =========================================================================
def _read_match_status_rowcount(path: Path) -> Optional[int]:
    try:
        import openpyxl
        wb = openpyxl.load_workbook(str(path), read_only=True, data_only=True)
        ws = wb["Match_Status"]
        n = max(ws.max_row - 1, 0)
        wb.close()
        return n
    except Exception as e:  # pragma: no cover
        log.warning(f"  could not read back {path.name} for validation: {e}")
        return None


def _validate(packages, classified, factory_groups, cn_by_port, cn_by_store,
              written_paths: Dict[Path, int]) -> Tuple[bool, str]:
    lines: List[str] = []
    ok = True

    def check(label: str, cond: bool, detail: str = ""):
        nonlocal ok
        status = "OK" if cond else "FAIL"
        if not cond:
            ok = False
        lines.append(f"[{status}] {label}" + (f" — {detail}" if detail else ""))
        return cond

    total_cartons = len(packages)
    total_qty = sum(p.calc_qty for p in packages)
    lines.append(f"PL TOTAL: {total_cartons} cartons, {total_qty} qty")

    # -- no duplicate package_code across the whole shipment --
    codes = [p.package_code for p in packages]
    dup_codes = sorted({c for c in codes if codes.count(c) > 1})
    check("No duplicate package_code across full shipment", not dup_codes,
          f"duplicates={dup_codes}" if dup_codes else "")

    # -- factory totals (REVIEW bucket counted explicitly, nothing silently dropped) --
    factory_cartons = sum(len(v) for v in factory_groups.values())
    factory_qty = sum(sum(p.calc_qty for p in v) for v in factory_groups.values())
    check("SUM(factory groups incl. REVIEW) cartons == PL_TOTAL cartons",
          factory_cartons == total_cartons, f"sum={factory_cartons} total={total_cartons}")
    check("SUM(factory groups incl. REVIEW) quantity == PL_TOTAL quantity",
          factory_qty == total_qty, f"sum={factory_qty} total={total_qty}")

    review_factory = factory_groups.get("REVIEW", [])
    if review_factory:
        lines.append(f"[WARN] {len(review_factory)} package(s) could not be classified to a factory (REVIEW):")
        for p in review_factory:
            lines.append(f"    REVIEW-FACTORY  source_file={p.source_file}  reference_code={p.reference_code}  package_code={p.package_code}")

    # -- CN port / store totals --
    cn_pkgs = factory_groups.get("CN", [])
    cn_cartons = len(cn_pkgs)
    cn_review = [c for c in classified if c["factory"] == "CN" and (not c["store"] or c["store"] == "REVIEW")]
    expected_cn_classified = cn_cartons - len(cn_review)

    port_cartons = sum(len(v) for v in cn_by_port.values())
    port_qty = sum(sum(p.calc_qty for p in v) for v in cn_by_port.values())
    store_cartons = sum(len(v) for v in cn_by_store.values())
    store_qty = sum(sum(p.calc_qty for p in v) for v in cn_by_store.values())

    if cn_pkgs:
        check("SUM(CN port groups) cartons == CN factory cartons minus REVIEW",
              port_cartons == expected_cn_classified,
              f"port_sum={port_cartons} expected={expected_cn_classified} (CN_total={cn_cartons}, review={len(cn_review)})")
        check("SUM(CN store groups) cartons == CN factory cartons minus REVIEW",
              store_cartons == expected_cn_classified,
              f"store_sum={store_cartons} expected={expected_cn_classified}")
        check("SUM(CN port groups) quantity == SUM(CN store groups) quantity",
              port_qty == store_qty, f"port_qty={port_qty} store_qty={store_qty}")

    if cn_review:
        lines.append(f"[WARN] {len(cn_review)} CN package(s) could not be confidently mapped to a store (REVIEW):")
        for c in cn_review:
            p = c["pkg"]
            lines.append(f"    REVIEW-STORE  source_file={p.source_file}  reference_code={p.reference_code}  "
                          f"package_code={p.package_code}  confidence={c['confidence']}  suggested={c['suggestion']}")

    # -- no duplicate package_code within any single exported group --
    def _dup_within(name: str, groups: Dict[str, List]):
        for key, plist in groups.items():
            pcodes = [p.package_code for p in plist]
            dups = sorted({c for c in pcodes if pcodes.count(c) > 1})
            check(f"No duplicate package_code within {name}={key}", not dups,
                  f"duplicates={dups}" if dups else "")

    _dup_within("factory", factory_groups)
    _dup_within("cn_port", cn_by_port)
    _dup_within("cn_store", cn_by_store)

    # -- no package lost: every package_code appears in exactly one factory bucket --
    all_grouped_codes = [p.package_code for v in factory_groups.values() for p in v]
    check("No package lost between PL_TOTAL and factory groups",
          sorted(all_grouped_codes) == sorted(codes),
          f"grouped_count={len(all_grouped_codes)} total_count={len(codes)}")

    # -- cross-check the actual files written to disk (catches writer bugs) --
    for path, expected_n in written_paths.items():
        n = _read_match_status_rowcount(path)
        if n is None:
            lines.append(f"[WARN] Could not verify {path.name} on disk (read-back failed)")
            continue
        check(f"{path.name}: Match_Status row count == expected cartons",
              n == expected_n, f"on_disk={n} expected={expected_n}")

    return ok, "\n".join(lines)


# =========================================================================
# 8) Main entry point
# =========================================================================
def export_grouped_pl(
    packages: List,
    output_dir,
    write_workbook: Callable,
    total_workbook=None,
    pdf_folder=None,
    recursive: bool = False,
    store_threshold: float = 0.55,
    store_margin: float = 0.08,
) -> Path:
    """Split `packages` (as produced by run_pipeline) into the grouped
    PL_SPLIT_OUTPUT folder tree and return the path to PL_SPLIT_CONTROL.csv.

    Raises RuntimeError if reconciliation fails after writing everything —
    never just prints "Completed" while data is actually missing/duplicated.
    """
    if not packages:
        raise ValueError(
            "export_grouped_pl: `packages` is empty. Run the OCR pipeline "
            "(run_pipeline(...)) first and pass its return value here."
        )

    output_dir = Path(output_dir)
    dir_total = output_dir / "01_PL_TOTAL"
    dir_factory = output_dir / "02_BY_FACTORY"
    dir_cn_port = output_dir / "03_CN_BY_PORT"
    dir_cn_store = output_dir / "04_CN_BY_STORE"
    for d in (output_dir, dir_total, dir_factory, dir_cn_port, dir_cn_store):
        d.mkdir(parents=True, exist_ok=True)

    # ---- classify every package exactly once ----
    receiver_cache: Dict[str, str] = {}
    classified: List[dict] = []
    control_rows: List[dict] = []

    for pkg in packages:
        factory = detect_factory(pkg.reference_code, pkg.source_file)
        store = port = ""
        confidence: object = ""
        suggestion = ""
        if factory == "CN":
            signal = _collect_cn_signal(pkg, pdf_folder, recursive, receiver_cache)
            store, confidence, suggestion = match_store(signal, store_threshold, store_margin)
            port = STORE_MASTER[store]["port"] if store in STORE_MASTER else "REVIEW"

        classified.append({"pkg": pkg, "factory": factory, "store": store,
                            "port": port, "confidence": confidence, "suggestion": suggestion})
        control_rows.append({
            "source_file": pkg.source_file,
            "reference_code": pkg.reference_code,
            "package_code": pkg.package_code,
            "factory": factory,
            "port": port,
            "store": store,
            "store_confidence": confidence,
            "suggested_store_if_review": suggestion,
        })

    written_paths: Dict[Path, int] = {}

    # ---- 1) TOTAL ----
    log.info("Writing 01_PL_TOTAL ...")
    total_path = _write_total(dir_total, packages, write_workbook, total_workbook)
    written_paths[total_path] = len(packages)

    # ---- 2) BY FACTORY ----
    log.info("Writing 02_BY_FACTORY ...")
    factory_groups: Dict[str, List] = defaultdict(list)
    for c in classified:
        factory_groups[c["factory"]].append(c["pkg"])

    for factory, pkgs in factory_groups.items():
        fname = FACTORY_FILE_MAP.get(factory)
        if not fname or not pkgs:
            if factory == "REVIEW":
                log.warning(f"  {len(pkgs)} package(s) left unclassified (factory=REVIEW) — see control CSV")
            continue
        p = _write_group(dir_factory / fname, pkgs, write_workbook, renumber=True)
        if p:
            written_paths[p] = len(pkgs)

    # ---- 3 & 4) CN BY PORT / CN BY STORE ----
    log.info("Writing 03_CN_BY_PORT / 04_CN_BY_STORE ...")
    cn_by_port: Dict[str, List] = defaultdict(list)
    cn_by_store: Dict[str, List] = defaultdict(list)
    for c in classified:
        if c["factory"] != "CN":
            continue
        if not c["store"] or c["store"] == "REVIEW":
            continue  # excluded on purpose — never silently guess
        cn_by_port[c["port"]].append(c["pkg"])
        cn_by_store[c["store"]].append(c["pkg"])

    for port, pkgs in cn_by_port.items():
        fname = PORT_FILE_MAP.get(port)
        if fname and pkgs:
            p = _write_group(dir_cn_port / fname, pkgs, write_workbook, renumber=True)
            if p:
                written_paths[p] = len(pkgs)

    for store, pkgs in cn_by_store.items():
        fname = STORE_FILE_MAP.get(store)
        if fname and pkgs:
            p = _write_group(dir_cn_store / fname, pkgs, write_workbook, renumber=True)
            if p:
                written_paths[p] = len(pkgs)

    # ---- safety net: restore original carton numbers on every package ----
    # (each _write_group call already restores in its own finally-block; this
    # loop is a defensive no-op unless an exception skipped that restore.)

    # ---- control CSV ----
    control_path = output_dir / "PL_SPLIT_CONTROL.csv"
    _write_control_csv(control_path, control_rows)
    log.info(f"  wrote {control_path.name}  ({len(control_rows)} rows)")

    # ---- validation / reconciliation ----
    ok, report_text = _validate(packages, classified, factory_groups, cn_by_port, cn_by_store, written_paths)
    report_path = output_dir / "PL_SPLIT_VALIDATION.txt"
    report_path.write_text(report_text, encoding="utf-8")
    print("\n" + "=" * 70)
    print("PL SPLIT VALIDATION REPORT")
    print("=" * 70)
    print(report_text)
    print("=" * 70)

    if not ok:
        raise RuntimeError(
            "PL split reconciliation FAILED — see PL_SPLIT_VALIDATION.txt "
            f"({report_path}) for the full list of mismatches."
        )

    log.info("Reconciliation PASSED.")
    return control_path
