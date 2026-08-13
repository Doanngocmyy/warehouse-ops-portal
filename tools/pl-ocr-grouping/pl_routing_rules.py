#!/usr/bin/env python3
"""
pl_routing_rules.py — PL OCR V21 user-provided routing engine.

Replaces the previous hardcoded Store/Port inference (STORE_MASTER +
_KEC_* fuzzy alias matching in pl_group_export.py) with an explicit
Country | Port | Store rule table the USER supplies (spec "PL OCR V21
FINAL MASTER AUDIT + IMPLEMENTATION PROMPT", sections 2-13).

This module is the ONE canonical implementation (spec section 49) of:
  - routing-rule validation/normalization (sections 3-6)
  - Shipping Mark BODY tokenization (sections 7-8)
  - rule matching: exact / partial-unique / unique-fallback / conflict
    (sections 10-13)

pl_ocr_core.py / pl_group_export.py consume RouteMatch objects produced
here; they must never re-derive Country/Port/Store with their own
independent logic once routing_rules are supplied (section 46: manual
routing rules are authoritative; legacy alias logic may only be used as
a backward-compatible fallback when no routing_rules are supplied, and
must never silently override an explicit manual route).

No import of pl_ocr_core / pl_group_export here — this module is a leaf
dependency, importable stand-alone (and independently unit-testable).
"""
from __future__ import annotations
import re
from dataclasses import dataclass, field
from typing import List, Optional, Sequence, Set

# ── Country ──────────────────────────────────────────────────────────────
# Spec section 3: exactly 2 letters, no finite whitelist -- any real 2-letter
# market/destination code is accepted.
RE_COUNTRY_CODE = re.compile(r'^[A-Z]{2}$')


def normalize_country(value: object) -> str:
    return str(value or "").strip().upper()


def is_valid_country_code(value: str) -> bool:
    return bool(RE_COUNTRY_CODE.fullmatch(value or ""))


# ── Port ─────────────────────────────────────────────────────────────────
# Spec section 4: optional, no finite whitelist either -- normalize only.
def normalize_port(value: object) -> str:
    return str(value or "").strip().upper()


# ── Store ────────────────────────────────────────────────────────────────
# Spec section 5: optional; normalize for MATCHING only (never mutates a
# real display value) by lowercasing and collapsing whitespace / "-" / "_".
_STORE_SEP_RE = re.compile(r'[\s\-_]+')


def normalize_store(value: object) -> str:
    text = str(value or "").strip().lower()
    return _STORE_SEP_RE.sub('', text)


# ── Routing rule ─────────────────────────────────────────────────────────
@dataclass(frozen=True)
class RoutingRule:
    country: str   # normalized 2-letter code, e.g. "CN"
    port: str      # normalized ("" == intentionally blank, spec section 4)
    store: str      # ORIGINAL casing (for display / OR List matching)

    @property
    def store_key(self) -> str:
        return normalize_store(self.store)

    def __repr__(self) -> str:
        return f"RoutingRule({self.country}|{self.port or 'blank'}|{self.store or 'blank'})"


@dataclass
class RuleValidationResult:
    rules: List[RoutingRule]
    errors: List[str]
    warnings: List[str]

    @property
    def is_valid(self) -> bool:
        return not self.errors


def validate_routing_rules(raw_rows: Sequence[dict]) -> RuleValidationResult:
    """raw_rows: [{"country": ..., "port": ..., "store": ...}, ...] exactly
    as typed into the routing-rule table UI (spec section 45). Validates +
    normalizes per sections 3-6:
      - Country Code required, exactly 2 letters (no whitelist)
      - Port optional
      - Store optional
      - exact duplicate rows rejected
      - multiple rows sharing the same Country are explicitly VALID
        (Store/Port disambiguate)
    """
    rules: List[RoutingRule] = []
    errors: List[str] = []
    warnings: List[str] = []
    seen: Set = set()

    for i, row in enumerate(raw_rows or [], start=1):
        raw_country = row.get("country") if isinstance(row, dict) else getattr(row, "country", "")
        raw_port = row.get("port") if isinstance(row, dict) else getattr(row, "port", "")
        raw_store = row.get("store") if isinstance(row, dict) else getattr(row, "store", "")

        country = normalize_country(raw_country)
        if not is_valid_country_code(country):
            errors.append(
                f"Routing rule #{i}: Country Code {raw_country!r} must be exactly 2 letters (e.g. CN, SG, JP)."
            )
            continue

        port = normalize_port(raw_port)
        store = str(raw_store or "").strip()
        rule = RoutingRule(country=country, port=port, store=store)

        dup_key = (rule.country, rule.port, rule.store_key)
        if dup_key in seen:
            errors.append(
                f"Routing rule #{i}: duplicate of an earlier rule "
                f"({country} | {port or 'blank'} | {store or 'blank'})."
            )
            continue
        seen.add(dup_key)
        rules.append(rule)

    return RuleValidationResult(rules=rules, errors=errors, warnings=warnings)


# ── Shipping Mark tokenization ────────────────────────────────────────────
# Spec section 7-8: COUNTRY-TOTALQTY-BODY...-SUFFIX, separators -/_/whitespace
# all equivalent for matching. Port/Store are NOT assumed to sit at fixed
# BODY positions -- callers match BODY_TOKENS against the routing rules
# instead of indexing into them.
_SEP_RE = re.compile(r'[\s\-_]+')


@dataclass
class ShipmarkTokens:
    raw: str
    country: str
    total_qty: str
    body_tokens: List[str]
    suffix: str


def tokenize_shipping_mark(mark: str) -> ShipmarkTokens:
    raw = str(mark or "")
    tokens = [t for t in _SEP_RE.split(raw.strip()) if t]
    if not tokens:
        return ShipmarkTokens(raw=raw, country="", total_qty="", body_tokens=[], suffix="")

    first = tokens[0]
    country = ""
    rest = tokens
    if len(first) >= 2 and first[:2].isalpha():
        country = first[:2].upper()
        remainder = first[2:]
        rest = ([remainder] if remainder else []) + tokens[1:]

    body = list(rest)
    total_qty = ""
    if body and body[0].isdigit():
        total_qty = body[0]
        body = body[1:]

    if body:
        suffix = body[-1]
        body_tokens = body[:-1]
    else:
        suffix = ""
        body_tokens = []

    return ShipmarkTokens(raw=raw, country=country, total_qty=total_qty,
                           body_tokens=body_tokens, suffix=suffix)


# ── Rule matching ──────────────────────────────────────────────────────────
# Diagnostic method labels (spec sections 10-13); RouteMatch.status is always
# "MATCHED" or "REVIEW" -- callers branch on status, use `method`/`reason`
# purely for audit trails (Raw_Data / PL_SPLIT_CONTROL diagnostics).
METHOD_EXACT = "ROUTE_EXACT_MATCH"
METHOD_PARTIAL = "ROUTE_PARTIAL_MATCH"
METHOD_UNIQUE_FALLBACK = "ROUTE_UNIQUE_FALLBACK"
METHOD_CONFLICT = "ROUTE_CONFLICT"
METHOD_NO_MATCH = "ROUTE_NO_MATCH"
METHOD_AMBIGUOUS = "ROUTE_AMBIGUOUS"
METHOD_NO_RULE_FOR_COUNTRY = "ROUTE_NO_RULE_FOR_COUNTRY"
# v21 gap-fix (static audit): distinct from METHOD_NO_RULE_FOR_COUNTRY --
# that one means "we resolved a Country Code but nobody wrote a rule for
# it"; this one means "we never even resolved a Country Code for this
# package at all" (pkg.country is blank). Two different failure modes,
# two different audit reasons -- never collapsed into one generic REVIEW.
METHOD_COUNTRY_UNRESOLVED = "ROUTE_COUNTRY_UNRESOLVED"


@dataclass
class RouteMatch:
    status: str             # "MATCHED" | "REVIEW"
    method: str
    country: str
    port: str                # resolved Port ("" is valid when status MATCHED)
    store: str                # resolved Store, original casing ("" valid)
    rule: Optional[RoutingRule] = None
    candidates: List[RoutingRule] = field(default_factory=list)
    reason: str = ""


def _body_has_port(port: str, body_tokens: Sequence[str]) -> bool:
    if not port:
        return False
    upper_tokens = {t.upper() for t in body_tokens if t}
    return port.upper() in upper_tokens


def _body_has_store(store_key: str, body_norm_tokens: Set[str], joined_norm: str) -> bool:
    if not store_key:
        return False
    if store_key in body_norm_tokens:
        return True
    # bigram-joining (spec section 7 note: Store may be split across two
    # adjacent BODY tokens, e.g. "SH" + "Taikooli") -- compare against the
    # fully-joined BODY string too so a multi-word Store rule can still
    # match a Shipmark that ran it together without a separator.
    return store_key in joined_norm


def match_route(country: str, body_tokens: Sequence[str], rules: Sequence[RoutingRule]) -> RouteMatch:
    """Core V21 matching algorithm (spec sections 10-13). `country` is the
    ALREADY-resolved package destination (from pl_ocr_core's existing
    detect_shipment_country()/resolve_package_country() -- this module does
    not re-derive it); `body_tokens` are the Shipmark BODY tokens between
    the country/qty prefix and the trailing factory/origin suffix (see
    tokenize_shipping_mark, or caller-supplied equivalents e.g. filename
    tokens / receiver-text tokens).

    Algorithm (evidence-based, not positional):
      1. port_rules  = candidate rules whose OWN (non-blank) Port literally
         appears as a BODY token.
      2. store_rules = candidate rules whose OWN (non-blank) Store is found
         in the BODY (exact normalized token, or the whole normalized BODY
         join for a multi-token Store name).
      3. If both non-empty: their intersection must be exactly one rule
         (Country+Port+Store all agree) -- 0 rules in common is a genuine
         CONTRADICTION (spec section 12, never silently resolved), >1 is
         a malformed-rules edge case, also REVIEW.
      4. If only one of the two is non-empty and uniquely identifies a
         single rule, that's spec section 11's "partial but unique" match
         (OCR missed one BODY field) -- allowed, but flagged as PARTIAL
         (vs EXACT) purely for audit when the winning rule actually HAD a
         second field that just wasn't found in the BODY. A rule whose
         second field is blank BY DESIGN (e.g. CN|blank|Tmall) still comes
         back as a full/EXACT match -- a rule can't be "partially missing"
         evidence for a field it never declared.
      5. Zero signal at all: unique-country fallback (section 13) only
         when exactly one rule exists for this Country; otherwise REVIEW.
    """
    country = normalize_country(country)
    candidates = [r for r in rules if r.country == country]
    if not candidates:
        return RouteMatch(status="REVIEW", method=METHOD_NO_RULE_FOR_COUNTRY, country=country, port="", store="",
                           reason=f"No routing rule defined for Country Code {country!r}.")

    body_tokens = [t for t in body_tokens if t]
    body_norm_tokens = {normalize_store(t) for t in body_tokens}
    body_norm_tokens.discard("")
    joined_norm = normalize_store("".join(body_tokens))

    def port_ok(r: RoutingRule) -> bool:
        return _body_has_port(r.port, body_tokens)

    def store_ok(r: RoutingRule) -> bool:
        return _body_has_store(r.store_key, body_norm_tokens, joined_norm)

    def fully_matches(r: RoutingRule) -> bool:
        return (not r.port or port_ok(r)) and (not r.store_key or store_ok(r))

    def resolve(r: RoutingRule, method_if_full: str, method_if_partial: str) -> RouteMatch:
        method = method_if_full if fully_matches(r) else method_if_partial
        return RouteMatch(status="MATCHED", method=method, country=country, port=r.port, store=r.store, rule=r)

    port_rules = [r for r in candidates if r.port and port_ok(r)]
    store_rules = [r for r in candidates if r.store_key and store_ok(r)]

    if port_rules and store_rules:
        intersect = [r for r in port_rules if r in store_rules]
        if len(intersect) == 1:
            return resolve(intersect[0], METHOD_EXACT, METHOD_PARTIAL)
        if not intersect:
            return RouteMatch(status="REVIEW", method=METHOD_CONFLICT, country=country, port="", store="",
                               candidates=port_rules + store_rules,
                               reason="Shipping Mark BODY signals contradict each other: the Port token found "
                                      f"belongs to {[repr(r) for r in port_rules]}, but the Store token found "
                                      f"belongs to {[repr(r) for r in store_rules]} -- these do not agree on a "
                                      "single routing rule.")
        return RouteMatch(status="REVIEW", method=METHOD_AMBIGUOUS, country=country, port="", store="",
                           candidates=intersect,
                           reason=f"Port+Store evidence matches {len(intersect)} routing rules at once "
                                  "(malformed/overlapping rule table).")

    if store_rules and not port_rules:
        if len(store_rules) == 1:
            return resolve(store_rules[0], METHOD_EXACT, METHOD_PARTIAL)
        return RouteMatch(status="REVIEW", method=METHOD_AMBIGUOUS, country=country, port="", store="",
                           candidates=store_rules,
                           reason="Shipping Mark BODY's Store signal matches more than one routing rule "
                                  f"for {country!r} -- " + "; ".join(repr(r) for r in store_rules) + ".")

    if port_rules and not store_rules:
        if len(port_rules) == 1:
            return resolve(port_rules[0], METHOD_EXACT, METHOD_PARTIAL)
        return RouteMatch(status="REVIEW", method=METHOD_AMBIGUOUS, country=country, port="", store="",
                           candidates=port_rules,
                           reason="Shipping Mark BODY's Port signal matches more than one routing rule "
                                  f"for {country!r} and no Store signal disambiguates -- " +
                                  "; ".join(repr(r) for r in port_rules) + ".")

    # Zero signal at all -- unique-country fallback (spec section 13).
    if len(candidates) == 1:
        r = candidates[0]
        return RouteMatch(status="MATCHED", method=METHOD_UNIQUE_FALLBACK, country=country, port=r.port,
                           store=r.store, rule=r)

    return RouteMatch(status="REVIEW", method=METHOD_NO_MATCH, country=country, port="", store="",
                       candidates=candidates,
                       reason=f"{len(candidates)} routing rules exist for {country!r} but no Port/Store signal in "
                              f"the Shipping Mark uniquely identifies one.")
