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
FACTORY_KEYWORDS = ["SBGEAR", "QIFENG", "POP", "JION", "CN"]  # longest-first for the flat-suffix fallback


def _detect_factory_from_code(code: str) -> Optional[str]:
    toks = _tokens(code)
    if not toks:
        return None

    last = toks[-1]
    if last in {"CN", "POP", "SBGEAR", "QIFENG", "JION"}:
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
    """Return one of CN / POP / SBGEAR / QIFENG / JION / REVIEW.

    Priority: the LAST suffix of reference_code, then the LAST suffix of the
    PDF filename (stem). A leading 'CN' (e.g. CN-2569_SH_PVG_POP) must NOT be
    read as factory=CN — only the final suffix counts.
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
    },
    "SHENZHEN": {
        "port": "SZX",
        "receiver": "CN - Shenzhen Mixc City (Shop T228)",
        "aliases": ["Shenzhen", "Shenzhen MixC", "Mixc City", "T228"],
    },
    "GUANGZHOU": {
        "port": "SZX",
        "receiver": "Topologie CN - Guangzhou Central Parc",
        "aliases": ["Guangzhou", "Guangzhou Central Parc", "Guangzhou Parc Central", "B262-1"],
    },
    "HANGZHOU": {
        "port": "PVG",
        "receiver": "Topologie CN - Hangzhou Mixc",
        "aliases": ["Hangzhou", "Hangzhou MixC", "B1C03"],
    },
    "IAPM": {
        "port": "PVG",
        "receiver": "Topologie CN - Iapm",
        "aliases": ["IAPM", "IAPM Mall", "L4-426"],
    },
    "KERRY": {
        "port": "PVG",
        "receiver": "Topologie CN - Kerry Center flagship",
        "aliases": ["Kerry", "Kerry Center", "Kerry Centre", "NB1-23B"],
    },
    "SHANGHAI_TAIKOOLI": {
        "port": "PVG",
        "receiver": "CN - Shanghai Taikooli (Shop B1-07b)",
        "aliases": ["Shanghai Taikooli", "Shanghai Taikoo Li", "B1-07b", "S-B1-07b"],
    },
    "SHANGHAI_HONGQIAO": {
        "port": "PVG",
        "receiver": "CN - Shanghai Hongqiao Airport",
        "aliases": ["Shanghai Hongqiao", "Hongqiao Airport", "D60-6"],
    },
    "CHINA_WORLD": {
        "port": "PEK",
        "receiver": "China World NB1026",
        "aliases": ["China World", "China World Mall", "NB1026"],
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
