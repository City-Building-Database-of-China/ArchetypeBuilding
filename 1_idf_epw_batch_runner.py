# -*- coding: utf-8 -*-
"""
Batch-match and optionally run EnergyPlus IDF files using eppy + runIDFs.

This repository release covers Shanghai only. Default paths are relative to the
script location (repository root) and require no local drive letters.

Repository layout (defaults):
  ArchetypeBuilding/
    1_idf_epw_batch_runner.py
    ready_idf/
      shang4hai3shi4/          # simulation-ready .idf files
    input/
      EPW/Shang4hai3shi4/       # baseline and RCP scenario .epw files
      GIS/Prototype/           # sample footprint shapefile
      Setting/                 # Schedule.xlsx, Static.xlsx
    result/                    # EnergyPlus outputs (created on run)
    result/reports/            # matching / run Excel reports

Set CHECK_ONLY=True to export a matching preview report without running EnergyPlus.
Set CHECK_ONLY=False to run EnergyPlus and export run reports.

Override paths with environment variables: IDF_ROOT, EPW_DIR, OUT_ROOT, REPORT_ROOT,
ENERGYPLUS_DIR, IDD_FILE, EPW_STEM, EP_VERSION, NUM_CPUS, CHECK_ONLY.
"""
import os
import re
import time
from typing import Dict, Iterable, List, Tuple, Optional, Set
import pandas as pd
from difflib import SequenceMatcher
from eppy.modeleditor import IDF
from eppy.runner.run_functions import runIDFs

# ======== Basic configuration ========
# All repository paths are relative to PROJECT_ROOT unless overridden by env vars.
PROJECT_ROOT = os.path.dirname(os.path.abspath(__file__))

ENERGYPLUS_DIR = os.environ.get("ENERGYPLUS_DIR", "")
IDD_FILE = os.environ.get(
    "IDD_FILE",
    os.path.join(ENERGYPLUS_DIR, "Energy+.idd") if ENERGYPLUS_DIR else "Energy+.idd",
)

IDF_ROOT = os.environ.get("IDF_ROOT", os.path.join(PROJECT_ROOT, "ready_idf"))
EPW_DIR = os.environ.get("EPW_DIR", os.path.join(PROJECT_ROOT, "input", "EPW", "Shang4hai3shi4"))
OUT_ROOT = os.environ.get("OUT_ROOT", os.path.join(PROJECT_ROOT, "result"))
REPORT_ROOT = os.environ.get("REPORT_ROOT", os.path.join(PROJECT_ROOT, "result", "reports"))

NUM_CPUS = int(os.environ.get("NUM_CPUS", "6"))
EP_VERSION = os.environ.get("EP_VERSION", "23-1-0")

# Set to True to generate only the EPW-IDF matching plan without running EnergyPlus.
CHECK_ONLY = os.environ.get("CHECK_ONLY", "False").strip().lower() in {"1", "true", "yes", "y"}

# Baseline EPW stem for Shanghai; override for RCP scenarios (e.g. SHANGHAISHI_RCP4.5_2040).
DEFAULT_EPW_STEM = os.environ.get("EPW_STEM", "Shanghai_2020")

# Optional Excel city mapping (not bundled in this Shanghai-only release).
EXCEL_PATH = os.environ.get("EXCEL_PATH", "")
EXCEL_SHEET_INDEX_1B = int(os.environ.get("EXCEL_SHEET_INDEX_1B", "2"))
EXCEL_COL_CITY_1B = int(os.environ.get("EXCEL_COL_CITY_1B", "1"))
EXCEL_COL_CODE_1B = int(os.environ.get("EXCEL_COL_CODE_1B", "3"))
EXCEL_COL_PINYIN_1B = int(os.environ.get("EXCEL_COL_PINYIN_1B", "6"))

# Shanghai folder name under ready_idf/ -> EPW file stem under input/EPW/Shang4hai3shi4/
FOLDER_TO_EPW_STEM: Dict[str, str] = {
    "shang4hai3shi4": DEFAULT_EPW_STEM,
}

# ======== Keyword fallback dictionary: folder_key -> English keywords ========
FOLDER_TO_KEYWORDS = {
    "shang4hai3shi4": ["Shanghai"],
}

# ======== Fuzzy matching parameters and alias bridge ========
FUZZY_RATIO_MIN      = 0.86   # SequenceMatcher threshold
SKELETON_RATIO_MIN   = 0.92   # consonant-skeleton threshold, stricter
MAX_FUZZY_CANDIDATES = 10     # score the top N most similar candidates

ALIAS_BRIDGE: Dict[str, List[str]] = {
    "akesu": ["aksu"], "akesushi": ["aksushi", "aksu"],
    "aba": ["barkam", "maerkang", "markang", "maerkangshi"],
    "haerbin": ["harbin"], "aerbin": ["harbin"],
    "kelamayi": ["karamay"],
    "wulumuqi": ["urumqi"], "urumuchi": ["urumqi"],
    "huhehaote": ["hohhot", "huhehot"],
    "lasa": ["lhasa"],
    "xiamen": ["amoy"],
    "guangzhou": ["canton"],
    "xianggang": ["hongkong", "hongkongsar"],
}

# ======== Pronunciation and normalization utilities ========
ADMIN_SUFFIXES = [
    "zizhizhou", "zizhixian", "zizhishi", "diqu",
    "sheng", "meng", "zhou", "xian", "shi", "qu"
]

def normalize_letters(s: str) -> str:
    if not isinstance(s, str):
        s = str(s) if s is not None else ""
    return re.sub(r"[^a-z]", "", s.lower())

def normalize_alnum(s: str) -> str:
    if not isinstance(s, str):
        s = str(s) if s is not None else ""
    return re.sub(r"[^a-z0-9]", "", s.lower())

def remove_vowels(s: str) -> str:
    return re.sub(r"[aeiou]", "", s)

def generate_variants_from_pinyin_key(pinyin_key: str) -> List[str]:
    base = normalize_letters(pinyin_key)
    variants = [base]
    for suf in ADMIN_SUFFIXES:
        if base.endswith(suf) and len(base) > len(suf) + 1:
            variants.append(base[:-len(suf)])
    seen = set(); out = []
    for v in variants:
        if v and v not in seen:
            seen.add(v); out.append(v)
    return out

def generate_tokens_from_city_name(city_name: str) -> List[str]:
    tokens = []
    base = normalize_letters(city_name)
    if base:
        tokens.append(base)
        for suf in ADMIN_SUFFIXES:
            if base.endswith(suf) and len(base) > len(suf) + 1:
                tokens.append(base[:-len(suf)])
    seen = set(); out = []
    for t in tokens:
        if t and t not in seen:
            seen.add(t); out.append(t)
    return out

# ======== File and directory utilities ========
def safe_makedirs(path: str) -> None:
    os.makedirs(path, exist_ok=True)

def make_eplaunch_options(idf: IDF) -> Dict:
    idf_full = idf.idfname
    idf_stem = os.path.splitext(os.path.basename(idf_full))[0]
    city = os.path.basename(os.path.dirname(idf_full))
    out_dir = os.path.join(OUT_ROOT, city, idf_stem)
    safe_makedirs(out_dir)
    return {
        "ep_version": EP_VERSION,
        "output_prefix": os.path.basename(idf_full),
        "output_suffix": "C",
        "output_directory": out_dir,
        "readvars": True,
        "expandobjects": True,
    }

def list_city_subfolders(root: str) -> List[str]:
    try:
        entries = os.listdir(root)
    except FileNotFoundError:
        return []
    return sorted(os.path.join(root, n) for n in entries if os.path.isdir(os.path.join(root, n)))

def collect_idf_files(folder: str) -> List[str]:
    try:
        entries = os.listdir(folder)
    except FileNotFoundError:
        return []
    return sorted(os.path.join(folder, f) for f in entries if f.lower().endswith(".idf"))

def build_epw_map(epw_dir: str) -> Dict[str, str]:
    mp: Dict[str, str] = {}
    for root, _, files in os.walk(epw_dir):
        for f in files:
            if f.lower().endswith(".epw"):
                stem = os.path.splitext(f)[0]
                mp[stem] = os.path.join(root, f)
    return mp

def index_epw_files_with_norm(epw_dir: str) -> List[Tuple[str, str, str]]:
    out: List[Tuple[str, str, str]] = []
    for root, _, files in os.walk(epw_dir):
        for f in files:
            if f.lower().endswith(".epw"):
                stem = os.path.splitext(f)[0]
                out.append((stem, normalize_letters(stem), os.path.join(root, f)))
    return out

def epw_tokens(stem: str) -> Set[str]:
    raw = stem.lower()
    parts = re.split(r"[^a-z0-9]+", raw)
    toks = set()
    for p in parts:
        if not p:
            continue
        toks.add(p)
        toks.add(normalize_letters(p))
    return {t for t in toks if t}

def epw_skeleton_tokens(stem: str) -> Set[str]:
    return {remove_vowels(t) for t in epw_tokens(stem) if t}

# ======== Read Excel columns: city name / code / pinyin key ========
def read_excel_city_pinyin_pairs(path: str, sheet_1b: int,
                                 col_city_1b: int, col_pinyin_1b: int,
                                 col_code_1b: Optional[int] = None) -> pd.DataFrame:
    if not os.path.isfile(path):
        raise FileNotFoundError(f"Excel file does not exist: {path}")
    df = pd.read_excel(path, sheet_name=sheet_1b - 1, header=0)
    need_cols = [col_city_1b, col_pinyin_1b] + ([col_code_1b] if col_code_1b else [])
    if df.shape[1] < max(c for c in need_cols if c):
        raise IndexError(f"The target Excel sheet has too few columns; current column count: {df.shape[1]} columns")
    out = {
        "city_name": df.iloc[:, col_city_1b - 1].astype(str).str.strip(),
        "pinyin_key": df.iloc[:, col_pinyin_1b - 1].astype(str).str.strip(),
    }
    if col_code_1b:
        out["code"] = df.iloc[:, col_code_1b - 1].astype(str).str.strip()
    else:
        out["code"] = ""
    sub = pd.DataFrame(out).replace({"nan": "", "None": ""})
    sub = sub[(sub["city_name"] != "") & (sub["pinyin_key"] != "")]
    return sub

def read_pinyin_keys_set_from_excel(path: str, sheet_1b: int, col_pinyin_1b: int) -> Set[str]:
    df = read_excel_city_pinyin_pairs(path, sheet_1b, EXCEL_COL_CITY_1B, col_pinyin_1b, EXCEL_COL_CODE_1B)
    return set(df["pinyin_key"].tolist())

# ======== EPW selection with seven matching routes; Excel code has highest priority ========
def find_epw_by_excel_code(folder_name: str,
                           epw_index: List[Tuple[str, str, str]],
                           excel_pairs_df: pd.DataFrame) -> Optional[Tuple[str, str]]:
    """
    Method #0, highest priority: match the Excel city/station code against EPW filenames.
      - Locate rows where pinyin_key == folder_name.
      - Read the code field and build matching tokens, prioritizing numeric tokens and then alphanumeric tokens.
      - Check whether each normalized EPW stem contains any generated token.
    """
    rows = excel_pairs_df[excel_pairs_df["pinyin_key"] == folder_name]
    if rows.empty or "code" not in rows.columns:
        return None

    def code_tokens(code: str) -> List[str]:
        # 1) Numeric token, preferred
        digits = "".join(ch for ch in code if ch.isdigit())
        toks: List[str] = []
        if digits:
            toks.append(digits)
        # 2) Alphanumeric normalization, e.g. "110000BEIJINGSHI_2020" -> "110000beijingshi2020"
        alnum = normalize_alnum(code)
        if alnum and alnum not in toks:
            toks.append(alnum)
        # 3) Optional split tokens such as "110000" or "2020"; usually not needed.
        # parts = [p for p in re.findall(r"[A-Za-z]+|\d+", code) if len(p) >= 3]
        # for p in parts:
        #     q = normalize_alnum(p)
        #     if q and q not in toks:
        #         toks.append(q)
        return toks

    codes = [c for c in rows["code"].tolist() if isinstance(c, str) and c.strip()]
    if not codes:
        return None

    # Try each code in token-priority order.
    for code in codes:
        toks = code_tokens(code)
        if not toks:
            continue

        hits: List[str] = []
        for stem, _, full in epw_index:
            stem_norm = normalize_alnum(stem)
            if any(t in stem_norm for t in toks):
                hits.append(full)

        if hits:
            # Prefer 2020, CSWD, IWEC, TMY, newer modification time, and shorter filenames.
            def score(fp: str) -> Tuple:
                fn_up = os.path.basename(fp).upper()
                try:
                    mtime = os.stat(fp).st_mtime
                except OSError:
                    mtime = 0
                return ("2020" in fn_up, "CSWD" in fn_up, "IWEC" in fn_up, "TMY" in fn_up, mtime, -len(fn_up))
            hits.sort(key=score, reverse=True)
            return hits[0], f"excel_code:{code}"

    return None


def find_epw_by_excel_city_to_epwname(folder_name: str,
                                      epw_index: List[Tuple[str, str, str]],
                                      excel_pairs_df: pd.DataFrame) -> Optional[Tuple[str, str]]:
    rows = excel_pairs_df[excel_pairs_df["pinyin_key"] == folder_name]
    if rows.empty:
        return None
    tokens_ordered: List[str] = []
    for city in rows["city_name"].tolist():
        tokens_ordered += generate_tokens_from_city_name(city)
    for pkey in rows["pinyin_key"].tolist():
        tokens_ordered += generate_variants_from_pinyin_key(pkey)
    seen = set(); tokens = []
    for t in tokens_ordered:
        if t and t not in seen:
            seen.add(t); tokens.append(t)
    candidates: List[str] = []
    used_token = ""
    for tok in tokens:
        hits = [full for stem, norm, full in epw_index if tok in norm]
        if hits:
            candidates.extend(hits)
            used_token = tok
            break
    if not candidates:
        return None
    def score(fp: str) -> Tuple:
        fn_up = os.path.basename(fp).upper()
        try:
            mtime = os.stat(fp).st_mtime
        except OSError:
            mtime = 0
        return ("2020" in fn_up, "CSWD" in fn_up, "IWEC" in fn_up, "TMY" in fn_up, mtime, -len(fn_up))
    candidates.sort(key=score, reverse=True)
    return candidates[0], f"excel_city:{used_token}"

def find_epw_by_excel_pronunciation(folder_name: str,
                                    epw_index: List[Tuple[str, str, str]],
                                    excel_keys: Set[str]) -> Optional[Tuple[str, str]]:
    if folder_name not in excel_keys:
        return None
    variants = generate_variants_from_pinyin_key(folder_name)
    candidates: List[str] = []
    used_variant = ""
    for v in variants:
        hits = [full for stem, norm, full in epw_index if v in norm]
        if hits:
            candidates.extend(hits)
            used_variant = v
            break
    if not candidates:
        return None
    def score(fp: str) -> Tuple:
        fn_up = os.path.basename(fp).upper()
        try:
            mtime = os.stat(fp).st_mtime
        except OSError:
            mtime = 0
        return ("2020" in fn_up, "CSWD" in fn_up, "IWEC" in fn_up, "TMY" in fn_up, mtime, -len(fn_up))
    candidates.sort(key=score, reverse=True)
    return candidates[0], f"excel:{used_variant}"

def find_epw_by_folder_pronunciation(folder_name: str,
                                     epw_index: List[Tuple[str, str, str]]) -> Optional[Tuple[str, str]]:
    variants = generate_variants_from_pinyin_key(folder_name)
    candidates: List[str] = []
    used_variant = ""
    for v in variants:
        hits = [full for stem, norm, full in epw_index if v in norm]
        if hits:
            candidates.extend(hits)
            used_variant = v
            break
    if not candidates:
        return None
    def score(fp: str) -> Tuple:
        fn_up = os.path.basename(fp).upper()
        try:
            mtime = os.stat(fp).st_mtime
        except OSError:
            mtime = 0
        return ("2020" in fn_up, "CSWD" in fn_up, "IWEC" in fn_up, "TMY" in fn_up, mtime, -len(fn_up))
    candidates.sort(key=score, reverse=True)
    return candidates[0], f"folder:{used_variant}"

def find_epw_by_explicit_mapping(folder_name: str, epw_map: Dict[str, str]) -> Optional[Tuple[str, str]]:
    epw_stem = FOLDER_TO_EPW_STEM.get(folder_name)
    if epw_stem:
        epw_full = epw_map.get(epw_stem)
        if epw_full:
            return epw_full, f"explicit:{epw_stem}"
        else:
            print(f"[WARN] EPW specified by exact mapping does not exist: {epw_stem}.epw")
    return None

def expand_aliases(tokens: List[str]) -> List[str]:
    out = []
    seen = set()
    for t in tokens:
        if t and t not in seen:
            seen.add(t); out.append(t)
        for a in ALIAS_BRIDGE.get(t, []):
            if a not in seen:
                seen.add(a); out.append(a)
    return out

def find_epw_by_fuzzy(folder_name: str,
                      epw_index: List[Tuple[str, str, str]],
                      excel_pairs_df: pd.DataFrame) -> Optional[Tuple[str, str]]:
    cand = generate_variants_from_pinyin_key(folder_name)
    rows = excel_pairs_df[excel_pairs_df["pinyin_key"] == folder_name]
    if not rows.empty:
        for city in rows["city_name"].tolist():
            cand += generate_tokens_from_city_name(city)
        for pkey in rows["pinyin_key"].tolist():
            cand += generate_variants_from_pinyin_key(pkey)
        for code in rows.get("code", []):
            if isinstance(code, str) and code.strip():
                parts = [p for p in re.split(r"[^a-z0-9]+", code.lower()) if p]
                cand += [normalize_letters(p) for p in parts]

    cand = expand_aliases([t for t in cand if t])
    cand = [t for t in cand if len(t) >= 3]
    if not cand:
        return None

    epw_token_cache = []
    for stem, _, full in epw_index:
        toks = epw_tokens(stem)
        skel = {remove_vowels(x) for x in toks if x}
        epw_token_cache.append((stem, full, toks, skel))

    best = None  # (score, rank_tuple, full, reason)
    for tok in cand:
        tok_skel = remove_vowels(tok)
        scored_candidates = []
        for stem, full, toks, skel in epw_token_cache:
            if tok in toks:
                scored_candidates.append((1.0, stem, full, f"fuzzy:{tok}==token"))
                continue
            local_best = 0.0; local_best_t = ""
            for t in toks:
                r = SequenceMatcher(None, tok, t).ratio()
                if r > local_best:
                    local_best = r; local_best_t = t
            local_best_skel = 0.0; local_best_skel_t = ""
            for t in skel:
                r2 = SequenceMatcher(None, tok_skel, t).ratio()
                if r2 > local_best_skel:
                    local_best_skel = r2; local_best_skel_t = t
            score = max(local_best, local_best_skel * 0.98)
            reason = f"fuzzy:{tok}~{local_best_t}({local_best:.2f})|skel:{tok_skel}~{local_best_skel_t}({local_best_skel:.2f})"
            scored_candidates.append((score, stem, full, reason))

        scored_candidates.sort(key=lambda x: x[0], reverse=True)
        for score, stem, full, reason in scored_candidates[:MAX_FUZZY_CANDIDATES]:
            fn_up = os.path.basename(full).upper()
            accept = (score >= FUZZY_RATIO_MIN)
            if not accept:
                if SequenceMatcher(None, remove_vowels(tok), remove_vowels(stem.lower())).ratio() >= SKELETON_RATIO_MIN:
                    accept = True
                    reason += "|accept=skel_gate"
            if accept:
                rank = ("2020" in fn_up, "CSWD" in fn_up, "IWEC" in fn_up, "TMY" in fn_up, -len(fn_up))
                pack = (score, rank, full, reason)
                if (best is None) or (pack > (best[0], best[1], best[2], best[3])):
                    best = pack
    if best:
        _, _, full, reason = best
        return full, reason
    return None

def find_epw_by_keyword_mapping(folder_name: str, epw_map: Dict[str, str]) -> Optional[Tuple[str, str]]:
    kws = FOLDER_TO_KEYWORDS.get(folder_name, [])
    if not kws:
        return None
    candidates: List[str] = []
    for stem, full in epw_map.items():
        name_low = stem.lower()
        if any(kw.lower() in name_low for kw in kws):
            candidates.append(full)
    if not candidates:
        return None
    if len(candidates) == 1:
        return candidates[0], f"keyword:{kws[0]}"
    def score(fp: str) -> tuple:
        fn = os.path.basename(fp).upper()
        try:
            mtime = os.stat(fp).st_mtime
        except OSError:
            mtime = 0
        return ("2020" in fn, "CSWD" in fn, "IWEC" in fn, "TMY" in fn, mtime, -len(fn))
    candidates.sort(key=score, reverse=True)
    return candidates[0], f"keyword:{'/'.join(kws)}"

def choose_best_epw_for_folder(folder_name: str,
                               epw_index: List[Tuple[str, str, str]],
                               epw_map: Dict[str, str],
                               excel_pairs_df: pd.DataFrame,
                               excel_keys: Set[str]) -> Optional[Tuple[str, str]]:
    r = find_epw_by_excel_code(folder_name, epw_index, excel_pairs_df)                 # 0)
    if r: return r
    r = find_epw_by_excel_city_to_epwname(folder_name, epw_index, excel_pairs_df)      # 1)
    if r: return r
    r = find_epw_by_excel_pronunciation(folder_name, epw_index, excel_keys)            # 2)
    if r: return r
    r = find_epw_by_folder_pronunciation(folder_name, epw_index)                       # 3)
    if r: return r
    r = find_epw_by_explicit_mapping(folder_name, epw_map)                             # 4)
    if r: return r
    r = find_epw_by_fuzzy(folder_name, epw_index, excel_pairs_df)                      # 5)
    if r: return r
    r = find_epw_by_keyword_mapping(folder_name, epw_map)                              # 6)
    if r: return r
    return None

# ======== Run-status detection and report writers ========
def detect_success(out_dir: str) -> Tuple[bool, str]:
    err_path = os.path.join(out_dir, "eplusout.err")
    if os.path.isfile(err_path):
        try:
            with open(err_path, "r", encoding="utf-8", errors="ignore") as fh:
                txt = fh.read()
            ok = ("Completed Successfully" in txt) or ("EnergyPlus Completed Successfully" in txt)
            if ok:
                return True, "Completed Successfully"
            tail = "\n".join(txt.splitlines()[-20:])
            return False, f"ERR tail:\n{tail}"
        except Exception as e:
            return False, f"ERR read failed: {e}"
    if os.path.isfile(os.path.join(out_dir, "eplusout.sql")):
        return True, "No ERR, but SQL exists"
    try:
        errs = [f for f in os.listdir(out_dir) if f.lower().endswith(".err")]
        if errs:
            errs.sort(key=lambda fn: os.stat(os.path.join(out_dir, fn)).st_mtime, reverse=True)
            p = os.path.join(out_dir, errs[0])
            with open(p, "r", encoding="utf-8", errors="ignore") as fh:
                txt = fh.read()
            ok = ("Completed Successfully" in txt)
            return (True, f"{errs[0]} says success") if ok else (False, f"{errs[0]} tail:\n" + "\n".join(txt.splitlines()[-20:]))
    except Exception:
        pass
    return False, "No ERR/SQL found"

def write_match_plan_report(plan_records: List[dict], report_root: str) -> str:
    os.makedirs(report_root, exist_ok=True)
    df = pd.DataFrame(plan_records)
    df_unmatched = df[df["epw_path"].isna() | (df["epw_path"] == "")]
    summary = pd.DataFrame({
        "metric": ["city_folders", "with_idfs", "mapped_epw", "no_epw", "total_idfs"],
        "value":  [
            df["city_folder"].nunique(),
            int((df["idf_count"] > 0).sum()),
            int((df["epw_path"].notna() & (df["epw_path"] != "")).sum()),
            int(len(df_unmatched)),
            int(df["idf_count"].sum()),
        ],
    })
    stamp = time.strftime("%Y%m%d_%H%M%S")
    xlsx_path = os.path.join(report_root, f"epw_idf_match_plan_{stamp}.xlsx")
    with pd.ExcelWriter(xlsx_path, engine="openpyxl") as writer:
        df.sort_values(["epw_path"], na_position="first").to_excel(writer, index=False, sheet_name="plan")
        df_unmatched.to_excel(writer, index=False, sheet_name="unmatched")
        summary.to_excel(writer, index=False, sheet_name="summary")
    print(f"[REPORT] Matching preview report generated: {xlsx_path}")
    return xlsx_path

def write_run_report(all_records: List[Dict], report_root: str) -> str:
    df = pd.DataFrame(all_records)
    df["ts"] = pd.Timestamp.now().strftime("%Y-%m-%d %H:%M:%S")
    df_success  = df[df["run_status"] == "success"].copy()
    df_failures = df[df["run_status"] != "success"].copy()
    summary = pd.DataFrame({
        "metric": ["total_pairs", "success", "failures", "unique_city_folders", "unique_epw_used"],
        "value":  [len(df), len(df_success), len(df_failures), df["city_folder"].nunique(), df["epw_name"].nunique()]
    })
    stamp = time.strftime("%Y%m%d_%H%M%S")
    xlsx_path = os.path.join(report_root, f"epw_idf_run_report_{stamp}.xlsx")
    with pd.ExcelWriter(xlsx_path, engine="openpyxl") as writer:
        df_success.to_excel(writer, index=False, sheet_name="success")
        df_failures.to_excel(writer, index=False, sheet_name="failures")
        df.to_excel(writer, index=False, sheet_name="all")
        summary.to_excel(writer, index=False, sheet_name="summary")
    print(f"[REPORT] Run report generated: {xlsx_path}")
    return xlsx_path

# ======== EnergyPlus run wrapper ========
def run_one_folder(iddfile: str, epwfile: str, idf_folder: str, num_cpus: int,
                   match_info: str) -> List[Dict]:
    idf_paths = collect_idf_files(idf_folder)
    if not idf_paths:
        return []
    IDF.setiddname(iddfile)
    idf_objs: List[IDF] = [IDF(p, epwfile) for p in idf_paths]
    runs_list: List[Tuple[IDF, Dict]] = []
    for idf in idf_objs:
        opts = make_eplaunch_options(idf)
        runs_list.append((idf, opts))
    print(f"    Preparing to run {len(runs_list)} IDF files; example mapping: {os.path.basename(idf_paths[0])} -> {os.path.basename(epwfile)}")
    runIDFs(((idf, opts) for (idf, opts) in runs_list), num_cpus)

    records: List[Dict] = []
    for idf, opts in runs_list:
        out_dir = opts.get("output_directory", os.path.dirname(idf.idfname))
        ok, msg = detect_success(out_dir)
        records.append({
            "city_folder": os.path.basename(os.path.dirname(idf.idfname)),
            "idf_path": idf.idfname,
            "idf_name": os.path.basename(idf.idfname),
            "epw_path": epwfile,
            "epw_name": os.path.basename(epwfile),
            "output_dir": out_dir,
            "run_status": "success" if ok else "failed",
            "message": msg,
            "match_info": match_info,
        })
    return records

# ======== Main workflow ========
def main():
    start = time.time()
    print(f"[INFO] PROJECT_ROOT = {PROJECT_ROOT}")
    print(f"[INFO] IDF_ROOT      = {IDF_ROOT}")
    print(f"[INFO] EPW_DIR       = {EPW_DIR}")
    print(f"[INFO] OUT_ROOT      = {OUT_ROOT}")
    print(f"[INFO] EPW_STEM      = {DEFAULT_EPW_STEM}")
    if not CHECK_ONLY and not os.path.isfile(IDD_FILE):
        print(f"[ERROR] IDD file not found: {IDD_FILE}")
        print("Set ENERGYPLUS_DIR or IDD_FILE to your local EnergyPlus install.")
        return
    safe_makedirs(OUT_ROOT)
    safe_makedirs(REPORT_ROOT)

    epw_map   = build_epw_map(EPW_DIR)
    epw_index = index_epw_files_with_norm(EPW_DIR)
    if not epw_map:
        print(f"[ERROR] No .epw files found in EPW_DIR: {EPW_DIR}");  return
    print(f"[INFO] EPW files indexed: {len(epw_index)}")

    excel_pairs_df = pd.DataFrame(columns=["city_name", "pinyin_key", "code"])
    excel_keys: Set[str] = set()
    if EXCEL_PATH and os.path.isfile(EXCEL_PATH):
        try:
            excel_pairs_df = read_excel_city_pinyin_pairs(
                EXCEL_PATH, EXCEL_SHEET_INDEX_1B,
                EXCEL_COL_CITY_1B, EXCEL_COL_PINYIN_1B,
                EXCEL_COL_CODE_1B,
            )
            excel_keys = set(excel_pairs_df["pinyin_key"].tolist())
        except Exception as e:
            print(f"[WARN] Failed to read Excel; using FOLDER_TO_EPW_STEM only: {e}")
    else:
        print("[INFO] No cities_info.xlsx; using FOLDER_TO_EPW_STEM for EPW matching.")

    city_folders = list_city_subfolders(IDF_ROOT)
    if not city_folders:
        print(f"[ERROR] No city subfolders found in IDF_ROOT: {IDF_ROOT}");  return

    if CHECK_ONLY:
        plan_records: List[dict] = []
        for folder in city_folders:
            folder_name = os.path.basename(folder)
            idf_files   = collect_idf_files(folder)
            idf_count   = len(idf_files)
            chosen = choose_best_epw_for_folder(folder_name, epw_index, epw_map,
                                                excel_pairs_df, excel_keys)
            if chosen:
                epw_full, match_info = chosen
                epw_name = os.path.basename(epw_full)
            else:
                epw_full, epw_name, match_info = "", "", ""
            plan_records.append({
                "city_folder": folder_name,
                "idf_count": idf_count,
                "idf_samples": " | ".join([os.path.basename(p) for p in idf_files[:3]]),
                "epw_name": epw_name,
                "epw_path": epw_full,
                "match_info": match_info,
            })
        write_match_plan_report(plan_records, REPORT_ROOT)
        print("\n[INFO] CHECK_ONLY=True: generated only the matching preview; runIDFs was not executed.")
        return

    total_idf = 0
    all_records: List[Dict] = []
    for folder in city_folders:
        folder_name = os.path.basename(folder)
        chosen = choose_best_epw_for_folder(folder_name, epw_index, epw_map,
                                            excel_pairs_df, excel_keys)
        if not chosen:
            print(f"[SKIP] No EPW matched for folder={folder_name} (check Excel, exact mapping, or keywords)")
            continue
        epw_full, match_info = chosen
        idf_files = collect_idf_files(folder)
        if not idf_files:
            print(f"[SKIP] No IDF files in folder: {folder}")
            continue
        print(f"[RUN] {folder_name} -> {os.path.basename(epw_full)} | IDFs: {len(idf_files)} | CPUs: {NUM_CPUS} | match={match_info}")
        records = run_one_folder(IDD_FILE, epw_full, folder, NUM_CPUS, match_info)
        all_records.extend(records)
        total_idf += len(idf_files)

    if all_records:
        write_run_report(all_records, REPORT_ROOT)

    elapsed = time.time() - start
    succ = sum(1 for r in all_records if r.get("run_status") == "success") if all_records else 0
    fail = sum(1 for r in all_records if r.get("run_status") != "success") if all_records else 0
    print(f"\nCompleted: found {total_idf} IDF files; success {succ}; failed {fail}; elapsed {elapsed:.1f}s")

if __name__ == "__main__":
    main()
