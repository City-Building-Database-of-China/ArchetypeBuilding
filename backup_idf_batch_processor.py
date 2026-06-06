# -*- coding: utf-8 -*-
"""
Three-stage batch processor for EnergyPlus IDF files.

This script combines the original writeIDF-1/2/3 workflow and supports
multi-process execution plus skip-if-up-to-date output checks.

Stage 1
    Write Schedule:Compact objects. Function-specific schedules are preferred,
    and climate-zone schedules are used as a fallback.
Stage 2
    Add PEOPLE, LIGHTS, ELECTRICEQUIPMENT, infiltration, HVAC template objects,
    and output settings.
Stage 3
    Refresh insulation layers and update SimpleGlazing U-factor/SHGC values
    using the static parameter workbook.

Key features
------------
- `MAX_WORKERS` controls parallel worker count.
- Each IDF file is processed as one task. Worker initializers load the IDD and
  required Excel/template data once per process to reduce repeated I/O.
- Existing outputs are skipped when they are newer than or equal to the source
  file timestamp.
- Windows multiprocessing is protected by `if __name__ == "__main__":`.
- Name normalization converts `YYYY.0` to `YYYY` and replaces `nan` tokens with
  the year parsed from the file name.
- Construction references are repaired by exact matching, `YYYY.0`/`YYYY`
  fallback, and floor-year fallback within the same climate zone, building type,
  and envelope part.
- When `KEEP_INTERMEDIATE_OUTPUTS` is False, all stages run in memory and only
  final IDFs are written to `FINAL_DIR`; when True, each stage writes its own
  output directory.
- Only files whose file-name year token is exactly `nan` are skipped.
- File names ending in `.idf` and `.IDF` are both supported.

Expected input file-name pattern
--------------------------------
    <city>_<func_index>_<archetype>_<year>_<scenario>.idf

Dependencies
------------
    pip install eppy pandas openpyxl

Configuration
-------------
The script is GitHub-ready and avoids personal absolute paths. You can configure
paths with environment variables, or use the default relative project layout:

    project/
      data/
        idf/
        Schedule1114.xlsx
        ClimateZone.xlsx
        Static-10.xlsx
        Insulation_Thickness_ByZone_config.xlsx
        template_outputs.idf
      output/

Supported environment variables:
    ENERGYPLUS_DIR, IDD_PATH, DATA_DIR, IDF_ROOT, SCHEDULE_XLS, CLIMATE_XLS,
    STATIC_XLS, INSUL_XLS, TEMPLATE_OUTPUTS_IDF, OUT_BASE
"""

from __future__ import annotations
import os
import re
from pathlib import Path
from typing import Dict, List, Optional, Tuple, Union, Set

import pandas as pd
from eppy.modeleditor import IDF
from concurrent.futures import ProcessPoolExecutor, as_completed

# ================== Path configuration ==================
PROJECT_ROOT = Path(__file__).resolve().parent

def _env_path(var_name: str, default: Union[str, Path]) -> Path:
    """Return a path from an environment variable or a project-relative default."""
    return Path(os.getenv(var_name, str(default))).expanduser()

ENERGYPLUS_DIR = _env_path("ENERGYPLUS_DIR", r"C:\EnergyPlusV23-2-0")
IDD_PATH = _env_path("IDD_PATH", ENERGYPLUS_DIR / "Energy+.idd")

DATA_DIR = _env_path("DATA_DIR", PROJECT_ROOT / "data")
IDF_ROOT = _env_path("IDF_ROOT", DATA_DIR / "idf")

SCHEDULE_XLS = _env_path("SCHEDULE_XLS", DATA_DIR / "Schedule.xlsx")
CLIMATE_XLS = _env_path("CLIMATE_XLS", DATA_DIR / "ClimateZone.xlsx")  # Optional.
STATIC_XLS = _env_path("STATIC_XLS", DATA_DIR / "Static.xlsx")  # People/lights/equipment and window U/SHGC data.
INSUL_XLS = _env_path("INSUL_XLS", DATA_DIR / "Insulation_Thickness_ByZone_config.xlsx")

TEMPLATE_OUTPUTS_IDF = _env_path("TEMPLATE_OUTPUTS_IDF", DATA_DIR / "template_outputs.idf")

OUT_BASE = _env_path("OUT_BASE", PROJECT_ROOT / "output")
STAGE1_DIR = Path(OUT_BASE) / "_stage1_with_schedules"
STAGE2_DIR = Path(OUT_BASE) / "_stage2_with_loads_hvac"
STAGE3_DIR = Path(OUT_BASE) / "_stage3_with_insulation_window"
FINAL_DIR = Path(OUT_BASE) / "_final_idf"

# Runtime switches.
DRY_RUN_STAGE1 = False
DRY_RUN_STAGE2 = False
DRY_RUN_STAGE3 = False
SHOW_LOG = True
MAX_WORKERS = 6

KEEP_INTERMEDIATE_OUTPUTS = False
# ========================================================

# ================== Common utilities ==================
def _norm(s: str) -> str:
    return re.sub(r"[^a-z0-9]+", "", str(s).strip().lower())

def _all_idf_files(root: Path) -> List[Path]:
    """Collect both .idf and .IDF files."""
    return list(root.rglob("*.idf")) + list(root.rglob("*.IDF"))

def compute_out_path(idf_path: Path, idf_root: Path, out_root: Path) -> Path:
    """Map an IDF path under idf_root to the corresponding path under out_root."""
    try:
        rel = idf_path.relative_to(idf_root)
        out_dir = out_root / rel.parent
    except Exception:
        out_dir = out_root
    out_dir.mkdir(parents=True, exist_ok=True)
    return out_dir / idf_path.name

def is_up_to_date(in_path: Path, idf_root: Path, out_root: Path) -> bool:
    out_path = compute_out_path(in_path, idf_root, out_root)
    try:
        return out_path.exists() and out_path.stat().st_mtime >= in_path.stat().st_mtime
    except Exception:
        return False

# ================== File-name parsing and skip rules ==================
# <city>_<func_index>_<archetype>_<year>_Sx.idf
# The year token may be 2005, 2005.0, etc. Skip only when the year token is exactly 'nan'.

def _parse_int_from_token(token: str) -> int:
    """
    Robust integer parsing:
    - Supports '2005', '2005.0', ' 2005 ', '2005.0abc', and 'abc2005.0xyz'.
    - Try float conversion first, then fall back to the first numeric token found by regex.
    """
    if token is None:
        raise ValueError("Empty token cannot be parsed as an integer")
    s = str(token).strip()
    try:
        return int(float(s))
    except Exception:
        m = re.search(r"-?\d+(?:\.\d+)?", s)
        if not m:
            raise ValueError(f"Cannot parse an integer from '{token}'")
        return int(float(m.group(0)))

def parse_filename_parts(idf_path: Union[str, Path]) -> Dict[str, Union[str, int]]:
    stem = Path(idf_path).stem  # Example: city_5_4_2005.0_S0
    parts = stem.split("_")
    if len(parts) < 5:
        raise ValueError(f"File name must contain at least five underscore-separated tokens: {stem}")
    # The city token may contain underscores, so parse from the right.
    city = "_".join(parts[:-4]) if len(parts) > 5 else parts[0]
    func_index = _parse_int_from_token(parts[-4])
    archetype  = parts[-3]  # Keep as-is.
    year       = _parse_int_from_token(parts[-2])  # Allow 2005.0-style tokens.
    scenario   = parts[-1]
    return {"city": city, "func_index": func_index, "archetype": archetype, "year": year, "scenario": scenario}

def _year_token_is_nan(p: Union[str, Path]) -> bool:
    """
    Return True only when the file-name pattern <city>_<func_index>_<archetype>_<year>_<scenario>
    has a <year> token that is exactly "nan" case-insensitively.
    Example: city_0_1_nan_S0.idf -> True
         ji3nan2shi4_5_4_2005.0_S0.idf -> False
    """
    try:
        stem = Path(p).stem
        parts = stem.split("_")
        if len(parts) < 5:
            return False
        return parts[-2].strip().lower() == "nan"
    except Exception:
        return False

# ============ Name normalization: remove .0 year suffixes and replace nan with file year ============

YEAR_FLOAT_RE = re.compile(r"(?P<y>(19|20)\d{2})\.0\b", re.IGNORECASE)

def normalize_year_token_to_int(txt: str) -> str:
    if not txt: return txt
    return YEAR_FLOAT_RE.sub(lambda m: m.group("y"), txt)

def replace_nan_year_token(txt: str, file_year: int) -> str:
    if not txt: return txt
    def repl_token(token: str) -> str:
        return str(file_year) if token.strip().lower() == "nan" else token
    if "_" in txt:
        parts = txt.split("_")
        parts = [repl_token(p) for p in parts]
        txt = "_".join(parts)
    if "-" in txt:
        parts = txt.split("-")
        parts = [repl_token(p) for p in parts]
        txt = "-".join(parts)
    return txt

def normalize_any_name(name: str, file_year: int) -> str:
    s = str(name or "").strip()
    if not s: return s
    s = normalize_year_token_to_int(s)
    s = replace_nan_year_token(s, file_year)
    return s

def get_layer_names_ordered(cons) -> List[str]:
    names: List[str] = []
    if hasattr(cons, "Outside_Layer"):
        names.append(str(getattr(cons, "Outside_Layer") or "").strip())
    i = 2
    while hasattr(cons, f"Layer_{i}"):
        names.append(str(getattr(cons, f"Layer_{i}") or "").strip())
        i += 1
    while names and names[-1] == "":
        names.pop()
    return names

def rebuild_construction_with_layers(idf: IDF, cons, names: List[str]):
    if not names:
        return cons
    old_name = str(cons.Name)
    idf.removeidfobject(cons)
    kwargs = {"Name": old_name, "Outside_Layer": names[0]}
    for i, nm in enumerate(names[1:], start=2):
        kwargs[f"Layer_{i}"] = nm
    return idf.newidfobject("CONSTRUCTION", **kwargs)

def rename_construction_and_update_refs(idf: IDF, old: str, new: str):
    if old == new: return
    for o in idf.idfobjects["CONSTRUCTION"]:
        if str(o.Name).strip() == old:
            o.Name = new
            break
    for s in idf.idfobjects.get("BUILDINGSURFACE:DETAILED", []):
        if str(s.Construction_Name).strip() == old:
            s.Construction_Name = new
    for f in idf.idfobjects.get("FENESTRATIONSURFACE:DETAILED", []):
        if str(f.Construction_Name).strip() == old:
            f.Construction_Name = new

def rename_windowmat_and_update_refs(idf: IDF, old: str, new: str):
    if old == new: return
    for w in idf.idfobjects.get("WINDOWMATERIAL:SIMPLEGLAZINGSYSTEM", []):
        if str(w.Name).strip() == old:
            w.Name = new
            break
    for cons in list(idf.idfobjects.get("CONSTRUCTION", [])):
        layers = get_layer_names_ordered(cons)
        changed = False
        for i, nm in enumerate(layers):
            if str(nm).strip() == old:
                layers[i] = new
                changed = True
        if changed:
            rebuild_construction_with_layers(idf, cons, layers)

def normalize_names_everywhere(idf: IDF, file_year: int):
    for cons in list(idf.idfobjects.get("CONSTRUCTION", [])):
        old = str(cons.Name)
        new = normalize_any_name(old, file_year)
        if new != old:
            rename_construction_and_update_refs(idf, old, new)
    for w in list(idf.idfobjects.get("WINDOWMATERIAL:SIMPLEGLAZINGSYSTEM", [])):
        old = str(w.Name)
        new = normalize_any_name(old, file_year)
        if new != old:
            rename_windowmat_and_update_refs(idf, old, new)
    for s in idf.idfobjects.get("BUILDINGSURFACE:DETAILED", []):
        if getattr(s, "Construction_Name", None) is not None:
            s.Construction_Name = normalize_any_name(s.Construction_Name, file_year)
    for f in idf.idfobjects.get("FENESTRATIONSURFACE:DETAILED", []):
        if getattr(f, "Construction_Name", None) is not None:
            f.Construction_Name = normalize_any_name(f.Construction_Name, file_year)

# ============ Construction-reference repair with floor-year fallback ============

CZ_ANY = "__ANY__"
CZ_TOKEN = r"(?:SCZ\d+|CZ|HSCWZ|HSWWZ|TZ)"

# Construction-name pattern. Window/Win is included as an envelope part.
CONS_RE = re.compile(
    rf"(?P<cz>{CZ_TOKEN})\s*[_\-]\s*"
    r"(?P<btype>Residential|Public(?:&|and)?Industrial|Public|Industrial)\s*[_\-]\*?"
    r"(?P<year>(?:19|20)\d{2})(?:\.0)?\s*[_\-]\s*"
    r"(?P<part>Roof|Wall|ExtWall|IntWall|Floor|FloorSlab|FloorGround|Win|Window)",
    re.IGNORECASE,
)


def _normalize_year_token(name: str) -> str:
    return re.sub(r'_(\d{4})\.0(?=_|$)', r'_\1', name)

def _floatify_year_token(name: str) -> str:
    return re.sub(r'_(\d{4})(?=_|$)', r'_\1.0', name)

def _canon_btype(s: str) -> str:
    low = s.lower()
    return "Residential" if ("resi" in low) else "Public&Industrial"

def _canon_part(s: str) -> Optional[str]:
    low = s.lower().replace(" ", "")
    if any(k in low for k in ["roof"]): return "Roof"
    if "intwall" in low: return "IntWall"
    if "extwall" in low or ("wall" in low and "int" not in low): return "ExtWall"
    if any(k in low for k in ["floorslab", "slab"]): return "FloorSlab"
    if any(k in low for k in ["floorground", "ground"]): return "FloorGround"
    if "floor" in low: return "Floor"
    if "win" in low or "window" in low: return "Win"   # Include window constructions.
    if "wall" in low: return "Wall"
    return None

def parse_from_name(name: str) -> Tuple[Optional[str], Optional[str], Optional[int], Optional[str]]:
    m = CONS_RE.search(str(name))
    if not m: return None, None, None, None
    cz = m.group("cz").upper().strip()
    bt = _canon_btype(m.group("btype"))
    yr = int(m.group("year"))
    env = _canon_part(m.group("part"))
    return cz, bt, yr, env

def _build_cons_index(existing_names: Set[str]):
    """Build index: {(cz, building_type, part): [(year, name), ...]} sorted by year."""
    idx: Dict[Tuple[str,str,str], List[Tuple[int,str]]] = {}
    for nm in existing_names:
        cz, bt, yr, env = parse_from_name(nm)
        if all([cz, bt, yr, env]):
            key = (cz, bt, env)
            idx.setdefault(key, []).append((int(yr), nm))
            # Allow ExtWall/IntWall to fall back to Wall during floor-year matching.
            if env in ("ExtWall", "IntWall"):
                key2 = (cz, bt, "Wall")
                idx.setdefault(key2, []).append((int(yr), nm))
    for k in list(idx.keys()):
        idx[k] = sorted(set(idx[k]), key=lambda x: x[0])
    return idx

def _floor_map_construction_name(cname: str, existing_index) -> Optional[str]:
    """If no exact construction exists, choose the maximum candidate year <= target year by (cz, bt, part)."""
    cz, bt, yr, env = parse_from_name(cname)
    if not all([cz, bt, yr, env]):
        return None
    for env_try in ([env] + (["Wall"] if env in ("ExtWall", "IntWall") else [])):
        key = (cz, bt, env_try)
        seq = existing_index.get(key, [])
        if not seq:
            continue
        cand = [pair for pair in seq if pair[0] <= int(yr)]
        if cand:
            return max(cand, key=lambda x: x[0])[1]
    return None

def fix_construction_refs(idf: IDF) -> Tuple[int, Set[str]]:
    """
    Align surface/window Construction_Name values with existing Construction names:
    - Try the original name first.
    - Then try _YYYY.0 -> _YYYY or _YYYY -> _YYYY.0.
    - If still unresolved, use floor-year matching by (cz, bt, part).
    Return: (number of repaired references, unresolved name set).
    """
    existing = {c.Name for c in idf.idfobjects['CONSTRUCTION']}
    index = _build_cons_index(existing)
    fixed = 0
    unresolved: Set[str] = set()

    def map_name(cname: str):
        if cname in existing:
            return cname
        cand = _normalize_year_token(cname)
        if cand in existing:
            return cand
        cand2 = _floatify_year_token(cname)
        if cand2 in existing:
            return cand2
        cand3 = _floor_map_construction_name(_normalize_year_token(cname), index)
        if cand3 and cand3 in existing:
            return cand3
        return None

    for key in ('BUILDINGSURFACE:DETAILED', 'FENESTRATIONSURFACE:DETAILED'):
        for o in idf.idfobjects.get(key, []):
            cname = getattr(o, "Construction_Name", "") or ""
            if not cname:
                continue
            mapped = map_name(cname)
            if mapped and mapped != cname:
                o.Construction_Name = mapped
                fixed += 1
            elif not mapped and cname not in existing:
                unresolved.add(cname)

    return fixed, unresolved

# ================== Stage 1: write schedules ==================
FUNC_INDEX_TO_COL: Dict[int, List[str]] = {
    0: ["Residential_1", "Residential"],
    1: ["Residential_2", "Residential"],
    2: ["Residential_3", "Residential"],
    3: ["Commercial"],
    4: ["Office"],
    5: ["Industrial", "Industry"],
    6: ["Transport"],
    7: ["Administrative", "Administration"],
}

def read_wide_sheets(path: str) -> Dict[str, pd.DataFrame]:
    xls = pd.ExcelFile(path)
    result: Dict[str, pd.DataFrame] = {}
    for name in xls.sheet_names:
        df = xls.parse(name)
        if df.empty:
            continue
        df.columns = [str(c).strip() for c in df.columns]
        if len(df.columns) >= 2 and df.columns[0].strip().lower() == "name":
            df = df.dropna(how="all").reset_index(drop=True)
            if not df.empty:
                result[str(name).strip()] = df
    if not result:
        raise ValueError("Schedule.xlsx does not contain any valid sheet whose first column is 'Name'.")
    return result

def read_city_to_zone(path: str) -> Dict[str, str]:
    try:
        df = pd.read_excel(path, sheet_name=0)
    except Exception:
        return {}
    if df.shape[1] < 5:
        return {}
    mapping: Dict[str, str] = {}
    for _, row in df.iterrows():
        city = str(row.iloc[0]).strip()
        zone = str(row.iloc[4]).strip()
        if city and zone and city.lower() != "nan" and zone.lower() != "nan":
            mapping[city] = zone
    return mapping

def find_best_column(df: pd.DataFrame, desired: List[str]) -> Optional[str]:
    existing = list(df.columns)[1:]
    for want in desired:
        for c in existing:
            if c == want: return c
        for c in existing:
            if c.lower() == want.lower(): return c
        wn = _norm(want)
        for c in existing:
            if _norm(c) == wn: return c
        for c in existing:
            if wn and wn in _norm(c): return c
    return None

def extract_schedule_from_sheet(df: pd.DataFrame, schedule_name: str, col_actual: str) -> Dict:
    names = df.iloc[:, 0].astype(str).str.strip()
    key_norm = names.str.lower().str.replace(" ", "").str.replace("-", "")
    tl_idx = key_norm[key_norm.isin({
        "schedule_typelimits_name", "scheduletype_limitsname",
        "scheduletype_limits_name", "scheduletypelimitsname",
        "schedule_type_limits_name"
    })].index
    type_limits_name = ""
    if len(tl_idx) > 0:
        type_limits_name = str(df.loc[tl_idx[0], col_actual]).strip()
    pairs: List[Tuple[int, int]] = []
    for i, label in enumerate(names):
        lab = str(label).strip()
        if lab.lower().startswith("field_"):
            try:
                n = int(lab.split("_", 1)[-1])
            except Exception:
                n = 10000 + i
            pairs.append((n, i))
    pairs.sort(key=lambda x: x[0])
    lines: List[str] = []
    for _, ridx in pairs:
        val = df.loc[ridx, col_actual]
        if pd.isna(val):
            continue
        sval = str(val).strip()
        if not sval:
            continue
        lines.append(sval)
    if (not type_limits_name) and (len(lines) == 0):
        return {}
    return {"ScheduleName": schedule_name, "TypeLimitsName": type_limits_name, "Lines": lines}

def collect_schedules_for_idf(sheet_map: Dict[str, pd.DataFrame], func_cols: List[str], zone_candidates: List[str]) -> List[Dict]:
    out: List[Dict] = []
    for sheet_name, df in sheet_map.items():
        col = find_best_column(df, func_cols)
        mode = "func"
        if not col and zone_candidates:
            col = find_best_column(df, zone_candidates)
            mode = "zone" if col else "none"
        if not col:
            if SHOW_LOG:
                print(f"  [MISS] {sheet_name}: no matching function column for {func_cols}" + (f"; no matching climate-zone column for {zone_candidates}" if zone_candidates else ""))
            continue
        sdef = extract_schedule_from_sheet(df, sheet_name, col)
        if not sdef:
            if SHOW_LOG:
                print(f"  [EMPTY] {sheet_name}: matched column '{col}' (mode={mode}), but the selected column has no valid fields")
            continue
        out.append(sdef)
    return out

def ensure_typelimits(idf: IDF, name: str):
    if not name:
        return
    exists = [o for o in idf.idfobjects["SCHEDULETYPELIMITS"] if o.Name.strip().lower() == name.strip().lower()]
    if exists:
        return
    obj = idf.newidfobject("SCHEDULETYPELIMITS")
    obj.Name = name

def upsert_schedule_compact(idf: IDF, name: str, typelimits: str, lines: List[str]):
    olds = [o for o in idf.idfobjects["SCHEDULE:COMPACT"] if o.Name.strip().lower() == name.strip().lower()]
    for o in olds:
        idf.removeidfobject(o)
    obj = idf.newidfobject("SCHEDULE:COMPACT")
    obj.Name = name
    obj.Schedule_Type_Limits_Name = typelimits if typelimits else ""
    for i, line in enumerate(lines, start=1):
        setattr(obj, f"Field_{i}", str(line))

# ---- Stage 1 parallel processing: worker initialization and per-file processing ----
_SHEET_MAP: Dict[str, pd.DataFrame] = {}
_CITY2ZONE: Dict[str, str] = {}

def _init_stage1(schedule_xls: str, climate_xls: Optional[str], idd_path: str):
    IDF.setiddname(idd_path)
    global _SHEET_MAP, _CITY2ZONE
    _SHEET_MAP = read_wide_sheets(schedule_xls)
    _CITY2ZONE = read_city_to_zone(climate_xls) if climate_xls else {}

def _stage1_worker(fp: str, idf_root_str: str, out_root_str: str, dry_run: bool, show_log: bool) -> Tuple[bool, str]:
    try:
        if _year_token_is_nan(fp):
            return True, f"[SKIP-YEAR-NAN] {Path(fp).name}"
        p = Path(fp); idf_root = Path(idf_root_str); out_root = Path(out_root_str)
        if (not dry_run) and is_up_to_date(p, idf_root, out_root):
            return True, f"[SKIP-UPTODATE] {p.name}"
        parts = parse_filename_parts(p)
        city = parts["city"]; func_index = parts["func_index"]
        func_cols = FUNC_INDEX_TO_COL.get(func_index, [])
        if isinstance(func_cols, str): func_cols = [func_cols]
        zone_txt = _CITY2ZONE.get(city) or { _norm(k): v for k, v in _CITY2ZONE.items() }.get(_norm(city))
        zone_candidates: List[str] = []
        if zone_txt:
            zone_candidates = [zone_txt, zone_txt.replace(" ", ""), zone_txt.replace("-", "")]
        schedules = collect_schedules_for_idf(_SHEET_MAP, func_cols, zone_candidates)
        if not schedules:
            return False, f"[SKIP] {p.name}: no writable schedule definitions were found"
        if dry_run:
            return True, f"[DRY-1] {p.name}: {len(schedules)} schedules"
        idf = IDF(str(p))
        normalize_names_everywhere(idf, int(parts["year"]))
        for s in schedules:
            ensure_typelimits(idf, s.get("TypeLimitsName", "").strip())
            upsert_schedule_compact(idf, s.get("ScheduleName", "").strip(), s.get("TypeLimitsName", "").strip(), s.get("Lines", []) or [])
        out_path = compute_out_path(p, idf_root, out_root)
        idf.saveas(str(out_path))
        return True, f"[OK-1] {p.name} -> {out_path}"
    except Exception as e:
        return False, f"[ERR-1] {fp}: {e}"

def stage1_write_schedules(idf_root: Path, out_root: Path, schedule_xls: str, climate_xls: Optional[str] = None):
    idf_files = _all_idf_files(Path(idf_root))
    total = len(idf_files); ok = skipped = errors = 0
    if MAX_WORKERS > 1:
        with ProcessPoolExecutor(max_workers=MAX_WORKERS, initializer=_init_stage1, initargs=(schedule_xls, climate_xls, IDD_PATH)) as ex:
            futures = [ex.submit(_stage1_worker, str(p), str(idf_root), str(out_root), DRY_RUN_STAGE1, SHOW_LOG) for p in idf_files]
            for fut in as_completed(futures):
                success, msg = fut.result()
                if SHOW_LOG: print(msg)
                if msg.startswith("[SKIP]") or msg.startswith("[SKIP-UPTODATE]") or msg.startswith("[SKIP-YEAR-NAN]"): skipped += 1
                elif success: ok += 1
                else: errors += 1
    else:
        _init_stage1(schedule_xls, climate_xls, IDD_PATH)
        for p in idf_files:
            success, msg = _stage1_worker(str(p), str(idf_root), str(out_root), DRY_RUN_STAGE1, SHOW_LOG)
            if SHOW_LOG: print(msg)
            if msg.startswith("[SKIP]") or msg.startswith("[SKIP-UPTODATE]") or msg.startswith("[SKIP-YEAR-NAN]"): skipped += 1
            elif success: ok += 1
            else: errors += 1
    print(f"\n[Stage 1] completed: total {total}, successful {ok}, skipped {skipped}, errors {errors}")

# ================== Stage 2:Loads + Infil + HVAC + Outputs ==================
THERMOSTAT_NAME = "HVACTemp"
HEAT_SP_SCH     = "Heating_SP_Schedule"
COOL_SP_SCH     = "Cooling_SP_Schedule"
HVAC_AVAIL_SCH  = "HVAC_ConditionedTime_Schedule"

SCH_PEOPLE_PRESENCE = "People_Schedule"
SCH_PEOPLE_ACTIVITY = "CN_GenericActivityLevel"
SCH_LIGHTS          = "Lights_Schedule"
SCH_EQUIP           = "ElectricEquipment_Schedule"
SCH_INFIL           = "AirEx_ConditionedTime_Schedule"

Function_name_list = {
    "Residential_1": 0, "Residential_2": 1, "Residential_3": 2,
    "Commercial": 3, "Office": 4, "Industry": 5, "Transport": 6, "Administration": 7
}
CODE_TO_FUNC = {v: k for k, v in Function_name_list.items()}
FUNC_TO_EXCEL_BT = {
    "Residential_1": "Residential",
    "Residential_2": "Residential",
    "Residential_3": "Residential",
    "Commercial":    "Commercial",
    "Office":        "Office",
    "Industry":      "Industrial",
    "Transport":     "Transport",
    "Administration":"Administrative",
}

people_df: pd.DataFrame = None
people_years: List[int] = []
lights_df: pd.DataFrame = None
lights_years: List[int] = []
equip_df: pd.DataFrame  = None
equip_years: List[int]  = []

_DEF_OUTVAR: List[Tuple[str, str, str]] = []  # (Key_Value, Variable_Name, Reporting_Frequency)
_DEF_OUTMETER: List[Tuple[str, str]] = []     # (Key_Name, Reporting_Frequency)

def load_sheet_single_or_multi_header(excel_path: str, sheet: str) -> Tuple[pd.DataFrame, List[int]]:
    df = pd.read_excel(excel_path, sheet_name=sheet)
    df = df.copy(); df.columns = [str(c).strip() for c in df.columns]
    if "BuildingType" in df.columns:
        years = []
        for c in df.columns[1:]:
            try: years.append(int(str(c).strip()))
            except Exception: pass
        if years:
            years = sorted(years)
            df["BuildingType"] = df["BuildingType"].astype(str).str.strip()
            return df[["BuildingType"] + [str(y) for y in years]].set_index("BuildingType"), years
    dfm = pd.read_excel(excel_path, sheet_name=sheet, header=[0, 1])
    first_col_candidates = [c for c in dfm.columns if str(c[0]).strip().lower() == "buildingtype"]
    first_col = first_col_candidates[0] if first_col_candidates else dfm.columns[0]
    bt = dfm[first_col].astype(str).str.strip().rename("BuildingType")
    years = []
    for _, sub in dfm.columns:
        try: years.append(int(str(sub).strip()))
        except Exception: pass
    years = sorted(set(years))
    if not years:
        raise ValueError(f"{sheet}: Excel has no year columns")
    out = pd.DataFrame({"BuildingType": bt})
    for y in years:
        block = dfm.xs(y, level=1, axis=1)
        col = block.columns[0]
        out[str(y)] = pd.to_numeric(block[col], errors="coerce")
    return out.set_index("BuildingType"), years

def _load_intensity_tables():
    global people_df, people_years, lights_df, lights_years, equip_df, equip_years
    people_df, people_years = load_sheet_single_or_multi_header(STATIC_XLS, "People")
    lights_df, lights_years = load_sheet_single_or_multi_header(STATIC_XLS, "Lights")
    equip_df,  equip_years  = load_sheet_single_or_multi_header(STATIC_XLS, "ElectricEquipment")

def closest_year_le(years_sorted: List[int], target: int) -> Optional[int]:
    cand = None
    for y in years_sorted:
        if y <= target: cand = y
        else: break
    return cand if cand is not None else (years_sorted[0] if years_sorted else None)

def get_intensity_safe(df: pd.DataFrame, years_sorted: List[int], bt_name: str, target_year: int, sheetname: str) -> Tuple[float, int]:
    if bt_name not in df.index:
        raise KeyError(f"{sheetname}: Excel does not contain BuildingType='{bt_name}'")
    y = closest_year_le(years_sorted, int(target_year))
    if y is not None:
        try:
            v = df.loc[bt_name, str(y)]
            if pd.notna(v):
                return float(v), int(y)
        except Exception:
            pass
    best_year, best_gap = None, 10**9
    for y2 in years_sorted:
        try:
            v2 = df.loc[bt_name, str(y2)]
            if pd.notna(v2):
                gap = abs(int(y2) - int(target_year))
                if gap < best_gap:
                    best_gap, best_year = gap, int(y2)
        except Exception:
            continue
    if best_year is not None:
        return float(df.loc[bt_name, str(best_year)]), best_year
    raise KeyError(f"{sheetname}: '{bt_name}' has empty values for all years")

def ensure_activity_schedule(idf: IDF):
    def normalize(name): return str(name).strip().replace(" ", "").lower()
    for o in list(idf.idfobjects.get("SCHEDULETYPELIMITS", [])):
        if normalize(o.Name) == "activitylevel":
            o.Name = "ActivityLevel"
            try: o.Numeric_Type = "Continuous"
            except Exception: pass
            try: o.Unit_Type = "Dimensionless"
            except Exception: pass
            break

def _remove_all(idf: IDF, objtype: str):
    for o in list(idf.idfobjects.get(objtype.upper(), [])):
        idf.removeidfobject(o)

def _new_hvac_template_thermostat(idf: IDF):
    _remove_all(idf, "HVACTemplate:Thermostat")
    idf.newidfobject(
        "HVACTEMPLATE:THERMOSTAT",
        Name=THERMOSTAT_NAME,
        Heating_Setpoint_Schedule_Name=HEAT_SP_SCH,
        Cooling_Setpoint_Schedule_Name=COOL_SP_SCH,
    )

def _new_ideal_loads_for_all_zones(idf: IDF) -> int:
    _remove_all(idf, "HVACTemplate:Zone:IdealLoadsAirSystem")
    zones = list(idf.idfobjects.get("ZONE", []))
    cnt = 0
    for z in zones:
        zn = str(z.Name).strip()
        idf.newidfobject(
            "HVACTEMPLATE:ZONE:IDEALLOADSAIRSYSTEM",
            Zone_Name=zn,
            Template_Thermostat_Name=THERMOSTAT_NAME,
            System_Availability_Schedule_Name=HVAC_AVAIL_SCH,
            Heating_Availability_Schedule_Name="Heating_Period_Schedule",
            Cooling_Availability_Schedule_Name="Cooling_Period_Schedule",
        )
        cnt += 1
    return cnt

def ensure_hvac_templates(idf: IDF) -> int:
    _new_hvac_template_thermostat(idf)
    return _new_ideal_loads_for_all_zones(idf)

def remove_existing_for_zone(idf: IDF, objtype: str, zone_name: str):
    for o in list(idf.idfobjects.get(objtype.upper(), [])):
        nm = str(o.Name)
        if objtype.upper() == "PEOPLE" and nm == f"People_{zone_name}":
            idf.removeidfobject(o)
        elif objtype.upper() == "LIGHTS" and nm == f"Lights_{zone_name}":
            idf.removeidfobject(o)
        elif objtype.upper() == "ELECTRICEQUIPMENT" and nm == f"ElectricEquipment_{zone_name}":
            idf.removeidfobject(o)

def remove_infiltration_for_zone(idf: IDF, zone_name: str):
    for o in list(idf.idfobjects.get("ZONEINFILTRATION:DESIGNFLOWRATE", [])):
        if str(o.Zone_or_ZoneList_or_Space_or_SpaceList_Name).strip() == zone_name.strip():
            idf.removeidfobject(o)

def write_loads_for_all_zones(idf: IDF, bt_excel: str, year: int):
    p_val, _ = get_intensity_safe(people_df, people_years, bt_excel, year, "People")
    l_val, _ = get_intensity_safe(lights_df, lights_years, bt_excel, year, "Lights")
    e_val, _ = get_intensity_safe(equip_df,  equip_years,  bt_excel, year, "ElectricEquipment")
    zones = idf.idfobjects.get("ZONE", [])
    for z in zones:
        zn = str(z.Name).strip()
        remove_existing_for_zone(idf, "PEOPLE", zn)
        remove_existing_for_zone(idf, "LIGHTS", zn)
        remove_existing_for_zone(idf, "ELECTRICEQUIPMENT", zn)
        idf.newidfobject(
            "PEOPLE",
            Name=f"People_{zn}",
            Zone_or_ZoneList_or_Space_or_SpaceList_Name=zn,
            Number_of_People_Schedule_Name=SCH_PEOPLE_PRESENCE,
            Number_of_People_Calculation_Method="Area/Person",
            Floor_Area_per_Person=p_val,
            Activity_Level_Schedule_Name=SCH_PEOPLE_ACTIVITY,
        )
        idf.newidfobject(
            "LIGHTS",
            Name=f"Lights_{zn}",
            Zone_or_ZoneList_or_Space_or_SpaceList_Name=zn,
            Schedule_Name=SCH_LIGHTS,
            Design_Level_Calculation_Method="Watts/Area",
            Watts_per_Zone_Floor_Area=l_val,
        )
        idf.newidfobject(
            "ELECTRICEQUIPMENT",
            Name=f"ElectricEquipment_{zn}",
            Zone_or_ZoneList_or_Space_or_SpaceList_Name=zn,
            Schedule_Name=SCH_EQUIP,
            Design_Level_Calculation_Method="Watts/Area",
            Watts_per_Zone_Floor_Area=e_val,
        )

def write_infiltration_for_all_zones(idf: IDF) -> int:
    zones = idf.idfobjects.get("ZONE", [])
    count = 0
    for idx, z in enumerate(zones, start=1):
        zn = str(z.Name).strip()
        remove_infiltration_for_zone(idf, zn)
        idf.newidfobject(
            "ZONEINFILTRATION:DESIGNFLOWRATE",
            Name=f"Infil {idx}",
            Zone_or_ZoneList_or_Space_or_SpaceList_Name=zn,
            Schedule_Name=SCH_INFIL,
            Design_Flow_Rate_Calculation_Method="AirChanges/Hour",
            Air_Changes_per_Hour=1,
        )
        count += 1
    return count

def ensure_output_table_style(idf: IDF):
    objs = list(idf.idfobjects.get("OUTPUTCONTROL:TABLE:STYLE", []))
    if not objs:
        idf.newidfobject(
            "OUTPUTCONTROL:TABLE:STYLE",
            Column_Separator="Comma",
            Unit_Conversion="JtoKWH",
        )
    else:
        for o in objs:
            o.Column_Separator = "Comma"
            o.Unit_Conversion  = "JtoKWH"

def ensure_output_summary_reports(idf: IDF):
    objs = list(idf.idfobjects.get("OUTPUT:TABLE:SUMMARYREPORTS", []))
    if not objs:
        idf.newidfobject("OUTPUT:TABLE:SUMMARYREPORTS", Report_1_Name="AllSummary")
    else:
        for o in objs:
            o.Report_1_Name = "AllSummary"

def _apply_output_templates_from_cache(dst_idf: IDF):
    for o in list(dst_idf.idfobjects.get("OUTPUT:VARIABLE", [])):
        dst_idf.removeidfobject(o)
    for kv, vn, rf in _DEF_OUTVAR:
        dst_idf.newidfobject("OUTPUT:VARIABLE", Key_Value=kv, Variable_Name=vn, Reporting_Frequency=rf)
    for o in list(dst_idf.idfobjects.get("OUTPUT:METER:METERFILEONLY", [])):
        dst_idf.removeidfobject(o)
    for kn, rf in _DEF_OUTMETER:
        dst_idf.newidfobject("OUTPUT:METER:METERFILEONLY", Key_Name=kn, Reporting_Frequency=rf)

def _init_stage2(idd_path: str, template_outputs_idf: str):
    IDF.setiddname(idd_path)
    _load_intensity_tables()
    _DEF_OUTVAR.clear(); _DEF_OUTMETER.clear()
    try:
        if template_outputs_idf and os.path.isfile(template_outputs_idf):
            tmp = IDF(template_outputs_idf)
            for s in tmp.idfobjects.get("OUTPUT:VARIABLE", []):
                kv = getattr(s, "Key_Value", "*") or "*"
                vn = getattr(s, "Variable_Name", "") or ""
                rf = getattr(s, "Reporting_Frequency", "Hourly") or "Hourly"
                _DEF_OUTVAR.append((kv, vn, rf))
            for s in tmp.idfobjects.get("OUTPUT:METER:METERFILEONLY", []):
                kn = getattr(s, "Key_Name", "") or ""
                rf = getattr(s, "Reporting_Frequency", "Hourly") or "Hourly"
                _DEF_OUTMETER.append((kn, rf))
    except Exception:
        pass

def _stage2_worker(fp: str, in_root_str: str, out_root_str: str, dry_run: bool) -> Tuple[bool, str]:
    try:
        if _year_token_is_nan(fp):
            return True, f"[SKIP-YEAR-NAN] {Path(fp).name}"
        p = Path(fp); in_root = Path(in_root_str); out_root = Path(out_root_str)
        if (not dry_run) and is_up_to_date(p, in_root, out_root):
            return True, f"[SKIP-UPTODATE] {p.name}"
        parts = parse_filename_parts(p)
        func_code = int(parts["func_index"]) ; year = int(parts["year"])
        func_name = CODE_TO_FUNC.get(func_code)
        excel_bt  = FUNC_TO_EXCEL_BT.get(func_name)
        if dry_run:
            return True, f"[DRY-2] {p.name} func={func_code}({func_name}/{excel_bt}) year={year}"
        idf = IDF(str(p))

        normalize_names_everywhere(idf, year)
        fixed, unresolved = fix_construction_refs(idf)
        if SHOW_LOG and (fixed or unresolved):
            print(f"  [CONSREF] repaired={fixed}, unresolved={sorted(unresolved) if unresolved else []}")

        ensure_hvac_templates(idf)
        ensure_activity_schedule(idf)
        if excel_bt is not None and year is not None:
            write_loads_for_all_zones(idf, excel_bt, year)
        write_infiltration_for_all_zones(idf)
        ensure_output_table_style(idf)
        ensure_output_summary_reports(idf)
        _apply_output_templates_from_cache(idf)
        out_path = compute_out_path(p, in_root, out_root)
        idf.saveas(str(out_path))
        return True, f"[OK-2] {out_path}"
    except Exception as e:
        return False, f"[ERR-2] {fp}: {e}"

def stage2_write_loads_hvac_outputs(idf_root: Path, out_root: Path):
    files = _all_idf_files(Path(idf_root))
    ok = fail = 0
    if MAX_WORKERS > 1:
        with ProcessPoolExecutor(max_workers=MAX_WORKERS, initializer=_init_stage2, initargs=(IDD_PATH, TEMPLATE_OUTPUTS_IDF)) as ex:
            futures = [ex.submit(_stage2_worker, str(fp), str(idf_root), str(out_root), DRY_RUN_STAGE2) for fp in files]
            for fut in as_completed(futures):
                success, msg = fut.result();
                if SHOW_LOG: print(msg)
                if success: ok += 1
                else: fail += 1
    else:
        _init_stage2(IDD_PATH, TEMPLATE_OUTPUTS_IDF)
        for fp in files:
            success, msg = _stage2_worker(str(fp), str(idf_root), str(out_root), DRY_RUN_STAGE2)
            if SHOW_LOG: print(msg)
            if success: ok += 1
            else: fail += 1
    print(f"\n[Stage 2] completed: successful {ok}, failed {fail}")

# ================== Stage 3: insulation and window materials ==================
WIN_RE = re.compile(
    rf"(?P<cz>{CZ_TOKEN})\s*[_\-]\s*"
    r"(?P<btype>Residential|Public(?:&|and)?Industrial|Public|Industrial)\s*[_\-]\s*"
    r"(?P<year>(?:19|20)\d{2})(?:\.0)?",
    re.IGNORECASE,
)
INSUL_NAME_FLOAT  = re.compile(r"insulation\s*([0-9]*\.?[0-9]+)", re.IGNORECASE)
INSUL_NAME_STRICT = re.compile(r"^\s*insulation\s*([0-9]*\.?[0-9]+)\s*$", re.IGNORECASE)

THIN_LAYER_THRESHOLD_M   = 0.0002
TOL_MATCH_FLOAT          = 5e-4
TOL_THICKNESS_PHYS_ZERO  = 1e-9
INSUL_K, INSUL_RHO, INSUL_CP = 0.04, 40.0, 1500.0

PURGE_OLD_INSULATION = True
REPLACE_ALL          = True

def _normalize_col(x: str) -> str:
    return re.sub(r"\s+", "", str(x).strip().lower())

def _normalize_cz(cz_raw: Optional[str], allow_any=True) -> Optional[str]:
    if not cz_raw: return None
    s = str(cz_raw).strip().upper()
    m = re.match(r"^(SCZ)(\d+)$", s)
    if m: return f"SCZ{int(m.group(2))}"
    if s in {"CZ", "HSCWZ", "HSWWZ", "TZ"}: return CZ_ANY if allow_any else None
    return None

def q3(x: float) -> float: return round(float(x) + 1e-12, 3)
def q3s(x: float) -> str:  return f"{q3(x):.3f}"
def mat_name(t_m: float) -> str: return f"Insulation {q3s(t_m)}"

def _extract_insul_float(name: str) -> Optional[float]:
    m = INSUL_NAME_FLOAT.search(str(name) if name is not None else "")
    if not m: return None
    try: return float(m.group(1))
    except Exception: return None

def _canon_insul_name(name: str) -> str:
    s = str(name).strip()
    m = INSUL_NAME_STRICT.match(s)
    if not m: return s
    try: return mat_name(float(m.group(1)))
    except Exception: return s

def _melt_sheet(df: pd.DataFrame, btype_hint: str) -> Optional[pd.DataFrame]:
    df = df.copy(); df.columns = [_normalize_col(c) for c in df.columns]
    if not {"zone","insulationthickness"} <= set(df.columns): return None
    year_cols = [c for c in df.columns if re.fullmatch(r"\d{4}", c)]
    if not year_cols: return None
    long = df.melt(id_vars=["zone","insulationthickness"], value_vars=year_cols,
                   var_name="Year", value_name="thick_mm")
    from_part = df["insulationthickness"].apply(lambda x: _canon_part(str(x)))
    long["Envelope"]    = from_part
    long["ClimateZone"] = long["zone"].astype(str).str.upper().str.strip()
    long["Year"]        = pd.to_numeric(long["Year"], errors="coerce").astype("Int64")
    long["Thickness_m"] = pd.to_numeric(long["thick_mm"], errors="coerce") / 1000.0
    long = long.dropna(subset=["Envelope","Thickness_m"])
    long = long[["ClimateZone","Envelope","Year","Thickness_m"]]
    long["BType"] = btype_hint
    return long

def load_thickness_mapping(xlsx: Union[str, Path]) -> pd.DataFrame:
    xls = pd.ExcelFile(xlsx); parts=[]
    for sheet in xls.sheet_names:
        hint = "Residential" if ("residen" in sheet.lower()) else "Public&Industrial"
        chunk = _melt_sheet(xls.parse(sheet), hint)
        if chunk is not None and not chunk.empty: parts.append(chunk)
    if not parts: raise ValueError("No insulation-thickness records were parsed.")
    df = pd.concat(parts, ignore_index=True)
    df["Thickness_m"] = df["Thickness_m"].astype(float)
    return df

def _canon_window_metric(s: str) -> Optional[str]:
    low = str(s).lower().replace(" ", "")
    if low in ("winu","windowu","u","uw"): return "U"
    if low in ("winshgc","windowshgc","shgc"): return "SHGC"
    return None

def load_window_maps(xlsx_path: Union[str, Path]) -> pd.DataFrame:
    xls = pd.ExcelFile(xlsx_path); parts=[]
    for sheet in xls.sheet_names:
        hint_bt = _canon_btype(sheet)
        df = xls.parse(sheet); df.columns = [_normalize_col(c) for c in df.columns]
        zone_col = next((c for c in ["zone", "climatezone", "cz", "scz"] if c in df.columns), None)
        env_col  = next((c for c in ["envelope", "part", "category"] if c in df.columns), None)
        if not zone_col or not env_col: continue
        year_cols = [c for c in df.columns if re.fullmatch(r"\d{4}", c)]
        if not year_cols: continue
        melt = df.melt(id_vars=[zone_col,env_col], value_vars=year_cols,
                       var_name="Year", value_name="Value")
        melt["Metric"]      = melt[env_col].apply(_canon_window_metric)
        melt = melt[melt["Metric"].isin(["U","SHGC"])]
        melt["ClimateZone"] = melt[zone_col].astype(str).str.upper().str.strip()
        cz = melt["ClimateZone"].apply(lambda s: _normalize_cz(s, allow_any=True) or CZ_ANY)
        melt["ClimateZone"] = cz
        melt["BType"]       = hint_bt
        melt["Year"]        = pd.to_numeric(melt["Year"], errors="coerce").astype("Int64")
        melt["Value"]       = pd.to_numeric(melt["Value"], errors="coerce")
        melt = melt.dropna(subset=["Value"])
        wide = melt.pivot_table(index=["ClimateZone","BType","Year"], columns="Metric",
                                values="Value", aggfunc="mean").reset_index()
        parts.append(wide)
    if not parts: raise ValueError("Could not parse WinU/WinSHGC records from Static-10.xlsx.")
    win = pd.concat(parts, ignore_index=True)
    for col in ["U","SHGC"]:
        if col not in win.columns: win[col] = pd.NA
    win_any = win.copy(); win_any["BType"] = "*"
    return pd.concat([win, win_any], ignore_index=True)

def purge_old_insulation(idf: IDF) -> int:
    removed = 0
    for key in ["MATERIAL", "MATERIAL:NOMASS"]:
        for obj in list(idf.idfobjects[key]):
            if "insulation" in str(obj.Name).lower():
                idf.removeidfobject(obj); removed += 1
    return removed

def ensure_insulation(idf: IDF, t_nominal_m: float) -> Tuple[Optional[str], bool]:
    if float(t_nominal_m) < THIN_LAYER_THRESHOLD_M:
        return None, False
    name = mat_name(q3(t_nominal_m))
    for m in idf.idfobjects["MATERIAL"]:
        if str(m.Name).strip().lower() == name.lower():
            m.Conductivity  = INSUL_K
            m.Density       = INSUL_RHO
            m.Specific_Heat = INSUL_CP
            if hasattr(m, "Thermal_Absorptance"): m.Thermal_Absorptance = 0.9
            if hasattr(m, "Solar_Absorptance"):   m.Solar_Absorptance   = 0.7
            if hasattr(m, "Visible_Absorptance"): m.Visible_Absorptance = 0.7
            m.Thickness = float(q3(t_nominal_m))
            return m.Name, False
    idf.newidfobject(
        "MATERIAL",
        Name=name,
        Roughness="Rough",
        Thickness=float(q3(t_nominal_m)),
        Conductivity=INSUL_K,
        Density=INSUL_RHO,
        Specific_Heat=INSUL_CP,
        Thermal_Absorptance=0.9,
        Solar_Absorptance=0.7,
        Visible_Absorptance=0.7,
    )
    return name, True

def ensure_insulation_alias_exact(idf: IDF, alias_name: str) -> bool:
    nm = str(alias_name).strip()
    v = _extract_insul_float(nm)
    if v is None or v < THIN_LAYER_THRESHOLD_M:
        return False
    for mm in idf.idfobjects["MATERIAL"]:
        if str(mm.Name).strip().lower() == nm.lower():
            return False
    idf.newidfobject(
        "MATERIAL",
        Name=nm,
        Roughness="Rough",
        Thickness=float(q3(v)),
        Conductivity=INSUL_K,
        Density=INSUL_RHO,
        Specific_Heat=INSUL_CP,
        Thermal_Absorptance=0.9,
        Solar_Absorptance=0.7,
        Visible_Absorptance=0.7,
    )
    return True

def _get_material_obj(idf: IDF, mat_name: str):
    for key in ["MATERIAL", "MATERIAL:NOMASS", "WINDOWMATERIAL:SIMPLEGLAZINGSYSTEM"]:
        for m in idf.idfobjects.get(key, []):
            if str(m.Name).strip().lower() == str(mat_name).strip().lower():
                return m
    return None

def _mat_thickness(m) -> Optional[float]:
    if m is None: return None
    if hasattr(m, "Thickness"):
        try: return float(getattr(m, "Thickness") or 0.0)
        except Exception: return None
    return None

def remove_zero_like_layers_and_compact(idf: IDF, cons):
    names = get_layer_names_ordered(cons)
    keep, removed = [], 0
    for nm in names:
        s = str(nm).strip()
        rm = False
        v = _extract_insul_float(s)
        if v is not None and abs(v) <= TOL_MATCH_FLOAT:
            rm = True
        if not rm:
            mt = _get_material_obj(idf, s)
            t = _mat_thickness(mt)
            if t is not None and t <= TOL_THICKNESS_PHYS_ZERO:
                rm = True
        if rm: removed += 1
        else: keep.append(s)
    if not keep:
        return cons, removed
    new_cons = rebuild_construction_with_layers(idf, cons, keep)
    if removed and SHOW_LOG:
        print(f"  [CLEAN] {new_cons.Name}: removed {removed} zero-thickness layer(s) and rebuilt the construction")
    return new_cons, removed

def parse_win_name(name: str) -> Tuple[Optional[str], Optional[str], Optional[int]]:
    m = WIN_RE.search(str(name))
    if not m: return None, None, None
    cz = m.group("cz").upper().strip()
    bt = _canon_btype(m.group("btype"))
    yr = int(m.group("year"))
    return cz, bt, yr

def pick_thickness(mapping: pd.DataFrame, cz: Optional[str], bt: str, env: str, yr: int) -> Optional[float]:
    def _select(sub: pd.DataFrame) -> Optional[float]:
        if sub.empty: return None
        exact = sub[sub["Year"] == yr]
        if not exact.empty: return float(exact.iloc[0]["Thickness_m"])
        leq = sub[sub["Year"].astype("float").fillna(-1) <= yr]
        if not leq.empty: return float(leq.sort_values("Year").iloc[-1]["Thickness_m"])
        return float(sub.sort_values("Year").iloc[-1]["Thickness_m"])
    env_candidates = [env]
    if env in ("ExtWall", "IntWall"): env_candidates += ["Wall"]
    if env in ("FloorSlab", "FloorGround"): env_candidates += ["Floor"]
    for cz_try in [cz, CZ_ANY]:
        for env_try in env_candidates:
            sub = mapping[(mapping["BType"]==bt) & (mapping["Envelope"]==env_try)]
            if cz_try and cz_try != CZ_ANY:
                sub = sub[sub["ClimateZone"].str.upper()==cz_try.upper()]
            val = _select(sub)
            if val is not None: return val
    return None

def replace_or_delete_insulation(idf: IDF, cons, target_t_nominal: float, replace_all: bool=True):
    cons, _ = remove_zero_like_layers_and_compact(idf, cons)
    names = get_layer_names_ordered(cons)
    if float(target_t_nominal) < THIN_LAYER_THRESHOLD_M:
        keep, hits = [], []
        for idx, nm in enumerate(names, start=1):
            if "insulation" in str(nm).lower():
                hits.append(idx)
            else:
                keep.append(nm)
        changed = len(hits)
        if changed > 0 and keep:
            cons = rebuild_construction_with_layers(idf, cons, keep)
        cons, _ = remove_zero_like_layers_and_compact(idf, cons)
        return cons, changed, hits, None, 0
    target_name, created_now = ensure_insulation(idf, q3(target_t_nominal))
    target_name = _canon_insul_name(target_name)
    hits, changed = [], 0
    for idx, nm in enumerate(names, start=1):
        if "insulation" not in str(nm).lower():
            continue
        names[idx-1] = target_name
        hits.append(idx)
        changed += 1
        if not replace_all:
            break
    if changed > 0:
        cons = rebuild_construction_with_layers(idf, cons, names)
    cons, _ = remove_zero_like_layers_and_compact(idf, cons)
    return cons, changed, hits, target_name, int(bool(created_now))

def _sanitize(s: str) -> str:
    return re.sub(r"[^A-Za-z0-9]+", "", str(s)).lower()

def _set_window_field(obj, label: str, value: float) -> bool:
    try:
        obj.setfield(label, float(value)); return True
    except Exception:
        pass
    try:
        if _sanitize(label).startswith("u"):
            for attr in ["UFactor", "U_Factor"]:
                try: setattr(obj, attr, float(value)); return True
                except Exception: continue
        else:
            for attr in ["Solar_Heat_Gain_Coefficient", "SolarHeatGainCoefficient"]:
                try: setattr(obj, attr, float(value)); return True
                except Exception: continue
    except Exception:
        pass
    want = _sanitize(label)
    for fname in getattr(obj, "fieldnames", []):
        if fname.lower() == "name":
            continue
        if (_sanitize(fname) == want) or (want in _sanitize(fname)) or (_sanitize(fname) in want):
            try:
                obj.setfield(fname, float(value)); return True
            except Exception:
                continue
    return False

def update_windows_from_map(idf: IDF, winmap: pd.DataFrame) -> Tuple[int, int]:
    updated = 0; missed = 0
    try:
        objs = list(idf.idfobjects["WINDOWMATERIAL:SIMPLEGLAZINGSYSTEM"])
    except Exception:
        objs = []
    if SHOW_LOG:
        print(f"[WIN] SimpleGlazingSystem object count:{len(objs)}")
    for w in objs:
        cz, bt, yr = parse_win_name(str(w.Name))
        if not all([bt, yr]):
            missed += 1
            if SHOW_LOG:
                print(f"  [WIN][SKIP] {w.Name}: name does not match <CZ>_<BType>_<Year>_*")
            continue
        row = None
        for c,b in [(cz if cz and cz!=CZ_ANY else None, bt), (cz if cz and cz!=CZ_ANY else None, "*"), (CZ_ANY, bt), (CZ_ANY, "*")]:
            if c is None: continue
            cand = winmap[(winmap["ClimateZone"]==c) & (winmap["BType"]==b)]
            if cand.empty: continue
            exact = cand[cand["Year"]==yr]
            if not exact.empty: row = exact.iloc[0]; break
            leq = cand[cand["Year"].astype("float").fillna(-1) <= yr]
            if not leq.empty: row = leq.sort_values("Year").iloc[-1]; break
            row = cand.sort_values("Year").iloc[-1]; break
        if row is None:
            missed += 1
            if SHOW_LOG:
                print(f"  [WIN][MISS] {w.Name}: mapping is missing (cz={cz or 'ANY'}/{bt}/{yr})")
            continue
        ok_u = ok_s = True
        if "U" in row and not pd.isna(row["U"]):
            ok_u = _set_window_field(w, "U-Factor", float(row["U"]))
        if "SHGC" in row and not pd.isna(row["SHGC"]):
            ok_s = _set_window_field(w, "Solar Heat Gain Coefficient", float(row["SHGC"]))
        if SHOW_LOG:
            print(f"  [WIN][TRY] {w.Name}: set U={row.get('U', 'N/A')}({ok_u}), SHGC={row.get('SHGC', 'N/A')}({ok_s})")
        if ok_u and ok_s:
            updated += 1
        else:
            missed += 1
    return updated, missed

_thickness_map: Optional[pd.DataFrame] = None
_window_map: Optional[pd.DataFrame] = None

def _init_stage3(idd_path: str, insul_xls: str, static_xls: str):
    IDF.setiddname(idd_path)
    global _thickness_map, _window_map
    _thickness_map = load_thickness_mapping(insul_xls)
    _window_map    = load_window_maps(static_xls)

def _stage3_worker(fp: str, in_root_str: str, out_root_str: str, dry_run: bool) -> Tuple[bool, str]:
    try:
        if _year_token_is_nan(fp):
            return True, f"[SKIP-YEAR-NAN] {Path(fp).name}"
        p = Path(fp); in_root = Path(in_root_str); out_root = Path(out_root_str)
        if (not dry_run) and is_up_to_date(p, in_root, out_root):
            return True, f"[SKIP-UPTODATE] {p.name}"
        parts = parse_filename_parts(p)
        file_year = int(parts["year"])
        idf = IDF(str(p))

        normalize_names_everywhere(idf, file_year)

        fixed, unresolved = fix_construction_refs(idf)
        if SHOW_LOG and (fixed or unresolved):
            print(f"  [CONSREF] repaired={fixed}, unresolved={sorted(unresolved) if unresolved else []}")

        if PURGE_OLD_INSULATION:
            rm = purge_old_insulation(idf)
            if SHOW_LOG:
                print(f"  [MAT] removed old insulation materials: {rm}")
        for t in sorted({ float(t) for t in _thickness_map["Thickness_m"].dropna().tolist() if float(t) >= THIN_LAYER_THRESHOLD_M }):
            ensure_insulation(idf, t)
        for cons in idf.idfobjects["CONSTRUCTION"]:
            for nm in get_layer_names_ordered(cons):
                if "insulation" in str(nm).lower():
                    ensure_insulation_alias_exact(idf, nm)
        for cons in list(idf.idfobjects["CONSTRUCTION"]):
            cz, bt, yr, env = parse_from_name(str(cons.Name))
            if bt and not yr:
                yr = file_year
            if not all([bt, yr, env]):
                if SHOW_LOG:
                    print(f"  [SKIP] {cons.Name}: not matched <CZ>_<BType>_<Year>_<Part>")
                continue
            t_m = pick_thickness(_thickness_map, cz, bt, env, int(yr))
            if t_m is None:
                if SHOW_LOG:
                    print(f"  [MISS] {cons.Name}: no thickness mapping ({cz or 'ANY'}/{bt}/{env}/{yr})")
                continue
            cons_new, nrep, pos, target, _ = replace_or_delete_insulation(idf, cons, t_m, replace_all=True)
            if SHOW_LOG:
                print(f"    - {cons.Name}: pick={q3s(t_m)} -> {target or 'delete insulation layer'}  layers={pos or []}  changed={nrep}")
        wupd, wmiss = update_windows_from_map(idf, _window_map)
        if SHOW_LOG:
            print(f"  [WIN] updated window material objects:{wupd}; missed:{wmiss}")
        if dry_run:
            return True, f"[DRY-3] {p.name}"
        out_path = compute_out_path(p, in_root, out_root)
        idf.saveas(str(out_path))
        return True, f"[OK-3] saved -> {out_path}"
    except Exception as e:
        return False, f"[ERR-3] {fp}: {e}"

def stage3_refresh_insulation_and_windows(idf_root: Path, out_root: Path):
    files = _all_idf_files(Path(idf_root))
    ok = fail = 0
    if MAX_WORKERS > 1:
        with ProcessPoolExecutor(max_workers=MAX_WORKERS, initializer=_init_stage3, initargs=(IDD_PATH, INSUL_XLS, STATIC_XLS)) as ex:
            futures = [ex.submit(_stage3_worker, str(fp), str(idf_root), str(out_root), DRY_RUN_STAGE3) for fp in files]
            for fut in as_completed(futures):
                success, msg = fut.result();
                if SHOW_LOG: print(msg)
                if success: ok += 1
                else: fail += 1
    else:
        _init_stage3(IDD_PATH, INSUL_XLS, STATIC_XLS)
        for fp in files:
            success, msg = _stage3_worker(str(fp), str(idf_root), str(out_root), DRY_RUN_STAGE3)
            if SHOW_LOG: print(msg)
            if success: ok += 1
            else: fail += 1
    print(f"\n[Stage 3] completed: successful {ok}, failed {fail}")

# ================== Final-only parallel processing ==================
def _init_common(idd_path: str, schedule_xls: str, climate_xls: Optional[str],
                 static_xls: str, insul_xls: str, template_outputs_idf: str):
    IDF.setiddname(idd_path)
    global _SHEET_MAP, _CITY2ZONE, _thickness_map, _window_map
    _SHEET_MAP = read_wide_sheets(schedule_xls)
    _CITY2ZONE = read_city_to_zone(climate_xls) if climate_xls else {}
    _load_intensity_tables()
    _thickness_map = load_thickness_mapping(insul_xls)
    _window_map    = load_window_maps(static_xls)
    _DEF_OUTVAR.clear(); _DEF_OUTMETER.clear()
    try:
        if template_outputs_idf and os.path.isfile(template_outputs_idf):
            tmp = IDF(template_outputs_idf)
            for s in tmp.idfobjects.get("OUTPUT:VARIABLE", []):
                kv = getattr(s, "Key_Value", "*") or "*"
                vn = getattr(s, "Variable_Name", "") or ""
                rf = getattr(s, "Reporting_Frequency", "Hourly") or "Hourly"
                _DEF_OUTVAR.append((kv, vn, rf))
            for s in tmp.idfobjects.get("OUTPUT:METER:METERFILEONLY", []):
                kn = getattr(s, "Key_Name", "") or ""
                rf = getattr(s, "Reporting_Frequency", "Hourly") or "Hourly"
                _DEF_OUTMETER.append((kn, rf))
    except Exception:
        pass

def _apply_stage1_to_idf(idf: IDF, func_index: int, city: str) -> int:
    func_cols = FUNC_INDEX_TO_COL.get(func_index, [])
    if isinstance(func_cols, str):
        func_cols = [func_cols]
    zone_txt = _CITY2ZONE.get(city) or { _norm(k): v for k, v in _CITY2ZONE.items() }.get(_norm(city))
    zone_candidates: List[str] = []
    if zone_txt:
        zone_candidates = [zone_txt, zone_txt.replace(" ", ""), zone_txt.replace("-", "")]
    schedules = collect_schedules_for_idf(_SHEET_MAP, func_cols, zone_candidates)
    for s in schedules:
        ensure_typelimits(idf, s.get("TypeLimitsName", "").strip())
        upsert_schedule_compact(idf, s.get("ScheduleName", "").strip(), s.get("TypeLimitsName", "").strip(),
                                s.get("Lines", []) or [])
    return len(schedules)

def _final_worker(fp: str, idf_root_str: str, final_root_str: str, show_log: bool=False) -> Tuple[bool, str]:
    try:
        if _year_token_is_nan(fp):
            return True, f"[SKIP-YEAR-NAN] {Path(fp).name}"
        p = Path(fp); idf_root = Path(idf_root_str); final_root = Path(final_root_str)
        if is_up_to_date(p, idf_root, final_root):
            return True, f"[SKIP-UPTODATE] {p.name}"
        parts = parse_filename_parts(p)
        file_year = int(parts["year"])
        idf = IDF(str(p))

        normalize_names_everywhere(idf, file_year)
        fixed, unresolved = fix_construction_refs(idf)
        if show_log and (fixed or unresolved):
            print(f"[CONSREF] repaired={fixed}, unresolved={sorted(unresolved) if unresolved else []}")

        nsch = _apply_stage1_to_idf(idf, int(parts["func_index"]), str(parts["city"]))
        if show_log: print(f"[MEM-1] {p.name}: schedules={nsch}")

        func_name = CODE_TO_FUNC.get(int(parts["func_index"]))
        excel_bt  = FUNC_TO_EXCEL_BT.get(func_name)
        ensure_hvac_templates(idf)
        ensure_activity_schedule(idf)
        if excel_bt is not None and parts.get("year") is not None:
            write_loads_for_all_zones(idf, excel_bt, int(parts["year"]))
        write_infiltration_for_all_zones(idf)
        ensure_output_table_style(idf)
        ensure_output_summary_reports(idf)
        _apply_output_templates_from_cache(idf)

        if PURGE_OLD_INSULATION:
            purge_old_insulation(idf)
        for t in sorted({ float(t) for t in _thickness_map["Thickness_m"].dropna().tolist()
                          if float(t) >= THIN_LAYER_THRESHOLD_M }):
            ensure_insulation(idf, t)
        for cons in idf.idfobjects["CONSTRUCTION"]:
            for nm in get_layer_names_ordered(cons):
                if "insulation" in str(nm).lower():
                    ensure_insulation_alias_exact(idf, nm)
        for cons in list(idf.idfobjects["CONSTRUCTION"]):
            cz, bt, yr, env = parse_from_name(str(cons.Name))
            if bt and not yr:
                yr = file_year
            if not all([bt, yr, env]):
                continue
            t_m = pick_thickness(_thickness_map, cz, bt, env, int(yr))
            if t_m is None:
                continue
            replace_or_delete_insulation(idf, cons, t_m, replace_all=True)
        update_windows_from_map(idf, _window_map)

        final_path = compute_out_path(p, idf_root, final_root)
        final_path.parent.mkdir(parents=True, exist_ok=True)
        idf.saveas(str(final_path))
        return True, f"[OK-FINAL] {final_path}"
    except Exception as e:
        return False, f"[ERR-FINAL] {fp}: {e}"

def process_all_to_final(idf_root: Path, final_root: Path):
    files = _all_idf_files(Path(idf_root))
    total = len(files); ok = fail = 0
    if MAX_WORKERS > 1:
        with ProcessPoolExecutor(max_workers=MAX_WORKERS,
                                 initializer=_init_common,
                                 initargs=(IDD_PATH, SCHEDULE_XLS, CLIMATE_XLS, STATIC_XLS, INSUL_XLS, TEMPLATE_OUTPUTS_IDF)) as ex:
            futures = [ex.submit(_final_worker, str(p), str(idf_root), str(final_root), SHOW_LOG) for p in files]
            for fut in as_completed(futures):
                success, msg = fut.result()
                print(msg)
                if success: ok += 1
                else: fail += 1
    else:
        _init_common(IDD_PATH, SCHEDULE_XLS, CLIMATE_XLS, STATIC_XLS, INSUL_XLS, TEMPLATE_OUTPUTS_IDF)
        for p in files:
            success, msg = _final_worker(str(p), str(idf_root), str(final_root), SHOW_LOG)
            print(msg)
            if success: ok += 1
            else: fail += 1
    print(f"\n[FINAL] completed: total {total}, successful {ok}, failed {fail}")
    print(f"Output directory: {final_root}")

# ================== Main workflow ==================
def main():
    if not Path(IDD_PATH).exists():
        raise FileNotFoundError(f"IDD file not found: {IDD_PATH}")
    IDF.setiddname(IDD_PATH)

    if KEEP_INTERMEDIATE_OUTPUTS:
        print(f"\n===== Stage 1: Write schedules (workers={MAX_WORKERS}) =====")
        STAGE1_DIR.mkdir(parents=True, exist_ok=True)
        stage1_write_schedules(Path(IDF_ROOT), STAGE1_DIR, SCHEDULE_XLS, CLIMATE_XLS if CLIMATE_XLS else None)

        print(f"\n===== Stage 2: Loads + Infil + HVAC + Outputs (workers={MAX_WORKERS}) =====")
        STAGE2_DIR.mkdir(parents=True, exist_ok=True)
        stage2_write_loads_hvac_outputs(STAGE1_DIR, STAGE2_DIR)

        print(f"\n===== Stage 3: Insulation and window materials (workers={MAX_WORKERS}) =====")
        STAGE3_DIR.mkdir(parents=True, exist_ok=True)
        stage3_refresh_insulation_and_windows(STAGE2_DIR, STAGE3_DIR)

        print("\n=== All stages completed with intermediate outputs retained. ===")
    else:
        print(f"\n===== Integrated final-only processing (workers={MAX_WORKERS}) =====")
        FINAL_DIR.mkdir(parents=True, exist_ok=True)
        process_all_to_final(Path(IDF_ROOT), FINAL_DIR)
        print("\n=== All stages completed. Final IDF files only. ===")

if __name__ == "__main__":
    main()
