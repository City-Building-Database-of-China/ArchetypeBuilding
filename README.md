# ArchetypeBuilding

Prototype building energy modeling workflow and replication materials for **Shanghai**, China.

**Repository:** [City-Building-Database-of-China/ArchetypeBuilding](https://github.com/City-Building-Database-of-China/ArchetypeBuilding)

## Purpose

The data and code in this repository support the **peer review** of the manuscript (submitted; journal **TBD**):

**The hidden energy penalty of static building codes**

They are provided so reviewers and editors can verify methods, reproduce key workflow steps, and inspect Shanghai-level inputs and outputs described in the paper.

> **Scope.** This release covers **Shanghai only**. It includes a sample GIS footprint layer, simulation-ready IDF files, weather (EPW) files, building-parameter workbooks, and one batch simulation script. IDF post-processing and multi-city matching utilities are not required for the bundled ready-to-run models.

## Workflow overview

The bundled IDFs are already **simulation-ready**. Reviewers only need to run the batch script:

```
ready IDF  +  EPW  →  1_idf_epw_batch_runner.py  →  EnergyPlus results
```

| Stage | Location | Role |
|-------|----------|------|
| **Ready IDF** | `ready_idf/shang4hai3shi4/` | 113 archetype EnergyPlus models |
| **Weather (EPW)** | `input/EPW/Shang4hai3shi4/` | Baseline and RCP scenario weather files |
| **Batch simulation** | `1_idf_epw_batch_runner.py` | Match EPW, run EnergyPlus, write reports |
| **Results** | `result/shang4hai3shi4/` | One output folder per simulated IDF *(created locally, not committed)* |

Supporting data for GIS inspection and city-scale mapping described in the manuscript:

| Data | Location | Role |
|------|----------|------|
| **Prototype GIS** | `input/GIS/Prototype/310000_Shang4hai3shi4.*` | Sample footprint shapefile |
| **CityBuilding GIS** | `input/GIS/CityBuilding/310000_shang4hai3shi4.zip` | Full-city footprints *(reference only)* |
| **Settings** | `input/Setting/` | `Schedule.xlsx`, `Static.xlsx` |

## Data availability and legal notice

Please note that we are **prohibited** from distributing or uploading the original, precise geospatial datasets to public repositories under the *Surveying and Mapping Law of the People's Republic of China* and relevant national data security regulations.

In the distributed materials, explicit longitude and latitude coordinates are stripped or adjusted. This repository provides a **sample spatial dataset** together with the core simulation scripts.

## Repository layout

```
ArchetypeBuilding/
├── 1_idf_epw_batch_runner.py      # batch EPW matching + EnergyPlus runs
├── backup_idf_batch_processor.py  # optional IDF post-processing (not needed here)
├── ready_idf/
│   └── shang4hai3shi4/            # 113 simulation-ready .idf files
├── input/
│   ├── EPW/
│   │   └── Shang4hai3shi4/        # 7 .epw files
│   ├── GIS/
│   │   ├── Prototype/
│   │   │   └── 310000_Shang4hai3shi4.{shp,shx,dbf,prj,cpg}
│   │   └── CityBuilding/
│   │       └── 310000_shang4hai3shi4.zip
│   └── Setting/
│       ├── Schedule.xlsx
│       └── Static.xlsx
└── result/                          # local run outputs (do not commit)
    ├── shang4hai3shi4/<idf_stem>/   # EnergyPlus outputs
    └── reports/                     # Excel matching / run reports
```

> **Path names are case-sensitive on Linux/GitHub.** IDF inputs use `ready_idf/shang4hai3shi4/`; weather files use `input/EPW/Shang4hai3shi4/`.

### Weather (EPW)

All weather files are under `input/EPW/Shang4hai3shi4/`:

| File | Description |
|------|-------------|
| `Shanghai_2020.epw` | Baseline weather (**default**) |
| `SHANGHAISHI_RCP2.6_2040.epw` | RCP2.6, 2040 |
| `SHANGHAISHI_RCP2.6_2060.epw` | RCP2.6, 2060 |
| `SHANGHAISHI_RCP4.5_2040.epw` | RCP4.5, 2040 |
| `SHANGHAISHI_RCP4.5_2060.epw` | RCP4.5, 2060 |
| `SHANGHAISHI_RCP8.5_2040.epw` | RCP8.5, 2040 |
| `SHANGHAISHI_RCP8.5_2060.epw` | RCP8.5, 2060 |

### IDF file naming

Archetype IDFs under `ready_idf/shang4hai3shi4/` follow:

```
shang4hai3shi4_<func_index>_<archetype>_<year>_S0.idf
```

Example: `shang4hai3shi4_4_8_1995_S0.idf`

| Token | Meaning |
|-------|---------|
| `func_index` | Building function type (0–7) |
| `archetype` | Archetype index within that function |
| `year` | Construction / code vintage |
| `S0` | Scenario tag |

## Quick start

### 1. Clone and install Python dependencies

```bash
git clone https://github.com/City-Building-Database-of-China/ArchetypeBuilding.git
cd ArchetypeBuilding
pip install eppy pandas openpyxl
```

Install **EnergyPlus 23.1** locally. EnergyPlus is **not** included in this repository.

### 2. Set your EnergyPlus path

Windows PowerShell:

```powershell
$env:ENERGYPLUS_DIR = "C:\EnergyPlusV23-1-0"
# or
$env:IDD_FILE = "C:\EnergyPlusV23-1-0\Energy+.idd"
```

Linux / macOS example:

```bash
export ENERGYPLUS_DIR=/usr/local/EnergyPlus-23-1-0
```

### 3. Preview EPW–IDF matching (recommended)

```powershell
$env:CHECK_ONLY = "true"
py 1_idf_epw_batch_runner.py
```

This creates `result/reports/epw_idf_match_plan_*.xlsx` **without** running EnergyPlus.

Default mapping:

| IDF folder | EPW file |
|------------|----------|
| `ready_idf/shang4hai3shi4/` | `input/EPW/Shang4hai3shi4/Shanghai_2020.epw` |

### 4. Run EnergyPlus

```powershell
$env:CHECK_ONLY = "false"
py 1_idf_epw_batch_runner.py
```

Outputs:

- Simulation files: `result/shang4hai3shi4/<idf_stem>/`
- Run summary: `result/reports/epw_idf_run_report_*.xlsx`

### Optional: RCP climate scenarios

```powershell
$env:EPW_STEM = "SHANGHAISHI_RCP4.5_2040"
py 1_idf_epw_batch_runner.py
```

## Configuration

All data paths in `1_idf_epw_batch_runner.py` are **relative to the repository root** (`PROJECT_ROOT =` directory containing the script). No drive letters or usernames are hard-coded.

| Variable | Default | Purpose |
|----------|---------|---------|
| `IDF_ROOT` | `ready_idf` | Root folder with city subfolders of `.idf` files |
| `EPW_DIR` | `input/EPW/Shang4hai3shi4` | Folder containing `.epw` files |
| `OUT_ROOT` | `result` | EnergyPlus output root |
| `REPORT_ROOT` | `result/reports` | Excel report output |
| `EPW_STEM` | `Shanghai_2020` | EPW filename stem for folder `shang4hai3shi4` |
| `ENERGYPLUS_DIR` | *(unset)* | Local EnergyPlus install directory |
| `IDD_FILE` | `{ENERGYPLUS_DIR}/Energy+.idd` | Path to `Energy+.idd` |
| `EP_VERSION` | `23-1-0` | EnergyPlus version passed to eppy |
| `NUM_CPUS` | `6` | Parallel worker count |
| `CHECK_ONLY` | `false` | If `true`, only export matching report |

**Machine-specific vs portable paths**

- **Portable (in repo):** `ready_idf/`, `input/EPW/Shang4hai3shi4/`, `result/`
- **Machine-specific (you set locally):** `ENERGYPLUS_DIR` or `IDD_FILE`

## Scripts

| File | Required? | Description |
|------|-----------|-------------|
| `1_idf_epw_batch_runner.py` | **Yes** | Match Shanghai EPW to ready IDFs and run EnergyPlus in batch |
| `backup_idf_batch_processor.py` | No | Three-stage IDF post-processor (schedules, loads, insulation). **Not needed** when using the bundled simulation-ready IDFs |

## Generated files (not committed)

The `result/` folder is created at runtime and should **not** be pushed to GitHub. It may contain:

- `result/shang4hai3shi4/<idf_stem>/` — EnergyPlus outputs (`.csv`, `.sql`, `.err`, etc.)
- `result/reports/epw_idf_match_plan_*.xlsx` — matching preview from `CHECK_ONLY=true`
- `result/reports/epw_idf_run_report_*.xlsx` — success/failure summary after a full run

## Post-process: scale-up to city-wide buildings

The bundled IDFs are **prototype** models (one file per archetype). **City-wide scale-up** is described in the manuscript; this repository provides the GIS and IDF data used in that step, not a separate scale-up script.

| Data | Role |
|------|------|
| **CityBuilding** (`input/GIS/CityBuilding/310000_shang4hai3shi4.zip`) | All individual building footprints in Shanghai |
| **Prototype** (`input/GIS/Prototype/310000_Shang4hai3shi4.shp`) | Prototype buildings; geometry and archetype index |
| **`BuildID`** | Join key between Prototype and CityBuilding |
| **`LandNum`** + **`Cluster`** | Assign each building to the correct prototype type |

## Status

- Manuscript submitted; target journal **TBD**
- Shanghai prototype GIS, 113 simulation-ready IDFs, 7 EPW files, and batch simulation script are included
- Additional cities may be released separately from the broader China Building Energy Model Database
