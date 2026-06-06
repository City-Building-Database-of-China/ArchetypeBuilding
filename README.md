# ArchetypeBuilding

Prototype building energy modeling workflow and replication materials for **Shanghai**, China.

## Purpose

The data and code in this repository support the **peer review** of the manuscript (submitted; journal **TBD**):

**The hidden energy penalty of static building codes**

They are provided so reviewers and editors can verify methods, reproduce key workflow steps, and inspect Shanghai-level inputs and outputs described in the paper.

> **Scope.** This release covers **Shanghai only**. It includes prototype building GIS inputs, weather (EPW) files, and the core workflow scripts. Other cities from the broader China Building Energy Model Database are not included here.

## Workflow overview

Building models are produced and simulated through the following pipeline:

```
Shanghai GIS inputs  →  GIS2IDF  →  ready IDF  →  EnergyPlus simulation  →  results
```

| Stage | Description |
|-------|-------------|
| **GIS inputs** | Prototype building footprints and attributes, local weather (EPW), and building-parameter settings. **Replication uses the Prototype sample layer** (see below). |
| **GIS2IDF** | GIS data are turned into building EnergyPlus models (geometry and associated model setup). |
| **Ready IDF** | Per-prototype IDF files prepared for simulation. |
| **EnergyPlus simulation** | Batch runs using Shanghai weather file(s). |
| **Results** | Simulation outputs for analysis and comparison with published figures. |

## Data availability and legal notice

Please note that we are **prohibited** from distributing or uploading the original, precise geospatial datasets to public repositories under the *Surveying and Mapping Law of the People's Republic of China* and relevant national data security regulations.

In the distributed materials, explicit longitude and latitude coordinates are stripped or adjusted. This repository provides a **sample spatial dataset** together with the core simulation scripts.

Reviewers and future readers can run the code on this sample data to verify the logic and validity of the computational workflow without access to restricted full-resolution survey data.

## Data in this repository

| Category | Shanghai | Notes |
|----------|----------|-------|
| **GIS / prototype building data** | Included | Sample Prototype layer + optional full-city archive |
| **Weather (EPW)** | Included | Baseline and climate-scenario files under `input/EPW/SHANGHAI/` |
| **Workflow scripts** | Included | `1_GIS2IDF.py`, `2_BatchSimulation.py` |

### GIS inputs: Prototype vs CityBuilding

Two GIS products are provided under `input/GIS/`:

| Folder | Role | Use with bundled scripts? |
|--------|------|---------------------------|
| **`Prototype/`** | Public-release **sample** building footprints (explicit lon/lat stripped or adjusted). Matches the default paths in `1_GIS2IDF.py`. | **Yes** — use this for replication and peer review. |
| **`CityBuilding/`** | **Original full-city building GIS** from the study (all buildings in Shanghai). Supplied as a **ZIP archive** for reference only. | **No** — not used by the default workflow scripts. |

**CityBuilding ZIP archive** (under `input/GIS/CityBuilding/`):

| City | Archive |
|------|---------|
| Shanghai | `310000_shang4hai3shi4.zip` |

The archive is **not password-protected**; you can extract it directly with any standard ZIP utility.

It preserves the complete building layer used in the paper-scale analyses and is included for transparency and local inspection. **To run `1_GIS2IDF.py` and verify the computational workflow, point the script at shapefiles under `input/GIS/Prototype/`** (e.g. `310000SHANGHAISHI.shp`), not at the CityBuilding layers.

### Weather (EPW)

Shanghai weather files are under `input/EPW/SHANGHAI/`:

| File | Description |
|------|-------------|
| `Shanghai_2020.epw` | Baseline weather used for default replication runs |
| `SHANGHAISHI_RCP2.6_2040.epw` | RCP2.6 scenario, 2040 |
| `SHANGHAISHI_RCP2.6_2060.epw` | RCP2.6 scenario, 2060 |
| `SHANGHAISHI_RCP4.5_2040.epw` | RCP4.5 scenario, 2040 |
| `SHANGHAISHI_RCP4.5_2060.epw` | RCP4.5 scenario, 2060 |
| `SHANGHAISHI_RCP8.5_2040.epw` | RCP8.5 scenario, 2040 |
| `SHANGHAISHI_RCP8.5_2060.epw` | RCP8.5 scenario, 2060 |

### Repository layout

| Path | Contents |
|------|----------|
| **`1_GIS2IDF.py`** | Generate ready IDF files from **Prototype** GIS and input settings (default: Shanghai). |
| **`2_BatchSimulation.py`** | Run EnergyPlus in batch on ready IDF files (default: `demo/ready_idf/` → `demo/result/`). |
| **`input/GIS/Prototype/`** | Sample Shanghai prototype building footprints. **Input for `1_GIS2IDF.py`.** |
| **`input/GIS/CityBuilding/`** | Full-city Shanghai GIS packaged as `310000_shang4hai3shi4.zip`. |
| **`input/EPW/SHANGHAI/`** | Shanghai weather files (baseline and RCP scenarios). |
| **`input/Setting/`** | Non-geometry building parameters (`non_geomtry_data_all.xlsx`, `age_de/`). |
| **`demo/`** | Self-contained **demo package** (sample IDFs + pre-run results); see below. |
| **`ready_idf/`** | Full-city generated IDF files for Shanghai (`ready_idf/310000SHANGHAISHI/`). |
| **`result/`** | Simulation output folders for full-city or other runs (not used for the bundled demo). |

### Demo package (`demo/`)

A small **end-to-end example** lives under a single top-level folder so reviewers can find inputs and outputs in one place:

```
demo/
├── ready_idf/          # pre-generated IDF files (inputs to simulation)
│   ├── 310000SHANGHAISHI_1.idf
│   ├── 310000SHANGHAISHI_2.idf
│   └── 310000SHANGHAISHI_3.idf
└── result/             # pre-run EnergyPlus outputs (one subfolder per building)
    ├── 310000SHANGHAISHI_1/
    ├── 310000SHANGHAISHI_2/
    └── 310000SHANGHAISHI_3/
```

| Location | Contents |
|----------|----------|
| **`demo/ready_idf/`** | Three **pre-generated IDF** files for Shanghai prototype buildings, produced by the GIS2IDF workflow from `input/GIS/Prototype/`. |
| **`demo/result/`** | **Pre-run EnergyPlus results** for the same three buildings. Each subfolder (e.g. `310000SHANGHAISHI_1/`) holds standard outputs such as tabular summaries (`.csv`, `Table.htm`), SQLite (`.sql`), and logs (`.err`, `.eio`). |

**How to use the demo**

- **`2_BatchSimulation.py`** defaults to `demo/ready_idf/` → `demo/result/` (weather: `input/EPW/SHANGHAI/Shanghai_2020.epw`). Runtime stays short while exercising the batch script.
- Open **`demo/result/`** immediately to inspect a successful run, or re-run the batch and compare with the bundled outputs.
- Full-city ready IDFs are under `ready_idf/310000SHANGHAISHI/` when you need more than the three-building demo.

## Paths and local configuration

**Repository data paths (portable).** Neither script hard-codes your machine username, drive letter, or clone location. Both scripts set `base_dir` to the folder containing the `.py` file (`os.path.dirname(os.path.abspath(__file__))`), then build paths with `os.path.join(base_dir, "input", ...)`, `demo/`, `ready_idf/`, and `result/`. As long as you run the scripts from a normal clone of this repository, these folders resolve correctly on any OS.

**EnergyPlus install paths (machine-specific).** EnergyPlus itself is **not** shipped in this repo. `1_GIS2IDF.py` and `2_BatchSimulation.py` still use the **default Windows install layout** for version 23.1, for example `C:\EnergyPlusV23-1-0\Energy+.idd` and related files under `C:\EnergyPlusV{version}\`. If your installation is elsewhere—or you are on Linux or macOS—you must edit those EnergyPlus paths in the scripts (search for `EnergyPlusV` / `iddfile`) before running.

## Replication notes

1. Confirm inputs under `input/` (weather, settings). For GIS, use **`input/GIS/Prototype/`** only — do not point `1_GIS2IDF.py` at the CityBuilding ZIP or full-city layers. The default example uses `input/GIS/Prototype/310000SHANGHAISHI.shp`.
2. Install **EnergyPlus 23.1** and point the EnergyPlus paths in both scripts to your local `Energy+.idd` and install folder (see **Paths and local configuration** above).
3. Run **`1_GIS2IDF.py`** to produce IDFs under `ready_idf/` (default output folder: `ready_idf/310000SHANGHAISHI/`).
4. Run **`2_BatchSimulation.py`** to simulate IDFs under `demo/ready_idf/` and write outputs to `demo/result/`. Pre-computed results are already under `demo/result/` for comparison.
5. Compare outputs with the paper.

## Status

- Manuscript submitted; target journal **TBD**.
- This repository provides Shanghai prototype GIS inputs, EPW weather files, workflow scripts, full-city ready IDFs, and a self-contained `demo/` package.
- Additional cities and post-processing utilities may be released separately as the broader database is finalized.

## Post-process: scale-up to city-wide buildings

`1_GIS2IDF.py` builds **prototype** IDFs from `input/GIS/Prototype/` (one file per **`bh`**, e.g. `310000SHANGHAISHI_5.idf`). **City-wide scale-up** is described in the manuscript; this repository provides the GIS and IDF data used in that step, not a separate scale-up script.

| Data | Role |
|------|------|
| **CityBuilding** (unzip `input/GIS/CityBuilding/310000_shang4hai3shi4.zip`, read `.shp`) | All individual building footprints in Shanghai. |
| **Prototype** (`input/GIS/Prototype/310000SHANGHAISHI.shp`) | Prototype buildings; source of geometry and IDF index **`bh`**. |
| **`BuildID`** | Join key between Prototype and CityBuilding. |
| **`bh`** | Prototype index; matches IDF filename suffix and `1_GIS2IDF.py` output. |
| **`LandNum`** + **`Cluster`** | Together assign each building to the correct **`proptype`** (prototype type) for mapping prototype models to individual footprints. |
