# ArchetypeBuilding

Shanghai prototype building energy modeling materials for peer review of (submitted; journal **TBD**):

**The hidden energy penalty of static building codes**

[City-Building-Database-of-China/ArchetypeBuilding](https://github.com/City-Building-Database-of-China/ArchetypeBuilding)

## Contents

This release covers **Shanghai only**: 113 simulation-ready IDFs, weather files, sample GIS, parameter workbooks, and `1_idf_epw_batch_runner.py`.

| Script | Role |
|--------|------|
| `1_idf_epw_batch_runner.py` | Match EPW to ready IDFs and run EnergyPlus |
| `backup_idf_batch_processor.py` | Fill in or repair **incomplete or deficient IDFs** (schedules, loads, HVAC, insulation, windows from `input/Setting/`). Not needed for the bundled ready IDFs. |

```
ready_idf/shang4hai3shi4/  +  input/EPW/Shang4hai3shi4/
         →  1_idf_epw_batch_runner.py  →  result/shang4hai3shi4/
```

| Path | Contents |
|------|----------|
| `ready_idf/shang4hai3shi4/` | 113 simulation-ready `.idf` files |
| `input/EPW/Shang4hai3shi4/` | 7 `.epw` weather files (see below) |
| `input/GIS/Prototype/` | Sample footprint shapefile (`310000_Shang4hai3shi4.*`) |
| `input/GIS/CityBuilding/` | Full-city GIS zip (`310000_shang4hai3shi4.zip`, reference only) |
| `input/Setting/` | `Schedule.xlsx`, `Static.xlsx` |
| `result/` | Local simulation outputs (do not commit) |

IDF naming: `shang4hai3shi4_<func_index>_<archetype>_<year>_S0.idf`  
Example: `shang4hai3shi4_4_8_1995_S0.idf`

> Path names are case-sensitive on Linux/GitHub: IDF folder `shang4hai3shi4`, EPW folder `Shang4hai3shi4`.

## Weather (EPW)

All files are in `input/EPW/Shang4hai3shi4/`:

- **Baseline:** `Shanghai_2020.epw` — present-day / default weather for batch runs
- **Climate scenarios:** RCP **2.6**, **4.5**, and **8.5**, each for **2040** and **2060** (six files):

| RCP | 2040 | 2060 |
|-----|------|------|
| 2.6 | `SHANGHAISHI_RCP2.6_2040.epw` | `SHANGHAISHI_RCP2.6_2060.epw` |
| 4.5 | `SHANGHAISHI_RCP4.5_2040.epw` | `SHANGHAISHI_RCP4.5_2060.epw` |
| 8.5 | `SHANGHAISHI_RCP8.5_2040.epw` | `SHANGHAISHI_RCP8.5_2060.epw` |

The batch script defaults to `Shanghai_2020.epw`. To use a scenario file, set `EPW_STEM` to the filename stem (e.g. `SHANGHAISHI_RCP4.5_2040`).

## Replication

Dependencies: `pip install eppy pandas openpyxl`, plus **EnergyPlus 23.1** installed locally.

All data paths in the script are **relative to the repository root**. Set `ENERGYPLUS_DIR` or `IDD_FILE` to your local EnergyPlus install, then run `1_idf_epw_batch_runner.py`. Outputs go to `result/shang4hai3shi4/`; default weather is `Shanghai_2020.epw`.

Key environment variables: `IDF_ROOT`, `EPW_DIR`, `EPW_STEM`, `OUT_ROOT`, `NUM_CPUS`.

## Legal notice

We are **prohibited** from distributing precise geospatial datasets under the *Surveying and Mapping Law of the People's Republic of China* and related regulations. Distributed GIS materials use stripped or adjusted coordinates.

## City-wide scale-up (manuscript)

Prototype IDFs map to individual buildings via the GIS layers below (scale-up logic is described in the paper, not implemented as a separate script here):

| Data | Role |
|------|------|
| `input/GIS/CityBuilding/310000_shang4hai3shi4.zip` | All building footprints |
| `input/GIS/Prototype/310000_Shang4hai3shi4.shp` | Prototype footprints |
| `BuildID` | Join key between layers |
| `LandNum` + `Cluster` | Prototype type assignment |
