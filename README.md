# CRPC Region -- DRA Distress Status Visualization and AI Assistance

An interactive web map of economic distress status for the 252 census tracts across the
11 parishes of the Capital Region Planning Commission (CRPC) area in Louisiana, combined
with LODES job-density data, County Business Patterns business-establishment data, and an LLM-powered  AI assistant for natural-language querying across all three datasets.

This folder is a **self-contained, portable copy** of the project. It does not depend on
any files outside this directory -- copy the whole `CRPC Internship Project` folder to
another machine and it will run as-is (after installing the Python environment below).

The main file is **`visualize_crpc_region_with_filter22.py`** -- a single-file Flask app
that loads all datasets, builds the map, and serves the interactive dashboard.

---

## Demo Video
[![Capital Region Planning Commission CRPC Data Visualization and AI Assistance Internship Project](https://img.youtube.com/vi/EYTbgU9oy4c/maxresdefault.jpg)](https://www.youtube.com/watch?v=EYTbgU9oy4c)

*> Click the image above to watch the full project walkthrough on YouTube.*

## What it does

- **Interactive map** (Folium/Leaflet) of all 252 CRPC census tracts, color-coded by DRA
  distress status: Non-Distressed, Distressed by Unemployment Rate, Distressed by Per
  Capita Income Share, Distressed by Both, or No data.
- **Sidebar filters**: parish checkboxes, distress-status filter, and two range "knobs"
  (per-capita income, unemployment rate) with a 10-bin thermal green-to-red gradient
  scale -- all filters combine together and are honored by the map download.
- **Job Density Overlay** (per tract and per parish): LODES workplace-area-characteristics
  (WAC) 2023 job counts, broken out by all 20 NAICS sectors plus total jobs, shown as a
  color/size-coded overlay.
- **Business Establishments Overlay** (per parish): County Business Patterns 2023 data --
  total establishments plus micro (<10 employees) / small (10-99) / medium (100-499) /
  large (500+) employer counts, shown the same way as the job overlay.
- **AI Assistant chat box**: natural-language queries answered by a cloud LLM (Groq by
  default, or any OpenAI-compatible API). Understands income, unemployment, job-sector,
  and business-establishment questions, keeps conversation memory across turns, and comes
  with curated example prompts (including a "Show more suggested prompts" library grouped
  by topic).
- **Tract search bar** and an optional "Show Tract Numbers" label toggle.
- **600 DPI map download**: server-side render (matplotlib + geopandas + contextily) of
  the whole region with an "Active Filters" panel, honoring every filter/overlay that's
  currently enabled.
- **Full Screen** button that expands just the map pane (sidebar and chat stay hidden).

---

## Directory structure

```
CRPC Internship Project\
|-- visualize_crpc_region_with_filter22.py   Main file -- run this
|-- environment.yml                          Conda/mamba environment spec (recommended)
|-- requirements.txt                         Pip package list (fallback, see Setup notes)
|-- .env                                     LLM API key (GROQ_API_KEY) -- keep private
|-- README.md                                This file
|
|-- data\                                    Everything the app reads at runtime
|   |-- tiger_cache\                         Census Bureau cartographic tract boundaries
|   |   `-- cb_2025_22_tract_500k\           (shapefile set, auto-downloaded originally
|   |                                         from census.gov; cached here so no download
|   |                                         is needed on first run)
|   |-- distress_labeled\                    11 xlsx files, one per parish, with DRA
|   |                                         distress calculations and Yes/No labels
|   |-- ffiec\
|   |   `-- FFIEC_gov_CensusTractList2026_CRPC_Area_Only3.xlsx
|   |                                         FFIEC census tract list / GEOID lookup
|   |-- lodes\
|   |   `-- la_wac_S000_JT00_2023.csv        LODES WAC 2023 job counts (Louisiana,
|   |                                         block-level, aggregated to tract in-script)
|   `-- business\
|       `-- cbp23co_crpc_sba.csv             County Business Patterns 2023, pre-filtered
|                                             to the 11 CRPC parishes, with employer
|                                             size-class columns added
|
`-- source_code\                             Upstream ETL scripts that PRODUCED the files
    |                                        in data\, kept for provenance/reference.
    |                                        NOTE: these still use their original,
    |                                        machine-specific absolute paths (H:\stats
    |                                        america\...) -- they document how the data
    |                                        was built and are not wired to run from this
    |                                        folder as-is.
    |-- ffiec_processing\                    Scripts that built the FFIEC tract list
    |   |-- load_excel_columns.py
    |   |-- create_tract_decimal.py
    |   |-- add_geography_column.py
    |   |-- create_parish_dictionary.py
    |   |-- FFIEC_gov_CensusTractList2026_CRPC_Area_Only2.xlsx   (raw input)
    |   `-- FFIEC_govt_Parish_ALL_Tracts_Dictionary.json          (reference output)
    |
    |-- distress_labeling\                   Scripts that computed distress labels
    |   |-- label_distress_east_baton_rouge.py
    |   |-- label_distress_other_parishes.py
    |   |-- print_columns.py
    |   `-- raw_csv\                         11 unlabeled *_distress_download.csv inputs
    |
    `-- business_data_analysis\              Scripts that filtered/derived the CBP data
        |-- filter_louisiana.py              cbp23co.txt (national, not included -- 107MB)
        |                                     -> cbp23co_louisiana.csv
        |-- filter_crpc.py                   cbp23co_louisiana.csv -> cbp23co_crpc.csv
        |-- add_sba_columns.py               cbp23co_crpc.csv -> cbp23co_crpc_sba.csv
        |-- cbp23co_louisiana.csv            intermediate output
        |-- cbp23co_crpc.csv                 intermediate output
        `-- cbp23co_crpc_columns.txt         column reference
```

---

## Setup

### 1. Install the Python environment

The app needs a geospatial stack (GDAL/GEOS/PROJ via geopandas, rasterio, contextily)
that's much easier to get right with conda/mamba than plain pip.

**Recommended -- conda/mamba (matches the original `crpc` environment exactly):**

```powershell
mamba env create -f environment.yml
mamba activate crpc
```

If you don't have mamba, the same file works with plain conda:

```powershell
conda env create -f environment.yml
conda activate crpc
```

**Fallback -- pip:** `requirements.txt` is provided for reference, but plain
`pip install -r requirements.txt` on Windows often fails to build the GDAL/GEOS/PROJ
-dependent packages (rasterio, geopandas, contextily) without those system libraries
already present. Prefer the conda/mamba route above unless you already have a working
GDAL toolchain.

### 2. Configure the AI assistant (optional)

A `.env` file is already included with a working `GROQ_API_KEY` copied from the original
setup -- the AI chat box will work out of the box. If you need to use your own key instead:

1. Get a free key at https://console.groq.com
2. Edit `.env` and replace the `GROQ_API_KEY` value

To use a different OpenAI-compatible provider, set `LLM_API_KEY` / `LLM_BASE_URL` /
`LLM_MODEL` in the same `.env` file instead. Without a valid key, everything except the
AI chat box (map, filters, overlays, download) still works.

**Keep `.env` private** -- it contains a live API key. Don't commit it to a public repo
or share this folder publicly without removing/rotating the key first.

### 3. Run it

```powershell
python visualize_crpc_region_with_filter22.py
```

(Use the `crpc` environment's own `python.exe` if it's not activated on your PATH, e.g.
`C:\ProgramData\miniforge3\envs\crpc\python.exe visualize_crpc_region_with_filter22.py`.)

The script loads the data (~5-10 seconds), then starts a local web server. Once you see:

```
Total tracts loaded: 252
Running on http://127.0.0.1:5000
```

open **http://127.0.0.1:5000** in a browser.

It's a long-running Flask server, not a one-shot script -- it keeps running until you
stop it with Ctrl+C. Only run one instance at a time (it always binds port 5000).

---

## Notes / caveats

- **Internet is still required** even though the data is local: the map's basemap tiles
  (CartoDB Positron) load from a tile server, and the AI chat box calls the Groq API over
  the network. Everything else (filters, overlays, tract data) works fully offline.
- **Windows DLL handling**: the script auto-registers the conda environment's
  `Library\bin` directory (GDAL/PROJ/etc. native DLLs) at startup, so it works whether or
  not the `crpc` environment is "activated" -- launching it via the env's `python.exe`
  directly is sufficient.
- **Portability**: all data paths inside `visualize_crpc_region_with_filter22.py` are
  resolved relative to the script's own location (`Path(__file__).resolve().parent`), so
  moving this whole folder anywhere (a different drive letter, a different computer) keeps
  it working without editing any code.
- The `source_code\` scripts are for provenance only, as noted above -- if you want to
  regenerate `data\` from scratch, you'd need to restore their original absolute paths (or
  adapt them) and their raw source inputs (e.g. the 107MB national `cbp23co.txt`, which
  was intentionally left out of this folder).
