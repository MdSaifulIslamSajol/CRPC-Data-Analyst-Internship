"""
CRPC Region Distress Map -- Filter7
====================================

Builds on Filter6 (parish/distress/threshold filters, hover highlight, the
server-rendered 600 DPI download, and the two per-capita income / unemployment
range "knobs" with enable checkboxes, tick marks, and a scale). Filter7 adds a
thermal green -> yellow -> red gradient to both slider tracks (green at the
left, darkest red at the right), using heat-map colors distinct from the map's
distress palette.

Each slider is split into 10 equal-width bins between the dataset's min and max
value; ticking its checkbox activates it as a filter and selecting a bin shows
only the tracts whose value falls in that range. The income knob runs high ->
low (highest income on the left). The knobs combine with every other filter and
are honored by the high-resolution download too.

Adds an AI-powered natural-language query box below the map. Type a
request like "show me the lowest income tracts in East Baton Rouge"
and it is sent to a cloud LLM (Groq, OpenAI-compatible API), which
returns a structured filter. That filter is applied to the tract
data, the matching tracts are highlighted on the map, and a table of
results is rendered in a scrollable panel below the map.

Filter18 integrates the LODES WAC job-density data (already loaded for the
map overlay) into the AI query box: the LLM can now sort/filter tracts by
total jobs or by any of the 20 NAICS sector job counts (e.g. "highest
manufacturing jobs in Livingston Parish"), and the results table shows the
matching job count column.

Filter20 adds a "Full Screen" button next to the two download buttons
(bottom-left). It puts only #map-pane (the map plus its overlaid legend,
parish-filter panel, search bar, and download panel) into the browser's
Fullscreen API -- the left sidebar filters and the AI chat/results panes
are not part of #map-pane, so they stay hidden while fullscreen is active.

Filter21 adds a "Business Establishments Per Parish" overlay panel, built
the same way as "Job Density Overlay Per Parish": county-level business
counts (total establishments, micro/small/medium/large enterprise counts)
from the County Business Patterns extract are loaded, aggregated to one
value per parish, and shown as emoji markers sized/colored by value at
each parish's centroid. It has its own enable checkbox + 5-metric grid,
combines with the parish filter exactly like the job overlay does, and is
included in the 600 DPI / screen downloads.

Setup
-----
    pip install flask openai python-dotenv folium geopandas pandas requests openpyxl

    Get a free API key at https://console.groq.com, then create a file
    named ".env" next to this script (see ".env.example") containing:
        GROQ_API_KEY=gsk_...

    To use a different OpenAI-compatible provider instead, set
    LLM_API_KEY / LLM_BASE_URL / LLM_MODEL in the same .env file.

Run
---
    python visualize_crpc_region_with_filter5.py
    Then open http://127.0.0.1:5000 in a browser.
"""

import colorsys
import io
import json
import os
import re
import sys
import zipfile
from pathlib import Path

# On Windows, matplotlib / contextily / rasterio C-extensions depend on native
# DLLs that live in the conda env's "Library\bin" (freetype, libpng, GDAL, ...).
# When this script is launched via the env's python.exe *without* activating the
# env, that directory is not on the DLL search path and those extensions crash
# with a delay-load failure (0xc06d007f). Register the env's DLL dirs explicitly
# so the script works no matter how it is launched.
if os.name == "nt":
    for _sub in (r"Library\bin", r"Library\lib", r"Library\mingw-w64\bin", "DLLs"):
        _dll_dir = os.path.join(sys.prefix, _sub)
        if os.path.isdir(_dll_dir):
            os.add_dll_directory(_dll_dir)
            os.environ["PATH"] = _dll_dir + os.pathsep + os.environ.get("PATH", "")
    os.environ.setdefault("GDAL_DATA", os.path.join(sys.prefix, "Library", "share", "gdal"))
    os.environ.setdefault("PROJ_LIB", os.path.join(sys.prefix, "Library", "share", "proj"))

import contextily as ctx
import folium
import geopandas as gpd
import matplotlib.patches as mpatches
import matplotlib.patheffects as patheffects
import pandas as pd
import requests
from dotenv import load_dotenv
from flask import Flask, jsonify, render_template_string, request, send_file
from matplotlib.figure import Figure
from openai import OpenAI
from pyproj import Transformer
from shapely.geometry import mapping

load_dotenv()

# This copy is self-contained: all data lives under a data/ folder next to
# this script, so it runs on any machine/drive without depending on the
# original H:\stats america\... source locations.
BASE_DIR = Path(__file__).resolve().parent
DATA_DIR = BASE_DIR / "data"

OUT_DIR = BASE_DIR
CACHE_DIR = DATA_DIR / "tiger_cache"
CACHE_DIR.mkdir(parents=True, exist_ok=True)

LABEL_DIR = DATA_DIR / "distress_labeled"
FFIEC_XLSX = DATA_DIR / "ffiec" / "FFIEC_gov_CensusTractList2026_CRPC_Area_Only3.xlsx"

STATE_FIPS = "22"  # Louisiana

PARISHES = [
    {"fips": "005", "name": "ASCENSION PARISH", "slug": "ascension", "label": "Ascension Parish", "label_file": "Ascension_LA_Tract_distress_download_Labeled.xlsx"},
    {"fips": "033", "name": "EAST BATON ROUGE PARISH", "slug": "east_baton_rouge", "label": "East Baton Rouge Parish", "label_file": "East_Baton_Rouge_LA_Tract_distress_download_Labeled.xlsx"},
    {"fips": "037", "name": "EAST FELICIANA PARISH", "slug": "east_feliciana", "label": "East Feliciana Parish", "label_file": "East_Feliciana_LA_Tract_distress_download_Labeled.xlsx"},
    {"fips": "047", "name": "IBERVILLE PARISH", "slug": "iberville", "label": "Iberville Parish", "label_file": "Iberville_LA_Tract_distress_download_Labeled.xlsx"},
    {"fips": "063", "name": "LIVINGSTON PARISH", "slug": "livingston", "label": "Livingston Parish", "label_file": "Livingston_LA_Tract_distress_download_Labeled.xlsx"},
    {"fips": "077", "name": "POINTE COUPEE PARISH", "slug": "pointe_coupee", "label": "Pointe Coupee Parish", "label_file": "Pointe_Coupee_LA_Tract_distress_download_Labeled.xlsx"},
    {"fips": "091", "name": "ST. HELENA PARISH", "slug": "st_helena", "label": "St. Helena Parish", "label_file": "St._Helena_LA_Tract_distress_download_Labeled.xlsx"},
    {"fips": "105", "name": "TANGIPAHOA PARISH", "slug": "tangipahoa", "label": "Tangipahoa Parish", "label_file": "Tangipahoa_LA_Tract_distress_download_Labeled.xlsx"},
    {"fips": "117", "name": "WASHINGTON PARISH", "slug": "washington", "label": "Washington Parish", "label_file": "Washington_LA_Tract_distress_download_Labeled.xlsx"},
    {"fips": "121", "name": "WEST BATON ROUGE PARISH", "slug": "west_baton_rouge", "label": "West Baton Rouge Parish", "label_file": "West_Baton_Rouge_LA_Tract_distress_download_Labeled.xlsx"},
    {"fips": "125", "name": "WEST FELICIANA PARISH", "slug": "west_feliciana", "label": "West Feliciana Parish", "label_file": "West_Feliciana_LA_distress_download_Labeled.xlsx"},
]

CB_URL = f"https://www2.census.gov/geo/tiger/GENZ2025/shp/cb_2025_{STATE_FIPS}_tract_500k.zip"
CB_ZIP = CACHE_DIR / f"cb_2025_{STATE_FIPS}_tract_500k.zip"
CB_DIR = CACHE_DIR / f"cb_2025_{STATE_FIPS}_tract_500k"

COLOR_MAP = {
    "Non-Distressed": "#A9DFBF",
    "Distressed by Unemployment Rate": "#FFC7CE",
    "Distressed by Per Capita Income Share": "#FFD8A8",
    "Distressed by Both Unemployment Rate and Per Capita Income Share": "#FF4444",
    "No data": "#D9D9D9",
}

# Order shown in the bottom "Filter by Distress Status" bar (and color swatches).
DISTRESS_STATUSES = [
    {"slug": "both", "label": "Distressed by Both Unemployment Rate and Per Capita Income Share", "short": "Both", "color": "#FF4444"},
    {"slug": "unemployment", "label": "Distressed by Unemployment Rate", "short": "Unemployment Rate", "color": "#FFC7CE"},
    {"slug": "per_capita_income", "label": "Distressed by Per Capita Income Share", "short": "Per Capita Income Share", "color": "#FFD8A8"},
    {"slug": "non_distressed", "label": "Non-Distressed", "short": "Non-Distressed", "color": "#A9DFBF"},
    {"slug": "no_data", "label": "No data", "short": "No Data", "color": "#D9D9D9"},
]

# Extra "Filter by Threshold" control: lets the user isolate tracts that
# cross either of the two underlying DRA threshold columns directly,
# regardless of the overall Distress Status classification.
INCOME_THRESHOLD_CUTOFF = 60.0  # per capita income < 60% of U.S. average
UNEMPLOYMENT_THRESHOLD_CUTOFF = 2.0  # unemployment rate >= 2 pct pts above U.S. average

# ---------------------------------------------------------------------------
# 1. Get census tract boundaries (Census Bureau cartographic boundary file)
# ---------------------------------------------------------------------------
if not CB_ZIP.exists():
    resp = requests.get(CB_URL, timeout=120)
    resp.raise_for_status()
    CB_ZIP.write_bytes(resp.content)

if not CB_DIR.exists():
    with zipfile.ZipFile(CB_ZIP) as zf:
        zf.extractall(CB_DIR)

shp_path = next(CB_DIR.glob("*.shp"))
tracts = gpd.read_file(shp_path)
tracts["GEOID"] = tracts["GEOID"].astype(str)

# ---------------------------------------------------------------------------
# 2. Build a Geography -> GEOID lookup from the FFIEC census tract list
# ---------------------------------------------------------------------------
ffiec = pd.read_excel(FFIEC_XLSX)
ffiec["County name"] = ffiec["County name"].str.strip()


def classify(row):
    if row["Distressed by Both Unemployment Rate and Per Capita Income Share"] == "Yes":
        return "Distressed by Both Unemployment Rate and Per Capita Income Share"
    if row["Distressed by Unemployment Rate"] == "Yes":
        return "Distressed by Unemployment Rate"
    if row["Distressed by Per Capita Income Share"] == "Yes":
        return "Distressed by Per Capita Income Share"
    return "Non-Distressed"


fields = [
    "GEOID",
    "Geography",
    "2024 Unemployment Rate (5-Year ACS)",
    "Threshold Calculation",
    "2024 Per Capita Money Income (5-Year ACS)",
    "Threshold Calculation2",
    "Distressed by Unemployment Rate",
    "Distressed by Per Capita Income Share",
    "Distressed by Both Unemployment Rate and Per Capita Income Share",
    "Non-Distressed",
    "Distress Status",
    "Parish",
]

# ---------------------------------------------------------------------------
# 3. For each parish: merge tract geometry with its labeled distress data
# ---------------------------------------------------------------------------
parish_frames = []
for parish in PARISHES:
    parish_tracts = tracts[tracts["COUNTYFP"] == parish["fips"]].copy()

    lookup = ffiec[ffiec["County name"] == parish["name"]].copy()
    lookup["GEOID"] = lookup["FIPS code"].astype(str).str.zfill(11)
    geoid_by_geography = lookup.set_index("FFIEC_Geography")["GEOID"]

    labeled = pd.read_excel(LABEL_DIR / parish["label_file"])
    labeled = labeled[~labeled["Geography"].isin(["Region", "U.S."])].copy()
    labeled["GEOID"] = labeled["Geography"].map(geoid_by_geography)

    missing = labeled[labeled["GEOID"].isna()]
    if len(missing):
        print(f"{parish['name']}: tracts with no matching GEOID in the FFIEC list:")
        print(missing["Geography"].tolist())
    labeled = labeled.dropna(subset=["GEOID"])

    labeled["Distress Status"] = labeled.apply(classify, axis=1)
    labeled["Parish"] = parish["label"]

    merged = parish_tracts.merge(labeled[fields], on="GEOID", how="left")
    merged["Distress Status"] = merged["Distress Status"].fillna("No data")
    merged["Geography"] = merged["Geography"].fillna(merged["NAMELSAD"])
    merged["Parish"] = merged["Parish"].fillna(parish["label"])
    merged["fill_color"] = merged["Distress Status"].map(COLOR_MAP)
    merged["meets_income_threshold"] = merged["Threshold Calculation2"] < INCOME_THRESHOLD_CUTOFF
    merged["meets_unemployment_threshold"] = merged["Threshold Calculation"] >= UNEMPLOYMENT_THRESHOLD_CUTOFF
    # Numeric copies of the two metrics for the range knobs. The source income
    # column is currency-formatted text (e.g. "$32,034 "), so strip everything
    # but digits/decimal point before converting.
    merged["income_value"] = pd.to_numeric(
        merged["2024 Per Capita Money Income (5-Year ACS)"].astype(str).str.replace(r"[^0-9.]", "", regex=True),
        errors="coerce")
    merged["unemployment_value"] = pd.to_numeric(
        merged["2024 Unemployment Rate (5-Year ACS)"].astype(str).str.replace(r"[^0-9.\-]", "", regex=True),
        errors="coerce")
    merged = merged.to_crs(epsg=4326)
    parish_frames.append(merged)

    print(f"{parish['name']}: {len(merged)} tracts plotted, "
          f"{merged['Distress Status'].ne('No data').sum()} matched to distress labels")

# ---------------------------------------------------------------------------
# 4. Build the interactive map, fitted to the whole CRPC region
# ---------------------------------------------------------------------------
all_tracts = gpd.GeoDataFrame(pd.concat(parish_frames, ignore_index=True), crs=parish_frames[0].crs)

# ---------------------------------------------------------------------------
# LODES WAC: aggregate block-level job counts to census-tract level
# ---------------------------------------------------------------------------
LODES_CSV = DATA_DIR / "lodes" / "la_wac_S000_JT00_2023.csv"
JOB_COLS = ["C000"] + [f"CNS{i:02d}" for i in range(1, 21)]
_lodes_raw = pd.read_csv(LODES_CSV, usecols=["w_geocode"] + JOB_COLS, dtype={"w_geocode": str})
_lodes_raw["GEOID"] = _lodes_raw["w_geocode"].str[:11]
_tract_jobs = _lodes_raw.groupby("GEOID")[JOB_COLS].sum().reset_index()
for _i in range(len(parish_frames)):
    parish_frames[_i] = parish_frames[_i].merge(_tract_jobs, on="GEOID", how="left")
    for _col in JOB_COLS:
        parish_frames[_i][_col] = parish_frames[_i][_col].fillna(0).astype(int)
all_tracts = gpd.GeoDataFrame(pd.concat(parish_frames, ignore_index=True), crs=parish_frames[0].crs)

# ---------------------------------------------------------------------------
# County Business Patterns: county (parish) -level establishment counts
# ---------------------------------------------------------------------------
BUSINESS_CSV = DATA_DIR / "business" / "cbp23co_crpc_sba.csv"
BIZ_COLS = ["est", "n_micro_enterprise", "n_small_business", "n_medium_business", "n_large_employees"]
_biz_raw = pd.read_csv(BUSINESS_CSV, dtype={"fipstate": str, "fipscty": str, "naics": str})
_biz_totals = _biz_raw[_biz_raw["naics"] == "------"].set_index("fipscty")
BIZ_BY_FIPS = {fips: {col: int(_biz_totals.loc[fips, col]) for col in BIZ_COLS}
               for fips in _biz_totals.index}

minx, miny, maxx, maxy = all_tracts.total_bounds
center = [(miny + maxy) / 2, (minx + maxx) / 2]

m = folium.Map(location=center, tiles="cartodbpositron")
m.options["zoomDelta"] = 0.25
m.options["zoomSnap"] = 0.25
m.fit_bounds([[miny, minx], [maxy, maxx]], padding=(80, 80))

tooltip_fields = [
    "Parish",
    "Geography",
    "2024 Unemployment Rate (5-Year ACS)",
    "Threshold Calculation",
    "2024 Per Capita Money Income (5-Year ACS)",
    "Threshold Calculation2",
    "Distress Status",
]
tooltip_aliases = [
    "Parish:",
    "Tract:",
    "Unemployment rate (%):",
    "Unemployment rate vs. U.S. (pct pts):",
    "Per capita income:",
    "Per capita income vs. U.S. (%):",
    "Distress status:",
]

# ---------------------------------------------------------------------------
# 5. One FeatureGroup (layer) per parish: tracts + parish outline + name label
# ---------------------------------------------------------------------------
feature_group_names = {}
tract_layer_names = []
for parish, merged in zip(PARISHES, parish_frames):
    fg = folium.FeatureGroup(name=parish["label"], show=True)

    tract_layer = folium.GeoJson(
        merged,
        style_function=lambda feature: {
            "fillColor": feature["properties"]["fill_color"],
            "color": "#555555",
            "weight": 0.8,
            "fillOpacity": 0.75,
        },
        highlight_function=lambda feature: {"weight": 3.5, "color": "#FFFF00", "fillOpacity": 0.9, "opacity": 1},
        tooltip=folium.GeoJsonTooltip(
            fields=tooltip_fields,
            aliases=tooltip_aliases,
            localize=True,
            sticky=True,
        ),
    )
    tract_layer.add_to(fg)
    tract_layer_names.append(tract_layer.get_name())

    # Dissolve the parish's tracts into a single outline, drawn bold and
    # unfilled so neighboring parishes are visually separated.
    parish_outline = merged.dissolve()[["geometry"]]
    folium.GeoJson(
        parish_outline,
        style_function=lambda feature: {
            "fillOpacity": 0,
            "fill": False,
            "color": "#4FA3E3",
            "weight": 2.5,
            "opacity": 1,
        },
    ).add_to(fg)

    # Label the parish name at the centroid of its outline.
    centroid_3857 = parish_outline.to_crs(epsg=3857).geometry.centroid.iloc[0]
    centroid_4326 = gpd.GeoSeries([centroid_3857], crs=3857).to_crs(epsg=4326).iloc[0]
    folium.Marker(
        location=[centroid_4326.y, centroid_4326.x],
        icon=folium.DivIcon(
            icon_size=(160, 24),
            icon_anchor=(80, 12),
            html=(
                '<div style="font-size: 9pt; font-weight: bold; color: #1A1A1A; '
                'text-align: center; white-space: nowrap; '
                'text-shadow: -1px -1px 2px #fff, 1px -1px 2px #fff, '
                '-1px 1px 2px #fff, 1px 1px 2px #fff;">'
                f'{parish["label"]}</div>'
            ),
        ),
    ).add_to(fg)

    fg.add_to(m)
    feature_group_names[parish["slug"]] = fg.get_name()

# ---------------------------------------------------------------------------
# 6. Legend
# ---------------------------------------------------------------------------
legend_html = """
<div id="distress-legend" style="position: absolute; top: 10px; right: 10px; z-index: 9999;
            background-color: #fdf6ec; padding: 10px 14px; border: 1px solid #888;
            border-radius: 4px; font-size: 13px; line-height: 1.4; width: 173px; box-sizing: border-box;">
  <b>CRPC Region &mdash; DRA Distress Status</b><br>
  <span style="display:inline-block;width:14px;height:14px;background:#FF4444;
        border:1px solid #777;vertical-align:text-top;"></span> Distressed by Both Unemployment Rate and Per Capita Income Share<br>
  <span style="display:inline-block;width:14px;height:14px;background:#FFC7CE;
        border:1px solid #777;vertical-align:text-top;"></span> Distressed by Unemployment Rate<br>
  <span style="display:inline-block;width:14px;height:14px;background:#FFD8A8;
        border:1px solid #777;vertical-align:text-top;"></span> Distressed by Per Capita Income Share<br>
  <span style="display:inline-block;width:14px;height:14px;background:#A9DFBF;
        border:1px solid #777;vertical-align:text-top;"></span> Non-Distressed<br>
  <span style="display:inline-block;width:14px;height:14px;background:#D9D9D9;
        border:1px solid #777;vertical-align:text-top;"></span> No data
</div>
<script>
window.addEventListener('load', function () {
    var panel  = document.getElementById('parish-filter-panel');
    var legend = document.getElementById('distress-legend');
    if (panel && legend) {
        legend.style.top = (panel.offsetTop + panel.offsetHeight + 8) + 'px';
    }
});
</script>
"""
m.get_root().html.add_child(folium.Element(legend_html))

# ---------------------------------------------------------------------------
# 7. "Filter by Parishes" control panel: Show All + 11 parishes + Clear
# ---------------------------------------------------------------------------
checkbox_rows = "\n".join(
    f'  <label><input type="checkbox" class="parish-filter-cb" data-parish="{p["slug"]}" checked> {p["label"]}</label><br>'
    for p in PARISHES
)
layer_js_entries = ",\n".join(
    f'    "{p["slug"]}": {feature_group_names[p["slug"]]}' for p in PARISHES
)

parish_filter_html = f"""
<div id="parish-filter-panel" style="position: absolute; top: 10px; right: 10px; z-index: 9999;
            background-color: #e8f4fd; padding: 10px 14px; border: 1px solid #888;
            border-radius: 4px; font-size: 13px; line-height: 1.6; width: 173px; box-sizing: border-box;
            max-height: calc(100% - 20px); overflow-y: auto;">
  <b>Filter by Parishes</b>
  <hr style="margin:4px 0;">
  <label><input type="checkbox" id="parish-filter-all" checked> <b>Show All</b></label><br>
  <hr style="margin:4px 0;">
{checkbox_rows}
  <hr style="margin:4px 0;">
  <button id="parish-filter-clear" type="button">Clear</button>
</div>
<script>
document.addEventListener("DOMContentLoaded", function () {{
    var map = {m.get_name()};
    var parishLayers = {{
{layer_js_entries}
    }};

    var allCb = document.getElementById("parish-filter-all");
    var parishCbs = document.querySelectorAll(".parish-filter-cb");
    var clearBtn = document.getElementById("parish-filter-clear");

    function updateAllCheckbox() {{
        var checkedCount = 0;
        parishCbs.forEach(function (cb) {{ if (cb.checked) checkedCount++; }});
        allCb.checked = (checkedCount === parishCbs.length);
        allCb.indeterminate = (checkedCount > 0 && checkedCount < parishCbs.length);
    }}

    function setLayer(slug, visible) {{
        var layer = parishLayers[slug];
        if (!layer) return;
        if (visible) {{
            if (!map.hasLayer(layer)) map.addLayer(layer);
        }} else {{
            if (map.hasLayer(layer)) map.removeLayer(layer);
        }}
    }}

    function syncJobMarkers() {{
        if (window.CRPC && window.CRPC.redrawJobMarkers)       window.CRPC.redrawJobMarkers();
        if (window.CRPC && window.CRPC.redrawParishJobMarkers) window.CRPC.redrawParishJobMarkers();
        if (window.CRPC && window.CRPC.redrawBusinessMarkers)  window.CRPC.redrawBusinessMarkers();
        if (window.CRPC && window.CRPC.redrawTractLabels)      window.CRPC.redrawTractLabels();
    }}

    allCb.addEventListener("change", function () {{
        var checked = allCb.checked;
        allCb.indeterminate = false;
        parishCbs.forEach(function (cb) {{
            cb.checked = checked;
            setLayer(cb.getAttribute("data-parish"), checked);
        }});
        syncJobMarkers();
    }});

    parishCbs.forEach(function (cb) {{
        cb.addEventListener("change", function () {{
            setLayer(cb.getAttribute("data-parish"), cb.checked);
            updateAllCheckbox();
            syncJobMarkers();
        }});
    }});

    clearBtn.addEventListener("click", function () {{
        allCb.checked = false;
        allCb.indeterminate = false;
        parishCbs.forEach(function (cb) {{
            cb.checked = false;
            setLayer(cb.getAttribute("data-parish"), false);
        }});
        syncJobMarkers();
    }});
}});
</script>
"""
m.get_root().html.add_child(folium.Element(parish_filter_html))

# ---------------------------------------------------------------------------
# 7b. "Download Map (600 DPI)" control (bottom-left). The button calls the
#     server-side /download_map route, which renders the whole CRPC region to a
#     large static PNG with matplotlib + geopandas (crisp vector tracts, parish
#     outlines/labels, CartoDB Positron basemap) at 600 DPI -- no browser
#     capture, so nothing shifts. While rendering it shows a status message,
#     then triggers the download and shows a green check.
# ---------------------------------------------------------------------------
DOWNLOAD_DPI = 600
download_html = """
<div id="download-panel" style="position: absolute; bottom: 20px; left: 10px; z-index: 9999;
            background-color: #e8eaed; padding: 3px 4px; border: 1px solid #888;
            border-radius: 4px; font-size: 9px; box-shadow: 0 1px 4px rgba(0,0,0,0.3);
            display: flex; flex-direction: column; gap: 2px;">
  <div style="display: flex; flex-direction: row; gap: 3px;">
    <button id="download-map-btn" type="button"
            style="cursor: pointer; padding: 2px 4px; font-size: 9px; font-weight: bold;">
      &#11015; Download 1 (600 DPI)
    </button>
    <button id="download-screen-btn" type="button"
            style="cursor: pointer; padding: 2px 4px; font-size: 9px; font-weight: bold;">
      &#128247; Download 2 (Screen)
    </button>
    <button id="fullscreen-btn" type="button"
            style="cursor: pointer; padding: 2px 4px; font-size: 9px; font-weight: bold;">
      &#9974; Full Screen
    </button>
  </div>
  <div style="display: flex; flex-direction: row; gap: 3px;">
    <button id="download-clean-btn" type="button"
            style="cursor: pointer; padding: 2px 4px; font-size: 9px; font-weight: bold; width: 100%;">
      &#128444; Download 3 (Clean Map Only)
    </button>
  </div>
  <div id="download-status" style="font-size: 9px; min-height: 11px; color: #555;"></div>
</div>
<script>
document.addEventListener("DOMContentLoaded", function () {
    var btn = document.getElementById("download-map-btn");
    var screenBtn = document.getElementById("download-screen-btn");
    var cleanBtn = document.getElementById("download-clean-btn");
    var statusEl = document.getElementById("download-status");
    var fsBtn = document.getElementById("fullscreen-btn");
    var mapPaneEl = document.getElementById("map-pane");

    // Floating UI chrome drawn on top of the map (legend, parish filter,
    // title, search bar, tract-number toggle, this download panel itself) --
    // hidden during the "Download 3 (Clean)" capture so only the Leaflet map
    // and its data layers (tract colors, highlights, job markers) show.
    var overlayIds = ["map-title", "tract-search-bar", "distress-legend",
                       "parish-filter-panel", "download-panel", "tract-num-panel"];
    function setOverlaysHidden(hidden) {
        overlayIds.forEach(function (id) {
            var el = document.getElementById(id);
            if (el) el.style.visibility = hidden ? "hidden" : "";
        });
    }

    function isMapFullscreen() {
        var fsEl = document.fullscreenElement || document.webkitFullscreenElement;
        return fsEl === mapPaneEl;
    }

    fsBtn.addEventListener("click", function () {
        if (!isMapFullscreen()) {
            var req = mapPaneEl.requestFullscreen || mapPaneEl.webkitRequestFullscreen;
            if (req) req.call(mapPaneEl);
        } else {
            var exit = document.exitFullscreen || document.webkitExitFullscreen;
            if (exit) exit.call(document);
        }
    });

    function onFullscreenChange() {
        fsBtn.innerHTML = isMapFullscreen() ? "&#9974; Exit Full Screen" : "&#9974; Full Screen";
        // Leaflet caches the container size; force it to re-measure once the
        // fullscreen transition (and its layout reflow) has settled.
        setTimeout(function () {
            if (window.CRPC && window.CRPC.map) window.CRPC.map.invalidateSize();
        }, 150);
    }
    document.addEventListener("fullscreenchange", onFullscreenChange);
    document.addEventListener("webkitfullscreenchange", onFullscreenChange);

    function currentState() {
        var parishes = [];
        document.querySelectorAll(".parish-filter-cb").forEach(function (cb) {
            if (cb.checked) parishes.push(cb.getAttribute("data-parish"));
        });
        var statuses = [];
        document.querySelectorAll(".distress-filter-cb").forEach(function (cb) {
            if (cb.checked) statuses.push(cb.getAttribute("data-status"));
        });
        var threshold = "all";
        document.querySelectorAll(".threshold-filter-rb").forEach(function (rb) {
            if (rb.checked) threshold = rb.value;
        });
        var geoids = [];
        if (window.CRPC && window.CRPC.highlightedGeoids) {
            window.CRPC.highlightedGeoids.forEach(function (g) { geoids.push(g); });
        }
        var bounds = null;
        if (window.CRPC && window.CRPC.map) {
            var b = window.CRPC.map.getBounds();
            bounds = [[b.getSouth(), b.getWest()], [b.getNorth(), b.getEast()]];
        }
        var knob = (window.CRPC && window.CRPC.getKnobState)
            ? window.CRPC.getKnobState() : {};
        var jobTractEnableCb  = document.getElementById("jobs-enable");
        var jobTractSectorEl  = document.querySelector(".jobs-sector-item.jobs-selected");
        var jobTractNumsCb    = document.getElementById("jobs-show-nums");
        var jobParishEnableCb = document.getElementById("pjobs-enable");
        var jobParishSectorEl = document.querySelector(".pjobs-sector-item.pjobs-selected");
        var jobParishNumsCb   = document.getElementById("pjobs-show-nums");
        var bizEnableCb       = document.getElementById("biz-enable");
        var bizMetricEl       = document.querySelector(".biz-metric-item.biz-selected");
        var bizNumsCb         = document.getElementById("biz-show-nums");
        var tractNumCb        = document.getElementById("tract-num-enable");
        return {parishes: parishes, statuses: statuses, threshold: threshold,
                geoids: geoids, bounds: bounds,
                income_enabled: knob.income_enabled || false,
                unemployment_enabled: knob.unemployment_enabled || false,
                income_bin: knob.income_bin || 1,
                unemployment_bin: knob.unemployment_bin || 1,
                job_tract_enabled:   jobTractEnableCb  ? jobTractEnableCb.checked  : false,
                job_tract_sector:    jobTractSectorEl  ? jobTractSectorEl.getAttribute("data-key") : "C000",
                job_tract_show_nums: jobTractNumsCb    ? jobTractNumsCb.checked    : false,
                job_parish_enabled:   jobParishEnableCb ? jobParishEnableCb.checked  : false,
                job_parish_sector:    jobParishSectorEl ? jobParishSectorEl.getAttribute("data-key") : "C000",
                job_parish_show_nums: jobParishNumsCb   ? jobParishNumsCb.checked    : false,
                biz_parish_enabled:   bizEnableCb ? bizEnableCb.checked : false,
                biz_parish_metric:    bizMetricEl ? bizMetricEl.getAttribute("data-key") : "est",
                biz_parish_show_nums: bizNumsCb   ? bizNumsCb.checked   : false,
                show_tract_nums: tractNumCb ? tractNumCb.checked : false};
    }

    btn.addEventListener("click", function () {
        btn.disabled = true;
        statusEl.style.color = "#555";
        statusEl.textContent = "Rendering high-res map on the server… (can take ~15-40s)";
        fetch("/download_map", {
            method: "POST",
            headers: {"Content-Type": "application/json"},
            body: JSON.stringify(currentState())
        })
            .then(function (r) {
                if (!r.ok) { return r.text().then(function (t) { throw new Error(t || ("server error " + r.status)); }); }
                return r.blob();
            })
            .then(function (blob) {
                var url = URL.createObjectURL(blob);
                var a = document.createElement("a");
                a.href = url;
                a.download = "crpc_distress_map_600dpi.png";
                document.body.appendChild(a);
                a.click();
                document.body.removeChild(a);
                URL.revokeObjectURL(url);
                statusEl.style.color = "#137333";
                statusEl.innerHTML = "&#10004; Image downloaded";
                btn.disabled = false;
            })
            .catch(function (err) {
                statusEl.style.color = "#c00";
                statusEl.textContent = "Error: " + (err && err.message ? err.message : err);
                btn.disabled = false;
            });
    });

    screenBtn.addEventListener("click", function () {
        var mapPane = document.getElementById("map-pane");
        screenBtn.disabled = true;
        statusEl.style.color = "#555";
        statusEl.textContent = "Select this tab when prompted…";
        navigator.mediaDevices.getDisplayMedia({
            video: {
                displaySurface: "browser",
                width:  { ideal: 7680 },
                height: { ideal: 4320 },
                frameRate: { ideal: 1 }
            },
            preferCurrentTab: true
        }).then(function (stream) {
            var video = document.createElement("video");
            video.muted = true;
            video.srcObject = stream;
            video.onloadedmetadata = function () {
                video.play();
                setTimeout(function () {
                    var vw = video.videoWidth;
                    var vh = video.videoHeight;
                    var scaleX = vw / window.innerWidth;
                    var scaleY = vh / window.innerHeight;
                    // Crop canvas at native video resolution
                    var crop = document.createElement("canvas");
                    var cropCtx = crop.getContext("2d");
                    if (mapPane) {
                        var rect = mapPane.getBoundingClientRect();
                        crop.width  = Math.round(rect.width  * scaleX);
                        crop.height = Math.round(rect.height * scaleY);
                        cropCtx.drawImage(video,
                            Math.round(rect.left * scaleX), Math.round(rect.top * scaleY),
                            crop.width, crop.height,
                            0, 0, crop.width, crop.height);
                    } else {
                        crop.width = vw; crop.height = vh;
                        cropCtx.drawImage(video, 0, 0);
                    }
                    stream.getTracks().forEach(function (t) { t.stop(); });
                    // Upscale 2x with high-quality interpolation for print resolution
                    var hi = document.createElement("canvas");
                    hi.width  = crop.width  * 2;
                    hi.height = crop.height * 2;
                    var hiCtx = hi.getContext("2d");
                    hiCtx.imageSmoothingEnabled = true;
                    hiCtx.imageSmoothingQuality = "high";
                    hiCtx.drawImage(crop, 0, 0, hi.width, hi.height);
                    hi.toBlob(function (blob) {
                        var url = URL.createObjectURL(blob);
                        var a = document.createElement("a");
                        a.href = url;
                        a.download = "crpc_distress_map_screen.png";
                        document.body.appendChild(a);
                        a.click();
                        document.body.removeChild(a);
                        URL.revokeObjectURL(url);
                        statusEl.style.color = "#137333";
                        statusEl.innerHTML = "&#10004; Screen captured (hi-res)";
                        screenBtn.disabled = false;
                    }, "image/png");
                }, 300);
            };
        }).catch(function (err) {
            statusEl.style.color = "#c00";
            statusEl.textContent = "Capture error: " + (err && err.message ? err.message : err);
            screenBtn.disabled = false;
        });
    });

    cleanBtn.addEventListener("click", function () {
        var mapPane = document.getElementById("map-pane");
        cleanBtn.disabled = true;
        statusEl.style.color = "#555";
        statusEl.textContent = "Select this tab when prompted…";
        // Hide the legend/filter/title/search-bar/download chrome *before*
        // the capture stream is even requested, so every frame the stream
        // ever produces -- including the first one -- already shows the
        // clean state. (Hiding it only after the stream starts is unreliable:
        // the stream's frameRate is throttled for quality, so a frame grabbed
        // soon after hiding can still be a stale pre-hide frame.)
        setOverlaysHidden(true);
        navigator.mediaDevices.getDisplayMedia({
            video: {
                displaySurface: "browser",
                width:  { ideal: 7680 },
                height: { ideal: 4320 },
                frameRate: { ideal: 5 }
            },
            preferCurrentTab: true
        }).then(function (stream) {
            var video = document.createElement("video");
            video.muted = true;
            video.srcObject = stream;
            video.onloadedmetadata = function () {
                video.play();
                setTimeout(function () {
                    var vw = video.videoWidth;
                    var vh = video.videoHeight;
                    var scaleX = vw / window.innerWidth;
                    var scaleY = vh / window.innerHeight;
                    var crop = document.createElement("canvas");
                    var cropCtx = crop.getContext("2d");
                    if (mapPane) {
                        var rect = mapPane.getBoundingClientRect();
                        crop.width  = Math.round(rect.width  * scaleX);
                        crop.height = Math.round(rect.height * scaleY);
                        cropCtx.drawImage(video,
                            Math.round(rect.left * scaleX), Math.round(rect.top * scaleY),
                            crop.width, crop.height,
                            0, 0, crop.width, crop.height);
                    } else {
                        crop.width = vw; crop.height = vh;
                        cropCtx.drawImage(video, 0, 0);
                    }
                    stream.getTracks().forEach(function (t) { t.stop(); });
                    setOverlaysHidden(false);
                    // Upscale 2x with high-quality interpolation for print resolution
                    var hi = document.createElement("canvas");
                    hi.width  = crop.width  * 2;
                    hi.height = crop.height * 2;
                    var hiCtx = hi.getContext("2d");
                    hiCtx.imageSmoothingEnabled = true;
                    hiCtx.imageSmoothingQuality = "high";
                    hiCtx.drawImage(crop, 0, 0, hi.width, hi.height);
                    hi.toBlob(function (blob) {
                        var url = URL.createObjectURL(blob);
                        var a = document.createElement("a");
                        a.href = url;
                        a.download = "crpc_distress_map_clean.png";
                        document.body.appendChild(a);
                        a.click();
                        document.body.removeChild(a);
                        URL.revokeObjectURL(url);
                        statusEl.style.color = "#137333";
                        statusEl.innerHTML = "&#10004; Clean screenshot captured";
                        cleanBtn.disabled = false;
                    }, "image/png");
                }, 400);
            };
        }).catch(function (err) {
            setOverlaysHidden(false);
            statusEl.style.color = "#c00";
            statusEl.textContent = "Capture error: " + (err && err.message ? err.message : err);
            cleanBtn.disabled = false;
        });
    });
});
</script>
"""
m.get_root().html.add_child(folium.Element(download_html))

# ---------------------------------------------------------------------------
# 8. "Filter by Distress Status" (bottom-center) + "Filter by Threshold"
#    (bottom-right) control bars, plus AI-highlight support. Both filters
#    apply together (a tract must pass both to be shown). The tractStyle
#    function also dims/highlights tracts based on `highlightedGeoids`,
#    which the AI query box (section 11) populates. Everything that the
#    AI query script needs is exposed on `window.CRPC`.
# ---------------------------------------------------------------------------
distress_checkbox_items = "\n".join(
    '  <label style="display:flex; align-items:center; gap:4px;">'
    f'<input type="checkbox" class="distress-filter-cb" data-status="{d["label"]}" checked> '
    f'<span style="display:inline-block;width:12px;height:12px;background:{d["color"]};'
    'border:1px solid #777;flex-shrink:0;"></span> '
    f'{d["short"]}</label>'
    for d in DISTRESS_STATUSES
)
tract_layer_js_array = "[" + ", ".join(tract_layer_names) + "]"

# ---------------------------------------------------------------------------
# Range-knob inputs: two sliders (per-capita income and unemployment rate),
# each split into N_BINS equal-width bins between the dataset min and max.
# Selecting a bin shows only tracts whose value falls in that range.
# ---------------------------------------------------------------------------
INCOME_COL = "income_value"          # numeric per-capita income (parsed)
UNEMP_COL = "unemployment_value"     # numeric unemployment rate (parsed)
N_BINS = 10

_income_vals = pd.to_numeric(all_tracts[INCOME_COL], errors="coerce").dropna()
_unemp_vals = pd.to_numeric(all_tracts[UNEMP_COL], errors="coerce").dropna()
income_min, income_max = float(_income_vals.min()), float(_income_vals.max())
unemp_min, unemp_max = float(_unemp_vals.min()), float(_unemp_vals.max())
income_step = (income_max - income_min) / N_BINS
unemp_step = (unemp_max - unemp_min) / N_BINS

# Bin boundary values (N_BINS + 1 of them) used to draw the slider scale.
_income_bounds = [income_min + i * income_step for i in range(N_BINS + 1)]
_unemp_bounds = [unemp_min + i * unemp_step for i in range(N_BINS + 1)]
# A tick mark at every boundary; value labels at every other boundary (so the
# min sits at the left end and the max at the right end without crowding).
_knob_ticks_html = "".join("<span class='knob-tick'></span>" for _ in range(N_BINS + 1))
# Income knob runs high -> low (reversed): max at the left, min at the right.
income_scale_html = "".join(
    f"<span>${_income_bounds[i] / 1000:.0f}k</span>" for i in range(N_BINS, -1, -2)
)
unemp_scale_html = "".join(
    f"<span>{_unemp_bounds[i]:.0f}%</span>" for i in range(0, N_BINS + 1, 2)
)

# ---------------------------------------------------------------------------
# Thermal gradients for the slider tracks. The yellow midpoint is anchored to a
# meaningful value (not the data centre): the U.S. average per-capita income
# for the income knob, and a 1% unemployment rate for the unemployment knob.
# Left of yellow is green, right of yellow ramps yellow -> dark red.
# ---------------------------------------------------------------------------
US_AVG_INCOME = 44673.0     # U.S. average per-capita money income
UNEMP_YELLOW_AT = 1.0       # unemployment rate (%) that maps to yellow


def _heat_gradient(yellow_frac):
    """Green -> yellow (at yellow_frac) -> dark red, as a CSS linear-gradient."""
    yp = max(0.02, min(0.98, yellow_frac))
    stops = [(0.0, "#1a9850"), (0.55 * yp, "#66bd63"), (0.85 * yp, "#a6d96a"),
             (yp, "#ffff33")]
    for f, c in [(0.18, "#fdae61"), (0.45, "#f46d43"), (0.72, "#d73027"), (1.0, "#a50026")]:
        stops.append((yp + f * (1.0 - yp), c))
    return ("linear-gradient(to right, "
            + ", ".join(f"{c} {p * 100:.1f}%" for p, c in stops) + ")")


# Income slider is reversed (max income at the left end).
_income_yellow = (income_max - US_AVG_INCOME) / (income_max - income_min)
_unemp_yellow = (UNEMP_YELLOW_AT - unemp_min) / (unemp_max - unemp_min)
income_gradient = _heat_gradient(_income_yellow)
unemp_gradient = _heat_gradient(_unemp_yellow)

JOB_SECTOR_LABELS = {
    "C000":  "Total Jobs",
    "CNS01": "Agriculture, Forestry, Fishing & Hunting",
    "CNS02": "Mining, Quarrying & Oil and Gas Extraction",
    "CNS03": "Utilities",
    "CNS04": "Construction",
    "CNS05": "Manufacturing",
    "CNS06": "Wholesale Trade",
    "CNS07": "Retail Trade",
    "CNS08": "Transportation and Warehousing",
    "CNS09": "Information",
    "CNS10": "Finance and Insurance",
    "CNS11": "Real Estate and Rental and Leasing",
    "CNS12": "Professional, Scientific, and Technical Services",
    "CNS13": "Management of Companies and Enterprises",
    "CNS14": "Admin., Support & Waste Management",
    "CNS15": "Educational Services",
    "CNS16": "Health Care and Social Assistance",
    "CNS17": "Arts, Entertainment, and Recreation",
    "CNS18": "Accommodation and Food Services",
    "CNS19": "Other Services (excl. Public Administration)",
    "CNS20": "Public Administration",
}
JOB_RANGES = {}
for _jcol in JOB_COLS:
    _jvals = all_tracts[_jcol].dropna()
    _jmin = int(_jvals.min())
    _jmax = int(_jvals.max())
    _jstep = max(1, (_jmax - _jmin) // N_BINS)
    JOB_RANGES[_jcol] = {"min": _jmin, "max": _jmax, "step": _jstep}
JOB_RANGES_JSON = json.dumps(JOB_RANGES)
JOB_SECTOR_LABELS_JSON = json.dumps(JOB_SECTOR_LABELS)

JOB_EMOJIS = {
    "C000":  "\U0001f4bc",  # 💼 briefcase
    "CNS01": "\U0001f33e",  # 🌾 agriculture
    "CNS02": "⛏️",  # ⛏️ mining
    "CNS03": "⚡",      # ⚡ utilities
    "CNS04": "\U0001f3d7️",  # 🏗️ construction
    "CNS05": "\U0001f3ed",  # 🏭 manufacturing
    "CNS06": "\U0001f4e6",  # 📦 wholesale
    "CNS07": "\U0001f6d2",  # 🛒 retail
    "CNS08": "\U0001f69a",  # 🚚 transportation
    "CNS09": "\U0001f4bb",  # 💻 information
    "CNS10": "\U0001f3e6",  # 🏦 finance
    "CNS11": "\U0001f3e0",  # 🏠 real estate
    "CNS12": "\U0001f52c",  # 🔬 professional/technical
    "CNS13": "\U0001f3e2",  # 🏢 management
    "CNS14": "\U0001f9f9",  # 🧹 admin/support
    "CNS15": "\U0001f393",  # 🎓 education
    "CNS16": "\U0001f3e5",  # 🏥 health care
    "CNS17": "\U0001f3ad",  # 🎭 arts/entertainment
    "CNS18": "\U0001f37d️",  # 🍽️ food/accommodation
    "CNS19": "\U0001f527",  # 🔧 other services
    "CNS20": "\U0001f3db️",  # 🏛️ public administration
}
JOB_EMOJIS_JSON = json.dumps(JOB_EMOJIS)

# Tract centroids + all job column values — used by the JS marker layer.
_tracts_ll = all_tracts.to_crs(epsg=4326) if all_tracts.crs.to_epsg() != 4326 else all_tracts
_parish_label_to_slug = {p["label"]: p["slug"] for p in PARISHES}
JOB_DATA = []
for _, _row in _tracts_ll.iterrows():
    _c = _row.geometry.centroid
    _entry = {
        "geoid": _row["GEOID"],
        "parish": _parish_label_to_slug.get(_row["Parish"], ""),
        "lat": round(_c.y, 6),
        "lng": round(_c.x, 6),
        "jobs": {},
    }
    for _jc in JOB_COLS:
        _entry["jobs"][_jc] = int(_row[_jc])
    JOB_DATA.append(_entry)
JOB_DATA_JSON = json.dumps(JOB_DATA)
TRACT_POINTS = [{"geoid": e["geoid"], "parish": e["parish"], "lat": e["lat"], "lng": e["lng"]} for e in JOB_DATA]
TRACT_POINTS_JSON = json.dumps(TRACT_POINTS)

# Parish-level job totals (sum tracts per parish) + parish centroids
_parish_job_sums = {}
for _label, _group in all_tracts.groupby("Parish"):
    _parish_job_sums[_label] = {col: int(_group[col].sum()) for col in JOB_COLS}

_parish_dissolved = _tracts_ll.dissolve(by="Parish")
_parish_centroids = _parish_dissolved.geometry.centroid

PARISH_JOB_DATA = []
for _p in PARISHES:
    _label, _slug = _p["label"], _p["slug"]
    if _label not in _parish_centroids.index:
        continue
    _c = _parish_centroids[_label]
    PARISH_JOB_DATA.append({
        "parish": _slug,
        "lat": round(_c.y, 6),
        "lng": round(_c.x, 6),
        "jobs": _parish_job_sums.get(_label, {col: 0 for col in JOB_COLS}),
    })

PARISH_JOB_DATA_JSON = json.dumps(PARISH_JOB_DATA)

PARISH_JOB_RANGES = {}
for _jc in JOB_COLS:
    _vals = [e["jobs"].get(_jc, 0) for e in PARISH_JOB_DATA]
    PARISH_JOB_RANGES[_jc] = {"min": min(_vals) if _vals else 0, "max": max(_vals) if _vals else 0}
PARISH_JOB_RANGES_JSON = json.dumps(PARISH_JOB_RANGES)

# Parish-level business-establishment counts (County Business Patterns, already
# one row per parish -- no aggregation needed, just attach the parish centroid).
BIZ_METRIC_LABELS = {
    "est": "Total Establishments",
    "n_micro_enterprise": "Micro Enterprises (<10 employees)",
    "n_small_business": "Small Business (10-99 employees)",
    "n_medium_business": "Medium Business (100-499 employees)",
    "n_large_employees": "Large Employers (500+ employees)",
}
BIZ_METRIC_LABELS_JSON = json.dumps(BIZ_METRIC_LABELS)

PARISH_BIZ_DATA = []
for _p in PARISHES:
    _label, _slug, _fips = _p["label"], _p["slug"], _p["fips"]
    if _label not in _parish_centroids.index:
        continue
    _c = _parish_centroids[_label]
    PARISH_BIZ_DATA.append({
        "parish": _slug,
        "lat": round(_c.y, 6),
        "lng": round(_c.x, 6),
        "biz": BIZ_BY_FIPS.get(_fips, {col: 0 for col in BIZ_COLS}),
    })

PARISH_BIZ_DATA_JSON = json.dumps(PARISH_BIZ_DATA)

PARISH_BIZ_RANGES = {}
for _bc in BIZ_COLS:
    _vals = [e["biz"].get(_bc, 0) for e in PARISH_BIZ_DATA]
    PARISH_BIZ_RANGES[_bc] = {"min": min(_vals) if _vals else 0, "max": max(_vals) if _vals else 0}
PARISH_BIZ_RANGES_JSON = json.dumps(PARISH_BIZ_RANGES)

filters_html = f"""
<style>
  /* Separate rules per browser: a grouped selector containing the Firefox-only
     ::-moz-range-track would be dropped entirely by Chrome, so keep them apart. */
  #income-slider.heat {{ background: {income_gradient}; }}
  #income-slider.heat::-moz-range-track {{ background: {income_gradient}; }}
  #unemp-slider.heat {{ background: {unemp_gradient}; }}
  #unemp-slider.heat::-moz-range-track {{ background: {unemp_gradient}; }}
</style>
<div id="distress-filter-panel" class="sidebar-panel">
  <b>Filter by Distress Status</b>
  <hr style="margin:4px 0;">
  <label style="display:flex; align-items:center; gap:4px;">
    <input type="checkbox" id="distress-filter-all" checked> <b>Show All</b>
  </label>
{distress_checkbox_items}
  <hr style="margin:4px 0;">
  <button id="distress-filter-clear" type="button">Clear</button>
</div>

<div id="threshold-filter-panel" class="sidebar-panel">
  <b>Filter by Threshold</b>
  <hr style="margin:4px 0;">
  <label style="display:flex; align-items:center; gap:4px;">
    <input type="radio" name="threshold-filter" class="threshold-filter-rb" value="all" checked> Show All
  </label>
  <label style="display:flex; align-items:center; gap:4px;">
    <input type="radio" name="threshold-filter" class="threshold-filter-rb" value="income"> Per Capita Income &lt; 60% of U.S. Average
  </label>
  <label style="display:flex; align-items:center; gap:4px;">
    <input type="radio" name="threshold-filter" class="threshold-filter-rb" value="unemployment"> Unemployment Rate &ge; 2 pct pts above U.S. Average
  </label>
</div>

<div id="income-knob-panel" class="sidebar-panel knob-panel">
  <label class="knob-head">
    <input type="checkbox" class="knob-enable" id="income-enable"> <b>Filter by Per Capita Income</b>
  </label>
  <div class="knob-body" id="income-body">
    <input type="range" id="income-slider" class="heat" min="1" max="{N_BINS}" step="1" value="1">
    <div class="knob-ticks">{_knob_ticks_html}</div>
    <div class="knob-scale">{income_scale_html}</div>
    <div class="knob-current" id="income-slider-label">Show all (default)</div>
  </div>
</div>

<div id="unemp-knob-panel" class="sidebar-panel knob-panel">
  <label class="knob-head">
    <input type="checkbox" class="knob-enable" id="unemp-enable"> <b>Filter by Unemployment Rate</b>
  </label>
  <div class="knob-body" id="unemp-body">
    <input type="range" id="unemp-slider" class="heat" min="1" max="{N_BINS}" step="1" value="1">
    <div class="knob-ticks">{_knob_ticks_html}</div>
    <div class="knob-scale">{unemp_scale_html}</div>
    <div class="knob-current" id="unemp-slider-label">Show all (default)</div>
  </div>
</div>

<script>
document.addEventListener("DOMContentLoaded", function () {{
    var map = {m.get_name()};
    var tractLayers = {tract_layer_js_array};

    var distressAllCb = document.getElementById("distress-filter-all");
    var statusCbs = document.querySelectorAll(".distress-filter-cb");
    var distressClearBtn = document.getElementById("distress-filter-clear");

    var thresholdRbs = document.querySelectorAll(".threshold-filter-rb");

    // Tracks which Distress Status categories are currently selected, and
    // which threshold mode is active. The tract style function below reads
    // from these on every resetStyle() call, so hovering off a tract
    // restores it to the *filtered* state instead of the original
    // unfiltered style.
    var distressSelected = {{}};
    statusCbs.forEach(function (cb) {{
        distressSelected[cb.getAttribute("data-status")] = cb.checked;
    }});

    var thresholdMode = "all";

    // Range-knob state: bin 0 == "All"; bins 1..N_BINS select a value range.
    var INCOME_COL = {json.dumps(INCOME_COL)};
    var UNEMP_COL = {json.dumps(UNEMP_COL)};
    var N_BINS = {N_BINS};
    var INCOME_MIN = {income_min}, INCOME_MAX = {income_max}, INCOME_STEP = {income_step};
    var UNEMP_MIN = {unemp_min}, UNEMP_MAX = {unemp_max}, UNEMP_STEP = {unemp_step};
    // Each knob is only applied as a filter when its checkbox is enabled;
    // disabled (the default) means "show all" for that metric.
    var incomeEnabled = false, unempEnabled = false;
    var incomeBin = 1, unempBin = 1;

    // Set of GEOIDs returned by the most recent AI query (section 11).
    // When non-empty, matching tracts are emphasized and the rest dimmed.
    var highlightedGeoids = new Set();

    function thresholdOk(feature) {{
        if (thresholdMode === "income") {{
            return feature.properties["meets_income_threshold"] === true;
        }}
        if (thresholdMode === "unemployment") {{
            return feature.properties["meets_unemployment_threshold"] === true;
        }}
        return true;
    }}

    function inBin(v, bin, lo0, step, hi0) {{
        if (bin <= 0) return true;
        if (v === null || v === undefined) return false;
        v = Number(v);
        if (!isFinite(v)) return false;
        var lo = lo0 + (bin - 1) * step;
        var hi = (bin === N_BINS) ? hi0 : lo0 + bin * step;
        return v >= lo && v <= hi;
    }}

    function metricOk(feature) {{
        if (incomeEnabled &&
            !inBin(feature.properties[INCOME_COL], incomeBin, INCOME_MIN, INCOME_STEP, INCOME_MAX)) {{
            return false;
        }}
        if (unempEnabled &&
            !inBin(feature.properties[UNEMP_COL], unempBin, UNEMP_MIN, UNEMP_STEP, UNEMP_MAX)) {{
            return false;
        }}
        return true;
    }}

    var selectedGeoid = null;

    function tractStyle(feature) {{
        var distressOk = distressSelected[feature.properties["Distress Status"]] !== false;
        var visible = distressOk && thresholdOk(feature) && metricOk(feature);
        var style = {{
            fillColor: feature.properties.fill_color,
            color: visible ? "#555555" : "#c8a8f5",
            weight: 0.8,
            opacity: 1,
            fillOpacity: visible ? 0.75 : 0,
        }};
        if (highlightedGeoids.size > 0) {{
            if (highlightedGeoids.has(feature.properties.GEOID)) {{
                style.color = "#0033CC";
                style.weight = 2.45;
                style.opacity = 1;
                style.fillOpacity = Math.max(style.fillOpacity, 0.85);
            }} else {{
                style.fillOpacity = style.fillOpacity * 0.15;
                style.opacity = style.opacity * 0.25;
            }}
        }}
        if (selectedGeoid && feature.properties.GEOID === selectedGeoid) {{
            // Border itself stays as-is (so it doesn't collide with the parish's
            // blue outline when the tract sits on a parish edge); the selection
            // ring is instead drawn as a separate inset layer -- see
            // drawSelectedInset() below. Just boost the fill so it still pops.
            style.fillOpacity = Math.max(style.fillOpacity, 0.92);
        }}
        return style;
    }}

    tractLayers.forEach(function (geo) {{
        geo.options.style = tractStyle;
    }});

    function refreshTracts() {{
        tractLayers.forEach(function (geo) {{
            geo.eachLayer(function (layer) {{
                geo.resetStyle(layer);
            }});
        }});
    }}

    function updateDistressAllCheckbox() {{
        var checkedCount = 0;
        statusCbs.forEach(function (cb) {{ if (cb.checked) checkedCount++; }});
        distressAllCb.checked = (checkedCount === statusCbs.length);
        distressAllCb.indeterminate = (checkedCount > 0 && checkedCount < statusCbs.length);
    }}

    function applyDistressFilter() {{
        statusCbs.forEach(function (cb) {{
            distressSelected[cb.getAttribute("data-status")] = cb.checked;
        }});
        refreshTracts();
    }}

    distressAllCb.addEventListener("change", function () {{
        var checked = distressAllCb.checked;
        distressAllCb.indeterminate = false;
        statusCbs.forEach(function (cb) {{ cb.checked = checked; }});
        applyDistressFilter();
    }});

    statusCbs.forEach(function (cb) {{
        cb.addEventListener("change", function () {{
            updateDistressAllCheckbox();
            applyDistressFilter();
        }});
    }});

    distressClearBtn.addEventListener("click", function () {{
        distressAllCb.checked = false;
        distressAllCb.indeterminate = false;
        statusCbs.forEach(function (cb) {{ cb.checked = false; }});
        applyDistressFilter();
    }});

    thresholdRbs.forEach(function (rb) {{
        rb.addEventListener("change", function () {{
            if (rb.checked) {{
                thresholdMode = rb.value;
                refreshTracts();
            }}
        }});
    }});

    // --- Per-capita income / unemployment range knobs --------------------
    function fmtMoney(x) {{ return "$" + Math.round(x).toLocaleString(); }}
    function fmtPct(x) {{ return (Math.round(x * 10) / 10) + "%"; }}
    function binLabel(bin, lo0, step, hi0, money) {{
        if (bin <= 0) return "All";
        var lo = lo0 + (bin - 1) * step;
        var hi = (bin === N_BINS) ? hi0 : lo0 + bin * step;
        return money ? (fmtMoney(lo) + " – " + fmtMoney(hi))
                     : (fmtPct(lo) + " – " + fmtPct(hi));
    }}

    var incomeEnable = document.getElementById("income-enable");
    var incomeBody = document.getElementById("income-body");
    var incomeSlider = document.getElementById("income-slider");
    var incomeLabel = document.getElementById("income-slider-label");
    function updateIncome() {{
        incomeEnabled = incomeEnable.checked;
        // Income slider is reversed: left (value 1) = highest income bin.
        incomeBin = N_BINS + 1 - parseInt(incomeSlider.value, 10);
        incomeBody.classList.toggle("active", incomeEnabled);
        incomeLabel.textContent = incomeEnabled
            ? ("Showing " + binLabel(incomeBin, INCOME_MIN, INCOME_STEP, INCOME_MAX, true))
            : "Show all (default)";
        refreshTracts();
    }}
    incomeEnable.addEventListener("change", updateIncome);
    incomeSlider.addEventListener("input", updateIncome);

    var unempEnable = document.getElementById("unemp-enable");
    var unempBody = document.getElementById("unemp-body");
    var unempSlider = document.getElementById("unemp-slider");
    var unempLabel = document.getElementById("unemp-slider-label");
    function updateUnemp() {{
        unempEnabled = unempEnable.checked;
        unempBin = parseInt(unempSlider.value, 10);
        unempBody.classList.toggle("active", unempEnabled);
        unempLabel.textContent = unempEnabled
            ? ("Showing " + binLabel(unempBin, UNEMP_MIN, UNEMP_STEP, UNEMP_MAX, false))
            : "Show all (default)";
        refreshTracts();
    }}
    unempEnable.addEventListener("change", updateUnemp);
    unempSlider.addEventListener("input", updateUnemp);

    // Expose to the AI query script (section 11).
    window.CRPC = {{
        map: map,
        tractLayers: tractLayers,
        highlightedGeoids: highlightedGeoids,
        refreshTracts: refreshTracts,
        setSelectedGeoid: function (g) {{ selectedGeoid = g; }},
        getKnobState: function () {{
            return {{
                income_enabled: incomeEnabled,
                income_bin: incomeBin,
                unemployment_enabled: unempEnabled,
                unemployment_bin: unempBin,
            }};
        }},
    }};
}});
</script>
"""

# ---------------------------------------------------------------------------
# 9. Flat lookup tables used by the AI query box: bounding box per tract
#    (for map fitBounds/zoom) and the columns shown in the results table.
# ---------------------------------------------------------------------------
geoid_bounds = {}
for _, row in all_tracts.iterrows():
    b_minx, b_miny, b_maxx, b_maxy = row.geometry.bounds
    geoid_bounds[row["GEOID"]] = [[b_miny, b_minx], [b_maxy, b_maxx]]

parish_bounds = {}
for parish in PARISHES:
    pgdf = all_tracts[all_tracts["Parish"] == parish["label"]]
    if len(pgdf):
        px0, py0, px1, py1 = pgdf.total_bounds
        parish_bounds[parish["label"]] = [[py0, px0], [py1, px1]]

# Inset (buffered-inward) copy of each tract's boundary, used only to draw the
# "selected row" highlight ring a few meters inside the tract's true edge --
# tracts on a parish's outer edge would otherwise have their highlight ring
# drawn exactly on top of the parish's own blue outline, making both illegible.
INSET_BUFFER_METERS = -15
_inset_web = all_tracts.to_crs(epsg=3857)
_inset_buffered = _inset_web.geometry.buffer(INSET_BUFFER_METERS)
_inset_final = [orig if small.is_empty else small
                for small, orig in zip(_inset_buffered, _inset_web.geometry)]
_inset_ll = gpd.GeoSeries(_inset_final, index=_inset_web.index, crs=3857).to_crs(epsg=4326)
TRACT_INSET_GEOMS = {geoid: mapping(geom) for geoid, geom in zip(all_tracts["GEOID"], _inset_ll)}
TRACT_INSET_GEOMS_JSON = json.dumps(TRACT_INSET_GEOMS)

# Maps the friendly column names shown in the results table to the
# underlying DataFrame columns.
QUERY_COLUMNS = {
    "GEOID": "GEOID",
    "Parish": "Parish",
    "Tract": "Geography",
    "Per Capita Income ($)": "2024 Per Capita Money Income (5-Year ACS)",
    "Income vs US Avg (%)": "Threshold Calculation2",
    "Unemployment Rate (%)": "2024 Unemployment Rate (5-Year ACS)",
    "Unemployment vs US Avg (pct pts)": "Threshold Calculation",
    "Distress Status": "Distress Status",
}

# Metrics the AI can sort/filter by. Use the parsed numeric columns for income
# and unemployment (the raw ACS columns are text like "$32,034 ", which breaks
# nsmallest/nlargest sorting).
METRIC_COLUMNS = {
    "per_capita_income": "income_value",
    "income_pct_us_avg": "Threshold Calculation2",
    "unemployment_rate": "unemployment_value",
    "unemployment_pct_pts_us_avg": "Threshold Calculation",
}
METRIC_LABELS = {
    "per_capita_income": "Per Capita Income",
    "income_pct_us_avg": "Per Capita Income (% of U.S. Average)",
    "unemployment_rate": "Unemployment Rate (%)",
    "unemployment_pct_pts_us_avg": "Unemployment Rate vs. U.S. Average (pct pts)",
}

# LODES WAC job-density metrics: one per NAICS sector column (JOB_COLS), plus
# total jobs. Lets the AI query box sort/filter tracts by employment the same
# way it already does for income/unemployment.
JOB_METRIC_KEYS = [
    "total_jobs", "agriculture_jobs", "mining_jobs", "utilities_jobs",
    "construction_jobs", "manufacturing_jobs", "wholesale_trade_jobs",
    "retail_trade_jobs", "transportation_warehousing_jobs", "information_jobs",
    "finance_insurance_jobs", "real_estate_jobs", "professional_services_jobs",
    "management_jobs", "admin_support_jobs", "education_jobs",
    "health_care_jobs", "arts_entertainment_jobs", "food_accommodation_jobs",
    "other_services_jobs", "public_administration_jobs",
]
for _mk, _jc in zip(JOB_METRIC_KEYS, JOB_COLS):
    METRIC_COLUMNS[_mk] = _jc
    METRIC_LABELS[_mk] = JOB_SECTOR_LABELS[_jc] if _jc == "C000" else f"{JOB_SECTOR_LABELS[_jc]} Jobs"

# County Business Patterns (CBP) business-establishment metrics. Unlike income/
# unemployment/jobs, this data only exists at the parish level (no tract
# breakdown in the source), so it gets queried against BIZ_QUERY_DF -- a small
# one-row-per-parish frame -- instead of the tract-level all_tracts frame.
BIZ_METRIC_KEYS = [
    "total_establishments", "micro_enterprises", "small_businesses",
    "medium_businesses", "large_employers",
]
BIZ_METRIC_TO_COL = dict(zip(BIZ_METRIC_KEYS, BIZ_COLS))
for _bmk, _bc in zip(BIZ_METRIC_KEYS, BIZ_COLS):
    METRIC_COLUMNS[_bmk] = _bc
    METRIC_LABELS[_bmk] = BIZ_METRIC_LABELS[_bc]

BIZ_QUERY_DF = pd.DataFrame([
    dict({"Parish": _p["label"]}, **BIZ_BY_FIPS.get(_p["fips"], {c: 0 for c in BIZ_COLS}))
    for _p in PARISHES
])

PARISH_LABELS = [p["label"] for p in PARISHES]
DISTRESS_LABELS = [d["label"] for d in DISTRESS_STATUSES if d["slug"] != "no_data"]

# ---------------------------------------------------------------------------
# 10. AI query interpreter (Groq / any OpenAI-compatible cloud LLM API)
# ---------------------------------------------------------------------------
LLM_API_KEY = os.environ.get("LLM_API_KEY") or os.environ.get("GROQ_API_KEY")
LLM_BASE_URL = os.environ.get("LLM_BASE_URL", "https://api.groq.com/openai/v1")
LLM_MODEL = os.environ.get("LLM_MODEL", "llama-3.3-70b-versatile")

llm_client = OpenAI(api_key=LLM_API_KEY, base_url=LLM_BASE_URL) if LLM_API_KEY else None

JOB_METRIC_LIST = ", ".join(f'"{k}"' for k in JOB_METRIC_KEYS)
BIZ_METRIC_LIST = ", ".join(f'"{k}"' for k in BIZ_METRIC_KEYS)

SYSTEM_PROMPT = f"""You are a conversational assistant for a census-tract economic \
distress dataset covering the Capital Region Planning Commission (CRPC) area in Louisiana. \
The dataset also includes LODES workplace-area-characteristics (WAC) job counts per tract, \
broken out by NAICS sector (2023 data), and County Business Patterns (CBP) business \
establishment counts broken out by employer size, available only at the PARISH level (11 \
parishes total; there is no tract-level breakdown for business data). You maintain full \
conversation memory across turns.

You will receive the entire conversation history. Using the MOST RECENT user message together \
with all prior context, produce a JSON query spec. When the user says things like "show me more", \
"now focus on Tangipahoa", "change to unemployment", "make it 20 instead", or "only distressed \
ones" — update the relevant field(s) from the previous query and carry everything else forward \
unchanged.

Output a JSON object ONLY (no prose, no markdown fences) with exactly these fields:
  "parish": one of {PARISH_LABELS}, or null to search the whole region.
  "metric": one of "per_capita_income", "income_pct_us_avg", "unemployment_rate", \
"unemployment_pct_pts_us_avg", one of the job/employment metrics: {JOB_METRIC_LIST}, or one of \
the business/establishment metrics: {BIZ_METRIC_LIST}. \
Use "total_jobs" for overall employment, or the matching sector metric when the user names an \
industry (e.g. "manufacturing jobs" -> "manufacturing_jobs", "health care jobs" -> \
"health_care_jobs", "construction jobs" -> "construction_jobs"). Use "total_establishments" for \
overall business/establishment counts, or the matching size-class metric when the user names one \
(e.g. "small businesses" -> "small_businesses", "micro enterprises" or "businesses with fewer \
than 10 employees" -> "micro_enterprises", "medium-sized businesses" -> "medium_businesses", \
"large employers" -> "large_employers"). Default "per_capita_income".
  "order": "lowest", "highest", or "both". Use "highest" by default for job/employment and \
business/establishment metrics unless the user asks for the fewest/lowest.
  "n": integer 1-252 (total tracts in the dataset), default 10. Business/establishment metrics \
only have 11 parishes to rank, so results are naturally capped there regardless of "n".
  "distress_status": one of {DISTRESS_LABELS}, or null. Distress status is a tract-level label — \
always use null when "metric" is a business/establishment metric.
  "reply": a single conversational sentence acknowledging what you are showing \
(e.g. "Here are the 5 lowest income tracts in East Baton Rouge Parish." or "Here are the 10 \
tracts with the most manufacturing jobs in Livingston Parish." or "Here are the parishes ranked \
by number of small businesses."). This is shown directly to the user as your chat reply.

Respond with ONLY the JSON object, nothing else."""


def interpret_prompt_with_history(history):
    """Send the full conversation history to the LLM and parse its JSON spec + reply."""
    if llm_client is None:
        raise RuntimeError(
            "No LLM API key configured. Set GROQ_API_KEY (or LLM_API_KEY) as an "
            "environment variable before running this app."
        )

    messages = [{"role": "system", "content": SYSTEM_PROMPT}]
    messages.extend(history)

    try:
        resp = llm_client.chat.completions.create(
            model=LLM_MODEL, messages=messages, temperature=0,
            response_format={"type": "json_object"},
        )
    except Exception:
        resp = llm_client.chat.completions.create(
            model=LLM_MODEL, messages=messages, temperature=0,
        )

    content = resp.choices[0].message.content.strip()
    try:
        spec = json.loads(content)
    except json.JSONDecodeError:
        match = re.search(r"\{.*\}", content, re.S)
        if not match:
            raise RuntimeError(f"Could not parse LLM response as JSON: {content[:200]}")
        spec = json.loads(match.group(0))

    reply = str(spec.pop("reply", "") or "").strip()

    if spec.get("metric") not in METRIC_COLUMNS:
        spec["metric"] = "per_capita_income"
    if spec.get("order") not in ("lowest", "highest", "both"):
        spec["order"] = "both"
    try:
        spec["n"] = max(1, min(252, int(spec.get("n") or 10)))
    except (TypeError, ValueError):
        spec["n"] = 10
    if spec.get("parish") not in PARISH_LABELS:
        spec["parish"] = None
    if spec.get("distress_status") not in DISTRESS_LABELS:
        spec["distress_status"] = None
    return spec, reply


def interpret_prompt(user_prompt):
    """Legacy single-turn wrapper kept for the /api/query endpoint."""
    spec, _ = interpret_prompt_with_history([{"role": "user", "content": user_prompt}])
    return spec


def run_business_query(spec):
    """County Business Patterns data is parish-level only -- query BIZ_QUERY_DF
    (one row per parish) instead of the tract-level all_tracts frame."""
    df = BIZ_QUERY_DF
    if spec["parish"]:
        df = df[df["Parish"] == spec["parish"]]

    metric_col = BIZ_METRIC_TO_COL[spec["metric"]]
    result_columns = {"Parish": "Parish"}
    for _bmk, _bc in zip(BIZ_METRIC_KEYS, BIZ_COLS):
        result_columns[METRIC_LABELS[_bmk]] = _bc
    display_cols = list(result_columns.values())

    def to_rows(frame):
        renamed = frame[display_cols].rename(columns={v: k for k, v in result_columns.items()})
        renamed = renamed.where(pd.notnull(renamed), None)
        return renamed.to_dict("records")

    n = spec["n"]
    metric_label = METRIC_LABELS[spec["metric"]]
    groups = []
    if spec["order"] in ("lowest", "both"):
        lowest = df.nsmallest(n, metric_col)
        groups.append({"title": f"Lowest {len(lowest)} Parish(es) by {metric_label}", "rows": to_rows(lowest)})
    if spec["order"] in ("highest", "both"):
        highest = df.nlargest(n, metric_col)
        groups.append({"title": f"Highest {len(highest)} Parish(es) by {metric_label}", "rows": to_rows(highest)})
    return groups


def run_query(spec):
    if spec["metric"] in BIZ_METRIC_KEYS:
        return run_business_query(spec)

    df = all_tracts
    if spec["parish"]:
        df = df[df["Parish"] == spec["parish"]]
    if spec["distress_status"]:
        df = df[df["Distress Status"] == spec["distress_status"]]

    metric_col = METRIC_COLUMNS[spec["metric"]]
    df = df.dropna(subset=[metric_col])

    # Job/employment metrics aren't among the fixed QUERY_COLUMNS, so add the
    # selected sector's job count as an extra column in the results table.
    result_columns = dict(QUERY_COLUMNS)
    if spec["metric"] in JOB_METRIC_KEYS:
        result_columns[METRIC_LABELS[spec["metric"]]] = metric_col
    display_cols = list(result_columns.values())

    def to_rows(frame):
        renamed = frame[display_cols].rename(columns={v: k for k, v in result_columns.items()})
        renamed = renamed.where(pd.notnull(renamed), None)
        return renamed.to_dict("records")

    n = spec["n"]
    metric_label = METRIC_LABELS[spec["metric"]]
    groups = []
    if spec["order"] in ("lowest", "both"):
        lowest = df.nsmallest(n, metric_col)
        groups.append({"title": f"Lowest {len(lowest)} by {metric_label}", "rows": to_rows(lowest)})
    if spec["order"] in ("highest", "both"):
        highest = df.nlargest(n, metric_col)
        groups.append({"title": f"Highest {len(highest)} by {metric_label}", "rows": to_rows(highest)})
    return groups


# ---------------------------------------------------------------------------
# 10b. "Show Tract Numbers" toggle — bottom-right
tract_num_panel_html = """
<div id="tract-num-panel" style="position:absolute;bottom:20px;right:10px;z-index:9999;
     background:#f3e8fd;padding:8px 12px;border:1px solid #888;border-radius:4px;font-size:13px;
     width:173px;box-sizing:border-box;">
  <label style="cursor:pointer;user-select:none;">
    <input type="checkbox" id="tract-num-enable"> Show Tract Numbers
  </label>
</div>
"""
m.get_root().html.add_child(folium.Element(tract_num_panel_html))

# 11. Flask app: serves the map plus the AI query box and results table
# ---------------------------------------------------------------------------
m.get_root().render()
MAP_HEADER = m.get_root().header.render()
MAP_BODY = m.get_root().html.render()
MAP_SCRIPT = m.get_root().script.render()
GEOID_BOUNDS_JSON = json.dumps(geoid_bounds)
PARISH_BOUNDS_JSON = json.dumps(parish_bounds)

search_data = []
for _, row in all_tracts.iterrows():
    geo = str(row.get("Geography") or row.get("NAMELSAD") or "")
    search_data.append({
        "geoid": row["GEOID"],
        "label": geo,
        "parish": str(row.get("Parish") or ""),
    })
SEARCH_DATA_JSON = json.dumps(search_data)

_mb = all_tracts.total_bounds  # [minx, miny, maxx, maxy]
DEFAULT_BOUNDS_JSON = json.dumps([[_mb[1], _mb[0]], [_mb[3], _mb[2]]])

PAGE_TEMPLATE = """
<!DOCTYPE html>
<html>
<head>
<meta charset="utf-8">
<title>CRPC Region - DRA Distress Status Visualization | CRPC AI Assistant</title>
{{ map_header | safe }}
<style>
  html, body { margin: 0; padding: 0; overflow-y: auto; }
  body { display: flex; flex-direction: row; min-height: 100vh; align-items: flex-start; }
  #sidebar {
      width: 22%; min-width: 308px; flex-shrink: 0; height: 100vh; box-sizing: border-box;
      display: flex; flex-direction: column; border-right: 2px solid #888; background: #e4e7eb;
      overflow-y: auto; position: sticky; top: 0;
  }
  #right-pane {
      flex: 1 1 auto; display: flex; flex-direction: column; min-height: 100vh;
  }
  #map-pane { position: relative; height: 100vh; flex-shrink: 0; overflow: hidden; }
  /* Shift only the Leaflet map view itself -- title, search bar, legend,
     filter panels etc. are separate overlay elements and stay put. */
  .folium-map {
      transform: translateX(-55px);
  }
  #map-pane:fullscreen, #map-pane:-webkit-full-screen {
      width: 100vw; height: 100vh; background: #fff;
  }
  #map-title {
      position: absolute; top: 8px; left: 50%; transform: translateX(calc(-50% + 70px));
      z-index: 9999; background: rgba(255,255,255,0.92); padding: 6px 14px;
      border: 1px solid #888; border-radius: 4px; font-size: 15px; font-weight: bold;
      color: #404040; box-shadow: 0 1px 4px rgba(0,0,0,0.3); white-space: nowrap;
      pointer-events: none; max-width: calc(100% - 20px); overflow: hidden;
      text-overflow: ellipsis;
  }
  #results-pane {
      flex-shrink: 0; padding: 12px 16px; border-top: 2px solid #888; background: #eafaf1;
  }
  #results-header {
      display: flex; align-items: center; gap: 10px; margin-bottom: 8px;
      font-size: 13px; font-weight: bold; color: #404040;
  }
  #scroll-top-btn {
      margin-left: auto; padding: 4px 10px; font-size: 11px; cursor: pointer;
      background: #1a4e8a; color: #fff; border: none; border-radius: 12px;
  }
  #scroll-top-btn:hover { background: #0d3566; }
  #query-pane { padding: 10px; border-top: 1px solid #ddd; flex-shrink: 0; background: #eafaf1; }
  #query-pane form { display: flex; flex-direction: column; gap: 8px; }
  #ai-section-title {
      font-size: 15px; font-weight: bold; color: #1a4e8a; margin-bottom: 8px;
      padding-bottom: 6px; border-bottom: 2px solid #1a4e8a; letter-spacing: 0.4px;
  }
  #chat-history {
      max-height: 220px; overflow-y: auto; padding: 8px;
      background: #f0f2f4; border: 1px solid #c8cdd3; border-radius: 6px;
      margin-bottom: 8px; display: flex; flex-direction: column; gap: 6px;
  }
  #chat-history:empty::before {
      content: "Your conversation will appear here.";
      color: #aaa; font-size: 11px; display: block; text-align: center; padding: 10px 0;
  }
  .chat-bubble {
      max-width: 88%; padding: 7px 11px; border-radius: 14px;
      font-size: 12px; line-height: 1.45; word-wrap: break-word;
  }
  .chat-bubble.user {
      align-self: flex-end; background: #1a4e8a; color: #fff; border-bottom-right-radius: 3px;
  }
  .chat-bubble.assistant {
      align-self: flex-start; background: #d4dae0; color: #1a1a1a; border-bottom-left-radius: 3px;
  }
  #query-input {
      padding: 8px; font-size: 14px; width: 100%; box-sizing: border-box;
      min-height: 72px; resize: vertical; white-space: pre-wrap; word-wrap: break-word;
      overflow-y: auto; line-height: 1.5; font-family: inherit;
  }
  .form-buttons { display: flex; gap: 8px; }
  #query-submit, #query-reset { flex: 1; padding: 8px; font-size: 14px; cursor: pointer; }
  #query-status { font-size: 12px; color: #555; margin-top: 6px; min-height: 16px; }
  .examples { font-size: 11px; color: #777; margin-top: 6px; }
  .sidebar-panel { flex-shrink: 0; padding: 10px 14px; border-top: 1px solid #ddd; font-size: 13px; line-height: 1.5; }
  #distress-filter-panel { background: #e8f4fd; }
  #threshold-filter-panel { background: #e8eaed; }
  #income-knob-panel     { background: #e8f4fd; }
  #unemp-knob-panel      { background: #e8eaed; }
  #jobs-filter-panel          { background: #e8f4fd; }
  #parish-jobs-filter-panel   { background: #e8eaed; }
  #business-filter-panel      { background: #e8f4fd; }
  .sidebar-panel hr { border: none; border-top: 1px solid #ddd; }
  .knob-head { display: flex; align-items: center; gap: 6px; cursor: pointer; }
  .knob-body { margin-top: 8px; opacity: 0.4; pointer-events: none; transition: opacity 0.15s; }
  .knob-body.active { opacity: 1; pointer-events: auto; }
  .knob-body input[type="range"] { width: 100%; margin: 0; display: block; }
  .knob-ticks { display: flex; justify-content: space-between; align-items: flex-start; padding: 0 2px; margin-top: 1px; height: 7px; }
  .knob-tick { width: 1px; height: 7px; background: #999; }
  .knob-scale { display: flex; justify-content: space-between; font-size: 9px; color: #666; margin-top: 1px; }
  .knob-current { font-size: 12px; color: #137333; font-weight: bold; margin-top: 6px; }
  /* Thermal gradient slider track: green (left) -> yellow (mid) -> dark red
     (right). Heat-map colors, distinct from the map's distress palette. */
  input[type="range"].heat {
      -webkit-appearance: none; appearance: none; width: 100%; height: 10px;
      border-radius: 5px; outline: none; cursor: pointer;
      background: linear-gradient(to right,
          #1a9850 0%, #66bd63 14%, #a6d96a 28%, #ffff33 50%,
          #fdae61 68%, #f46d43 82%, #d73027 92%, #a50026 100%);
  }
  input[type="range"].heat::-webkit-slider-runnable-track {
      height: 10px; border-radius: 5px; background: transparent;
  }
  input[type="range"].heat::-webkit-slider-thumb {
      -webkit-appearance: none; appearance: none; width: 16px; height: 16px;
      margin-top: -3px; border-radius: 50%; background: #fff;
      border: 2px solid #222; box-shadow: 0 0 2px rgba(0,0,0,0.6); cursor: pointer;
  }
  input[type="range"].heat::-moz-range-track {
      height: 10px; border-radius: 5px;
      background: linear-gradient(to right,
          #1a9850 0%, #66bd63 14%, #a6d96a 28%, #ffff33 50%,
          #fdae61 68%, #f46d43 82%, #d73027 92%, #a50026 100%);
  }
  input[type="range"].heat::-moz-range-thumb {
      width: 16px; height: 16px; border-radius: 50%; background: #fff;
      border: 2px solid #222; box-shadow: 0 0 2px rgba(0,0,0,0.6); cursor: pointer;
  }
  .result-table-wrap { overflow-x: auto; margin-bottom: 16px; }
  .result-table { width: 100%; border-collapse: collapse; font-size: 11px; }
  .result-table caption { text-align: left; font-weight: bold; padding: 6px 0; font-size: 12px; }
  .result-table th, .result-table td { border: 1px solid #ccc; padding: 3px 6px; text-align: right; white-space: nowrap; }
  .result-table th:nth-child(1), .result-table td:nth-child(1),
  .result-table th:nth-child(2), .result-table td:nth-child(2) { text-align: left; }
  .result-table tbody tr:hover { background: #eef6ff; cursor: pointer; }
  #tract-search-bar {
      position: absolute; top: 8px; right: calc(50% + 240px);
      z-index: 9999; width: 300px;
  }
  #tract-search-input {
      width: 100%; padding: 8px 34px 8px 34px; font-size: 13px; box-sizing: border-box;
      border: 1px solid #aaa; border-radius: 20px; outline: none;
      box-shadow: 0 2px 8px rgba(0,0,0,0.18); background: rgba(255,255,255,0.97);
  }
  #tract-search-input:focus { border-color: #1a4e8a; box-shadow: 0 2px 10px rgba(26,78,138,0.25); }
  #tract-search-clear {
      display: none; position: absolute; right: 10px; top: 50%; transform: translateY(-50%);
      background: #bbb; color: #fff; border: none; border-radius: 50%;
      width: 18px; height: 18px; font-size: 11px; line-height: 18px; text-align: center;
      cursor: pointer; padding: 0;
  }
  #tract-search-clear:hover { background: #888; }
  #tract-search-icon {
      position: absolute; left: 11px; top: 50%; transform: translateY(-50%);
      font-size: 14px; color: #888; pointer-events: none;
  }
  #tract-search-dropdown {
      display: none; position: absolute; top: calc(100% + 4px); left: 0; right: 0;
      background: #fff; border: 1px solid #ccc; border-radius: 8px;
      box-shadow: 0 4px 12px rgba(0,0,0,0.15); max-height: 220px; overflow-y: auto;
  }
  .search-item {
      padding: 8px 14px; cursor: pointer; border-bottom: 1px solid #f0f0f0;
  }
  .search-item:last-child { border-bottom: none; }
  .search-item:hover, .search-item.active { background: #eef6ff; }
  .search-item-label { font-size: 12px; font-weight: bold; color: #1a1a1a; }
  .search-item-sub { font-size: 11px; color: #777; margin-top: 1px; }
  .search-no-results { padding: 10px 14px; font-size: 12px; color: #999; }
  .result-table tbody tr.selected { background: #fff3cd; outline: 2px solid #FFB347; outline-offset: -2px; }
  .prompt-suggestions { display: flex; flex-direction: column; gap: 5px; margin-top: 8px; }
  .prompt-chip {
      background: #fff; border: 1px solid #1a4e8a; border-radius: 14px;
      color: #1a4e8a; font-size: 11px; padding: 5px 10px; cursor: pointer;
      text-align: left; line-height: 1.4; transition: background 0.15s, color 0.15s;
  }
  .prompt-chip:hover { background: #1a4e8a; color: #fff; }
  .more-prompts-toggle {
      background: none; border: none; color: #1a4e8a; font-size: 11px; font-weight: bold;
      cursor: pointer; padding: 4px 0; text-align: left; margin-top: 6px; width: 100%;
  }
  .more-prompts-toggle:hover { text-decoration: underline; }
  .more-prompts-panel {
      display: none; max-height: 260px; overflow-y: auto; margin-top: 6px;
      padding: 8px; background: #fff; border: 1px solid #c8cdd3; border-radius: 6px;
  }
  .more-prompts-category {
      font-size: 11px; font-weight: bold; color: #404040; margin: 8px 0 4px;
  }
  .more-prompts-category:first-child { margin-top: 0; }
  .more-prompts-group { display: flex; flex-direction: column; gap: 4px; }
  .prompt-chip.more-chip { font-size: 10.5px; padding: 4px 8px; }
</style>
</head>
<body>
  <div id="sidebar">
    {{ filters_html | safe }}

    <div id="query-pane">
      <div id="ai-section-title">Ask CRPC AI Assistance</div>
      <div id="chat-history"></div>
      <form id="query-form">
        <textarea id="query-input" rows="3" placeholder="Ask me something or click a suggestion below..." autocomplete="off"></textarea>
        <div class="prompt-suggestions">
          <button type="button" class="prompt-chip">Show me the 5 lowest income tracts of East Baton Rouge Parish</button>
          <button type="button" class="prompt-chip">Show the highest unemployment tracts in Tangipahoa Parish</button>
          <button type="button" class="prompt-chip">Show tracts distressed by both unemployment and income in Ascension Parish</button>
          <button type="button" class="prompt-chip">Top 10 most distressed tracts across the entire CRPC region</button>
          <button type="button" class="prompt-chip">Which tracts have the most manufacturing jobs in Livingston Parish?</button>
        </div>
        <div class="form-buttons" style="margin-top:8px;">
          <button type="submit" id="query-submit">Ask</button>
          <button type="button" id="query-reset">Reset</button>
        </div>
        <button type="button" id="more-prompts-toggle" class="more-prompts-toggle">&#9656; Show more suggested prompts</button>
        <div id="more-prompts-panel" class="more-prompts-panel">
          <div class="more-prompts-category">Per Capita Income</div>
          <div class="more-prompts-group">
            <button type="button" class="prompt-chip more-chip">Show me the 10 lowest per capita income tracts in the CRPC region</button>
            <button type="button" class="prompt-chip more-chip">Which tracts have the highest per capita income in East Baton Rouge Parish?</button>
            <button type="button" class="prompt-chip more-chip">Show the 5 lowest income tracts in Ascension Parish</button>
            <button type="button" class="prompt-chip more-chip">What are the top 10 tracts by per capita income across the whole region?</button>
            <button type="button" class="prompt-chip more-chip">Show me tracts where per capita income is below 60% of the U.S. average</button>
            <button type="button" class="prompt-chip more-chip">Which tracts in Tangipahoa Parish have the lowest per capita income?</button>
            <button type="button" class="prompt-chip more-chip">Show the highest income tracts that are also distressed by unemployment</button>
            <button type="button" class="prompt-chip more-chip">List the 15 lowest per capita income tracts in Livingston Parish</button>
            <button type="button" class="prompt-chip more-chip">Show me tracts with the lowest per capita income relative to the U.S. average</button>
            <button type="button" class="prompt-chip more-chip">Compare the highest and lowest per capita income tracts in Washington Parish</button>
          </div>
          <div class="more-prompts-category">Unemployment Rate</div>
          <div class="more-prompts-group">
            <button type="button" class="prompt-chip more-chip">Show me the 10 highest unemployment rate tracts in the CRPC region</button>
            <button type="button" class="prompt-chip more-chip">Which tracts in Ascension Parish have the lowest unemployment rate?</button>
            <button type="button" class="prompt-chip more-chip">Top 5 tracts by unemployment rate in East Baton Rouge Parish</button>
            <button type="button" class="prompt-chip more-chip">Show me tracts with the highest unemployment rate compared to the U.S. average</button>
            <button type="button" class="prompt-chip more-chip">Show me the lowest unemployment rate tracts in Livingston Parish</button>
            <button type="button" class="prompt-chip more-chip">Which tracts are distressed by unemployment with the highest rates?</button>
            <button type="button" class="prompt-chip more-chip">Compare the highest and lowest unemployment tracts in Tangipahoa Parish</button>
            <button type="button" class="prompt-chip more-chip">Show me 20 tracts with the lowest unemployment rate in the region</button>
            <button type="button" class="prompt-chip more-chip">Which tracts in Pointe Coupee Parish have the highest unemployment rate?</button>
            <button type="button" class="prompt-chip more-chip">Show the 10 tracts with the highest unemployment rate that are also distressed by both income and unemployment</button>
          </div>
          <div class="more-prompts-category">Job &amp; Employment Data</div>
          <div class="more-prompts-group">
            <button type="button" class="prompt-chip more-chip">Which tracts have the most total jobs in the CRPC region?</button>
            <button type="button" class="prompt-chip more-chip">Show me the highest manufacturing jobs tracts in Livingston Parish</button>
            <button type="button" class="prompt-chip more-chip">Top 10 tracts by retail trade jobs in Tangipahoa Parish</button>
            <button type="button" class="prompt-chip more-chip">Which tracts have the most health care jobs in East Baton Rouge Parish?</button>
            <button type="button" class="prompt-chip more-chip">Show me the highest construction jobs tracts in the region</button>
            <button type="button" class="prompt-chip more-chip">Which tracts have the most agricultural jobs?</button>
            <button type="button" class="prompt-chip more-chip">Top 5 tracts for mining jobs in Iberville Parish</button>
            <button type="button" class="prompt-chip more-chip">Show me the highest finance and insurance jobs tracts</button>
            <button type="button" class="prompt-chip more-chip">Which tracts have the most educational services jobs in Ascension Parish?</button>
            <button type="button" class="prompt-chip more-chip">Show the top 10 tracts by professional services jobs</button>
            <button type="button" class="prompt-chip more-chip">Which tracts have the highest food and accommodation jobs?</button>
            <button type="button" class="prompt-chip more-chip">Show me the tracts with the most public administration jobs</button>
            <button type="button" class="prompt-chip more-chip">Top 10 tracts by wholesale trade jobs in the CRPC region</button>
            <button type="button" class="prompt-chip more-chip">Which tracts have the most transportation and warehousing jobs?</button>
            <button type="button" class="prompt-chip more-chip">Show me the highest real estate jobs tracts in West Baton Rouge Parish</button>
            <button type="button" class="prompt-chip more-chip">Which tracts have the most information sector jobs?</button>
            <button type="button" class="prompt-chip more-chip">Top 10 tracts by arts and entertainment jobs</button>
            <button type="button" class="prompt-chip more-chip">Show me the tracts with the most utilities jobs</button>
            <button type="button" class="prompt-chip more-chip">Which tracts have the highest management jobs in East Baton Rouge Parish?</button>
            <button type="button" class="prompt-chip more-chip">Show me the top 10 tracts by total jobs that are also non-distressed</button>
          </div>
          <div class="more-prompts-category">Business &amp; Establishments</div>
          <div class="more-prompts-group">
            <button type="button" class="prompt-chip more-chip">Which parish has the most business establishments in the CRPC region?</button>
            <button type="button" class="prompt-chip more-chip">Rank all parishes by total business establishments</button>
            <button type="button" class="prompt-chip more-chip">Which parish has the most micro enterprises with fewer than 10 employees?</button>
            <button type="button" class="prompt-chip more-chip">Show me the parishes with the fewest small businesses</button>
            <button type="button" class="prompt-chip more-chip">Which parish has the most medium-sized businesses?</button>
            <button type="button" class="prompt-chip more-chip">Which parish has the most large employers with 500 or more employees?</button>
            <button type="button" class="prompt-chip more-chip">Which parish has the fewest large employers?</button>
            <button type="button" class="prompt-chip more-chip">Show me the business establishment breakdown for East Baton Rouge Parish</button>
            <button type="button" class="prompt-chip more-chip">How many total business establishments are in Livingston Parish?</button>
            <button type="button" class="prompt-chip more-chip">Compare small businesses across all 11 CRPC parishes</button>
            <button type="button" class="prompt-chip more-chip">Which parish has the highest number of small businesses with 10 to 99 employees?</button>
            <button type="button" class="prompt-chip more-chip">Show me the parish with the lowest number of business establishments</button>
            <button type="button" class="prompt-chip more-chip">Which parishes have the most micro enterprises?</button>
            <button type="button" class="prompt-chip more-chip">Rank the parishes by number of large employers</button>
          </div>
        </div>
      </form>
      <div id="query-status"></div>
    </div>

    <div id="jobs-filter-panel" class="sidebar-panel knob-panel">
      <label class="knob-head">
        <input type="checkbox" class="knob-enable" id="jobs-enable"> <b>Job Density Overlay Per Tract</b>
      </label>
      <div class="knob-body" id="jobs-body">
        <div id="jobs-sector-grid">
          <div class="jobs-sector-item jobs-selected" data-key="C000" style="grid-column:1/-1;" title="Total Jobs">💼 Total Jobs</div>
          <div class="jobs-sector-item" data-key="CNS01" title="Agriculture, Forestry, Fishing &amp; Hunting">🌾 Agriculture</div>
          <div class="jobs-sector-item" data-key="CNS02" title="Mining, Quarrying &amp; Oil and Gas Extraction">⛏️ Mining</div>
          <div class="jobs-sector-item" data-key="CNS03" title="Utilities">⚡ Utilities</div>
          <div class="jobs-sector-item" data-key="CNS04" title="Construction">🏗️ Construction</div>
          <div class="jobs-sector-item" data-key="CNS05" title="Manufacturing">🏭 Manufacturing</div>
          <div class="jobs-sector-item" data-key="CNS06" title="Wholesale Trade">📦 Wholesale</div>
          <div class="jobs-sector-item" data-key="CNS07" title="Retail Trade">🛒 Retail</div>
          <div class="jobs-sector-item" data-key="CNS08" title="Transportation and Warehousing">🚚 Transport</div>
          <div class="jobs-sector-item" data-key="CNS09" title="Information">💻 Information</div>
          <div class="jobs-sector-item" data-key="CNS10" title="Finance and Insurance">🏦 Finance</div>
          <div class="jobs-sector-item" data-key="CNS11" title="Real Estate and Rental and Leasing">🏠 Real Estate</div>
          <div class="jobs-sector-item" data-key="CNS12" title="Professional, Scientific &amp; Technical Services">🔬 Professional</div>
          <div class="jobs-sector-item" data-key="CNS13" title="Management of Companies and Enterprises">🏢 Management</div>
          <div class="jobs-sector-item" data-key="CNS14" title="Admin., Support &amp; Waste Management">🧹 Admin/Support</div>
          <div class="jobs-sector-item" data-key="CNS15" title="Educational Services">🎓 Education</div>
          <div class="jobs-sector-item" data-key="CNS16" title="Health Care and Social Assistance">🏥 Health Care</div>
          <div class="jobs-sector-item" data-key="CNS17" title="Arts, Entertainment &amp; Recreation">🎭 Arts &amp; Entmt.</div>
          <div class="jobs-sector-item" data-key="CNS18" title="Accommodation and Food Services">🍽️ Food/Accom.</div>
          <div class="jobs-sector-item" data-key="CNS19" title="Other Services (excl. Public Administration)">🔧 Other Svcs.</div>
          <div class="jobs-sector-item" data-key="CNS20" title="Public Administration">🏛️ Public Admin.</div>
        </div>
        <div id="jobs-sector-name" style="font-size:11px;color:#1a4e8a;font-weight:bold;margin-top:5px;margin-bottom:4px;min-height:14px;word-wrap:break-word;">Total Jobs</div>
        <div id="jobs-legend" style="margin-bottom:6px;">
          <div style="height:10px;border-radius:4px;background:linear-gradient(to right,hsl(0,82%,44%),hsl(30,82%,44%),hsl(60,82%,44%),hsl(90,82%,44%),hsl(120,82%,44%));"></div>
          <div style="display:flex;justify-content:space-between;font-size:9px;color:#666;margin-top:2px;">
            <span id="jobs-leg-min">0</span><span style="color:#aaa;">Low → High</span><span id="jobs-leg-max">0</span>
          </div>
        </div>
        <div class="knob-current" id="jobs-status-label">Unselected</div>
        <label id="jobs-show-nums-wrap" style="font-size:12px;cursor:pointer;display:block;margin-top:5px;">
          <input type="checkbox" id="jobs-show-nums"> Show Numbers
        </label>
      </div>
    </div>

    <div id="parish-jobs-filter-panel" class="sidebar-panel knob-panel">
      <label class="knob-head">
        <input type="checkbox" class="knob-enable" id="pjobs-enable"> <b>Job Density Overlay Per Parish</b>
      </label>
      <div class="knob-body" id="pjobs-body">
        <div id="pjobs-sector-grid">
          <div class="pjobs-sector-item pjobs-selected" data-key="C000" style="grid-column:1/-1;" title="Total Jobs">💼 Total Jobs</div>
          <div class="pjobs-sector-item" data-key="CNS01" title="Agriculture, Forestry, Fishing &amp; Hunting">🌾 Agriculture</div>
          <div class="pjobs-sector-item" data-key="CNS02" title="Mining, Quarrying &amp; Oil and Gas Extraction">⛏️ Mining</div>
          <div class="pjobs-sector-item" data-key="CNS03" title="Utilities">⚡ Utilities</div>
          <div class="pjobs-sector-item" data-key="CNS04" title="Construction">🏗️ Construction</div>
          <div class="pjobs-sector-item" data-key="CNS05" title="Manufacturing">🏭 Manufacturing</div>
          <div class="pjobs-sector-item" data-key="CNS06" title="Wholesale Trade">📦 Wholesale</div>
          <div class="pjobs-sector-item" data-key="CNS07" title="Retail Trade">🛒 Retail</div>
          <div class="pjobs-sector-item" data-key="CNS08" title="Transportation and Warehousing">🚚 Transport</div>
          <div class="pjobs-sector-item" data-key="CNS09" title="Information">💻 Information</div>
          <div class="pjobs-sector-item" data-key="CNS10" title="Finance and Insurance">🏦 Finance</div>
          <div class="pjobs-sector-item" data-key="CNS11" title="Real Estate and Rental and Leasing">🏠 Real Estate</div>
          <div class="pjobs-sector-item" data-key="CNS12" title="Professional, Scientific &amp; Technical Services">🔬 Professional</div>
          <div class="pjobs-sector-item" data-key="CNS13" title="Management of Companies and Enterprises">🏢 Management</div>
          <div class="pjobs-sector-item" data-key="CNS14" title="Admin., Support &amp; Waste Management">🧹 Admin/Support</div>
          <div class="pjobs-sector-item" data-key="CNS15" title="Educational Services">🎓 Education</div>
          <div class="pjobs-sector-item" data-key="CNS16" title="Health Care and Social Assistance">🏥 Health Care</div>
          <div class="pjobs-sector-item" data-key="CNS17" title="Arts, Entertainment &amp; Recreation">🎭 Arts &amp; Entmt.</div>
          <div class="pjobs-sector-item" data-key="CNS18" title="Accommodation and Food Services">🍽️ Food/Accom.</div>
          <div class="pjobs-sector-item" data-key="CNS19" title="Other Services (excl. Public Administration)">🔧 Other Svcs.</div>
          <div class="pjobs-sector-item" data-key="CNS20" title="Public Administration">🏛️ Public Admin.</div>
        </div>
        <div id="pjobs-sector-name" style="font-size:11px;color:#1a4e8a;font-weight:bold;margin-top:5px;margin-bottom:4px;min-height:14px;word-wrap:break-word;">Total Jobs</div>
        <div id="pjobs-legend" style="margin-bottom:6px;">
          <div style="height:10px;border-radius:4px;background:linear-gradient(to right,hsl(0,82%,44%),hsl(30,82%,44%),hsl(60,82%,44%),hsl(90,82%,44%),hsl(120,82%,44%));"></div>
          <div style="display:flex;justify-content:space-between;font-size:9px;color:#666;margin-top:2px;">
            <span id="pjobs-leg-min">0</span><span style="color:#aaa;">Low → High</span><span id="pjobs-leg-max">0</span>
          </div>
        </div>
        <div class="knob-current" id="pjobs-status-label">Unselected</div>
        <label id="pjobs-show-nums-wrap" style="font-size:12px;cursor:pointer;display:block;margin-top:5px;">
          <input type="checkbox" id="pjobs-show-nums"> Show Numbers
        </label>
      </div>
    </div>

    <div id="business-filter-panel" class="sidebar-panel knob-panel">
      <label class="knob-head">
        <input type="checkbox" class="knob-enable" id="biz-enable"> <b>Business Establishments Per Parish</b>
      </label>
      <div class="knob-body" id="biz-body">
        <div id="biz-metric-grid">
          <div class="biz-metric-item biz-selected" data-key="est" style="grid-column:1/-1;" title="Total number of business establishments">🏪 Total Establishments</div>
          <div class="biz-metric-item" data-key="n_micro_enterprise" title="Establishments with fewer than 10 employees">🏠 Micro Enterprises</div>
          <div class="biz-metric-item" data-key="n_small_business" title="Establishments with 10-99 employees">🏬 Small Business</div>
          <div class="biz-metric-item" data-key="n_medium_business" title="Establishments with 100-499 employees">🏭 Medium Business</div>
          <div class="biz-metric-item" data-key="n_large_employees" title="Establishments with 500+ employees">🏢 Large Employers</div>
        </div>
        <div id="biz-metric-name" style="font-size:11px;color:#1a4e8a;font-weight:bold;margin-top:5px;margin-bottom:4px;min-height:14px;word-wrap:break-word;">Total Establishments</div>
        <div id="biz-legend" style="margin-bottom:6px;">
          <div style="height:10px;border-radius:4px;background:linear-gradient(to right,hsl(0,82%,44%),hsl(30,82%,44%),hsl(60,82%,44%),hsl(90,82%,44%),hsl(120,82%,44%));"></div>
          <div style="display:flex;justify-content:space-between;font-size:9px;color:#666;margin-top:2px;">
            <span id="biz-leg-min">0</span><span style="color:#aaa;">Low → High</span><span id="biz-leg-max">0</span>
          </div>
        </div>
        <div class="knob-current" id="biz-status-label">Unselected</div>
        <label id="biz-show-nums-wrap" style="font-size:12px;cursor:pointer;display:block;margin-top:5px;">
          <input type="checkbox" id="biz-show-nums"> Show Numbers
        </label>
      </div>
    </div>
  </div>

<style>
  #jobs-sector-grid, #pjobs-sector-grid, #biz-metric-grid {
      display: grid; grid-template-columns: 1fr 1fr; gap: 3px; margin-bottom: 4px;
      pointer-events: auto;
  }
  .jobs-sector-item, .pjobs-sector-item, .biz-metric-item {
      padding: 4px 5px; font-size: 11px; border: 1px solid #c8cdd3; border-radius: 3px;
      cursor: pointer; background: #fff; text-align: center; user-select: none;
      white-space: nowrap; overflow: hidden; text-overflow: ellipsis;
      transition: background 0.12s, color 0.12s;
  }
  .jobs-sector-item:hover, .pjobs-sector-item:hover, .biz-metric-item:hover { background: #ddeeff; border-color: #1a4e8a; }
  .jobs-sector-item.jobs-selected { background: #1a4e8a; color: #fff; border-color: #0d3566; }
  .pjobs-sector-item.pjobs-selected { background: #1a4e8a; color: #fff; border-color: #0d3566; }
  .biz-metric-item.biz-selected { background: #1a4e8a; color: #fff; border-color: #0d3566; }
</style>
<script>
(function () {
    var JOB_DATA   = {{ job_data_json | safe }};
    var JOB_RANGES = {{ job_ranges_json | safe }};
    var JOB_LABELS = {{ job_labels_json | safe }};

    var enableCb  = document.getElementById("jobs-enable");
    var body      = document.getElementById("jobs-body");
    var nameDiv   = document.getElementById("jobs-sector-name");
    var statusLbl = document.getElementById("jobs-status-label");
    var legMin    = document.getElementById("jobs-leg-min");
    var legMax    = document.getElementById("jobs-leg-max");
    var gridItems = document.querySelectorAll(".jobs-sector-item");
    var showNumsCb = document.getElementById("jobs-show-nums");

    var currentSector = "C000";
    var jobLayer = null;

    function fmtNum(x) { return Math.round(x).toLocaleString(); }

    function markerSize(val, sectorMax) {
        if (sectorMax <= 0 || val <= 0) return 0;
        var t = Math.min(1, val / sectorMax);
        return Math.round(11 + t * 21);
    }

    function jobColor(val, sectorMax) {
        var t = (sectorMax > 0) ? Math.min(1, val / sectorMax) : 0;
        var hue = Math.round(t * 120); // 0 = crystal red, 120 = crystal green
        return 'hsla(' + hue + ',82%,44%,0.90)';
    }

    function getEmoji(sector) {
        var el = document.querySelector(".jobs-sector-item[data-key='" + sector + "']");
        return el ? el.textContent.trim().split(" ")[0] : "●";
    }

    function getActiveSlugs() {
        var active = new Set();
        document.querySelectorAll(".parish-filter-cb").forEach(function (cb) {
            if (cb.checked) active.add(cb.getAttribute("data-parish"));
        });
        return active;
    }

    function drawMarkers(sector) {
        var map = window.CRPC && window.CRPC.map;
        if (!map) return;
        if (!jobLayer) { jobLayer = L.layerGroup().addTo(map); }
        jobLayer.clearLayers();
        var r = JOB_RANGES[sector];
        var emoji = getEmoji(sector);
        var showNums = showNumsCb && showNumsCb.checked;
        var activeSlugs = getActiveSlugs();
        // Build per-parish max so each tract's gradient is relative to its own parish's highest tract
        var parishMax = {};
        JOB_DATA.forEach(function (d) {
            if (!activeSlugs.has(d.parish)) return;
            var v = d.jobs[sector] || 0;
            if (!(d.parish in parishMax) || v > parishMax[d.parish]) parishMax[d.parish] = v;
        });
        // Legend shows the overall min/max across selected parishes
        var dynamicMax = 0, dynamicMin = Infinity;
        JOB_DATA.forEach(function (d) {
            if (!activeSlugs.has(d.parish)) return;
            var v = d.jobs[sector] || 0;
            if (v > dynamicMax) dynamicMax = v;
            if (v > 0 && v < dynamicMin) dynamicMin = v;
        });
        if (dynamicMax === 0) { dynamicMax = r.max; dynamicMin = r.min; }
        if (dynamicMin === Infinity) dynamicMin = 0;
        legMin.textContent = fmtNum(dynamicMin);
        legMax.textContent = fmtNum(dynamicMax);
        JOB_DATA.forEach(function (d) {
            if (!activeSlugs.has(d.parish)) return;
            var val = d.jobs[sector] || 0;
            if (val <= 0) return;
            var pMax = parishMax[d.parish] || dynamicMax;
            var sz = markerSize(val, pMax);
            var numHtml = showNums
                ? '<div style="display:inline-block;font-size:8px;font-weight:bold;color:#fff;'
                  + 'background:' + jobColor(val, pMax) + ';border-radius:2px;'
                  + 'padding:0 2px;line-height:1.3;white-space:nowrap;'
                  + 'text-shadow:0 0 3px rgba(0,0,0,0.55);margin-top:1px;">'
                  + fmtNum(val) + '</div>'
                : '';
            var icon = L.divIcon({
                className: "",
                html: '<div style="display:inline-block;transform:translate(-50%,-50%);text-align:center;">'
                    + '<span style="font-size:' + sz + 'px;line-height:1;display:block;'
                    + 'filter:drop-shadow(0 0 2px rgba(255,255,255,0.85));" title="' + fmtNum(val) + ' jobs">'
                    + emoji + '</span>'
                    + numHtml
                    + '</div>',
                iconSize: [0, 0],
                iconAnchor: [0, 0]
            });
            L.marker([d.lat, d.lng], {icon: icon, interactive: false}).addTo(jobLayer);
        });
    }

    function clearMarkers() { if (jobLayer) jobLayer.clearLayers(); }

    function updateLegend(sector) {
        var r = JOB_RANGES[sector];
        legMin.textContent = fmtNum(r.min);
        legMax.textContent = fmtNum(r.max);
    }

    function selectSector(key) {
        currentSector = key;
        gridItems.forEach(function (el) {
            el.classList.toggle("jobs-selected", el.getAttribute("data-key") === key);
        });
        nameDiv.textContent = JOB_LABELS[key] || key;
        updateLegend(key);
        if (enableCb.checked) drawMarkers(key);
    }

    function updateJob() {
        body.classList.toggle("active", enableCb.checked);
        if (!enableCb.checked) {
            statusLbl.textContent = "Unselected";
            clearMarkers();
        } else {
            statusLbl.textContent = "Active";
            drawMarkers(currentSector);
        }
    }

    gridItems.forEach(function (el) {
        el.addEventListener("click", function () { selectSector(el.getAttribute("data-key")); });
    });
    enableCb.addEventListener("change", updateJob);
    showNumsCb.addEventListener("change", function () {
        if (enableCb.checked) drawMarkers(currentSector);
    });

    updateLegend("C000");

    document.addEventListener("DOMContentLoaded", function () {
        if (window.CRPC) {
            window.CRPC.redrawJobMarkers = function () {
                if (enableCb.checked) drawMarkers(currentSector);
            };
        }
    });
})();
</script>

<script>
(function () {
    var PARISH_JOB_DATA   = {{ parish_job_data_json | safe }};
    var PARISH_JOB_RANGES = {{ parish_job_ranges_json | safe }};
    var JOB_LABELS        = {{ job_labels_json | safe }};

    var enableCb   = document.getElementById("pjobs-enable");
    var body       = document.getElementById("pjobs-body");
    var nameDiv    = document.getElementById("pjobs-sector-name");
    var statusLbl  = document.getElementById("pjobs-status-label");
    var legMin     = document.getElementById("pjobs-leg-min");
    var legMax     = document.getElementById("pjobs-leg-max");
    var gridItems  = document.querySelectorAll(".pjobs-sector-item");
    var showNumsCb = document.getElementById("pjobs-show-nums");

    var currentSector = "C000";
    var parishJobLayer = null;

    function fmtNum(x) { return Math.round(x).toLocaleString(); }

    function markerSize(val, sectorMax) {
        if (sectorMax <= 0 || val <= 0) return 0;
        var t = Math.min(1, val / sectorMax);
        return Math.round(14 + t * 28);
    }

    function jobColor(val, sectorMax) {
        var t = (sectorMax > 0) ? Math.min(1, val / sectorMax) : 0;
        return 'hsla(' + Math.round(t * 120) + ',82%,44%,0.90)';
    }

    function getEmoji(sector) {
        var el = document.querySelector(".pjobs-sector-item[data-key='" + sector + "']");
        return el ? el.textContent.trim().split(" ")[0] : "●";
    }

    function getActiveSlugs() {
        var s = new Set();
        document.querySelectorAll(".parish-filter-cb").forEach(function (cb) {
            if (cb.checked) s.add(cb.getAttribute("data-parish"));
        });
        return s;
    }

    function drawMarkers(sector) {
        var map = window.CRPC && window.CRPC.map;
        if (!map) return;
        if (!parishJobLayer) { parishJobLayer = L.layerGroup().addTo(map); }
        parishJobLayer.clearLayers();
        var r = PARISH_JOB_RANGES[sector];
        var emoji = getEmoji(sector);
        var showNums = showNumsCb && showNumsCb.checked;
        var activeSlugs = getActiveSlugs();
        PARISH_JOB_DATA.forEach(function (d) {
            if (!activeSlugs.has(d.parish)) return;
            var val = d.jobs[sector] || 0;
            if (val <= 0) return;
            var sz = markerSize(val, r.max);
            var numHtml = showNums
                ? '<div style="display:inline-block;font-size:10px;font-weight:bold;color:#fff;'
                  + 'background:' + jobColor(val, r.max) + ';border-radius:3px;'
                  + 'padding:1px 5px;line-height:1.5;white-space:nowrap;'
                  + 'text-shadow:0 0 3px rgba(0,0,0,0.55);margin-top:2px;">'
                  + fmtNum(val) + '</div>'
                : '';
            var icon = L.divIcon({
                className: "",
                html: '<div style="display:inline-block;transform:translate(-50%,-50%);text-align:center;">'
                    + '<span style="font-size:' + sz + 'px;line-height:1;display:block;'
                    + 'filter:drop-shadow(0 0 2px rgba(255,255,255,0.85));" title="' + fmtNum(val) + ' jobs">'
                    + emoji + '</span>'
                    + numHtml + '</div>',
                iconSize: [0, 0],
                iconAnchor: [0, 0],
            });
            L.marker([d.lat, d.lng], {icon: icon, interactive: false}).addTo(parishJobLayer);
        });
    }

    function clearMarkers() { if (parishJobLayer) parishJobLayer.clearLayers(); }

    function updateLegend(sector) {
        var r = PARISH_JOB_RANGES[sector];
        legMin.textContent = fmtNum(r.min);
        legMax.textContent = fmtNum(r.max);
    }

    function selectSector(key) {
        currentSector = key;
        gridItems.forEach(function (el) {
            el.classList.toggle("pjobs-selected", el.getAttribute("data-key") === key);
        });
        nameDiv.textContent = JOB_LABELS[key] || key;
        updateLegend(key);
        if (enableCb.checked) drawMarkers(key);
    }

    function updateJob() {
        body.classList.toggle("active", enableCb.checked);
        if (!enableCb.checked) {
            statusLbl.textContent = "Unselected";
            clearMarkers();
        } else {
            statusLbl.textContent = "Active";
            drawMarkers(currentSector);
        }
    }

    gridItems.forEach(function (el) {
        el.addEventListener("click", function () { selectSector(el.getAttribute("data-key")); });
    });
    enableCb.addEventListener("change", updateJob);
    showNumsCb.addEventListener("change", function () {
        if (enableCb.checked) drawMarkers(currentSector);
    });

    updateLegend("C000");

    document.addEventListener("DOMContentLoaded", function () {
        if (window.CRPC) {
            window.CRPC.redrawParishJobMarkers = function () {
                if (enableCb.checked) drawMarkers(currentSector);
            };
        }
    });
})();
</script>

<script>
(function () {
    var PARISH_BIZ_DATA   = {{ parish_biz_data_json | safe }};
    var PARISH_BIZ_RANGES = {{ parish_biz_ranges_json | safe }};
    var BIZ_LABELS        = {{ biz_metric_labels_json | safe }};

    var enableCb   = document.getElementById("biz-enable");
    var body       = document.getElementById("biz-body");
    var nameDiv    = document.getElementById("biz-metric-name");
    var statusLbl  = document.getElementById("biz-status-label");
    var legMin     = document.getElementById("biz-leg-min");
    var legMax     = document.getElementById("biz-leg-max");
    var gridItems  = document.querySelectorAll(".biz-metric-item");
    var showNumsCb = document.getElementById("biz-show-nums");

    var currentMetric = "est";
    var bizLayer = null;

    function fmtNum(x) { return Math.round(x).toLocaleString(); }

    function markerSize(val, metricMax) {
        if (metricMax <= 0 || val <= 0) return 0;
        var t = Math.min(1, val / metricMax);
        return Math.round(14 + t * 28);
    }

    function bizColor(val, metricMax) {
        var t = (metricMax > 0) ? Math.min(1, val / metricMax) : 0;
        return 'hsla(' + Math.round(t * 120) + ',82%,44%,0.90)';
    }

    function getEmoji(metric) {
        var el = document.querySelector(".biz-metric-item[data-key='" + metric + "']");
        return el ? el.textContent.trim().split(" ")[0] : "●";
    }

    function getActiveSlugs() {
        var s = new Set();
        document.querySelectorAll(".parish-filter-cb").forEach(function (cb) {
            if (cb.checked) s.add(cb.getAttribute("data-parish"));
        });
        return s;
    }

    function drawMarkers(metric) {
        var map = window.CRPC && window.CRPC.map;
        if (!map) return;
        if (!bizLayer) { bizLayer = L.layerGroup().addTo(map); }
        bizLayer.clearLayers();
        var r = PARISH_BIZ_RANGES[metric];
        var emoji = getEmoji(metric);
        var showNums = showNumsCb && showNumsCb.checked;
        var activeSlugs = getActiveSlugs();
        PARISH_BIZ_DATA.forEach(function (d) {
            if (!activeSlugs.has(d.parish)) return;
            var val = d.biz[metric] || 0;
            if (val <= 0) return;
            var sz = markerSize(val, r.max);
            var numHtml = showNums
                ? '<div style="display:inline-block;font-size:10px;font-weight:bold;color:#fff;'
                  + 'background:' + bizColor(val, r.max) + ';border-radius:3px;'
                  + 'padding:1px 5px;line-height:1.5;white-space:nowrap;'
                  + 'text-shadow:0 0 3px rgba(0,0,0,0.55);margin-top:2px;">'
                  + fmtNum(val) + '</div>'
                : '';
            var icon = L.divIcon({
                className: "",
                html: '<div style="display:inline-block;transform:translate(-50%,-50%) translateY(16px);text-align:center;">'
                    + '<span style="font-size:' + sz + 'px;line-height:1;display:block;'
                    + 'filter:drop-shadow(0 0 2px rgba(255,255,255,0.85));" title="' + fmtNum(val) + '">'
                    + emoji + '</span>'
                    + numHtml + '</div>',
                iconSize: [0, 0],
                iconAnchor: [0, 0],
            });
            L.marker([d.lat, d.lng], {icon: icon, interactive: false}).addTo(bizLayer);
        });
    }

    function clearMarkers() { if (bizLayer) bizLayer.clearLayers(); }

    function updateLegend(metric) {
        var r = PARISH_BIZ_RANGES[metric];
        legMin.textContent = fmtNum(r.min);
        legMax.textContent = fmtNum(r.max);
    }

    function selectMetric(key) {
        currentMetric = key;
        gridItems.forEach(function (el) {
            el.classList.toggle("biz-selected", el.getAttribute("data-key") === key);
        });
        nameDiv.textContent = BIZ_LABELS[key] || key;
        updateLegend(key);
        if (enableCb.checked) drawMarkers(key);
    }

    function updateBiz() {
        body.classList.toggle("active", enableCb.checked);
        if (!enableCb.checked) {
            statusLbl.textContent = "Unselected";
            clearMarkers();
        } else {
            statusLbl.textContent = "Active";
            drawMarkers(currentMetric);
        }
    }

    gridItems.forEach(function (el) {
        el.addEventListener("click", function () { selectMetric(el.getAttribute("data-key")); });
    });
    enableCb.addEventListener("change", updateBiz);
    showNumsCb.addEventListener("change", function () {
        if (enableCb.checked) drawMarkers(currentMetric);
    });

    updateLegend("est");

    document.addEventListener("DOMContentLoaded", function () {
        if (window.CRPC) {
            window.CRPC.redrawBusinessMarkers = function () {
                if (enableCb.checked) drawMarkers(currentMetric);
            };
        }
    });
})();
</script>

  <div id="right-pane">
    <div id="map-pane">
      <div id="map-title">CRPC Region &ndash; DRA Distress Status Visualization and AI Assistance</div>
      <div id="tract-search-bar">
        <span id="tract-search-icon">&#128269;</span>
        <input type="text" id="tract-search-input" placeholder="Search tract number or name..." autocomplete="off" spellcheck="false">
        <button id="tract-search-clear" type="button" title="Clear search">&#10005;</button>
        <div id="tract-search-dropdown"></div>
      </div>
      {{ map_body | safe }}
    </div>

    <div id="results-pane">
      <div id="results-header" style="display:none;">
        Results
        <button id="scroll-top-btn" type="button" onclick="window.scrollTo({top:0,behavior:'smooth'})">&#8593; Back to map</button>
      </div>
      <p style="color:#777; font-size:12px;">Results will appear here after you ask a question above.</p>
    </div>
  </div>

<script>{{ map_script | safe }}</script>
<script>
(function () {
    var TRACT_POINTS = {{ tract_points_json | safe }};
    var enableCb    = document.getElementById("tract-num-enable");
    var labelLayer  = null;

    function tractNum(geoid) {
        var n     = parseInt(geoid.slice(5, 11), 10);
        var major = Math.floor(n / 100);
        var minor = n % 100;
        return major + "." + String(minor).padStart(2, "0");
    }

    function getActiveSlugs() {
        var s = new Set();
        document.querySelectorAll(".parish-filter-cb").forEach(function (cb) {
            if (cb.checked) s.add(cb.getAttribute("data-parish"));
        });
        return s;
    }

    function drawLabels() {
        var map = window.CRPC && window.CRPC.map;
        if (!map) return;
        if (!labelLayer) { labelLayer = L.layerGroup().addTo(map); }
        labelLayer.clearLayers();
        var active = getActiveSlugs();
        TRACT_POINTS.forEach(function (d) {
            if (!active.has(d.parish)) return;
            var icon = L.divIcon({
                className: "",
                html: '<div style="display:inline-block;transform:translate(-50%,-130%);'
                    + 'font-size:8px;font-weight:bold;color:#111;'
                    + 'background:rgba(255,255,255,0.88);padding:0 1px;'
                    + 'border-radius:1px;white-space:nowrap;pointer-events:none;line-height:1.2;">'
                    + tractNum(d.geoid) + '</div>',
                iconSize: [0, 0],
                iconAnchor: [0, 0],
            });
            L.marker([d.lat, d.lng], {icon: icon, interactive: false}).addTo(labelLayer);
        });
    }

    function clearLabels() { if (labelLayer) labelLayer.clearLayers(); }
    function update() { if (enableCb.checked) drawLabels(); else clearLabels(); }

    enableCb.addEventListener("change", update);

    document.addEventListener("DOMContentLoaded", function () {
        if (window.CRPC) {
            window.CRPC.redrawTractLabels = function () { if (enableCb.checked) drawLabels(); };
        }
    });
})();
</script>
<script>var GEOID_BOUNDS = {{ geoid_bounds_json | safe }};</script>
<script>var PARISH_BOUNDS = {{ parish_bounds_json | safe }};</script>
<script>var SEARCH_DATA = {{ search_data_json | safe }};</script>
<script>
(function () {
    var input    = document.getElementById("tract-search-input");
    var dropdown = document.getElementById("tract-search-dropdown");
    var clearBtn = document.getElementById("tract-search-clear");
    var activeIdx = -1;

    var DEFAULT_BOUNDS = {{ default_bounds_json | safe }};

    function showClear(v) { clearBtn.style.display = v ? "block" : "none"; }

    function resetSearch() {
        input.value = "";
        showClear(false);
        dropdown.style.display = "none";
        activeIdx = -1;
        if (!window.CRPC) return;
        window.CRPC.setSelectedGeoid(null);
        window.CRPC.refreshTracts();
        window.CRPC.map.fitBounds(DEFAULT_BOUNDS, {padding: [80, 80]});
    }

    clearBtn.addEventListener("click", resetSearch);

    function normalize(s) { return (s || "").toLowerCase(); }

    function score(item, q) {
        var g = normalize(item.geoid);
        var l = normalize(item.label);
        var p = normalize(item.parish);
        if (g === q || l === q) return 3;
        if (g.startsWith(q) || l.startsWith(q)) return 2;
        if (g.includes(q) || l.includes(q) || p.includes(q)) return 1;
        return 0;
    }

    function search(q) {
        q = normalize(q.trim());
        if (!q) return [];
        var results = [];
        for (var i = 0; i < SEARCH_DATA.length; i++) {
            var s = score(SEARCH_DATA[i], q);
            if (s > 0) results.push({item: SEARCH_DATA[i], score: s});
        }
        results.sort(function (a, b) { return b.score - a.score; });
        return results.slice(0, 12).map(function (r) { return r.item; });
    }

    function selectItem(item) {
        input.value = item.label || item.geoid;
        dropdown.style.display = "none";
        activeIdx = -1;

        if (!window.CRPC) return;
        window.CRPC.setSelectedGeoid(item.geoid);
        window.CRPC.refreshTracts();

        var pb = item.parish && PARISH_BOUNDS[item.parish];
        if (pb) {
            window.CRPC.map.fitBounds(pb, {padding: [24, 24]});
        } else {
            var b = GEOID_BOUNDS[item.geoid];
            if (b) window.CRPC.map.fitBounds(b, {maxZoom: 10});
        }
    }

    function renderDropdown(items) {
        dropdown.innerHTML = "";
        activeIdx = -1;
        if (!items.length) {
            var none = document.createElement("div");
            none.className = "search-no-results";
            none.textContent = "No tracts found.";
            dropdown.appendChild(none);
        } else {
            items.forEach(function (item, idx) {
                var div = document.createElement("div");
                div.className = "search-item";
                var label = document.createElement("div");
                label.className = "search-item-label";
                label.textContent = item.label || item.geoid;
                var sub = document.createElement("div");
                sub.className = "search-item-sub";
                sub.textContent = item.parish + "  ·  GEOID: " + item.geoid;
                div.appendChild(label);
                div.appendChild(sub);
                div.addEventListener("mousedown", function (e) {
                    e.preventDefault();
                    selectItem(item);
                });
                dropdown.appendChild(div);
            });
        }
        dropdown.style.display = "block";
    }

    input.addEventListener("input", function () {
        var q = input.value.trim();
        showClear(input.value.length > 0);
        if (!q) { dropdown.style.display = "none"; return; }
        renderDropdown(search(q));
    });

    input.addEventListener("keydown", function (e) {
        var items = dropdown.querySelectorAll(".search-item");
        if (e.key === "ArrowDown") {
            e.preventDefault();
            activeIdx = Math.min(activeIdx + 1, items.length - 1);
        } else if (e.key === "ArrowUp") {
            e.preventDefault();
            activeIdx = Math.max(activeIdx - 1, 0);
        } else if (e.key === "Enter") {
            e.preventDefault();
            if (activeIdx >= 0 && items[activeIdx]) items[activeIdx].dispatchEvent(new MouseEvent("mousedown"));
            return;
        } else if (e.key === "Escape") {
            dropdown.style.display = "none"; return;
        }
        items.forEach(function (el, i) {
            el.classList.toggle("active", i === activeIdx);
            if (i === activeIdx) el.scrollIntoView({block: "nearest"});
        });
    });

    document.addEventListener("click", function (e) {
        if (!document.getElementById("tract-search-bar").contains(e.target)) {
            dropdown.style.display = "none";
        }
    });
})();
</script>
<script>
document.addEventListener("DOMContentLoaded", function () {
    var highlightedGeoids = window.CRPC.highlightedGeoids;
    var selectedTr = null;
    var conversationHistory = [];

    var form = document.getElementById("query-form");
    var input = document.getElementById("query-input");
    var statusEl = document.getElementById("query-status");
    var resultsPane = document.getElementById("results-pane");
    var resetBtn = document.getElementById("query-reset");
    var chatHistoryEl = document.getElementById("chat-history");
    var suggestionsEl = document.querySelector(".prompt-suggestions");
    var morePromptsToggle = document.getElementById("more-prompts-toggle");
    var morePromptsPanel = document.getElementById("more-prompts-panel");

    function addChatBubble(role, text) {
        var bubble = document.createElement("div");
        bubble.className = "chat-bubble " + role;
        bubble.textContent = text;
        chatHistoryEl.appendChild(bubble);
        chatHistoryEl.scrollTop = chatHistoryEl.scrollHeight;
    }

    // Suggestion chips: fill textarea on click (covers both the always-visible
    // chips and the ones inside the "Show more suggested prompts" panel)
    document.querySelectorAll(".prompt-chip").forEach(function (chip) {
        chip.addEventListener("click", function () {
            input.value = chip.textContent.trim();
            input.focus();
        });
    });

    // "Show more suggested prompts" collapsible panel
    morePromptsToggle.addEventListener("click", function () {
        var expanded = morePromptsPanel.style.display === "block";
        morePromptsPanel.style.display = expanded ? "none" : "block";
        morePromptsToggle.innerHTML = (expanded ? "&#9656; Show more suggested prompts"
                                                 : "&#9662; Hide suggested prompts");
    });

    function buildTable(group) {
        var allCols = group.rows.length ? Object.keys(group.rows[0]) : [];
        var cols = allCols.filter(function (c) { return c !== "GEOID"; });
        var wrap = document.createElement("div");
        wrap.className = "result-table-wrap";
        var table = document.createElement("table");
        table.className = "result-table";

        var caption = document.createElement("caption");
        caption.textContent = group.title;
        table.appendChild(caption);

        var thead = document.createElement("thead");
        var headRow = document.createElement("tr");
        cols.forEach(function (c) {
            var th = document.createElement("th");
            th.textContent = c;
            headRow.appendChild(th);
        });
        thead.appendChild(headRow);
        table.appendChild(thead);

        var tbody = document.createElement("tbody");
        if (!group.rows.length) {
            var emptyRow = document.createElement("tr");
            var emptyCell = document.createElement("td");
            emptyCell.colSpan = cols.length || 1;
            emptyCell.textContent = "No matching tracts.";
            emptyRow.appendChild(emptyCell);
            tbody.appendChild(emptyRow);
        }
        group.rows.forEach(function (row) {
            var tr = document.createElement("tr");
            cols.forEach(function (c) {
                var td = document.createElement("td");
                var val = row[c];
                if (typeof val === "number") {
                    val = Math.round(val * 100) / 100;
                }
                td.textContent = (val === null || val === undefined) ? "" : val;
                tr.appendChild(td);
            });
            tr.addEventListener("click", function () {
                if (!row["GEOID"]) return;
                if (selectedTr) selectedTr.classList.remove("selected");
                tr.classList.add("selected");
                selectedTr = tr;
                window.CRPC.setSelectedGeoid(row["GEOID"]);
                window.CRPC.refreshTracts();
            });
            tbody.appendChild(tr);
        });
        table.appendChild(tbody);
        wrap.appendChild(table);
        return wrap;
    }

    var resultsHeader = document.getElementById("results-header");

    function renderResults(groups) {
        resultsPane.innerHTML = "";
        resultsPane.appendChild(resultsHeader);
        resultsHeader.style.display = "flex";
        highlightedGeoids.clear();
        groups.forEach(function (group) {
            group.rows.forEach(function (row) { if (row["GEOID"]) highlightedGeoids.add(row["GEOID"]); });
            resultsPane.appendChild(buildTable(group));
        });
        window.CRPC.refreshTracts();
        // Scroll results into view after a short delay so the DOM has painted
        setTimeout(function () {
            resultsPane.scrollIntoView({behavior: "smooth", block: "start"});
        }, 120);
    }

    form.addEventListener("submit", function (e) {
        e.preventDefault();
        var prompt = input.value.trim();
        if (!prompt) return;

        addChatBubble("user", prompt);
        conversationHistory.push({role: "user", content: prompt});
        input.value = "";

        // Hide the initial suggestion chips once the conversation starts, but keep the
        // "Show more suggested prompts" toggle/panel available throughout the chat so
        // the user can keep picking from the 40 examples mid-conversation.
        if (suggestionsEl) suggestionsEl.style.display = "none";

        statusEl.textContent = "Thinking...";

        fetch("/api/chat", {
            method: "POST",
            headers: {"Content-Type": "application/json"},
            body: JSON.stringify({history: conversationHistory})
        })
        .then(function (r) { return r.json(); })
        .then(function (data) {
            statusEl.textContent = "";
            if (data.error) {
                var errMsg = "Sorry, I ran into an error: " + data.error;
                addChatBubble("assistant", errMsg);
                conversationHistory.push({role: "assistant", content: errMsg});
                return;
            }
            var reply = data.reply || "Here are the results.";
            addChatBubble("assistant", reply);
            conversationHistory.push({role: "assistant", content: reply});
            renderResults(data.groups);
        })
        .catch(function (err) {
            statusEl.textContent = "";
            var errMsg = "Network error: " + err;
            addChatBubble("assistant", errMsg);
            conversationHistory.push({role: "assistant", content: errMsg});
        });
    });

    resetBtn.addEventListener("click", function () {
        input.value = "";
        statusEl.textContent = "";
        conversationHistory = [];
        chatHistoryEl.innerHTML = "";
        resultsHeader.style.display = "none";
        resultsPane.innerHTML = '<p style="color:#777; font-size:12px;">Results will appear here after you ask a question above.</p>';
        highlightedGeoids.clear();
        if (selectedTr) { selectedTr.classList.remove("selected"); selectedTr = null; }
        window.CRPC.setSelectedGeoid(null);
        window.CRPC.refreshTracts();
        if (suggestionsEl) suggestionsEl.style.display = "";
        if (morePromptsToggle) {
            morePromptsToggle.style.display = "";
            morePromptsToggle.innerHTML = "&#9656; Show more suggested prompts";
        }
        if (morePromptsPanel) morePromptsPanel.style.display = "none";
        window.scrollTo({top: 0, behavior: "smooth"});
    });
});
</script>
<script>
(function () {
    // Draws the "selected row" highlight as a ring a few meters inside the
    // tract's true boundary, instead of styling the tract's own edge --
    // tracts that sit on a parish's outer edge would otherwise have their
    // highlight drawn exactly on top of the parish's blue outline.
    var TRACT_INSET_GEOMS = {{ tract_inset_geoms_json | safe }};
    var insetLayer = null;

    function drawSelectedInset(geoid) {
        if (insetLayer) { window.CRPC.map.removeLayer(insetLayer); insetLayer = null; }
        var geom = geoid && TRACT_INSET_GEOMS[geoid];
        if (!geom) return;
        insetLayer = L.geoJSON(geom, {
            style: { color: "#FFB347", weight: 2.5, fill: false, opacity: 1 },
            interactive: false
        }).addTo(window.CRPC.map);
    }

    document.addEventListener("DOMContentLoaded", function () {
        if (!window.CRPC) return;
        var origSetSelectedGeoid = window.CRPC.setSelectedGeoid;
        window.CRPC.setSelectedGeoid = function (geoid) {
            origSetSelectedGeoid(geoid);
            drawSelectedInset(geoid);
        };
    });
})();
</script>
</body>
</html>
"""

app = Flask(__name__)


@app.route("/")
def index():
    return render_template_string(
        PAGE_TEMPLATE,
        map_header=MAP_HEADER,
        map_body=MAP_BODY,
        map_script=MAP_SCRIPT,
        geoid_bounds_json=GEOID_BOUNDS_JSON,
        parish_bounds_json=PARISH_BOUNDS_JSON,
        search_data_json=SEARCH_DATA_JSON,
        default_bounds_json=DEFAULT_BOUNDS_JSON,
        filters_html=filters_html,
        job_ranges_json=JOB_RANGES_JSON,
        job_data_json=JOB_DATA_JSON,
        job_labels_json=JOB_SECTOR_LABELS_JSON,
        tract_points_json=TRACT_POINTS_JSON,
        parish_job_data_json=PARISH_JOB_DATA_JSON,
        parish_job_ranges_json=PARISH_JOB_RANGES_JSON,
        parish_biz_data_json=PARISH_BIZ_DATA_JSON,
        parish_biz_ranges_json=PARISH_BIZ_RANGES_JSON,
        biz_metric_labels_json=BIZ_METRIC_LABELS_JSON,
        tract_inset_geoms_json=TRACT_INSET_GEOMS_JSON,
    )


@app.route("/api/query", methods=["POST"])
def api_query():
    data = request.get_json(silent=True) or {}
    user_prompt = (data.get("prompt") or "").strip()
    if not user_prompt:
        return jsonify({"error": "Empty prompt."}), 400

    try:
        spec = interpret_prompt(user_prompt)
        groups = run_query(spec)
    except Exception as exc:
        return jsonify({"error": str(exc)}), 500

    total = sum(len(g["rows"]) for g in groups)
    unit = "parish(es)" if spec["metric"] in BIZ_METRIC_KEYS else "tract(s)"
    summary_bits = []
    if spec["parish"]:
        summary_bits.append(spec["parish"])
    if spec["distress_status"]:
        summary_bits.append(spec["distress_status"])
    summary = f"Showing {total} {unit}"
    if summary_bits:
        summary += " -- " + ", ".join(summary_bits)

    return jsonify({"groups": groups, "summary": summary, "spec": spec})


@app.route("/api/chat", methods=["POST"])
def api_chat():
    data = request.get_json(silent=True) or {}
    history = data.get("history") or []
    if not any(m.get("role") == "user" for m in history):
        return jsonify({"error": "No user message found in history."}), 400

    try:
        spec, reply = interpret_prompt_with_history(history)
        groups = run_query(spec)
    except Exception as exc:
        return jsonify({"error": str(exc)}), 500

    total = sum(len(g["rows"]) for g in groups)
    if not reply:
        bits = []
        if spec["parish"]:
            bits.append(spec["parish"])
        if spec["distress_status"]:
            bits.append(spec["distress_status"])
        unit = "parish(es)" if spec["metric"] in BIZ_METRIC_KEYS else "tract(s)"
        reply = f"Showing {total} {unit}"
        if bits:
            reply += " — " + ", ".join(bits)

    return jsonify({"groups": groups, "reply": reply, "spec": spec})


# ---------------------------------------------------------------------------
# 12. Server-side static map render: draws the CRPC region to a large
#     high-resolution PNG with matplotlib + geopandas (vector tracts stay crisp
#     at any DPI), with parish outlines/labels, a CartoDB Positron basemap, and
#     a legend. No browser capture involved, so nothing shifts. The render
#     honors the current UI state posted from the browser: which parishes and
#     distress statuses are checked, the threshold mode, the AI-highlighted
#     tracts, and the current map view (zoom/pan) bounds.
# ---------------------------------------------------------------------------
# Reproject once to Web Mercator so the basemap tiles line up.
_static_web = all_tracts.to_crs(epsg=3857)
_static_outlines = [
    (parish["slug"], parish["label"], merged.dissolve()[["geometry"]].to_crs(epsg=3857))
    for parish, merged in zip(PARISHES, parish_frames)
]
_to_merc = Transformer.from_crs(4326, 3857, always_xy=True)


def _job_color_rgba(val, max_val):
    """Replicate JS jobColor: red (t=0) → green (t=1) via HSL hue 0→120."""
    t = min(1.0, val / max_val) if max_val > 0 else 0.0
    hue = t * 120.0 / 360.0
    r, g, b = colorsys.hls_to_rgb(hue, 0.44, 0.82)
    return (r, g, b, 0.90)


def _tract_num_str(geoid):
    n = int(geoid[5:11])
    return f"{n // 100}.{n % 100:02d}"


def _draw_job_overlay(ax, job_entries, sector, show_nums, active_slugs, merc_transform, pt_scale):
    """Draw job density markers (circle + optional number) on the matplotlib axes."""
    parish_max = {}
    for d in job_entries:
        if d["parish"] not in active_slugs:
            continue
        v = d["jobs"].get(sector, 0)
        if v > parish_max.get(d["parish"], 0):
            parish_max[d["parish"]] = v
    for d in job_entries:
        if d["parish"] not in active_slugs:
            continue
        val = d["jobs"].get(sector, 0)
        if val <= 0:
            continue
        pm = parish_max.get(d["parish"], 1) or 1
        t = min(1.0, val / pm)
        color = _job_color_rgba(val, pm)
        sz = pt_scale * (1 + t * 2)
        mx, my = merc_transform.transform(d["lng"], d["lat"])
        ax.plot(mx, my, "o", markersize=sz, color=color[:3], alpha=color[3],
                markeredgewidth=0.3, markeredgecolor="white", zorder=5)
        if show_nums:
            ax.annotate(
                f"{val:,}",
                xy=(mx, my), ha="center", va="bottom",
                fontsize=4, fontweight="bold", color="white",
                bbox=dict(boxstyle="square,pad=0.1", fc=color[:3], ec="none", alpha=0.90),
                zorder=6,
            )


def _draw_business_overlay(ax, biz_entries, metric, show_nums, active_slugs, merc_transform, pt_scale):
    """Draw business-establishment markers (circle + optional number) on the matplotlib axes."""
    parish_max = {}
    for d in biz_entries:
        if d["parish"] not in active_slugs:
            continue
        v = d["biz"].get(metric, 0)
        if v > parish_max.get(d["parish"], 0):
            parish_max[d["parish"]] = v
    for d in biz_entries:
        if d["parish"] not in active_slugs:
            continue
        val = d["biz"].get(metric, 0)
        if val <= 0:
            continue
        pm = parish_max.get(d["parish"], 1) or 1
        t = min(1.0, val / pm)
        color = _job_color_rgba(val, pm)
        sz = pt_scale * (1 + t * 2)
        mx, my = merc_transform.transform(d["lng"], d["lat"])
        ax.plot(mx, my, "o", markersize=sz, color=color[:3], alpha=color[3],
                markeredgewidth=0.3, markeredgecolor="white", zorder=5)
        if show_nums:
            ax.annotate(
                f"{val:,}",
                xy=(mx, my), ha="center", va="bottom",
                fontsize=4, fontweight="bold", color="white",
                bbox=dict(boxstyle="square,pad=0.1", fc=color[:3], ec="none", alpha=0.90),
                zorder=6,
            )


def _draw_filter_panel(pax, *, selected_labels, statuses, threshold,
                       income_enabled, income_bin,
                       unemployment_enabled, unemployment_bin, geoids):
    """Draw a sidebar-style summary of the active filters into axes `pax`."""
    pax.set_xlim(0, 1)
    pax.set_ylim(0, 1)
    pax.axis("off")
    pax.add_patch(mpatches.Rectangle((0, 0), 1, 1, facecolor="#fafafa",
                                     edgecolor="#888888", linewidth=1.5, clip_on=False))

    cur = {"y": 0.945}
    LEFT = 0.07
    LINE = 0.026

    def header(txt):
        cur["y"] -= LINE * 0.5
        pax.text(0.04, cur["y"], txt, fontsize=11, fontweight="bold", va="top")
        cur["y"] -= LINE * 1.25

    def item(txt, color=None, size=9, indent=0.0):
        x = LEFT + indent
        if color is not None:
            pax.add_patch(mpatches.Rectangle((x, cur["y"] - 0.016), 0.03, 0.016,
                                             facecolor=color, edgecolor="#777777",
                                             linewidth=0.6))
            pax.text(x + 0.05, cur["y"], txt, fontsize=size, va="top")
        else:
            pax.text(x, cur["y"], txt, fontsize=size, va="top")
        cur["y"] -= LINE

    pax.text(0.5, 0.99, "Active Filters", fontsize=15, fontweight="bold",
             ha="center", va="top")

    header("Filter by Parishes")
    if len(selected_labels) >= len(PARISHES):
        item("All parishes (%d)" % len(PARISHES))
    elif not selected_labels:
        item("None selected")
    else:
        for lbl in selected_labels:
            item("✓ " + lbl, size=8)

    header("Filter by Distress Status")
    for d in DISTRESS_STATUSES:
        checked = (statuses is None) or (d["label"] in statuses)
        item(("[x] " if checked else "[  ] ") + d["short"], color=d["color"], size=8)

    header("Filter by Threshold")
    tmap = {
        "all": "Show All",
        "income": "Per Capita Income < 60% U.S. Avg",
        "unemployment": "Unemployment ≥ 2 pct pts > U.S. Avg",
    }
    item(tmap.get(threshold, "Show All"), size=8)

    header("Filter by Per Capita Income")
    if income_enabled:
        lo = income_min + (income_bin - 1) * income_step
        hi = income_max if income_bin >= N_BINS else income_min + income_bin * income_step
        item("Showing ${:,.0f} – ${:,.0f}".format(lo, hi))
    else:
        item("Show all (off)")

    header("Filter by Unemployment Rate")
    if unemployment_enabled:
        lo = unemp_min + (unemployment_bin - 1) * unemp_step
        hi = unemp_max if unemployment_bin >= N_BINS else unemp_min + unemployment_bin * unemp_step
        item("Showing {:.1f}% – {:.1f}%".format(lo, hi))
    else:
        item("Show all (off)")

    if geoids:
        header("AI Query Highlight")
        item("%d tract(s) highlighted" % len(geoids))


def render_static_map(state, dpi=DOWNLOAD_DPI):
    state = state or {}

    # --- selections (default to "everything" when a key is absent) -----------
    parish_slugs = state.get("parishes")
    if parish_slugs is None:
        selected_labels = [p["label"] for p in PARISHES]
        selected_slugs = {p["slug"] for p in PARISHES}
    else:
        selected_slugs = set(parish_slugs)
        selected_labels = [p["label"] for p in PARISHES if p["slug"] in selected_slugs]

    statuses = state.get("statuses")
    threshold = state.get("threshold", "all")
    geoids = set(state.get("geoids") or [])
    bounds = state.get("bounds")  # [[south, west], [north, east]] in lat/lng

    # --- filter the tracts to match the on-screen view -----------------------
    # gdf_parish = all tracts from selected parishes (used for border-only rendering)
    gdf_parish = _static_web[_static_web["Parish"].isin(selected_labels)]
    gdf = gdf_parish.copy()
    if statuses is not None:
        gdf = gdf[gdf["Distress Status"].isin(statuses)]
    if threshold == "income":
        gdf = gdf[gdf["meets_income_threshold"] == True]  # noqa: E712
    elif threshold == "unemployment":
        gdf = gdf[gdf["meets_unemployment_threshold"] == True]  # noqa: E712

    # Range knobs: only applied when the matching checkbox was enabled.
    income_bin = int(state.get("income_bin") or 1)
    unemployment_bin = int(state.get("unemployment_bin") or 1)
    if state.get("income_enabled") and income_bin > 0:
        lo = income_min + (income_bin - 1) * income_step
        hi = income_max if income_bin >= N_BINS else income_min + income_bin * income_step
        vals = pd.to_numeric(gdf[INCOME_COL], errors="coerce")
        gdf = gdf[(vals >= lo) & (vals <= hi)]
    if state.get("unemployment_enabled") and unemployment_bin > 0:
        lo = unemp_min + (unemployment_bin - 1) * unemp_step
        hi = unemp_max if unemployment_bin >= N_BINS else unemp_min + unemployment_bin * unemp_step
        vals = pd.to_numeric(gdf[UNEMP_COL], errors="coerce")
        gdf = gdf[(vals >= lo) & (vals <= hi)]

    # Tracts in selected parishes that didn't pass the filters (border-only, like Leaflet)
    gdf_hidden = gdf_parish[~gdf_parish.index.isin(gdf.index)]

    # --- figure extent: current map view if provided, else region bounds -----
    if bounds:
        (south, west), (north, east) = bounds
        x0, y0 = _to_merc.transform(west, south)
        x1, y1 = _to_merc.transform(east, north)
        minx, maxx = sorted((x0, x1))
        miny, maxy = sorted((y0, y1))
    else:
        minx, miny, maxx, maxy = _static_web.total_bounds
        pad_x = (maxx - minx) * 0.02
        pad_y = (maxy - miny) * 0.02
        minx, maxx, miny, maxy = minx - pad_x, maxx + pad_x, miny - pad_y, maxy + pad_y

    span_x = maxx - minx
    span_y = maxy - miny
    aspect = span_x / span_y if span_y else 1.0
    height_in = 11.0
    width_in = max(6.0, min(22.0, height_in * aspect))

    panel_w = 3.6
    fig = Figure(figsize=(width_in + panel_w, height_in))
    gs = fig.add_gridspec(1, 2, width_ratios=[panel_w, width_in], wspace=0.015)
    pax = fig.add_subplot(gs[0, 0])
    ax = fig.add_subplot(gs[0, 1])

    _draw_filter_panel(
        pax,
        selected_labels=selected_labels,
        statuses=statuses,
        threshold=threshold,
        income_enabled=bool(state.get("income_enabled")),
        income_bin=income_bin,
        unemployment_enabled=bool(state.get("unemployment_enabled")),
        unemployment_bin=unemployment_bin,
        geoids=geoids,
    )

    # --- draw tracts: hidden first as lavender borders, then visible filled --------
    # Mirrors the Leaflet tractStyle: filtered-out tracts show lavender border + no fill
    if len(gdf_hidden):
        gdf_hidden.plot(ax=ax, facecolor="none", edgecolor="#c8a8f5", linewidth=0.4)

    if len(gdf):
        if geoids:
            hi = gdf[gdf["GEOID"].isin(geoids)]
            lo = gdf[~gdf["GEOID"].isin(geoids)]
            if len(lo):
                lo.plot(ax=ax, color=lo["fill_color"].tolist(),
                        edgecolor="#999999", linewidth=0.2, alpha=0.2)
            if len(hi):
                hi.plot(ax=ax, color=hi["fill_color"].tolist(),
                        edgecolor="#0033CC", linewidth=2.0)
        else:
            gdf.plot(ax=ax, color=gdf["fill_color"].tolist(),
                     edgecolor="#555555", linewidth=0.4)

    # --- parish outlines + labels, only for selected parishes ----------------
    for slug, label, outline in _static_outlines:
        if slug not in selected_slugs:
            continue
        outline.boundary.plot(ax=ax, color="#4FA3E3", linewidth=1.4)
        c = outline.geometry.centroid.iloc[0]
        ax.annotate(
            label, xy=(c.x, c.y), ha="center", va="center",
            fontsize=9, fontweight="bold", color="#1A1A1A",
            path_effects=[patheffects.withStroke(linewidth=2.5, foreground="white")],
        )

    ax.set_xlim(minx, maxx)
    ax.set_ylim(miny, maxy)

    try:
        ctx.add_basemap(ax, source=ctx.providers.CartoDB.Positron,
                        crs=_static_web.crs, attribution=False)
    except Exception as exc:  # offline / tile fetch failure -> plain background
        print("Static map: basemap fetch failed, rendering without it:", exc)

    ax.set_axis_off()
    ax.set_title("CRPC Region — DRA Distress Status", fontsize=14, fontweight="bold")

    # --- job density per tract overlay ---------------------------------------
    if state.get("job_tract_enabled"):
        _draw_job_overlay(
            ax, JOB_DATA,
            sector=state.get("job_tract_sector", "C000"),
            show_nums=bool(state.get("job_tract_show_nums")),
            active_slugs=selected_slugs,
            merc_transform=_to_merc,
            pt_scale=5,
        )

    # --- job density per parish overlay -------------------------------------
    if state.get("job_parish_enabled"):
        _draw_job_overlay(
            ax, PARISH_JOB_DATA,
            sector=state.get("job_parish_sector", "C000"),
            show_nums=bool(state.get("job_parish_show_nums")),
            active_slugs=selected_slugs,
            merc_transform=_to_merc,
            pt_scale=9,
        )

    # --- business establishments per parish overlay --------------------------
    if state.get("biz_parish_enabled"):
        _draw_business_overlay(
            ax, PARISH_BIZ_DATA,
            metric=state.get("biz_parish_metric", "est"),
            show_nums=bool(state.get("biz_parish_show_nums")),
            active_slugs=selected_slugs,
            merc_transform=_to_merc,
            pt_scale=9,
        )

    # --- tract number labels ------------------------------------------------
    if state.get("show_tract_nums"):
        for d in TRACT_POINTS:
            if d["parish"] not in selected_slugs:
                continue
            mx, my = _to_merc.transform(d["lng"], d["lat"])
            ax.annotate(
                _tract_num_str(d["geoid"]),
                xy=(mx, my), ha="center", va="top",
                fontsize=3.5, fontweight="bold", color="#111",
                bbox=dict(boxstyle="square,pad=0.08", fc="white", ec="none", alpha=0.88),
                zorder=7,
            )

    # --- legend: only the statuses actually shown ----------------------------
    shown_statuses = set(gdf["Distress Status"].unique()) if len(gdf) else set()
    legend_items = [d for d in DISTRESS_STATUSES
                    if (statuses is None or d["label"] in (statuses or []))
                    and (not shown_statuses or d["label"] in shown_statuses)]
    if not legend_items:
        legend_items = DISTRESS_STATUSES
    handles = [
        mpatches.Patch(facecolor=d["color"], edgecolor="#777777", label=d["label"])
        for d in legend_items
    ]
    leg = ax.legend(
        handles=handles, title="DRA Distress Status", loc="lower right",
        fontsize=11, title_fontsize=12, framealpha=1.0, markerscale=1.6,
        borderpad=0.8, labelspacing=0.6, fancybox=True,
    )
    leg.get_frame().set_edgecolor("#444444")

    buf = io.BytesIO()
    fig.savefig(buf, format="png", dpi=dpi, bbox_inches="tight", pad_inches=0.1)
    buf.seek(0)
    return buf


@app.route("/download_map", methods=["GET", "POST"])
def download_map():
    state = request.get_json(silent=True) if request.method == "POST" else None
    try:
        buf = render_static_map(state, dpi=DOWNLOAD_DPI)
    except Exception as exc:
        return str(exc), 500
    return send_file(
        buf,
        mimetype="image/png",
        as_attachment=True,
        download_name="crpc_distress_map_600dpi.png",
    )


if __name__ == "__main__":
    print()
    print("Total tracts loaded:", len(all_tracts))
    if llm_client is None:
        print("WARNING: GROQ_API_KEY (or LLM_API_KEY) is not set.")
        print("         The AI query box will return an error until you set it.")
    print()
    print("Open http://127.0.0.1:5000 in your browser.")
    app.run(host="127.0.0.1", port=5000)
