# Assistive Device Demand Dashboard

## Live Dashboard

[![Open Dashboard](https://img.shields.io/badge/Streamlit-Dashboard-red?logo=streamlit&style=for-the-badge)](https://assistive-device-dashboard-r2syjajgp3broag4y2zea3.streamlit.app/)

Streamlit dashboard for reviewing assistive device demand across three live assessment sources:

- school STRIDE form responses
- bedridden individual assessment responses
- institute demand from a published Google Sheet

The app cleans the live sources into request-level rows, combines them for analysis, and supports source-aware filtering across schools, bedridden records, and institutes.

## Project Files

- `app.py` - Streamlit dashboard
- `clean.py` - cleaning logic for school and bedridden sources
- `data/cleaned_data.csv` - cleaned school request data
- `data/cleaned_bedridden_data.csv` - cleaned bedridden request data
- `data/DEVICE_INFORMATION_CATALOG_FINAL.xlsx` - device catalog used by the size chart
- `.streamlit/config.toml` - Streamlit theme configuration

## Data Sources

### 1. School sheet

Loaded from the published Google Sheet configured in `clean.py`.

Current cleaning includes:

- district standardization
- school-name standardization by `School ID`
- gender normalization
- disability normalization
- request reshaping from:
  - `Device Priority 1`
  - `Device Priority 2`
  - `Device Priority 3`
  - `Common Requirement` / `Common Requirements`

Output:

- request-level rows in `data/cleaned_data.csv`

### 2. Bedridden assessment sheet

Loaded from the published Google Sheet configured in `clean.py`.

Current cleaning includes:

- `Primary Medical Condition` -> `Disability`
- `Preference 1`, `Preference 2`, `Perference 3` -> device request rows
- district, gender, and device-name normalization

Output:

- cleaned export in `data/cleaned_bedridden_data.csv`

Export columns:

- `Name`
- `Age`
- `Gender`
- `Contact No`
- `Address`
- `District`
- `Disability`
- `Device`
- `Other requirement`

### 3. Institutes

Institute demand is loaded from a published Google Sheet configured in `app.py`.

Current sheet shape:

- `Institution`
- `District`
- one column per device
- `Total`

The app reshapes that source into:

- `Institute`
- `District`
- `Device`
- `Device Category`
- `Requests`

This source is treated as institute-level demand, not person-level assessment data.

## Dashboard Sections

- `Demand overview`
  - district demand
  - device demand
  - priority distribution
  - device category distribution

- `Institutes`
  - institute device demand
  - institute data table for viewing

- `Bedridden`
  - bedridden district demand
  - bedridden device demand
  - cleaned bedridden table for viewing

- `Learner profile`
  - disability profile
  - social category profile
  - gender distribution

- `Size chart`
  - catalog mapping for measurement-based and fixed devices

- `Filtered data`
  - current filtered request rows

## Filters and Scope

The sidebar supports:

- `Population`
  - `Combined`
  - `Schools`
  - `Bedridden`
  - `Institutes`

- `KPI view`
  - `Auto`
  - `Schools`
  - `Bedridden`
  - `Institutes`

- collapsed slicers for:
  - Districts
  - Schools
  - Device categories
  - Devices
  - Institutes
  - Priorities
  - Genders
  - Ranked rows

Notes:

- `Schools` is dependent on selected districts
- `Population = Combined` uses school + bedridden records in the people views
- institute demand is also included in combined KPI request totals where applicable
- new institutes added to the live institute sheet will appear in the institute filter after manual refresh or hourly refresh

## KPI Behavior

The KPI cards can follow the selected population automatically or be forced to:

- school view
- bedridden view
- institute view

This affects:

- device request totals
- population coverage cards
- priority card wording
- most-needed-device logic

## Size Chart Logic

The size chart uses the catalog in `data/DEVICE_INFORMATION_CATALOG_FINAL.xlsx`.

Measurement-based sizing is used only for:

- `utensil holder`
- `palm pen holder`
- `toothbrush holder`

For those devices:

- palm length and width from cleaned school data are mapped to the nearest catalog `ND` size

All other devices are treated as fixed-measurement items or mapped to the closest available fixed catalog code.

Bedridden records do not currently provide palm measurements, so they mainly contribute to fixed-device size outputs.

## Refresh and Caching

- Streamlit auto-refresh: hourly
- data cache TTL: hourly

## Run Locally

Install dependencies:

```bash
pip install -r requirements.txt
```

Run the dashboard:

```bash
streamlit run app.py
```

If port `8501` is already in use:

```bash
streamlit run app.py --server.port 8502
```

## Deployment Notes

Keep the device catalog inside the repository at:

```text
data/DEVICE_INFORMATION_CATALOG_FINAL.xlsx
```

This is required for hosted deployments such as Streamlit Community Cloud, because local Windows paths are not available there.

## Dependencies

The app currently requires:

- `streamlit`
- `pandas`
- `plotly`
- `streamlit-autorefresh`
- `openpyxl`

`openpyxl` is required for reading the Excel device catalog.
