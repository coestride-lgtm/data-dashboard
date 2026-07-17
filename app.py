from datetime import datetime
from html import escape
from io import BytesIO
from pathlib import Path
import re

import pandas as pd
import plotly.express as px
import requests
import streamlit as st
from streamlit_autorefresh import st_autorefresh

from clean import load_and_clean_bedridden_data, load_and_clean_data


st.set_page_config(
    page_title="Assistive Device Dashboard",
    layout="wide",
    initial_sidebar_state="expanded",
)

REFRESH_INTERVAL_MS = 3600000
DATA_CACHE_TTL_SECONDS = 3600
PROJECT_DIR = Path(__file__).resolve().parent
CDC_PATH = PROJECT_DIR / "data" / "CDC.txt"
LAST_YEAR_SHEET_URL = (
    "https://innovationstrategiccouncil-my.sharepoint.com/:x:/g/personal/"
    "stride_kdisc_kerala_gov_in/IQB5AwKgFLXdQZ3xlHHthLDoAZdXsVA45jLSFXfcC_PAqTw?download=1"
)
CATALOG_PATHS = [
    PROJECT_DIR / "data" / "DEVICE_INFORMATION_CATALOG_FINAL.xlsx",
    Path(r"C:\Users\hp\Downloads\DEVICE_INFORMATION_CATALOG_FINAL.xlsx"),
]
CATALOG_PATH = next((path for path in CATALOG_PATHS if path.exists()), CATALOG_PATHS[0])

COLORS = {
    "ink": "#0f172a",
    "muted": "#475569",
    "line": "#dbeafe",
    "soft": "#eff6ff",
    "surface": "#ffffff",
    "blue": "#1d4ed8",
    "blue_dark": "#1e3a8a",
    "blue_light": "#60a5fa",
    "cyan": "#0891b2",
    "pink": "#ff69b4",
    "amber": "#f59e0b",
    "green": "#16a34a",
    "slate": "#94a3b8",
}

BLUE_SCALE = ["#dbeafe", "#bfdbfe", "#93c5fd", "#60a5fa", "#2563eb", "#1e3a8a"]
GENDER_COLORS = {
    "Male": "#1f77b4",
    "Female": "#ff69b4",
    "Other": "#94a3b8",
}
PRIORITY_COLORS = {
    "1": COLORS["amber"],
    "2": COLORS["blue"],
    "3": COLORS["blue_light"],
    "Common": COLORS["slate"],
}
CATEGORY_COLORS = {
    "Assistive": COLORS["blue"],
    "Cognitive": COLORS["cyan"],
    "Mobility": COLORS["green"],
}
DEVICE_CATEGORY_MAP = {
    "wheelchair": "Mobility",
    "crutches": "Mobility",
    "walking aid": "Mobility",
    "prosthetic limb": "Mobility",
    "orthotic device": "Mobility",
    "visual aid": "Assistive",
    "button aid": "Assistive",
    "braille": "Assistive",
    "hearing aid": "Assistive",
    "reading bar": "Assistive",
    "palm pen holder": "Assistive",
    "utensil holder": "Assistive",
    "toothbrush holder": "Assistive",
    "adaptive pencil grip": "Assistive",
    "braille kit": "Assistive",
    "communication device": "Cognitive",
    "maze": "Cognitive",
    "tetris": "Cognitive",
    "low profile switch": "Assistive",
    "communication board": "Cognitive",
    "speech device": "Cognitive",
}
DEVICE_NAME_MAP = {
    "wheel chair": "wheelchair",
    "wheel-chair": "wheelchair",
    "hearing aid device": "hearing aid",
    "hearing machine": "hearing aid",
    "walking stick": "walking aid",
    "walker": "walking aid",
    "low switch profile": "low profile switch",
    "tooth brush holder": "toothbrush holder",
    "tooth  brush adapter": "toothbrush holder",
    "tooth brush adapter": "toothbrush holder",
    "toothbrush adapter": "toothbrush holder",
    "reading bar - curved": "reading bar",
}


st.markdown(
    f"""
    <style>
        .stApp {{
            background: #ffffff;
            color: {COLORS["ink"]};
        }}

        .block-container {{
            padding-top: 2rem;
            padding-bottom: 2.75rem;
        }}

        [data-testid="stSidebar"] {{
            background: linear-gradient(180deg, #eff6ff 0%, #ffffff 58%);
            border-right: 1px solid #bfdbfe;
        }}

        [data-testid="stSidebar"] * {{
            color: {COLORS["ink"]};
        }}

        h1, h2, h3, h4, h5, h6, p, label, span {{
            color: {COLORS["ink"]};
            letter-spacing: 0;
        }}

        hr {{
            border-color: #dbeafe;
            margin: 1.5rem 0;
        }}

        div[data-testid="stButton"] button,
        div[data-testid="stDownloadButton"] button {{
            background: #3b82f6;
            border: 1px solid #3b82f6;
            border-radius: 8px;
            color: #ffffff;
            font-weight: 700;
        }}

        div[data-testid="stButton"] button:hover,
        div[data-testid="stDownloadButton"] button:hover {{
            background: #2563eb;
            border-color: #2563eb;
            color: #ffffff;
        }}

        [data-testid="stSidebar"] [data-testid="stExpander"] {{
            background: #ffffff;
            border: 1px solid #bfdbfe;
            border-radius: 8px;
            box-shadow: 0 8px 24px rgba(29, 78, 216, 0.08);
            margin-bottom: 0.65rem;
            overflow: hidden;
        }}

        [data-testid="stSidebar"] details summary {{
            background: #eff6ff;
            border-radius: 8px;
            color: {COLORS["blue_dark"]};
            font-weight: 800;
            padding: 0.15rem 0.2rem;
        }}

        [data-testid="stSidebar"] details[open] summary {{
            border-bottom: 1px solid #dbeafe;
            border-radius: 8px 8px 0 0;
            margin-bottom: 0.55rem;
        }}

        [data-testid="stSidebar"] div[data-baseweb="select"] > div {{
            background: #ffffff;
            border-color: #93c5fd;
        }}

        [data-testid="stSidebar"] div[data-baseweb="select"] svg {{
            fill: {COLORS["blue_dark"]};
            color: {COLORS["blue_dark"]};
        }}

        [data-testid="stSidebar"] div[data-baseweb="select"] span,
        [data-testid="stSidebar"] div[data-baseweb="tag"] span {{
            color: {COLORS["blue_dark"]};
        }}

        [data-testid="stSidebar"] div[data-baseweb="tag"],
        [data-testid="stSidebar"] div[data-baseweb="tag"] > span,
        [data-testid="stSidebar"] div[data-baseweb="tag"] > div {{
            background-color: #2563eb !important;
            border: 1px solid #2563eb;
            border-radius: 6px;
            color: #ffffff !important;
        }}

        [data-testid="stSidebar"] div[data-baseweb="tag"] svg,
        [data-testid="stSidebar"] div[data-baseweb="tag"] path {{
            fill: #ffffff;
            color: #ffffff;
        }}

        [data-testid="stSidebar"] div[data-baseweb="tag"]:hover,
        [data-testid="stSidebar"] div[data-baseweb="tag"]:hover > span,
        [data-testid="stSidebar"] div[data-baseweb="tag"]:hover > div {{
            background-color: #1d4ed8 !important;
            border-color: #1d4ed8;
        }}

        [data-testid="stSidebar"] div[data-baseweb="tag"] button,
        [data-testid="stSidebar"] div[data-baseweb="tag"] [role="button"] {{
            background: transparent !important;
            color: #ffffff !important;
        }}

        [data-testid="stSidebar"] div[data-baseweb="tag"] *,
        [data-testid="stSidebar"] div[data-baseweb="tag"] span,
        [data-testid="stSidebar"] div[data-baseweb="tag"] div {{
            color: #ffffff !important;
            -webkit-text-fill-color: #ffffff !important;
        }}

        [data-testid="stSidebar"] div[data-baseweb="tag"] svg,
        [data-testid="stSidebar"] div[data-baseweb="tag"] svg *,
        [data-testid="stSidebar"] div[data-baseweb="tag"] path {{
            color: #ffffff !important;
            fill: #ffffff !important;
            stroke: #ffffff !important;
        }}

        [data-testid="stSidebar"] [data-baseweb="tag"],
        [data-testid="stSidebar"] [data-baseweb="tag"] *,
        [data-testid="stSidebar"] [data-baseweb="tag"] span,
        [data-testid="stSidebar"] [data-baseweb="tag"] div,
        [data-testid="stSidebar"] [data-baseweb="tag"] [title],
        [data-testid="stSidebar"] div[data-baseweb="select"] [data-baseweb="tag"],
        [data-testid="stSidebar"] div[data-baseweb="select"] [data-baseweb="tag"] *,
        [data-testid="stSidebar"] div[data-baseweb="select"] span[title] {{
            color: #ffffff !important;
            fill: #ffffff !important;
            stroke: #ffffff !important;
            -webkit-text-fill-color: #ffffff !important;
        }}

        [data-testid="stSidebar"] [data-baseweb="tag"] svg,
        [data-testid="stSidebar"] [data-baseweb="tag"] path,
        [data-testid="stSidebar"] div[data-baseweb="select"] [data-baseweb="tag"] svg,
        [data-testid="stSidebar"] div[data-baseweb="select"] [data-baseweb="tag"] path {{
            color: #ffffff !important;
            fill: #ffffff !important;
            stroke: #ffffff !important;
        }}

        [data-testid="stSidebar"] div[role="listbox"],
        [data-testid="stSidebar"] ul {{
            background: #ffffff;
            border-color: #bfdbfe;
        }}

        [data-testid="stSidebar"] li {{
            color: {COLORS["ink"]};
        }}

        div[data-baseweb="select"] > div {{
            background: #ffffff;
            border-color: #bfdbfe;
        }}

        div[data-baseweb="select"] span,
        div[data-baseweb="select"] input,
        div[data-baseweb="select"] div[role="button"] {{
            color: {COLORS["ink"]} !important;
            -webkit-text-fill-color: {COLORS["ink"]} !important;
            opacity: 1 !important;
        }}

        [data-testid="stMetric"] {{
            background: #ffffff;
            border: 1px solid #dbeafe;
            border-radius: 8px;
            padding: 1rem;
        }}

        [data-testid="stExpander"] {{
            background: #ffffff;
            border: 1px solid #dbeafe;
            border-radius: 8px;
        }}

        .app-hero {{
            border-bottom: 1px solid #dbeafe;
            padding: 0 0 1.25rem 0;
            margin-bottom: 1rem;
        }}

        .eyebrow {{
            color: {COLORS["blue"]};
            font-size: 0.78rem;
            font-weight: 800;
            letter-spacing: 0.08em;
            text-transform: uppercase;
            margin-bottom: 0.35rem;
        }}

        .hero-title {{
            color: {COLORS["ink"]};
            font-size: 2.35rem;
            font-weight: 850;
            line-height: 1.05;
            margin: 0;
        }}

        .hero-copy {{
            color: {COLORS["muted"]};
            font-size: 1rem;
            line-height: 1.55;
            margin: 0.65rem 0 0 0;
            max-width: 68rem;
        }}

        .status-row {{
            display: flex;
            flex-wrap: wrap;
            gap: 0.55rem;
            margin-top: 1rem;
        }}

        .status-pill {{
            align-items: center;
            background: #eff6ff;
            border: 1px solid #bfdbfe;
            border-radius: 999px;
            color: {COLORS["blue_dark"]};
            display: inline-flex;
            font-size: 0.82rem;
            font-weight: 700;
            gap: 0.4rem;
            padding: 0.35rem 0.7rem;
        }}

        .section-title {{
            color: {COLORS["ink"]};
            font-size: 1.25rem;
            font-weight: 850;
            margin: 1.25rem 0 0.35rem 0;
        }}

        .section-copy {{
            color: {COLORS["muted"]};
            font-size: 0.94rem;
            margin: 0 0 0.9rem 0;
        }}

        .metric-card {{
            background: linear-gradient(135deg, #1d4ed8 0%, #0284c7 100%);
            border: 1px solid #bfdbfe;
            border-radius: 8px;
            box-shadow: 0 10px 28px rgba(29, 78, 216, 0.14);
            min-height: 150px;
            padding: 1.1rem;
        }}

        .metric-label {{
            color: rgba(255, 255, 255, 0.86);
            font-size: 0.78rem;
            font-weight: 800;
            letter-spacing: 0.06em;
            text-transform: uppercase;
        }}

        .metric-value {{
            color: #ffffff;
            font-size: 2rem;
            font-weight: 850;
            line-height: 1.1;
            margin-top: 0.55rem;
        }}

        .metric-note {{
            color: rgba(255, 255, 255, 0.84);
            font-size: 0.86rem;
            line-height: 1.35;
            margin-top: 0.6rem;
        }}

        .insight-card {{
            background: #ffffff;
            border: 1px solid #dbeafe;
            border-radius: 8px;
            box-shadow: 0 8px 24px rgba(15, 23, 42, 0.05);
            min-height: 132px;
            padding: 1rem;
        }}

        .insight-label {{
            color: {COLORS["muted"]};
            font-size: 0.78rem;
            font-weight: 800;
            letter-spacing: 0.06em;
            text-transform: uppercase;
        }}

        .insight-value {{
            color: {COLORS["ink"]};
            font-size: 1.35rem;
            font-weight: 850;
            line-height: 1.2;
            margin-top: 0.45rem;
        }}

        .insight-note {{
            color: {COLORS["muted"]};
            font-size: 0.9rem;
            line-height: 1.4;
            margin-top: 0.45rem;
        }}

        .empty-state {{
            background: #eff6ff;
            border: 1px solid #bfdbfe;
            border-radius: 8px;
            color: {COLORS["blue_dark"]};
            padding: 1.2rem;
        }}

        .sidebar-title {{
            color: {COLORS["blue_dark"]};
            font-size: 1.05rem;
            font-weight: 850;
            margin-bottom: 0.15rem;
        }}

        .sidebar-copy {{
            color: {COLORS["muted"]};
            font-size: 0.86rem;
            line-height: 1.4;
            margin-bottom: 0.85rem;
        }}
    </style>
    """,
    unsafe_allow_html=True,
)


st_autorefresh(interval=REFRESH_INTERVAL_MS, key="autorefresh")


@st.cache_data(ttl=DATA_CACHE_TTL_SECONDS)
def load_data_streamlit():
    return load_and_clean_data()


@st.cache_data(ttl=DATA_CACHE_TTL_SECONDS)
def load_bedridden_data_streamlit():
    return load_and_clean_bedridden_data()


def normalize_district_name(value):
    district = clean_text(value, "").strip().title()
    district_map = {
        "Thiruvananthapuram": "Trivandrum",
        "Tvm": "Trivandrum",
        "Kazargod": "Kasaragod",
    }
    return district_map.get(district, district)


def read_excel_workbook(source_url):
    response = requests.get(source_url, timeout=30)
    response.raise_for_status()
    return BytesIO(response.content)


@st.cache_data(ttl=DATA_CACHE_TTL_SECONDS)
def load_last_year_data_streamlit(source_url):
    unit_columns = ["District", "Device", "Device Category", "Requests", "Data Source", "Record Type"]
    learner_columns = [
        "District",
        "Student ID",
        "Student Name",
        "disability_cleaned",
        "School_Name",
        "Palm Length Cleaned",
        "Palm Width Cleaned",
        "Device",
        "Device Category",
        "Gender",
        "Social Category",
        "Priority",
        "Record Type",
        "Data Source",
    ]

    try:
        workbook = read_excel_workbook(source_url)
        pivot = pd.read_excel(workbook, sheet_name="PIVOT DATA", header=None)
    except Exception:
        return pd.DataFrame(columns=unit_columns), pd.DataFrame(columns=learner_columns)

    header_matches = pivot[pivot.iloc[:, 0].astype(str).str.strip().str.upper() == "DISTRICT"]
    if header_matches.empty:
        return pd.DataFrame(columns=unit_columns), pd.DataFrame(columns=learner_columns)

    header_idx = header_matches.index[0]
    header = pivot.iloc[header_idx].fillna("").astype(str).str.strip().tolist()
    district_end_matches = pivot.loc[
        header_idx + 1 :, 0
    ].astype(str).str.strip().str.upper().eq("GRAND TOTAL")
    end_idx = district_end_matches[district_end_matches].index.min()
    rows = pivot.iloc[header_idx + 1 : end_idx].copy()
    rows.columns = header
    rows = rows[rows["DISTRICT"].notna()]
    rows = rows[rows["DISTRICT"].astype(str).str.strip().str.upper() != "GRAND TOTAL"]
    rows = rows[rows["DISTRICT"].astype(str).str.strip() != ""]

    device_cols = [col for col in rows.columns if col not in {"DISTRICT", "GRAND TOTAL", ""}]
    unit_rows = []
    for _, row in rows.iterrows():
        district = normalize_district_name(row.get("DISTRICT"))
        if not district:
            continue
        for device_col in device_cols:
            requests = pd.to_numeric(row.get(device_col), errors="coerce")
            if pd.isna(requests) or requests <= 0:
                continue
            device = normalize_device_name(device_col)
            unit_rows.append(
                {
                    "District": district,
                    "Device": device,
                    "Device Category": DEVICE_CATEGORY_MAP.get(device, "Other"),
                    "Requests": int(requests),
                    "Data Source": "Phase 1",
                    "Record Type": "Last Year",
                }
            )
    unit_df = pd.DataFrame(unit_rows, columns=unit_columns)

    try:
        workbook.seek(0)
        learner_df = pd.read_excel(workbook, sheet_name="ALL 498 CONSOLIDATED NEW")
    except Exception:
        return unit_df, pd.DataFrame(columns=learner_columns)
    learner_df.columns = learner_df.columns.astype(str).str.strip()
    learner_df = learner_df.rename(
        columns={
            "DISTRICT": "District",
            "STUDENT_ID": "Student ID",
            "STUDENT NAME": "Student Name",
            "DISABILITY": "disability_cleaned",
            "BUDS SCHOOL": "School_Name",
            "LENGTH": "Palm Length Cleaned",
            "WIDTH": "Palm Width Cleaned",
            "PREFERENCE": "Device",
        }
    )
    expected_cols = [
        "District",
        "Student ID",
        "Student Name",
        "disability_cleaned",
        "School_Name",
        "Palm Length Cleaned",
        "Palm Width Cleaned",
        "Device",
    ]
    learner_df = learner_df[[col for col in expected_cols if col in learner_df.columns]].copy()
    learner_df = learner_df[learner_df["Device"].notna()].copy()
    learner_df["District"] = learner_df["District"].map(normalize_district_name)
    learner_df["Device"] = learner_df["Device"].map(normalize_device_name)
    learner_df = learner_df[learner_df["Device"].notna() & learner_df["District"].ne("")]
    learner_df["Device Category"] = learner_df["Device"].map(lambda device: DEVICE_CATEGORY_MAP.get(device, "Other"))
    learner_df["Gender"] = "Unknown"
    learner_df["Social Category"] = "Unknown"
    learner_df["Priority"] = "Collected"
    learner_df["Record Type"] = "Last Year"
    learner_df["Data Source"] = "Phase 1"

    return unit_df, learner_df.reindex(columns=learner_columns)


@st.cache_data(ttl=DATA_CACHE_TTL_SECONDS)
def load_device_catalog(catalog_path):
    path = Path(catalog_path)
    if not path.exists():
        return {}

    item_master = pd.read_excel(path, sheet_name="Item Code Master")
    device_codes = pd.read_excel(path, sheet_name="Device Codes")
    size_raw = pd.read_excel(path, sheet_name="Size Code Master", header=None)

    item_master.columns = item_master.columns.astype(str).str.strip()
    device_codes.columns = device_codes.columns.astype(str).str.strip()

    nd_ref = size_raw.iloc[3:23, 0:4].copy()
    nd_ref.columns = ["Size ID", "Length (cm)", "Width (cm)", "Size Code"]
    nd_ref["Size ID"] = nd_ref["Size ID"].map(lambda value: clean_text(value, "").replace("\xa0", "").strip())
    nd_ref["Length (cm)"] = nd_ref["Length (cm)"].map(numeric_value)
    nd_ref["Width (cm)"] = nd_ref["Width (cm)"].map(numeric_value)
    nd_ref["Size Code"] = nd_ref["Size Code"].map(numeric_code)
    nd_ref = nd_ref.dropna(subset=["Size ID", "Length (cm)", "Width (cm)", "Size Code"])

    ndl_ref = size_raw.iloc[27:37, 0:3].copy()
    ndl_ref.columns = ["Size ID", "Length (cm)", "Size Code"]
    ndl_ref["Size ID"] = ndl_ref["Size ID"].map(lambda value: clean_text(value, "").replace("\xa0", "").strip())
    ndl_ref["Length (cm)"] = ndl_ref["Length (cm)"].map(numeric_value)
    ndl_ref["Size Code"] = ndl_ref["Size Code"].map(numeric_code)
    ndl_ref = ndl_ref.dropna(subset=["Size ID", "Length (cm)", "Size Code"])

    rb_ref = size_raw.iloc[39:41, 0:3].copy()
    rb_ref.columns = ["Size ID", "Paper Size", "Size Code"]
    rb_ref["Size ID"] = rb_ref["Size ID"].map(lambda value: clean_text(value, "").strip())
    rb_ref["Paper Size"] = rb_ref["Paper Size"].map(lambda value: clean_text(value, "").strip())
    rb_ref["Size Code"] = rb_ref["Size Code"].map(numeric_code)
    rb_ref = rb_ref.dropna(subset=["Size ID", "Size Code"])

    item_lookup = {}
    for _, row in item_master.iterrows():
        device_name = clean_text(row.get("Device"), "")
        if device_name:
            item_lookup[normalize_key(device_name)] = device_name

    for item_name in device_codes["Item"].dropna().unique():
        item_lookup.setdefault(normalize_key(item_name), clean_text(item_name))

    item_lookup.update(
        {
            "adaptivepencilgrip": "Assistive Pencil Grip",
            "assistivepencilgrip": "Assistive Pencil Grip",
            "toothbrushholder": "Tooth Brush Adapter",
            "toothbrushadapter": "Tooth Brush Adapter",
            "braille": "Braille Device",
            "brailledevice": "Braille Device",
        }
    )

    return {
        "item_master": item_master,
        "device_codes": device_codes,
        "nd_ref": nd_ref,
        "ndl_ref": ndl_ref,
        "rb_ref": rb_ref,
        "item_lookup": item_lookup,
    }


def fmt_number(value):
    try:
        return f"{int(value):,}"
    except (TypeError, ValueError):
        return "0"


def fmt_percent(part, whole):
    if not whole:
        return "0%"
    return f"{(part / whole) * 100:.1f}%"


def clean_text(value, fallback="Not available"):
    if value is None or pd.isna(value) or str(value).strip() == "":
        return fallback
    return str(value)


def safe_html(value):
    return escape(clean_text(value))


def normalize_key(value):
    text = clean_text(value, "").lower()
    return re.sub(r"[^a-z0-9]+", "", text)


def numeric_code(value):
    digits = re.sub(r"\D", "", clean_text(value, ""))
    return int(digits) if digits else None


def numeric_value(value):
    match = re.search(r"\d+(?:\.\d+)?", clean_text(value, ""))
    return float(match.group(0)) if match else None


def value_counts_frame(df, source_col, label_col, limit=None):
    if df.empty or source_col not in df.columns:
        return pd.DataFrame(columns=[label_col, "Requests"])
    counts = df[source_col].fillna("Unknown").astype(str).value_counts().reset_index()
    counts.columns = [label_col, "Requests"]
    if limit:
        counts = counts.head(limit)
    return counts


def demand_counts_frame(df, source_col, label_col, limit=None):
    if df.empty or source_col not in df.columns:
        return pd.DataFrame(columns=[label_col, "Requests"])

    request_source = df["Requests"] if "Requests" in df.columns else pd.Series(1, index=df.index)
    request_values = pd.to_numeric(request_source, errors="coerce").fillna(1)
    demand_df = df.assign(_Requests=request_values)
    counts = (
        demand_df.assign(**{source_col: demand_df[source_col].fillna("Unknown").astype(str)})
        .groupby(source_col, as_index=False)["_Requests"]
        .sum()
        .rename(columns={source_col: label_col, "_Requests": "Requests"})
        .sort_values("Requests", ascending=False)
    )
    counts["Requests"] = counts["Requests"].astype(int)
    if limit:
        counts = counts.head(limit)
    return counts


def learner_key_frame(df):
    if df.empty:
        return df.copy()

    keyed = df.copy()
    record_type = keyed.get("Record Type", pd.Series("", index=keyed.index)).fillna("").astype(str)
    school_id = keyed.get("School ID", pd.Series("", index=keyed.index)).fillna("").map(normalize_key)
    school_name = keyed.get("School_Name", pd.Series("", index=keyed.index)).fillna("").map(normalize_key)
    district = keyed.get("District", pd.Series("", index=keyed.index)).fillna("").map(normalize_key)
    student_id = keyed.get("Student ID", pd.Series("", index=keyed.index)).fillna("").map(normalize_key)
    student = keyed.get("Student Name", pd.Series("", index=keyed.index)).fillna("").map(normalize_key)
    contact = keyed.get("Contact No", pd.Series("", index=keyed.index)).fillna("").map(normalize_key)

    school_scope = school_id.where(school_id.ne(""), school_name)
    person_scope = contact.where(contact.ne(""), student_id.where(student_id.ne(""), student))
    fallback = pd.Series(keyed.index, index=keyed.index).astype(str)
    person_scope = person_scope.where(person_scope.ne(""), fallback)

    keyed["_Learner Key"] = record_type + "|" + district + "|" + school_scope + "|" + person_scope
    return keyed


def distinct_learner_counts_frame(df, source_col, label_col, limit=None):
    if df.empty or source_col not in df.columns:
        return pd.DataFrame(columns=[label_col, "Learners"])

    learner_rows = learner_key_frame(df).drop_duplicates("_Learner Key")
    counts = learner_rows[source_col].fillna("Unknown").astype(str).value_counts().reset_index()
    counts.columns = [label_col, "Learners"]
    if limit:
        counts = counts.head(limit)
    return counts


def normalize_device_name(value):
    device = clean_text(value, "").strip().lower()
    return DEVICE_NAME_MAP.get(device, device)


def display_device_name(value):
    return clean_text(value, "").replace("cdc", "CDC").title()


@st.cache_data(ttl=DATA_CACHE_TTL_SECONDS)
def load_cdc_data(cdc_path):
    path = Path(cdc_path)
    if not path.exists():
        return pd.DataFrame(columns=["Institute", "District", "Device", "Device Category", "Requests", "Source"])

    records = []
    for raw_line in path.read_text(encoding="utf-8").splitlines():
        line = raw_line.strip()
        if not line or line.lower() == "cdc" or line.lower().startswith("total count"):
            continue
        match = re.match(r"^(.*?)\s*-\s*(\d+)\s*$", line)
        if not match:
            continue
        device = normalize_device_name(match.group(1))
        requests = int(match.group(2))
        records.append(
            {
                "Institute": "CDC",
                "District": "Trivandrum",
                "Device": device,
                "Device Category": DEVICE_CATEGORY_MAP.get(device, "Other"),
                "Requests": requests,
                "Source": "Institute",
            }
        )

    return pd.DataFrame(records)


def style_chart(fig, height=380, show_colorbar=False):
    fig.update_layout(
        template="plotly_white",
        paper_bgcolor="#ffffff",
        plot_bgcolor="#ffffff",
        font=dict(color=COLORS["ink"], family="Arial, sans-serif", size=13),
        height=height,
        margin=dict(l=12, r=12, t=18, b=20),
        hoverlabel=dict(bgcolor="#0f172a", font_color="#ffffff", bordercolor="#0f172a"),
        legend=dict(
            orientation="h",
            yanchor="bottom",
            y=1.02,
            xanchor="right",
            x=1,
            font=dict(color=COLORS["ink"]),
        ),
        coloraxis_showscale=show_colorbar,
    )
    fig.update_xaxes(
        gridcolor="#dbeafe",
        zerolinecolor="#bfdbfe",
        linecolor="#bfdbfe",
        tickfont=dict(color=COLORS["muted"]),
        title_font=dict(color=COLORS["ink"]),
    )
    fig.update_yaxes(
        gridcolor="#dbeafe",
        zerolinecolor="#bfdbfe",
        linecolor="#bfdbfe",
        tickfont=dict(color=COLORS["muted"]),
        title_font=dict(color=COLORS["ink"]),
    )
    return fig


def render_metric(label, value, note):
    st.markdown(
        f"""
        <div class="metric-card">
            <div class="metric-label">{safe_html(label)}</div>
            <div class="metric-value">{safe_html(value)}</div>
            <div class="metric-note">{safe_html(note)}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def render_insight(label, value, note):
    st.markdown(
        f"""
        <div class="insight-card">
            <div class="insight-label">{safe_html(label)}</div>
            <div class="insight-value">{safe_html(value)}</div>
            <div class="insight-note">{safe_html(note)}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def make_horizontal_bar(df, x_col, y_col, color_col="Requests", height=380, xaxis_title="Requests"):
    chart_df = df.sort_values(x_col, ascending=True).copy()
    chart_df["Display label"] = chart_df.apply(
        lambda row: f"{row[y_col]} ({fmt_number(row[x_col])})",
        axis=1,
    )
    fig = px.bar(
        chart_df,
        x=x_col,
        y="Display label",
        orientation="h",
        color=color_col,
        color_continuous_scale=BLUE_SCALE,
        text=x_col,
        custom_data=[y_col, x_col],
    )
    fig.update_traces(
        marker_line_width=0,
        texttemplate="%{text:,}",
        textposition="outside",
        cliponaxis=False,
        hovertemplate=f"{y_col}: %{{customdata[0]}}<br>{xaxis_title}: %{{customdata[1]:,}}<extra></extra>",
    )
    fig.update_layout(xaxis_title=xaxis_title, yaxis_title=None)
    fig.update_yaxes(categoryorder="array", categoryarray=chart_df["Display label"].tolist())
    return style_chart(fig, height=height)


def nearest_size_row(reference_df, length=None, width=None):
    if reference_df.empty:
        return None
    if length is None and width is None:
        return None
    scored = reference_df.copy()
    scored["_distance"] = 0.0
    if length is not None and pd.notna(length) and "Length (cm)" in scored.columns:
        scored["_distance"] += (scored["Length (cm)"].astype(float) - float(length)).abs()
    if width is not None and pd.notna(width) and "Width (cm)" in scored.columns:
        scored["_distance"] += (scored["Width (cm)"].astype(float) - float(width)).abs()
    if scored["_distance"].isna().all():
        return None
    return scored.sort_values("_distance").iloc[0]


def catalog_device_code(catalog, item_name, size_code):
    device_codes = catalog.get("device_codes", pd.DataFrame())
    if device_codes.empty or size_code is None:
        return {}

    matches = device_codes[
        (device_codes["Item"].map(normalize_key) == normalize_key(item_name))
        & (device_codes["Size Code"].map(numeric_code) == int(size_code))
    ]
    if matches.empty:
        return {}

    record = matches.iloc[0]
    return {
        "Unique Device Code": clean_text(record.get("Unique Device Code"), ""),
        "Area": clean_text(record.get("Area"), ""),
        "Sub-area": clean_text(record.get("Sub-area"), ""),
    }


def fixed_catalog_info(catalog, item_name):
    device_codes = catalog.get("device_codes", pd.DataFrame())
    if device_codes.empty:
        return {
            "size_code": 999,
            "size_id": "Fixed",
            "size_label": "Fixed - measurement not required",
            "code_info": {},
        }

    matches = device_codes[device_codes["Item"].map(normalize_key) == normalize_key(item_name)].copy()
    if matches.empty:
        return {
            "size_code": 999,
            "size_id": "Fixed",
            "size_label": "Fixed - measurement not required",
            "code_info": {},
        }

    fixed_matches = matches[matches["Size Code"].map(numeric_code) == 999]
    if not fixed_matches.empty:
        record = fixed_matches.iloc[0]
        size_code = 999
        size_id = "Fixed"
    elif len(matches) == 1:
        record = matches.iloc[0]
        size_code = numeric_code(record.get("Size Code"))
        size_id = f"Fixed-{int(size_code):03d}" if size_code is not None else "Fixed"
    else:
        record = matches.sort_values("Size Code").iloc[0]
        size_code = numeric_code(record.get("Size Code"))
        size_id = "Fixed"

    code_info = {
        "Unique Device Code": clean_text(record.get("Unique Device Code"), ""),
        "Area": clean_text(record.get("Area"), ""),
        "Sub-area": clean_text(record.get("Sub-area"), ""),
    }
    return {
        "size_code": size_code,
        "size_id": size_id,
        "size_label": "Fixed - measurement not required",
        "code_info": code_info,
    }


def assign_catalog_size(row, catalog, device_measurements=None):
    item_lookup = catalog.get("item_lookup", {})
    cleaned_device = clean_text(row.get("Device"), "")
    item_name = item_lookup.get(normalize_key(cleaned_device), cleaned_device.title())
    item_key = normalize_key(item_name)
    length = pd.to_numeric(row.get("Palm Length Cleaned"), errors="coerce")
    width = pd.to_numeric(row.get("Palm Width Cleaned"), errors="coerce")
    length = None if pd.isna(length) else float(length)
    width = None if pd.isna(width) else float(width)
    measurement_fill = (device_measurements or {}).get(cleaned_device, {})
    if length is None:
        fill_length = measurement_fill.get("Palm Length Cleaned")
        length = None if pd.isna(fill_length) else float(fill_length)
    if width is None:
        fill_width = measurement_fill.get("Palm Width Cleaned")
        width = None if pd.isna(fill_width) else float(fill_width)

    if item_key in {"utensilholder", "palmpenholder", "toothbrushadapter"}:
        size_system = "Palm length and width"
        if length is None or width is None:
            size_id = "Missing"
            size_code = None
            size_label = "Measurement missing"
            catalog_measure = {"Catalog Length (cm)": None, "Catalog Width (cm)": None}
        else:
            size_row = nearest_size_row(catalog.get("nd_ref", pd.DataFrame()), length=length, width=width)
            size_id = clean_text(size_row["Size ID"]) if size_row is not None else "ND"
            size_code = int(size_row["Size Code"]) if size_row is not None else None
            size_label = (
                f"{size_id} - {size_row['Length (cm)']:g} x {size_row['Width (cm)']:g} cm"
                if size_row is not None
                else "Measurement size"
            )
            catalog_measure = {
                "Catalog Length (cm)": float(size_row["Length (cm)"]) if size_row is not None else None,
                "Catalog Width (cm)": float(size_row["Width (cm)"]) if size_row is not None else None,
            }
    elif item_key == "buttonaid":
        size_row = catalog.get("ndl_ref", pd.DataFrame())
        size_row = size_row[size_row["Size Code"] == 9].iloc[0] if not size_row.empty else None
        size_system = "Fixed length"
        size_id = clean_text(size_row["Size ID"]) if size_row is not None else "NDL-09"
        size_code = 9
        size_label = f"{size_id} - 9 cm"
        catalog_measure = {"Catalog Length (cm)": 9.0, "Catalog Width (cm)": None}
    else:
        fixed_info = fixed_catalog_info(catalog, item_name)
        size_system = "Fixed measurement"
        size_id = fixed_info["size_id"]
        size_code = fixed_info["size_code"]
        size_label = fixed_info["size_label"]
        catalog_measure = {"Catalog Length (cm)": None, "Catalog Width (cm)": None}
        code_info = fixed_info["code_info"]

    if item_key in {"utensilholder", "palmpenholder", "toothbrushadapter", "buttonaid"}:
        code_info = catalog_device_code(catalog, item_name, size_code)
    code_info.setdefault("Unique Device Code", "")
    code_info.setdefault("Area", "")
    code_info.setdefault("Sub-area", "")
    return {
        "Device": item_name,
        "Dashboard Device": cleaned_device,
        "Size System": size_system,
        "Size ID": size_id,
        "Size Code": f"{int(size_code):03d}" if size_code is not None else "",
        "Size Label": size_label,
        "Palm Length Cleaned": length,
        "Palm Width Cleaned": width,
        **catalog_measure,
        **code_info,
    }


def build_size_chart_data(filtered_data, catalog):
    if filtered_data.empty or not catalog:
        return pd.DataFrame()

    device_measurements = (
        filtered_data[filtered_data["Device"].isin(["utensil holder", "palm pen holder", "toothbrush holder"])]
        .groupby("Device")[["Palm Length Cleaned", "Palm Width Cleaned"]]
        .median()
        .to_dict("index")
    )

    rows = [
        assign_catalog_size(row, catalog, device_measurements)
        for _, row in filtered_data.iterrows()
    ]
    assigned = pd.DataFrame(rows)
    group_cols = [
        "Device",
        "Dashboard Device",
        "Size System",
        "Size ID",
        "Size Code",
        "Size Label",
        "Unique Device Code",
        "Area",
        "Sub-area",
        "Catalog Length (cm)",
        "Catalog Width (cm)",
    ]
    summary = (
        assigned.groupby(group_cols, dropna=False)
        .agg(
            Requests=("Device", "size"),
            Avg_Palm_Length_cm=("Palm Length Cleaned", "mean"),
            Avg_Palm_Width_cm=("Palm Width Cleaned", "mean"),
        )
        .reset_index()
        .sort_values(["Device", "Requests"], ascending=[True, False])
    )
    summary["Avg Palm Length (cm)"] = summary["Avg_Palm_Length_cm"].round(2)
    summary["Avg Palm Width (cm)"] = summary["Avg_Palm_Width_cm"].round(2)
    return summary.drop(columns=["Avg_Palm_Length_cm", "Avg_Palm_Width_cm"])


def render_slicer(title, options, key, placeholder=None):
    option_list = list(options)
    previous_options_key = f"_{key}_options"
    previous_options = st.session_state.get(previous_options_key, [])
    if key not in st.session_state:
        st.session_state[key] = option_list
    elif set(st.session_state[key]) == set(previous_options):
        st.session_state[key] = option_list
    else:
        st.session_state[key] = [item for item in st.session_state[key] if item in option_list]
    st.session_state[previous_options_key] = option_list

    with st.sidebar.expander(title, expanded=False):
        return st.multiselect(
            "Choose options",
            options=option_list,
            key=key,
            label_visibility="collapsed",
            placeholder=placeholder or f"Select {title.lower()}",
        )


school_df = load_data_streamlit().copy()
school_df["Data Source"] = "Schools"
school_df["Record Type"] = "School"
school_df["Name"] = school_df["Student Name"]
school_df["Age"] = pd.NA
school_df["Contact No"] = pd.NA
school_df["Address"] = pd.NA
school_df["Other requirement"] = pd.NA
bedridden_df = load_bedridden_data_streamlit().copy()
people_df = pd.concat([school_df, bedridden_df], ignore_index=True, sort=False)
catalog = load_device_catalog(str(CATALOG_PATH))
cdc_df = load_cdc_data(str(CDC_PATH))
last_year_units_df, last_year_learners_df = load_last_year_data_streamlit(LAST_YEAR_SHEET_URL)

if "kpi_basis" not in st.session_state:
    st.session_state["kpi_basis"] = "Auto"
if "analysis_scope" not in st.session_state:
    st.session_state["analysis_scope"] = "Combined"
if "dashboard_year" not in st.session_state:
    st.session_state["dashboard_year"] = "All phases"
else:
    st.session_state["dashboard_year"] = {
        "Current year": "Phase 2",
        "Last year collected": "Phase 1",
        "All years": "Phases",
        "Phases": "All phases",
    }.get(st.session_state["dashboard_year"], st.session_state["dashboard_year"])

st.sidebar.markdown(
    """
    <div class="sidebar-title">Filters</div>
    """,
    unsafe_allow_html=True,
)

if st.sidebar.button("Refresh data"):
    st.cache_data.clear()
    st.rerun()

st.sidebar.divider()

dashboard_choice = st.sidebar.radio(
    "Phase",
    options=["All phases", "Phase 2", "Phase 1"],
    key="dashboard_year",
)
dashboard_year = {"All phases": "Phases", "Phase 2": "Phase 2", "Phase 1": "Phase 1"}[dashboard_choice]

if dashboard_year == "Phase 1":
    districts = sorted(last_year_units_df["District"].dropna().unique()) if not last_year_units_df.empty else []
    selected_districts = render_slicer("Districts", districts, "last_year_districts_filter", "Choose districts")

    categories = sorted(last_year_units_df["Device Category"].dropna().unique()) if not last_year_units_df.empty else []
    selected_categories = render_slicer(
        "Device categories",
        categories,
        "last_year_categories_filter",
        "Choose categories",
    )

    devices = sorted(last_year_units_df["Device"].dropna().unique()) if not last_year_units_df.empty else []
    selected_devices = render_slicer("Devices", devices, "last_year_devices_filter", "Choose devices")

    with st.sidebar.expander("Ranked rows", expanded=False):
        top_n = st.slider("Rows per chart", min_value=5, max_value=20, value=10, step=1, key="last_year_top_n")

    ly_units = last_year_units_df[
        (last_year_units_df["District"].isin(selected_districts))
        & (last_year_units_df["Device Category"].isin(selected_categories))
        & (last_year_units_df["Device"].isin(selected_devices))
    ].copy()
    ly_learners = last_year_learners_df[
        (last_year_learners_df["District"].isin(selected_districts))
        & (last_year_learners_df["Device Category"].isin(selected_categories))
        & (last_year_learners_df["Device"].isin(selected_devices))
    ].copy()

    last_year_total = int(ly_units["Requests"].sum()) if not ly_units.empty else 0
    last_year_students = int(ly_learners["Student ID"].dropna().nunique()) if "Student ID" in ly_learners else 0
    last_year_schools = int(ly_learners["School_Name"].dropna().nunique()) if "School_Name" in ly_learners else 0
    last_year_districts = int(ly_units["District"].dropna().nunique()) if not ly_units.empty else 0
    last_year_top_device = (
        display_device_name(ly_units.groupby("Device")["Requests"].sum().idxmax())
        if not ly_units.empty
        else "No data"
    )
    latest_refresh = datetime.now().strftime("%d %b %Y, %I:%M %p")

    st.markdown(
        f"""
            <div class="app-hero">
                <div class="eyebrow">Phase 1 collected devices</div>
                <h1 class="hero-title">Assistive Device Collection Dashboard</h1>
                <div class="status-row">
                    <span class="status-pill">Source: SharePoint workbook</span>
                    <span class="status-pill">Updated {safe_html(latest_refresh)}</span>
                    <span class="status-pill">{fmt_number(last_year_total)} device units</span>
                    <span class="status-pill">{fmt_number(last_year_students)} students</span>
                </div>
            </div>
        """,
        unsafe_allow_html=True,
    )

    if ly_units.empty:
        st.markdown(
            """
            <div class="empty-state">
                No Phase 1 collected device records match the current filters.
            </div>
            """,
            unsafe_allow_html=True,
        )
        st.stop()

    metric_1, metric_2, metric_3, metric_4 = st.columns(4)
    with metric_1:
        render_metric("Device units", fmt_number(last_year_total), "Phase 1 collected")
    with metric_2:
        render_metric("Students", fmt_number(last_year_students), "Consolidated records")
    with metric_3:
        render_metric("Schools", fmt_number(last_year_schools), f"{fmt_number(last_year_districts)} districts")
    with metric_4:
        render_metric("Top device", last_year_top_device, "By collected units")

    st.markdown('<div class="section-title">Last year overview</div>', unsafe_allow_html=True)
    district_counts = (
        ly_units.groupby("District", as_index=False)["Requests"].sum().sort_values("Requests", ascending=False).head(top_n)
    )
    device_counts = (
        ly_units.groupby("Device", as_index=False)["Requests"].sum().sort_values("Requests", ascending=False).head(top_n)
    )
    category_counts = (
        ly_units.groupby("Device Category", as_index=False)["Requests"].sum().sort_values("Requests", ascending=False)
    )

    left, right = st.columns(2)
    with left:
        st.markdown("#### District collection")
        st.plotly_chart(make_horizontal_bar(district_counts, "Requests", "District"), width="stretch")
    with right:
        st.markdown("#### Device collection")
        st.plotly_chart(make_horizontal_bar(device_counts, "Requests", "Device"), width="stretch")

    left, right = st.columns(2)
    with left:
        st.markdown("#### Device category distribution")
        fig = px.pie(
            category_counts,
            names="Device Category",
            values="Requests",
            color="Device Category",
            color_discrete_map=CATEGORY_COLORS,
            hole=0.58,
        )
        fig.update_traces(
            marker=dict(line=dict(color="#ffffff", width=2)),
            texttemplate="%{label}<br>%{percent}",
            textfont=dict(color=COLORS["ink"], size=13),
            hovertemplate="%{label}<br>%{value:,} device units<extra></extra>",
        )
        st.plotly_chart(style_chart(fig, height=350), width="stretch")
    with right:
        st.markdown("#### Disability profile")
        disability_counts = distinct_learner_counts_frame(ly_learners, "disability_cleaned", "Disability", top_n)
        st.plotly_chart(
            make_horizontal_bar(
                disability_counts,
                "Learners",
                "Disability",
                color_col="Learners",
                xaxis_title="Learners",
            ),
            width="stretch",
        )

    st.markdown("#### Phase 1 collected data")
    display_cols = [
        col
        for col in [
            "District",
            "School_Name",
            "Student ID",
            "Student Name",
            "disability_cleaned",
            "Device",
            "Device Category",
        ]
        if col in ly_learners.columns
    ]
    st.dataframe(ly_learners[display_cols].rename(columns={"disability_cleaned": "Disability"}), hide_index=True, width="stretch")
    st.stop()

analysis_scope = st.sidebar.selectbox(
    "Population",
    options=["Combined", "Schools", "Bedridden", "Institutes"],
    key="analysis_scope",
)

with st.sidebar.expander("KPI view", expanded=False):
    kpi_basis = st.selectbox(
        "Basis",
        options=["Auto", "Schools", "Bedridden", "Institutes"],
        key="kpi_basis",
        label_visibility="collapsed",
    )

source_base_df = {
    "Combined": people_df,
    "Schools": school_df,
    "Bedridden": bedridden_df,
    "Institutes": cdc_df.rename(columns={"Requests": "Request Count"}),
}[analysis_scope]
if dashboard_year == "Phases":
    source_base_df = pd.concat([source_base_df, last_year_learners_df], ignore_index=True, sort=False)

districts = sorted(source_base_df["District"].dropna().unique())
selected_districts = render_slicer("Districts", districts, "districts_filter", "Choose districts")

district_scope = source_base_df[source_base_df["District"].isin(selected_districts)] if selected_districts else source_base_df.iloc[0:0]
schools = sorted(district_scope["School_Name"].dropna().unique()) if analysis_scope in {"Combined", "Schools"} else []
current_district_scope = tuple(selected_districts)
if st.session_state.get("_district_scope") != current_district_scope:
    st.session_state["schools_filter"] = schools
    st.session_state["_district_scope"] = current_district_scope
selected_schools = (
    render_slicer("Schools", schools, "schools_filter", "Choose schools")
    if analysis_scope in {"Combined", "Schools"}
    else []
)

categories = sorted(source_base_df["Device Category"].dropna().unique())
selected_categories = render_slicer(
    "Device categories",
    categories,
    "categories_filter",
    "Choose categories",
)

devices = sorted(source_base_df["Device"].dropna().unique())
selected_devices = render_slicer("Devices", devices, "devices_filter", "Choose devices")

institutes = sorted(cdc_df["Institute"].dropna().unique()) if not cdc_df.empty else []
selected_institutes = (
    render_slicer("Institutes", institutes, "institutes_filter", "Choose institutes")
    if institutes
    else []
)

priorities = sorted(source_base_df["Priority"].dropna().unique(), key=str)
selected_priorities = render_slicer("Priorities", priorities, "priorities_filter", "Choose priorities")

genders = sorted(source_base_df["Gender"].dropna().unique())
selected_genders = render_slicer("Genders", genders, "genders_filter", "Choose genders")

with st.sidebar.expander("Ranked rows", expanded=False):
    top_n = st.slider("Rows per chart", min_value=5, max_value=20, value=10, step=1)

school_filtered_df = school_df[
    (school_df["District"].isin(selected_districts))
    & (school_df["School_Name"].isin(selected_schools if selected_schools else school_df["School_Name"].dropna().unique()))
    & (school_df["Device Category"].isin(selected_categories))
    & (school_df["Device"].isin(selected_devices))
    & (school_df["Priority"].isin(selected_priorities))
    & (school_df["Gender"].isin(selected_genders))
].copy()

bedridden_filtered_df = bedridden_df[
    (bedridden_df["District"].isin(selected_districts))
    & (bedridden_df["Device Category"].isin(selected_categories))
    & (bedridden_df["Device"].isin(selected_devices))
    & (bedridden_df["Priority"].isin(selected_priorities))
    & (bedridden_df["Gender"].isin(selected_genders))
].copy()

filtered_df = {
    "Combined": pd.concat([school_filtered_df, bedridden_filtered_df], ignore_index=True, sort=False),
    "Schools": school_filtered_df,
    "Bedridden": bedridden_filtered_df,
    "Institutes": school_df.iloc[0:0],
}[analysis_scope].copy()
size_chart_df = build_size_chart_data(filtered_df, catalog)

cdc_filtered = cdc_df[
    (cdc_df["District"].isin(selected_districts))
    & (cdc_df["Device Category"].isin(selected_categories))
    & (cdc_df["Device"].isin(selected_devices))
    & (cdc_df["Institute"].isin(selected_institutes if selected_institutes else institutes))
].copy()

include_last_year_units = (
    dashboard_year != "Phases"
    or ("Collected" in set(selected_priorities) and "Unknown" in set(selected_genders))
)
last_year_current_filter = (
    last_year_units_df[
        (last_year_units_df["District"].isin(selected_districts))
        & (last_year_units_df["Device Category"].isin(selected_categories))
        & (last_year_units_df["Device"].isin(selected_devices))
    ].copy()
    if include_last_year_units
    else last_year_units_df.iloc[0:0].copy()
)
last_year_filtered_requests = (
    int(last_year_current_filter["Requests"].sum()) if not last_year_current_filter.empty else 0
)
last_year_profile_filter = last_year_learners_df[
    (last_year_learners_df["District"].isin(selected_districts))
    & (last_year_learners_df["School_Name"].isin(selected_schools if selected_schools else last_year_learners_df["School_Name"].dropna().unique()))
    & (last_year_learners_df["Device Category"].isin(selected_categories))
    & (last_year_learners_df["Device"].isin(selected_devices))
    & (last_year_learners_df["Priority"].isin(selected_priorities))
    & (last_year_learners_df["Gender"].isin(selected_genders))
].copy()

current_demand_df = filtered_df.copy()
if not current_demand_df.empty:
    current_demand_df["Requests"] = 1
    current_demand_df["Phase"] = "Phase 2"
institute_demand_df = cdc_filtered.copy()
if not institute_demand_df.empty:
    institute_demand_df["Priority"] = "Institute"
    institute_demand_df["Phase"] = "Phase 2"
last_year_demand_df = last_year_current_filter.copy()
if not last_year_demand_df.empty:
    last_year_demand_df["Priority"] = "Phase 1 collected"
    last_year_demand_df["Phase"] = "Phase 1"

all_year_demand_df = (
    pd.concat(
        [
            current_demand_df,
            institute_demand_df if analysis_scope in {"Combined", "Institutes"} else institute_demand_df.iloc[0:0],
            last_year_demand_df,
        ],
        ignore_index=True,
        sort=False,
    )
    if dashboard_year == "Phases"
    else pd.DataFrame()
)
all_year_profile_df = (
    pd.concat([filtered_df, last_year_profile_filter], ignore_index=True, sort=False)
    if dashboard_year == "Phases"
    else filtered_df
)
demand_view_df = all_year_demand_df if dashboard_year == "Phases" else filtered_df
profile_view_df = all_year_profile_df if dashboard_year == "Phases" else filtered_df
all_year_total_requests = int(all_year_demand_df["Requests"].sum()) if not all_year_demand_df.empty else 0

cdc_requests = int(cdc_filtered["Requests"].sum()) if not cdc_filtered.empty else 0

school_requests = len(school_filtered_df)
bedridden_requests = len(bedridden_filtered_df)
analysis_requests = len(filtered_df)

analysis_device_counts = filtered_df["Device"].value_counts() if not filtered_df.empty else pd.Series(dtype="int64")
school_device_counts = school_filtered_df["Device"].value_counts() if not school_filtered_df.empty else pd.Series(dtype="int64")
bedridden_device_counts = bedridden_filtered_df["Device"].value_counts() if not bedridden_filtered_df.empty else pd.Series(dtype="int64")
cdc_device_counts = cdc_filtered.groupby("Device")["Requests"].sum() if not cdc_filtered.empty else pd.Series(dtype="int64")
combined_device_counts = analysis_device_counts.add(cdc_device_counts, fill_value=0).sort_values(ascending=False)
institute_kpi_df = cdc_filtered.copy()
institute_requests = int(institute_kpi_df["Requests"].sum()) if not institute_kpi_df.empty else 0
institute_device_counts = (
    institute_kpi_df.groupby("Device")["Requests"].sum().sort_values(ascending=False)
    if not institute_kpi_df.empty
    else pd.Series(dtype="int64")
)
institute_count = int(institute_kpi_df["Institute"].nunique()) if not institute_kpi_df.empty else 0
institute_districts = int(institute_kpi_df["District"].nunique()) if not institute_kpi_df.empty else 0

total_schools = filtered_df["School_Name"].dropna().nunique()
total_districts = filtered_df["District"].nunique()
total_bedridden_people = filtered_df.loc[filtered_df["Record Type"] == "Bedridden", "Student Name"].dropna().nunique()
analysis_priority_numeric = pd.to_numeric(filtered_df["Priority"], errors="coerce")
priority_one = int((analysis_priority_numeric == 1).sum()) if analysis_requests else 0
school_priority_one = int((pd.to_numeric(school_filtered_df["Priority"], errors="coerce") == 1).sum()) if school_requests else 0
bedridden_priority_one = int((pd.to_numeric(bedridden_filtered_df["Priority"], errors="coerce") == 1).sum()) if bedridden_requests else 0
school_filter_active = bool(schools) and set(selected_schools) != set(schools)

if kpi_basis == "Institutes":
    total_requests = institute_requests
    top_device = display_device_name(institute_device_counts.idxmax()) if not institute_device_counts.empty else "No data"
    scope_count_value = institute_count
    scope_count_label = "institutes"
    priority_value = institute_districts
elif kpi_basis == "Schools":
    total_requests = school_requests
    top_device = (
        display_device_name(school_device_counts.idxmax())
        if not school_device_counts.empty
        else "No data"
    )
    scope_count_value = school_filtered_df["School_Name"].dropna().nunique()
    scope_count_label = "schools"
    priority_value = school_priority_one
elif kpi_basis == "Bedridden":
    total_requests = bedridden_requests
    top_device = (
        display_device_name(bedridden_device_counts.idxmax())
        if not bedridden_device_counts.empty
        else "No data"
    )
    scope_count_value = bedridden_filtered_df["Student Name"].dropna().nunique()
    scope_count_label = "bedridden records"
    priority_value = bedridden_priority_one
else:
    if analysis_scope == "Institutes":
        total_requests = institute_requests
        active_device_counts = institute_device_counts
        top_device = display_device_name(active_device_counts.idxmax()) if not active_device_counts.empty else "No data"
        scope_count_value = institute_count
        scope_count_label = "institutes"
    else:
        total_requests = analysis_requests + (institute_requests if analysis_scope == "Combined" else 0)
        active_device_counts = combined_device_counts if analysis_scope == "Combined" else analysis_device_counts
        top_device = display_device_name(active_device_counts.idxmax()) if not active_device_counts.empty else "No data"
        if analysis_scope == "Bedridden":
            scope_count_value = total_bedridden_people
            scope_count_label = "bedridden records"
        elif analysis_scope == "Schools":
            scope_count_value = school_filtered_df["School_Name"].dropna().nunique()
            scope_count_label = "schools"
        else:
            scope_count_value = filtered_df["Student Name"].dropna().nunique()
            scope_count_label = "records"
    priority_value = priority_one

if dashboard_year == "Phases":
    total_requests = all_year_total_requests
    display_total_requests = all_year_total_requests
    all_year_device_counts = demand_counts_frame(all_year_demand_df, "Device", "Device")
    all_year_district_counts = demand_counts_frame(all_year_demand_df, "District", "District")
    top_device = display_device_name(all_year_device_counts.iloc[0]["Device"]) if not all_year_device_counts.empty else "No data"
    top_district = all_year_district_counts.iloc[0]["District"] if not all_year_district_counts.empty else "No data"
    scope_count_value = int(learner_key_frame(profile_view_df)["_Learner Key"].nunique()) if not profile_view_df.empty else 0
    scope_count_label = "learners"
    total_districts = int(all_year_demand_df["District"].dropna().nunique()) if not all_year_demand_df.empty else 0
else:
    display_total_requests = total_requests
    top_district = (
        (cdc_filtered["District"].value_counts().idxmax() if not cdc_filtered.empty else "No data")
        if analysis_scope == "Institutes"
        else (
            filtered_df["District"].value_counts().idxmax()
            if analysis_requests and not filtered_df["District"].dropna().empty
            else "No data"
        )
    )
latest_refresh = datetime.now().strftime("%d %b %Y, %I:%M %p")
status_scope_pill = (
    f'<span class="status-pill">{fmt_number(institute_count)} institutes</span>'
    if kpi_basis == "Institutes" or analysis_scope == "Institutes"
    else f'<span class="status-pill">{fmt_number(scope_count_value)} {scope_count_label}</span>'
)

st.markdown(
    f"""
        <div class="app-hero">
            <div class="eyebrow">Assistive device dashboard</div>
            <h1 class="hero-title">Assistive Device Demand Dashboard</h1>
            <div class="status-row">
                <span class="status-pill">Refresh: hourly</span>
                <span class="status-pill">Updated {safe_html(latest_refresh)}</span>
                <span class="status-pill">{fmt_number(display_total_requests)} total devices</span>
                {status_scope_pill}
            </div>
        </div>
    """,
    unsafe_allow_html=True,
)

if (
    dashboard_year == "Phases"
    and all_year_demand_df.empty
) or (
    dashboard_year != "Phases"
    and ((analysis_scope == "Institutes" and cdc_filtered.empty) or (analysis_scope != "Institutes" and filtered_df.empty and cdc_filtered.empty))
):
    st.markdown(
        """
        <div class="empty-state">
            No records match the current filters.
        </div>
        """,
        unsafe_allow_html=True,
    )
    st.stop()

metric_1, metric_2, metric_3, metric_4 = st.columns(4)
with metric_1:
    if dashboard_year == "Phases":
        current_year_count = all_year_total_requests - last_year_filtered_requests
        device_kpi_note = (
            f"{fmt_number(current_year_count)} Phase 2 + {fmt_number(last_year_filtered_requests)} Phase 1"
        )
    else:
        device_kpi_note = "Selected Phase 2 demand"
    if kpi_basis == "Institutes":
        render_metric(
            "Device count",
            fmt_number(display_total_requests),
            device_kpi_note,
        )
    elif kpi_basis == "Bedridden":
        render_metric("Device count", fmt_number(display_total_requests), device_kpi_note)
    elif kpi_basis == "Schools":
        render_metric("Device count", fmt_number(display_total_requests), device_kpi_note)
    else:
        render_metric("Device count", fmt_number(display_total_requests), device_kpi_note)
with metric_2:
    if dashboard_year == "Phases":
        render_metric("Learners covered", fmt_number(scope_count_value), f"{fmt_number(total_districts)} districts")
    elif kpi_basis == "Institutes":
        render_metric("Institutes covered", fmt_number(institute_count), f"{fmt_number(institute_districts)} districts")
    elif kpi_basis == "Bedridden":
        render_metric("Individuals covered", fmt_number(scope_count_value), f"{fmt_number(total_districts)} districts")
    elif kpi_basis == "Auto" and analysis_scope == "Institutes":
        render_metric("Institutes covered", fmt_number(scope_count_value), f"{fmt_number(institute_districts)} districts")
    elif kpi_basis == "Auto" and analysis_scope == "Bedridden":
        render_metric("Individuals covered", fmt_number(scope_count_value), f"{fmt_number(total_districts)} districts")
    elif kpi_basis == "Auto" and analysis_scope == "Combined":
        render_metric("Records covered", fmt_number(scope_count_value), f"{fmt_number(total_districts)} districts")
    else:
        render_metric("Schools covered", fmt_number(total_schools), f"{fmt_number(total_districts)} districts")
with metric_3:
    if dashboard_year == "Phases":
        render_metric("Phases combined", "2", "Phase 2 + Phase 1")
    elif kpi_basis == "Institutes":
        render_metric("Institute districts", fmt_number(institute_districts), "Filtered institute locations")
    elif kpi_basis == "Bedridden":
        render_metric(
            "Priority 1 needs",
            fmt_number(priority_value),
            f"{fmt_percent(priority_value, bedridden_requests)} of bedridden requests.",
        )
    elif kpi_basis == "Auto" and analysis_scope == "Institutes":
        render_metric("Institute districts", fmt_number(institute_districts), "Filtered institute locations")
    else:
        render_metric(
            "Priority 1 needs",
            fmt_number(priority_value),
            f"{fmt_percent(priority_value, total_requests)} of selected requests.",
        )
with metric_4:
    if dashboard_year == "Phases":
        render_metric("Most needed device", top_device, "Combined phase demand")
    elif kpi_basis == "Institutes":
        render_metric("Most needed device", top_device, "Institute demand")
    elif kpi_basis == "Bedridden":
        render_metric("Most needed device", top_device, "Bedridden demand")
    elif kpi_basis == "Schools":
        render_metric("Most needed device", top_device, "School demand")
    else:
        render_metric(
            "Most needed device",
            top_device,
            "Combined school, bedridden, and institute demand"
            if analysis_scope == "Combined"
            else ("Institute demand" if analysis_scope == "Institutes" else f"{analysis_scope.lower()} demand"),
        )

st.markdown('<div class="section-title">Executive focus</div>', unsafe_allow_html=True)

learner_gender_counts = distinct_learner_counts_frame(profile_view_df, "Gender", "Gender")
female_count = int(learner_gender_counts.loc[learner_gender_counts["Gender"] == "Female", "Learners"].sum())
male_count = int(learner_gender_counts.loc[learner_gender_counts["Gender"] == "Male", "Learners"].sum())
learner_total = int(learner_key_frame(profile_view_df)["_Learner Key"].nunique()) if not profile_view_df.empty else 0
category_counts = demand_counts_frame(demand_view_df, "Device Category", "Category")
leading_category = category_counts.iloc[0]["Category"] if not category_counts.empty else "No data"
leading_category_count = int(category_counts.iloc[0]["Requests"]) if not category_counts.empty else 0
district_counts_focus = demand_counts_frame(demand_view_df, "District", "District")
district_count = int(district_counts_focus.iloc[0]["Requests"]) if not district_counts_focus.empty else 0

focus_1, focus_2, focus_3 = st.columns(3)
with focus_1:
    render_insight(
        "Demand concentration",
        top_district,
        f"{fmt_number(district_count)} requests",
    )
with focus_2:
    render_insight(
        "Leading category",
        leading_category,
        f"{fmt_percent(leading_category_count, total_requests)} of selected demand.",
    )
with focus_3:
    render_insight(
        "Gender distribution",
        f"{fmt_number(male_count)} male / {fmt_number(female_count)} female",
        f"Female share: {fmt_percent(female_count, learner_total)}.",
    )

overview_tab, institutes_tab, bedridden_tab, profile_tab, size_tab, data_tab = st.tabs(
    ["Demand overview", "Institutes", "Bedridden", "Learner profile", "Size chart", "Filtered data"]
)

with overview_tab:
    st.markdown('<div class="section-title">Demand overview</div>', unsafe_allow_html=True)

    district_counts = demand_counts_frame(demand_view_df, "District", "District", top_n)
    device_counts = demand_counts_frame(demand_view_df, "Device", "Device", top_n)

    left, right = st.columns(2)
    with left:
        st.markdown("#### District demand")
        st.plotly_chart(make_horizontal_bar(district_counts, "Requests", "District"), width="stretch")

    with right:
        st.markdown("#### Device demand")
        st.plotly_chart(make_horizontal_bar(device_counts, "Requests", "Device"), width="stretch")

    priority_counts = demand_counts_frame(demand_view_df, "Priority", "Priority")
    priority_counts["Priority"] = priority_counts["Priority"].astype(str)
    priority_counts["Priority label"] = priority_counts["Priority"].map(
        lambda value: "Common requirement" if str(value) == "Common" else f"Priority {value}"
    )
    priority_counts = priority_counts.sort_values(
        "Priority",
        key=lambda s: s.map(lambda value: 4 if str(value) == "Common" else pd.to_numeric(value, errors="coerce")),
    )

    category_counts = demand_counts_frame(demand_view_df, "Device Category", "Category")

    left, right = st.columns(2)
    with left:
        st.markdown("#### Priority Distribution")
        fig = px.bar(
            priority_counts,
            x="Priority label",
            y="Requests",
            color="Priority",
            color_discrete_map=PRIORITY_COLORS,
            text="Requests",
            hover_data={"Requests": ":,", "Priority": False, "Priority label": True},
        )
        fig.update_traces(texttemplate="%{text:,}", textposition="outside", marker_line_width=0)
        fig.update_layout(xaxis_title=None, yaxis_title="Requests", showlegend=False)
        st.plotly_chart(style_chart(fig, height=350), width="stretch")

    with right:
        st.markdown("#### Device category distribution")
        fig = px.pie(
            category_counts,
            names="Category",
            values="Requests",
            color="Category",
            color_discrete_map=CATEGORY_COLORS,
            hole=0.58,
        )
        fig.update_traces(
            marker=dict(line=dict(color="#ffffff", width=2)),
            texttemplate="%{label}<br>%{percent}",
            textfont=dict(color=COLORS["ink"], size=13),
            hovertemplate="%{label}<br>%{value:,} requests<extra></extra>",
        )
        st.plotly_chart(style_chart(fig, height=350), width="stretch")

with bedridden_tab:
    st.markdown('<div class="section-title">Bedridden</div>', unsafe_allow_html=True)

    if bedridden_filtered_df.empty:
        st.markdown(
            """
            <div class="empty-state">
                No bedridden records match the current filters.
            </div>
            """,
            unsafe_allow_html=True,
        )
    else:
        bedridden_people = bedridden_filtered_df["Student Name"].dropna().nunique()
        bedridden_top_device = display_device_name(
            bedridden_device_counts.idxmax() if not bedridden_device_counts.empty else "No data"
        )

        bed_1, bed_2, bed_3 = st.columns(3)
        with bed_1:
            render_insight("Individuals", fmt_number(bedridden_people), "Bedridden records")
        with bed_2:
            render_insight("Districts", fmt_number(bedridden_filtered_df["District"].nunique()), "Current filters")
        with bed_3:
            render_insight("Top device", bedridden_top_device, f"{fmt_number(bedridden_requests)} total requests")

        left, right = st.columns(2)
        with left:
            st.markdown("#### Bedridden district demand")
            district_counts = value_counts_frame(bedridden_filtered_df, "District", "District", top_n)
            st.plotly_chart(make_horizontal_bar(district_counts, "Requests", "District"), width="stretch")
        with right:
            st.markdown("#### Bedridden device demand")
            device_counts = value_counts_frame(bedridden_filtered_df, "Device", "Device", top_n)
            st.plotly_chart(make_horizontal_bar(device_counts, "Requests", "Device"), width="stretch")

        st.markdown("#### Cleaned bedridden data")
        bedridden_display = bedridden_filtered_df[
            ["Name", "Age", "Gender", "Contact No", "Address", "District", "disability_cleaned", "Device", "Other requirement"]
        ].rename(columns={"disability_cleaned": "Disability"})
        st.dataframe(bedridden_display, hide_index=True, width="stretch")
        st.download_button(
            label="Download cleaned bedridden CSV",
            data=bedridden_display.to_csv(index=False).encode("utf-8"),
            file_name="cleaned_bedridden_data.csv",
            mime="text/csv",
        )

with institutes_tab:
    st.markdown('<div class="section-title">Institutes</div>', unsafe_allow_html=True)

    if cdc_filtered.empty:
        st.markdown(
            """
            <div class="empty-state">
                No institute records match the current district, category, or device filters.
            </div>
            """,
            unsafe_allow_html=True,
        )
    else:
        institute_count = cdc_filtered["Institute"].nunique()
        institute_requests = int(cdc_filtered["Requests"].sum())
        institute_top_device = display_device_name(
            cdc_filtered.sort_values("Requests", ascending=False).iloc[0]["Device"]
        )

        ins_1, ins_2, ins_3 = st.columns(3)
        with ins_1:
            render_insight("Institutes", fmt_number(institute_count), "Included in this view")
        with ins_2:
            render_insight("District", "Trivandrum", "Institute location")
        with ins_3:
            render_insight("Device requests", fmt_number(institute_requests), f"Top device: {institute_top_device}")

        institute_device_counts = cdc_filtered.sort_values("Requests", ascending=False).copy()
        institute_device_counts["Device"] = institute_device_counts["Device"].map(display_device_name)

        left, right = st.columns([1.15, 0.85])
        with left:
            st.markdown("#### Institute device demand")
            chart_height = max(320, min(620, 120 + (len(institute_device_counts) * 34)))
            st.plotly_chart(
                make_horizontal_bar(institute_device_counts, "Requests", "Device", height=chart_height),
                width="stretch",
            )

        with right:
            st.markdown("#### Institute reference")
            st.dataframe(
                cdc_filtered[["Institute", "District", "Device", "Device Category", "Requests"]]
                .assign(Device=lambda frame: frame["Device"].map(display_device_name))
                .sort_values("Requests", ascending=False),
                hide_index=True,
                width="stretch",
            )

with profile_tab:
    st.markdown('<div class="section-title">Learner profile</div>', unsafe_allow_html=True)

    disability_counts = distinct_learner_counts_frame(profile_view_df, "disability_cleaned", "Disability", top_n)
    social_counts = distinct_learner_counts_frame(profile_view_df, "Social Category", "Social category", top_n)

    left, right = st.columns(2)
    with left:
        st.markdown("#### Disability profile")
        st.plotly_chart(
            make_horizontal_bar(
                disability_counts,
                "Learners",
                "Disability",
                color_col="Learners",
                xaxis_title="Learners",
            ),
            width="stretch",
        )

    with right:
        st.markdown("#### Social category")
        st.plotly_chart(
            make_horizontal_bar(
                social_counts,
                "Learners",
                "Social category",
                color_col="Learners",
                xaxis_title="Learners",
            ),
            width="stretch",
        )

    left, right = st.columns(2)
    with left:
        st.markdown("#### Gender distribution")
        fig = px.pie(
            learner_gender_counts,
            names="Gender",
            values="Learners",
            color="Gender",
            color_discrete_map=GENDER_COLORS,
            hole=0.58,
        )
        fig.update_traces(
            marker=dict(line=dict(color="#ffffff", width=2)),
            texttemplate="%{label}<br>%{percent}",
            textfont=dict(color=COLORS["ink"], size=13),
            hovertemplate="%{label}<br>%{value:,} learners<extra></extra>",
        )
        st.plotly_chart(style_chart(fig, height=350), width="stretch")

with size_tab:
    st.markdown('<div class="section-title">Size chart</div>', unsafe_allow_html=True)

    if not catalog:
        st.markdown(
            f"""
            <div class="empty-state">
                Device catalog not found at {safe_html(str(CATALOG_PATH))}.
            </div>
            """,
            unsafe_allow_html=True,
        )
    elif size_chart_df.empty:
        st.markdown(
            """
            <div class="empty-state">
                No size records are available for the current filter selection.
            </div>
            """,
            unsafe_allow_html=True,
        )
    else:
        size_device_count = size_chart_df["Device"].nunique()
        size_code_count = size_chart_df["Unique Device Code"].replace("", pd.NA).dropna().nunique()
        top_size_row = size_chart_df.sort_values("Requests", ascending=False).iloc[0]

        size_metric_1, size_metric_2, size_metric_3 = st.columns(3)
        with size_metric_1:
            render_insight(
                "Sized devices",
                fmt_number(size_device_count),
                "Current filters",
            )
        with size_metric_2:
            render_insight(
                "Catalog codes",
                fmt_number(size_code_count),
                "Unique codes",
            )
        with size_metric_3:
            render_insight(
                "Top size demand",
                top_size_row["Size ID"],
                f"{top_size_row['Device']} with {fmt_number(top_size_row['Requests'])} requests.",
            )

        device_options = sorted(size_chart_df["Device"].dropna().unique())
        selected_size_device = st.selectbox(
            "Device",
            options=device_options,
            index=0,
        )

        selected_device_sizes = (
            size_chart_df[size_chart_df["Device"] == selected_size_device]
            .sort_values("Requests", ascending=False)
            .copy()
        )

        left, right = st.columns([1.15, 0.85])
        with left:
            st.markdown(f"#### {selected_size_device} size demand")
            chart_height = max(340, min(680, 110 + (len(selected_device_sizes) * 34)))
            st.plotly_chart(
                make_horizontal_bar(
                    selected_device_sizes,
                    "Requests",
                    "Size Label",
                    height=chart_height,
                ),
                width="stretch",
            )

        with right:
            st.markdown("#### Catalog reference")
            display_columns = [
                "Size ID",
                "Size Code",
                "Unique Device Code",
                "Size System",
                "Catalog Length (cm)",
                "Catalog Width (cm)",
                "Requests",
            ]
            st.dataframe(
                selected_device_sizes[display_columns],
                hide_index=True,
                width="stretch",
            )

        st.markdown("#### All device-size requirements")
        all_size_display = size_chart_df[
            [
                "Device",
                "Size ID",
                "Size Code",
                "Size Label",
                "Unique Device Code",
                "Requests",
                "Avg Palm Length (cm)",
                "Avg Palm Width (cm)",
            ]
        ].sort_values(["Device", "Requests"], ascending=[True, False])
        st.dataframe(all_size_display, hide_index=True, width="stretch")

with data_tab:
    st.markdown('<div class="section-title">Filtered data</div>', unsafe_allow_html=True)

    data_display_df = all_year_demand_df if dashboard_year == "Phases" else filtered_df
    st.dataframe(data_display_df, hide_index=True, width="stretch")

    csv = data_display_df.to_csv(index=False).encode("utf-8")
    st.download_button(
        label="Download filtered CSV",
        data=csv,
        file_name="filtered_device_data.csv",
        mime="text/csv",
    )
