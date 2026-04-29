import re

import pandas as pd

SCHOOL_SHEET_URL = "https://docs.google.com/spreadsheets/d/e/2PACX-1vTh9RBFi-hMFzAmCekGFnHjhgn0ZMTZVuqSHYHsV3qanHKbXN29QMPHKBGFDkKS_ioXsrJLU-zvgBSV/pub?gid=139800674&single=true&output=csv"
BEDRIDDEN_SHEET_URL = "https://docs.google.com/spreadsheets/d/e/2PACX-1vSyfjs_rI1G74Zlgzr1J6ty2O6FMWMU0NKqwnYOkOLx9UMQRet9br945ikiLowAxL6o5-mVOiSlBHzA/pub?output=csv"


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

DEVICE_MAP = {
    "wheel chair": "wheelchair",
    "wheel-chair": "wheelchair",
    "hearing aid device": "hearing aid",
    "hearing machine": "hearing aid",
    "walking stick": "walking aid",
    "walker": "walking aid",
    "low switch profile": "low profile switch",
    "tooth brush holder": "toothbrush holder",
}


def normalize_device_name(value):
    device = str(value).strip().lower()
    if not device or device in {"nan", "not applicable", "none"}:
        return pd.NA
    return DEVICE_MAP.get(device, device)


def split_common_requirements(value):
    text = str(value).strip()
    if not text or text.lower() in {"nan", "none", "not applicable"}:
        return []
    return [
        part.strip()
        for part in re.split(r"[;,/]\s*|\n+", text)
        if part.strip()
    ]


def clean_disability(text):
    text = str(text).strip().lower()

    if "cp" in text or "cerebral" in text:
        return "Cerebral Palsy"
    if "id" in text or "intellectual" in text or "idd" in text or "ied" in text or "mr" in text or "mental" in text:
        return "Intellectual Disability"
    if "vision" in text or "visual" in text or "blind" in text:
        return "Visual Impairment"
    if "hearing" in text or "deaf" in text:
        return "Hearing Impairment"
    if "speech" in text or "dumb" in text or "communication" in text:
        return "Speech Impairment"
    if "learning" in text or "ld" in text:
        return "Learning Disability"
    if "autism" in text or "asd" in text:
        return "Autism"
    if "adhd" in text or "attention" in text:
        return "ADHD"
    if "down syndrome" in text or "ds" in text or "downs" in text or "down's" in text or "down" in text:
        return "Down Syndrome"
    if "dwarfism" in text or "dwarf" in text:
        return "Dwarfism"
    if "seizure" in text or "epilepsy" in text:
        return "Epilepsy"
    if "neurological" in text or "glutoria" in text:
        return "Neurological Condition"
    if "bed ridden" in text or "bedridden" in text:
        return "Bedridden"
    if "emotionally unstable" in text or "emotional disorder" in text:
        return "Emotional Disorder"
    if "global development delay" in text or "gdd" in text:
        return "Global Developmental Delay"
    if "locomotor" in text:
        return "Locomotor Disability"
    return "Other"


def clean_gender_series(series):
    gender_map = {
        "male": "Male",
        "female": "Female",
        "m": "Male",
        "f": "Female",
        "nale": "Male",
        "nalee": "Male",
        "femal": "Female",
    }
    normalized = (
        series.astype(str)
        .str.strip()
        .str.lower()
        .str.extract(r"(male|female|m|f|nale|nalee|femal)", expand=False)
    )
    return normalized.map(gender_map).fillna("Other")


def clean_district_series(series):
    cleaned = series.astype(str).str.split("/").str[0].str.strip().str.title()
    district_map = {"Thiruvananthapuram": "Trivandrum", "Tvm": "Trivandrum"}
    return cleaned.replace(district_map)

def load_and_clean_data(path= None):
    # ---------------- LOAD DATA ----------------
    df = pd.read_csv(SCHOOL_SHEET_URL)
    df.columns = df.columns.str.strip()
    

# Remove extra description inside brackets
    df.columns = df.columns.str.replace(r"\(.*\)", "", regex=True)

# Remove newlines
    df.columns = df.columns.str.replace("\n", " ")

    df.columns = df.columns.str.strip()

    # ---------------- USE EXISTING SCHOOL CODE ----------------
    df['School ID'] = df['School ID'].astype(str).str.strip().str.upper()
    school_id = "School ID"   
    # Create standard school name using School ID
    df['School Name'] = df['School Name'].astype(str).str.strip().str.title() 
    school_map = df.groupby(school_id)['School Name'].agg(lambda x: x.value_counts().index[0])
    
# Apply standardized name
    df['School_Name'] = df[school_id].map(school_map)

    # ---------------- CLEAN MEASUREMENTS ----------------
    df['Palm Width'] = pd.to_numeric(df['Palm Width'], errors='coerce')
    df['Palm Length'] = pd.to_numeric(df['Palm Length'], errors='coerce')

    # Keep measurements within plausible human-hand ranges while preserving true sub-2 cm widths.
    df['Palm Width Cleaned'] = df['Palm Width'].where(df['Palm Width'].between(1.0, 5.0)).round(2)
    df['Palm Length Cleaned'] = df['Palm Length'].where(df['Palm Length'].between(4.0, 15.0)).round(2)

    #cleaning district names and standardizing them
    df['District'] = clean_district_series(df['District'])

    #cleaning gender values and standardizing them
    df['Gender'] = clean_gender_series(df['Gender'])

    
    # ---------------- CLEAN DISABILITY (NOT USED IN DASHBOARD) ----------------
    disability_col = [col for col in df.columns if "primary disability" in col.lower()][0]

    df["disability_cleaned"] = df[disability_col].apply(clean_disability) 
    
    
    df["disability_cleaned"].value_counts()

    # ---------------- CLEAN DEVICE NAMES ----------------
    
    # ---------------- DEVICE RESHAPING ----------------
    device_cols = ['Device Priority 1', 'Device Priority 2', 'Device Priority 3']
    common_requirement_col = next(
        (
            col
            for col in df.columns
            if col.strip().lower() in {"common requirement", "common requirements"}
        ),
        None,
    )

    df_devices = df.melt(
        id_vars=[
            'District',
            'School_Name',
            school_id,
            'Student Name',
            'Gender',
            'Social Category',
            'disability_cleaned',
            'Palm Width Cleaned',
            'Palm Length Cleaned',
        ],
        value_vars=device_cols,
        var_name="Priority",
        value_name='Device'
    )
    df_devices["Priority"] = df_devices["Priority"].str.replace("Device Priority ", "")
    df_devices["Priority"] = df_devices["Priority"].replace({"Common Requirements": "Common"})
    df_devices['Device'] = df_devices['Device'].map(normalize_device_name)
    df_devices = df_devices[df_devices['Device'].notna()&(df_devices['Device'] != '')&(df_devices['Device'] != 'nan')&(df_devices['Device'] != 'none')]

    if common_requirement_col:
        common_rows = df[
            [
                'District',
                'School_Name',
                school_id,
                'Student Name',
                'Gender',
                'Social Category',
                'disability_cleaned',
                'Palm Width Cleaned',
                'Palm Length Cleaned',
                common_requirement_col,
            ]
        ].copy()
        common_rows[common_requirement_col] = common_rows[common_requirement_col].apply(split_common_requirements)
        common_rows = common_rows.explode(common_requirement_col)
        common_rows = common_rows.rename(columns={common_requirement_col: "Device"})
        common_rows["Priority"] = "Common"
        common_rows["Device"] = common_rows["Device"].map(normalize_device_name)
        common_rows = common_rows[common_rows["Device"].notna()]
        df_devices = pd.concat([df_devices, common_rows[df_devices.columns]], ignore_index=True)

    measurement_based_devices = {
        'utensil holder',
        'palm pen holder',
        'toothbrush holder',
    }
    fixed_measurement_mask = ~df_devices['Device'].isin(measurement_based_devices)
    df_devices.loc[fixed_measurement_mask, ['Palm Width Cleaned', 'Palm Length Cleaned']] = pd.NA

    
    df_devices['Device Category'] = df_devices['Device'].map(DEVICE_CATEGORY_MAP)

    # ---------------- SAVE CLEANED FILE ----------------
    try:
        df_devices.to_csv("data/cleaned_data.csv", index=False)
    except PermissionError:
        pass

    return df_devices


def load_and_clean_bedridden_data(path=None):
    df = pd.read_csv(path or BEDRIDDEN_SHEET_URL)
    df.columns = df.columns.str.strip()

    name_col = next(col for col in df.columns if col.lower().startswith("name/"))
    age_col = next(col for col in df.columns if col.lower().startswith("age/"))
    gender_col = next(col for col in df.columns if col.lower().startswith("gender/"))
    address_col = next(col for col in df.columns if col.lower().startswith("address/"))
    contact_col = next(col for col in df.columns if col.lower().startswith("contact number/"))
    condition_col = next(col for col in df.columns if col.lower().startswith("primary medical condition/"))
    district_col = next(col for col in df.columns if col.lower().startswith("district/"))
    pref_3_col = next(col for col in df.columns if "perference 3" in col.lower() or "preference 3" in col.lower())
    other_req_col = next(col for col in df.columns if col.strip().lower() == "other requirement")

    df["District"] = clean_district_series(df[district_col])
    df["Name"] = df[name_col].astype(str).str.strip()
    df["Age"] = pd.to_numeric(df[age_col], errors="coerce")
    df["Gender"] = clean_gender_series(df[gender_col])
    df["Contact No"] = df[contact_col].astype(str).str.strip()
    df["Address"] = df[address_col].astype(str).str.strip()
    df["disability_cleaned"] = df[condition_col].apply(clean_disability)
    df["Other requirement"] = df[other_req_col].replace({"nan": pd.NA}).astype("string")

    preference_map = {
        "Preference 1": "1",
        "Preference 2": "2",
        pref_3_col: "3",
    }
    preference_cols = list(preference_map.keys())
    melted = df.melt(
        id_vars=[
            "District",
            "Name",
            "Age",
            "Gender",
            "Contact No",
            "Address",
            "disability_cleaned",
            "Other requirement",
        ],
        value_vars=preference_cols,
        var_name="Priority",
        value_name="Device",
    )
    melted["Priority"] = melted["Priority"].map(preference_map)
    melted["Device"] = melted["Device"].map(normalize_device_name)
    melted = melted[melted["Device"].notna()].copy()
    melted["Device Category"] = melted["Device"].map(DEVICE_CATEGORY_MAP)
    melted["School_Name"] = pd.NA
    melted["School ID"] = pd.NA
    melted["Student Name"] = melted["Name"]
    melted["Social Category"] = pd.NA
    melted["Palm Width Cleaned"] = pd.NA
    melted["Palm Length Cleaned"] = pd.NA
    melted["Data Source"] = "Bedridden"
    melted["Record Type"] = "Bedridden"

    cleaned = melted[
        [
            "District",
            "School_Name",
            "School ID",
            "Student Name",
            "Gender",
            "Social Category",
            "disability_cleaned",
            "Palm Width Cleaned",
            "Palm Length Cleaned",
            "Priority",
            "Device",
            "Device Category",
            "Name",
            "Age",
            "Contact No",
            "Address",
            "Other requirement",
            "Data Source",
            "Record Type",
        ]
    ].copy()

    export_df = cleaned[
        ["Name", "Age", "Gender", "Contact No", "Address", "District", "disability_cleaned", "Device", "Other requirement"]
    ].rename(columns={"disability_cleaned": "Disability"})

    try:
        export_df.to_csv("data/cleaned_bedridden_data.csv", index=False)
    except PermissionError:
        pass

    return cleaned
# ---------------- RUN FILE ----------------
if __name__ == "__main__":
    df_devices=load_and_clean_data()
    print("✅Cleaned file saved as data/cleaned_data.csv")
    print(df_devices.head())
    print(df_devices['School ID'].unique())
