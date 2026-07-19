import pandas as pd
import numpy as np
import io
import os
import shutil
import datetime
import openpyxl
import re
import urllib.parse
import streamlit as st
import altair as alt
import scripts.comparable_assembly_logic as comparable_assembly_logic


DEFAULT_SAMPLE_SQUARES = 500

st.set_page_config(page_title="Historical Estimating Benchmark", layout="wide")

# Inject CSS for printer optimization
st.markdown("""
<style>
@media print {
    /* Hide interactive/non-content widgets */
    section[data-testid="stSidebar"],
    div.stButton,
    header,
    footer,
    [data-testid="stHeader"] {
        display: none !important;
    }

    /* Expand main container to use full page width and clear margins */
    section.main {
        width: 100% !important;
        padding: 0 !important;
        margin: 0 !important;
    }
    
    .block-container {
        padding: 0.2rem !important;
        max-width: 100% !important;
    }

    /* Prevent split-boundary breaks inside metric tiles and key panels */
    [data-testid="metric-container"],
    [data-testid="stMetricValue"],
    .element-container,
    div.stDataFrame,
    div[data-testid="stTable"],
    .stAlert {
        break-inside: avoid !important;
        page-break-inside: avoid !important;
    }

    /* Heading break rules */
    h1, h2, h3, h4, h5 {
        break-after: avoid !important;
        page-break-after: avoid !important;
        margin-top: 0.5rem !important;
        margin-bottom: 0.2rem !important;
    }
    
    h1 { font-size: 16pt !important; }
    h2 { font-size: 14pt !important; }
    h3 { font-size: 12pt !important; }
    h4 { font-size: 11pt !important; }

    /* Custom page break utility class */
    .print-page-break {
        break-before: page !important;
        page-break-before: always !important;
        height: 0 !important;
        margin: 0 !important;
        padding: 0 !important;
    }
    
    /* Force background colors to print */
    * {
        -webkit-print-color-adjust: exact !important;
        print-color-adjust: exact !important;
    }
    
    /* Premium typography */
    body {
        font-family: 'Inter', 'Segoe UI', sans-serif !important;
        color: #111827 !important;
        font-size: 10pt !important;
        line-height: 1.2 !important;
    }
    
    /* Force columns to stay horizontal in print instead of collapsing vertically */
    [data-testid="stHorizontalBlock"] {
        display: flex !important;
        flex-direction: row !important;
        flex-wrap: nowrap !important;
        width: 100% !important;
        gap: 0.5rem !important;
    }
    
    [data-testid="column"] {
        flex: 1 1 0px !important;
        min-width: 0 !important;
    }
    
    /* Ensure metric cards inside columns display correctly without overflow */
    [data-testid="metric-container"] {
        width: 100% !important;
        min-width: 0 !important;
        padding: 0.2rem !important;
    }
    
    [data-testid="stMetricValue"] {
        font-size: 1.2rem !important;
    }
    [data-testid="stMetricLabel"] {
        font-size: 0.8rem !important;
    }
    
    /* Compress tables */
    div.stDataFrame table, div[data-testid="stTable"] table {
        font-size: 9pt !important;
        line-height: 1.1 !important;
    }
    td, th {
        padding: 2px 4px !important;
    }
    
    /* Compress spacing */
    [data-testid="stVerticalBlock"] { gap: 0.2rem !important; }
    hr { margin: 0.2rem 0 !important; }
    
    /* Hide everything from audit tools onwards */
    div.element-container:has(#audit-tools-start),
    div.element-container:has(#audit-tools-start) ~ * {
        display: none !important;
    }
    
    /* Hide screen-only metrics and charts in print */
    [data-testid="metric-container"],
    div[data-testid="stVegaLiteChart"],
    div[data-testid="stArrowVegaLiteChart"] {
        display: none !important;
    }
    
    /* Hide screen-only dataframes in print */
    div.stDataFrame, [data-testid="stDataFrame"] {
        display: none !important;
    }
}

/* Screen vs Print utility classes */
@media screen {
    .print-only, .print-only-flex {
        display: none !important;
    }
}
@media print {
    .print-only {
        display: block !important;
    }
    .print-only-flex {
        display: flex !important;
    }
    .screen-only {
        display: none !important;
    }
}
</style>
""", unsafe_allow_html=True)

# File paths
from pathlib import Path
PROJECT_ROOT = Path(__file__).resolve().parent
SOURCE_DIR = PROJECT_ROOT / "data" / "source"
DASHBOARD_DATA_DIR = PROJECT_ROOT / "data" / "dashboard"
NEW_SOURCE_DIR = PROJECT_ROOT / "new source"

OFFICIAL_FILE = str(PROJECT_ROOT / "data" / "Demo_Master_List.xlsx")
DATA_DIR = str(DASHBOARD_DATA_DIR)
LOCAL_FILE = OFFICIAL_FILE

# Ensure directories exist
SOURCE_DIR.mkdir(parents=True, exist_ok=True)
DASHBOARD_DATA_DIR.mkdir(parents=True, exist_ok=True)

def make_google_maps_link(address, city):
    addr_str = str(address or '').strip()
    city_str = str(city or '').strip()
    if not addr_str or addr_str.lower() in ['nan', 'none', '']:
        return None
    query = addr_str
    if city_str and city_str.lower() not in ['nan', 'none', '']:
        query += f", {city_str}"
    return f"https://www.google.com/maps?q={urllib.parse.quote_plus(query)}&t=h"

st.markdown("<h1 class='screen-only' style='margin-bottom: 0px;'>Historical Estimating Benchmark Dashboard</h1>", unsafe_allow_html=True)
sticky_header_placeholder = st.empty()



# Column name mapping: new spreadsheet column -> old internal column name
# This allows all downstream code to continue using the old internal names
COLUMN_RENAME_MAP = {
    "Insulation R-Value": "Insulation Thickness/R-Value",
    "Scaffolding": "Scaffold",
    "01-0100 - General - Permits": "Permit",
    "01-0110 - General - Disposal": "Disposal",
    "01-0120 - General - Equipment Rental": "Equipment",
    "01-0160 - General - Fuel": "Fuel",
    "02-0110 - Labor - Roofing": "ST Cost",
    "02-0111 - Labor - Overtime": "OT Cost",
    "Type G - Benchmark": "Type G Total",
    "Type M - Benchmark": "Type M Total",
    "Total Cost": "Total Report Cost",
    # Material columns: new expanded names -> old short names
    "03-0100 - Materials - Fasteners": "03-0100 Fasteners",
    "03-0101 - Materials - Base": "03-0101 Base",
    "03-0102 - Materials - Ply": "03-0102 Ply",
    "03-0103 - Materials - Capsheet": "03-0103 Capsheet",
    "03-0104 - Materials - Asphalt": "03-0104 Asphalt",
    "03-0105 - Materials - Gravel": "03-0105 Gravel",
    "03-0106 - Materials - Shingles": "03-0106 Shingles",
    "03-0107 - Materials - Shingle Felt": "03-0107 Shingle Felt",
    "03-0108 - Materials - Lumber": "03-0108 Lumber",
    "03-0109 - Materials - Cold Adhesive": "03-0109 Cold Adhesive",
    "03-0110 - Material - Sheetmetal": "03-0110 Sheetmetal",
    "03-0111 - Material - Modified SBS": "03-0111 Modified SBS",
    "03-0112 - Materials-Coating": "03-0112 Coating",
    "03-0113 - Materials - Insulation": "03-0113 Insulation",
    "03-0115 - Materials - Tile": "03-0115 Tile",
    "03-0116 - Materials - Emulsion": "03-0116 Emulsion",
    "03-0117 - Materials - Polyester rolls": "03-0117 Polyester Rolls",
    "03-0118 - Material - Waterproofing": "03-0118 Waterproofing",
    "03-0120 - Material - Misc (other)": "03-0120 Misc (other)",
    "03-0120 - Special Cost Items": "03-0120 Special (Skylights/Hatches)",
    "03-0125 - Home Depot Purchases": "03-0125 Home Depot Purchases",
    "03-0130 - Material - Misc. Supplies": "03-0130 Misc. Supplies",
    "03-0130 - Pipe Support Blocks": "03-0130 Pipe Support Blocks",
    "03-0140 - Materials - Tax": "03-0140 Tax",
    "03-0150 - DENS DECK 4x8": "03-0150 DENS DECK 4x8",
    "03-0160 - DENS DECK 4x4": "03-0160 DENS DECK 4x4",
    "03-0201 - Materials - Single Ply Memb.": "03-0201 Single Ply Memb.",
    "03-0202 - Materials - Fleeceback Memb.": "03-0202 Fleeceback Memb.",
    "03-0203 - Materials - Single Ply Acces.": "03-0203 Single Ply Acces.",
}

# Reverse map: internal name -> original full source string (for display in cost breakdowns)
DISPLAY_NAME_MAP = {v: k for k, v in COLUMN_RENAME_MAP.items()}

# Cost breakdown columns for Tier 1 row detail pane
# Type G: post-rename internal names for all included General Condition codes
COST_BREAKDOWN_TYPE_G_COLS = [
    "Permit",       # 01-0100 - General - Permits
    "Disposal",     # 01-0110 - General - Disposal
    "Equipment",    # 01-0120 - General - Equipment Rental
    "Fuel",         # 01-0160 - General - Fuel
    # These 01- columns were not renamed — display name = original name
    "01-0145 - General - Other",
    "01-0165 - General - Bridges, Tolls, Prkg",
    "01-0170 - General - Propane",
    "01-0180 - General-Damage Repair",
    "01-0190 - General-Manufacturer Warranty",
    "01-0200 - General-Outside Services",
    "01-0210 - General - Asbestos Supplies",
    "01-0220 - General - Safety Supplies",
]

# Type M: post-rename internal names for all included Material codes
COST_BREAKDOWN_TYPE_M_COLS = [
    "03-0100 Fasteners", "03-0101 Base", "03-0102 Ply", "03-0103 Capsheet",
    "03-0104 Asphalt", "03-0105 Gravel",
    "03-0106 Shingles", "03-0107 Shingle Felt", "03-0108 Lumber",
    "03-0109 Cold Adhesive", "03-0110 Sheetmetal", "03-0111 Modified SBS",
    "03-0112 Coating", "03-0113 Insulation", "03-0115 Tile", "03-0116 Emulsion",
    "03-0117 Polyester Rolls",
    "03-0118 Waterproofing", "03-0120 Misc (other)", "03-0125 Home Depot Purchases",
    "03-0130 Misc. Supplies", "03-0140 Tax", "03-0150 DENS DECK 4x8",
    "03-0160 DENS DECK 4x4", "03-0201 Single Ply Memb.", "03-0202 Fleeceback Memb.",
    "03-0203 Single Ply Acces.",
]

# Explicitly excluded columns — NEVER appear in cost breakdown
COST_BREAKDOWN_EXCLUSIONS = {
    "Scaffold",
    "01-0140 Per Diem",
    "03-0120 Special (Skylights/Hatches)",
    "03-0130 Pipe Support Blocks",
}

# ---------------------------------------------------------------------------
# Spec Code acronym maps — composite identifier for the Spec Type column
# Format: [Spec] [R-Val] [CoverBoard] [Attachment] [Thickness] [Material]
# ---------------------------------------------------------------------------
_SPEC_TYPE_MAP = {
    "overlay": "OL",
    "tear-off": "TO",
    "tear off": "TO",
    "coating": "CT",
}
_INS_MAP = {
    "r-10": "R10", "r10": "R10",
    "r19": "R19",
    "r23": "R23",
    "r5.6": "R5.6",
}
_CB_MAP = {
    "densdeck": "DD",
    "hd polyiso": "HD",
    "cgf polyiso": "CGF",
    "fan fold": "FF",
    "gypsum": "GYP",
}
_ATT_MAP = {
    "fully adhered": "FA",
    "mechanically attached": "MA",
    "rhinobond": "RB",
}
_THICK_MAP = {
    "60 mil": "60", "60mil": "60",
    "80 mil": "80", "80mil": "80",
    "115 mil": "115", "115mil": "115",
    "135 mil": "135", "135mil": "135",
    "per diem": None,  # ignore
}
_MAT_MAP = {
    "tpo": "TPO",
    "pvc": "PVC",
    "fb tpo": "FB TPO",
    "fb pvc": "FB PVC",
}

_FIELD_TO_MAP = {
    "Spec Type": _SPEC_TYPE_MAP,
    "Insulation Thickness/R-Value": _INS_MAP,
    "Cover Board Type": _CB_MAP,
    "Roof Material Attachment": _ATT_MAP,
    "Cover Board Attachment": _ATT_MAP,
    "Roof Material Thickness": _THICK_MAP,
    "Roof Material Type": _MAT_MAP,
}

def normalize_field_value(field, raw_val):
    if not raw_val or str(raw_val).strip().lower() in ("", "nan", "none", "all"):
        return ""
    mapping = _FIELD_TO_MAP.get(field)
    if mapping:
        key = str(raw_val).strip().lower()
        if key in mapping:
            # We must convert to string because some mappings might map to None explicitly
            mapped_val = mapping[key]
            return str(mapped_val).lower() if mapped_val else ""
    return str(raw_val).strip().lower()

def _lookup_acronym(mapping, raw_val):
    """Return acronym for raw_val, original value as fallback, or None to skip."""
    if not raw_val or str(raw_val).strip().lower() in ("", "nan", "none", "all"):
        return None
    key = str(raw_val).strip().lower()
    if key in mapping:
        return mapping[key]  # may be None (explicitly ignored)
    return str(raw_val).strip()  # fallback: show as-is

def build_spec_code(job):
    """Build compact composite spec code from a job dict's filter fields."""
    parts = []
    for mapping, field in [
        (_SPEC_TYPE_MAP,  "Spec Type"),
        (_INS_MAP,        "Insulation Thickness/R-Value"),
        (_CB_MAP,         "Cover Board Type"),
        (_ATT_MAP,        "Roof Material Attachment"),
        (_THICK_MAP,      "Roof Material Thickness"),
        (_MAT_MAP,        "Roof Material Type"),
    ]:
        val = _lookup_acronym(mapping, job.get(field))
        if val:
            parts.append(val)
    return " ".join(parts) if parts else "—"

@st.cache_data(ttl=60) # Cache data for 60 seconds
def get_numeric_series(df, col_name, default=0.0):
    return pd.to_numeric(df.get(col_name, pd.Series(default, index=df.index)), errors='coerce').fillna(default)

def get_text_series(df, col_name, default=""):
    return df.get(col_name, pd.Series(default, index=df.index)).astype(str).fillna(default)

@st.cache_data(ttl=60)
def load_data():
    if not os.path.exists(LOCAL_FILE):
        return None, None, None, None
        
    try:
        df_all = pd.read_excel(LOCAL_FILE, sheet_name="FULL LIST")
    except Exception as e:
        st.error(f"Error loading local Excel file: {e}")
        return None, None, None, None

    # Rename columns from new format to old internal names
    df_all = df_all.rename(columns=COLUMN_RENAME_MAP)
    
    def clean_df(df):
        # Drop legacy Median/Sample rows if any exist at the bottom
        first_col = df.columns[0]
        df[first_col] = df[first_col].astype(str).str.strip()
        df = df[~df[first_col].str.lower().isin(["median", "sample"])]
        
        # List of text columns to preserve (not convert to numeric)
        text_cols = ["Job#", "Address", "City", "Spec Type", "Coating Spec", 
                     "Insulation Thickness/R-Value", "Insulation Attachment", 
                     "Cover Board Type", "Cover Board Thickness", "Cover Board Attachment",
                     "Roof Material Type", "Roof Material Thickness", "Roof Material Attachment",
                     "T/O Weight psf"]
        
        # Convert all non-text columns to numeric
        for col in df.columns:
            if col not in text_cols:
                if df[col].dtype == object:
                    try:
                        df[col] = pd.to_numeric(df[col])
                    except (ValueError, TypeError):
                        pass
                    
        # Replace NaN with 0 for calculation columns
        calc_cols = ["Contract", "Total Squares", "ST Cost", "OT Cost", 
                     "Permit", "Disposal", "Equipment", "Scaffold", "Fuel", 
                     "Type G Total", "Total Report Cost", 
                     "03-0120 Special (Skylights/Hatches)"]
        
        for col in calc_cols:
            if col in df.columns:
                df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0.0)
                
        # Fill material columns with 0
        mat_cols = [c for c in df.columns if str(c).startswith("03-")]
        for c in mat_cols:
            df[c] = pd.to_numeric(df[c], errors='coerce').fillna(0.0)
            
        # Clean Roof Material Thickness
        def clean_thickness(val):
            val_str = str(val).strip().lower()
            if "115" in val_str: return "115 mil"
            if "135" in val_str: return "135 mil"
            if "80" in val_str: return "80 mil"
            if "60" in val_str or "61" in val_str: return "60 mil"
            return ""
            
        if "Roof Material Thickness" in df.columns:
            df["Roof Material Thickness"] = df["Roof Material Thickness"].apply(clean_thickness)
            
        # Reconstruct computed fields
        if "Contract" in df.columns and "Total Squares" in df.columns:
            df["$/square"] = np.where(df["Total Squares"] > 0, df["Contract"] / df["Total Squares"], 0.0)
        
        # Total Labor = ST Cost (02-0110) + OT Cost (02-0111)
        if "ST Cost" in df.columns and "OT Cost" in df.columns:
            df["ST Cost"] = pd.to_numeric(df["ST Cost"], errors='coerce').fillna(0.0)
            df["OT Cost"] = pd.to_numeric(df["OT Cost"], errors='coerce').fillna(0.0)
            df["Total Labor"] = df["ST Cost"] + df["OT Cost"]
        else:
            df["Total Labor"] = 0.0
            
        # Estimator MD from Man Days
        official_md = pd.to_numeric(df.get("Man Days", pd.Series(np.nan, index=df.index)), errors='coerce')
        st_labor = pd.to_numeric(df.get("ST Labor", pd.Series(0.0, index=df.index)), errors='coerce').fillna(0.0)
        ot_labor = pd.to_numeric(df.get("OT Labor", pd.Series(0.0, index=df.index)), errors='coerce').fillna(0.0)
        calculated_md = (st_labor + ot_labor) / 16.0
        
        # Use cached/hardcoded value if available, otherwise compute it dynamically
        df["Estimator MD"] = official_md.fillna(calculated_md)
        sq = pd.to_numeric(df.get("Total Squares", pd.Series(0.0, index=df.index)), errors='coerce').fillna(0.0)
        
        # The spreadsheet formula for SQ/MD is corrupted (e.g. =O2/R2, dividing Contract by Report date).
        # We must completely ignore the spreadsheet's SQ/MD column and calculate it mathematically here.
        df["SQ/MD"] = np.where(df["Estimator MD"] > 0, sq / df["Estimator MD"], 0.0)
        
        df["$/MD"] = np.where(df["Estimator MD"] > 0, df["Total Labor"] / df["Estimator MD"], 0.0)
        if "Complete" in df.columns:
            df["Complete"] = pd.to_datetime(df["Complete"], errors="coerce")
        if "Report" in df.columns:
            df["Report"] = pd.to_datetime(df["Report"], errors="coerce")
        
        # TYPE G TOTAL
        # BM formula: =SUM(U:AG)-X
        # Decoded: sum all 01-xxxx columns (U to AG)
        #          MINUS Per Diem (X = 01-0140)
        # Note: Scaffold (AH) is outside the range.
        all_01_cols = [c for c in df.columns if str(c).startswith("01-")]
        per_diem_col = "01-0140 - General-Subsistence (Per Diem)"
        type_g_include = [c for c in all_01_cols if c != per_diem_col]
        # Include mapped aliases for cols that kept their original names
        for alias in ["Permit", "Disposal", "Fuel", "Equipment"]:
            if alias in df.columns and alias not in type_g_include:
                type_g_include.append(alias)
        computed_g = df[type_g_include].apply(pd.to_numeric, errors='coerce').fillna(0.0).sum(axis=1)

        if "Type G Total" not in df.columns:
            df["Type G Total"] = computed_g
        else:
            df["Type G Total"] = pd.to_numeric(df["Type G Total"], errors='coerce').fillna(0.0)
            # Use the dedicated benchmark column if filled in; otherwise use computed value
            df["Type G Total"] = np.where(df["Type G Total"] <= 0, computed_g, df["Type G Total"])

        # TYPE M TOTAL
        # BN formula: =SUM(AK:BJ)-BA-BD
        # Decoded: sum all 03-xxxx material codes (AK-BJ)
        #          MINUS 03-0120 Special Cost Items (BA)
        #          MINUS 03-0130 Pipe Support Blocks (BD)
        all_03_cols = [c for c in df.columns if str(c).startswith("03-")]
        # Exclude the two specific codes that BN subtracts out (using their renamed names and raw names just in case)
        type_m_exclude = {"03-0120 Special (Skylights/Hatches)", "03-0130 Pipe Support Blocks",
                          "03-0120 - Special Cost Items", "03-0130 - Pipe Support Blocks"}
        type_m_mat_cols = [c for c in all_03_cols if c not in type_m_exclude]
        computed_m = df[type_m_mat_cols].apply(pd.to_numeric, errors='coerce').fillna(0.0).sum(axis=1) if type_m_mat_cols else 0.0

        if "Type M Total" not in df.columns:
            df["Type M Total"] = computed_m
        else:
            df["Type M Total"] = pd.to_numeric(df["Type M Total"], errors='coerce').fillna(0.0)
            # Use the dedicated benchmark column if filled in; otherwise use computed value
            df["Type M Total"] = np.where(df["Type M Total"] <= 0, computed_m, df["Type M Total"])
                
        # Compute Total Report Cost if empty or 0
        computed_total = (
            df["Type G Total"].fillna(0.0) +
            df["Total Labor"].fillna(0.0) +
            df["Type M Total"].fillna(0.0)
        )
        sub_col = "04-All - Subcontractor" if "04-All - Subcontractor" in df.columns else None
        if sub_col:
            computed_total += pd.to_numeric(df[sub_col], errors='coerce').fillna(0.0)
            
        if "Total Report Cost" not in df.columns:
            df["Total Report Cost"] = computed_total
        else:
            df["Total Report Cost"] = pd.to_numeric(df["Total Report Cost"], errors='coerce').fillna(0.0)
            df["Total Report Cost"] = np.where(df["Total Report Cost"] <= 0, computed_total, df["Total Report Cost"])
            
        df["Material (W/O Skylights)"] = get_numeric_series(df, "Type M Total", 0.0) - get_numeric_series(df, "03-0120 Special (Skylights/Hatches)", 0.0)
        
        def get_clean_col(col_name):
            if col_name in df.columns:
                return pd.to_numeric(df[col_name], errors='coerce').fillna(0.0)
            return 0.0

        # Total Cost (Independent) and $/SQ Ind are computed AFTER the sidebar
        # labor rate is known (see recompute block below the sidebar setup).
        # Initialise as 0.0 placeholders here.
        df["Total Cost (Independent)"] = 0.0
        df["$/SQ Ind"] = 0.0
        
        return df

    df_all = clean_df(df_all)
    
    # Route jobs by Spec Type
    # TPO-PVC: Tear-Off and Overlay rows
    tpopvc_mask = df_all["Spec Type"].astype(str).str.strip().str.lower().isin(["tear-off", "tear off", "overlay"])
    df_tpopvc = df_all[tpopvc_mask].copy()
    
    # Coating: Coating rows
    coating_mask = df_all["Spec Type"].astype(str).str.strip().str.lower() == "coating"
    df_coating = df_all[coating_mask].copy()
    
    # Master List (macro-only / permit data): MASTER LIST rows
    master_list_mask = df_all["Spec Type"].astype(str).str.strip().str.upper() == "MASTER LIST"
    df_master_list_only = df_all[master_list_mask].copy()
    
    # df_master = all rows combined for permit/macro analysis (TPO + Coating + MASTER LIST)
    df_master = df_all.copy()
    
    return df_master, df_tpopvc, df_coating, df_master_list_only

df_master, df_tpopvc, df_coating, df_master_list_only = load_data()


def apply_eligibility_rule(df):
    if df is None or df.empty:
        return df, pd.DataFrame(columns=df.columns if df is not None else [])
    
    # Needs valid Complete and Report dates
    has_dates = df['Complete'].notna() & df['Report'].notna()
    # Report >= Complete + 30 days
    is_30_days = (df['Report'] - df['Complete']).dt.days >= 30
    
    eligible_mask = has_dates & is_30_days
    
    # Identify reasons for audit
    ineligible_df = df[~eligible_mask].copy()
    if not ineligible_df.empty:
        reasons = []
        for _, row in ineligible_df.iterrows():
            if pd.isna(row.get('Complete')):
                reasons.append("Missing Complete date")
            elif pd.isna(row.get('Report')):
                reasons.append("Missing Report date")
            else:
                gap = (row['Report'] - row['Complete']).days if pd.notna(row.get('Complete')) and pd.notna(row.get('Report')) else 0
                reasons.append(f"Gap < 30 days ({gap} days)")
        ineligible_df["Ineligible Reason"] = reasons
        
    return df[eligible_mask].copy(), ineligible_df

df_master_eligible, df_master_ineligible = None, None
df_tpopvc_eligible, df_tpopvc_ineligible = None, None
df_coating_eligible, df_coating_ineligible = None, None

if df_master is not None:
    df_master_eligible, df_master_ineligible = apply_eligibility_rule(df_master)
    df_tpopvc_eligible, df_tpopvc_ineligible = apply_eligibility_rule(df_tpopvc)
    df_coating_eligible, df_coating_ineligible = apply_eligibility_rule(df_coating)

if df_master is None:
    st.warning("⚠️ Local dashboard data not found. Please click 'Update Dashboard Data' above to initialize.")
    st.stop()

# Build base combo for allowances
df_all = pd.concat([df_tpopvc_eligible, df_coating_eligible], ignore_index=True)

# Permit Data (Uses Master List)
permit_mask = (pd.to_numeric(df_master_eligible["Permit"], errors='coerce').fillna(0) > 0) & (pd.to_numeric(df_master_eligible["Total Squares"], errors='coerce').fillna(0) > 0) & (~df_master_eligible["Job#"].astype(str).str.lower().isin(["median", "sample"]))
df_permit = df_master_eligible[permit_mask].copy()
if not df_permit.empty:
    df_permit["Permit Cost/SQ"] = df_permit["Permit"] / df_permit["Total Squares"]
    df_permit["Contract"] = pd.to_numeric(df_permit["Contract"], errors='coerce').fillna(0)
    df_permit["Permit % of Contract"] = np.where(df_permit["Contract"] > 0, (df_permit["Permit"] / df_permit["Contract"]) * 100.0, np.nan)

# Disposal Data (TPO-PVC Tear-Off only)
disp_mask = (df_tpopvc_eligible["Spec Type"].astype(str).str.strip().str.lower().isin(["tear-off", "tear off"])) & (pd.to_numeric(df_tpopvc_eligible["Disposal"], errors='coerce').fillna(0) > 0) & (pd.to_numeric(df_tpopvc_eligible["Total Squares"], errors='coerce').fillna(0) > 0) & (~df_tpopvc_eligible["Job#"].astype(str).str.lower().isin(["median", "sample"]))
df_disposal = df_tpopvc_eligible[disp_mask].copy()
if not df_disposal.empty:
    df_disposal["Disposal Cost/SQ"] = df_disposal["Disposal"] / df_disposal["Total Squares"]
    # Calculate tonnage and $/Ton
    df_disposal["T/O Weight psf"] = get_numeric_series(df_disposal, "T/O Weight psf", 0.0)
    df_disposal["Total Tons"] = (df_disposal["T/O Weight psf"] * df_disposal["Total Squares"] * 100) / 2000
    df_disposal["Disposal Cost/Ton"] = df_disposal.apply(lambda x: x["Disposal"] / x["Total Tons"] if x["Total Tons"] > 0 else np.nan, axis=1)

# Fuel Data (Uses Master List)
fuel_mask = (pd.to_numeric(df_master_eligible["Fuel"], errors='coerce').fillna(0) > 0) & (pd.to_numeric(df_master_eligible["Estimator MD"], errors='coerce').fillna(0) > 0) & (~df_master_eligible["Job#"].astype(str).str.lower().isin(["median", "sample"]))
df_fuel = df_master_eligible[fuel_mask].copy()
if not df_fuel.empty:
    df_fuel["Fuel Cost/MD"] = df_fuel["Fuel"] / df_fuel["Estimator MD"]
    df_fuel["City"] = df_fuel["City"].astype(str).str.strip().str.title()
    
    zone_1 = ["Hayward", "Newark", "Fremont", "San Leandro", "Union City", "San Lorenzo", "Castro Valley"]
    zone_2 = ["Redwood City", "San Mateo", "Palo Alto", "Menlo Park", "Foster City", "Brisbane", "South San Francisco", "Burlingame", "Belmont", "San Carlos"]
    zone_3 = ["San Jose", "Santa Clara", "Sunnyvale", "Milpitas", "Mountain View", "Cupertino", "Los Gatos", "Campbell"]
    zone_4 = ["Concord", "Walnut Creek", "Vallejo", "Pleasant Hill", "Martinez", "Pittsburg", "Antioch", "Dublin", "Pleasanton", "Livermore", "San Ramon"]
    zone_5 = ["Oakland", "Berkeley", "Alameda", "Emeryville", "Richmond", "San Pablo"]
    zone_6 = ["San Francisco"]

    def get_zone(city):
        if city in zone_1: return "Zone 1 (Local East Bay)"
        if city in zone_2: return "Zone 2 (Peninsula / Bridges)"
        if city in zone_3: return "Zone 3 (South Bay)"
        if city in zone_4: return "Zone 4 (Far East / North)"
        if city in zone_5: return "Zone 5 (Oakland / Berkeley)"
        if city in zone_6: return "Zone 6 (San Francisco)"
        return "Unzoned"

    df_fuel["Zone"] = df_fuel["City"].apply(get_zone)

# Helpers
def get_filter_options(df, col):
    if col not in df.columns:
        return []
        
    has_blanks = df[col].isna().any() or (df[col].astype(str).str.strip() == "").any()
    
    vals = df[col].dropna().astype(str).str.strip()
    vals = vals[vals != ""]
    
    unique_opts = {}
    for v in vals:
        norm_v = normalize_field_value(col, v)
        if not norm_v:
            continue
            
        if len(norm_v) <= 3 or norm_v in ["fb tpo", "fb pvc"]:
            display_v = norm_v.upper()
        else:
            display_v = norm_v.title()
            
        unique_opts[norm_v] = display_v
        
    options = sorted(list(unique_opts.values()))
    if has_blanks:
        options.append("[Blank]")
        
    return options

def filter_material_categories_for_selected_system(material_cols, selected_roof_material_type):
    if not selected_roof_material_type or "All" in selected_roof_material_type or "[Blank]" in selected_roof_material_type:
        return material_cols, False
        
    single_ply = ["TPO", "PVC", "EPDM"]
    is_single_ply = False
    for mat in selected_roof_material_type:
        if any(s.lower() in str(mat).lower() for s in single_ply):
            is_single_ply = True
            break
    
    if not is_single_ply:
        return material_cols, False
        
    shingle_keywords = [
        "shingle", "shingles", "comp shingle", "composition shingle", 
        "shingle underlayment", "felt", "starter", "ridge cap", 
        "hip and ridge", "ice and water", "ice & water", "synthetic underlayment"
    ]
    
    filtered_cols = []
    suppressed = False
    for col in material_cols:
        col_lower = col.lower()
        if any(kw in col_lower for kw in shingle_keywords):
            suppressed = True
        else:
            filtered_cols.append(col)
            
    return filtered_cols, suppressed

def job_matches_filter(job, field, target_list):
    if not target_list or "All" in target_list or "N/A" in target_list: 
        return True
        
    raw_val = job.get(field, "")
    val = normalize_field_value(field, raw_val)
    
    for t in target_list:
        t_clean = str(t).strip().lower()
        if t_clean == "[blank]":
            if val == "":
                return True
        else:
            t_norm = normalize_field_value(field, t)
            if val == t_norm:
                return True
            
    return False


def filter_jobs(df, filters, min_jobs, is_master=True, sample_squares=0.0, excluded_job_nos=None):
    if df is None or df.empty:
        return [], "No Data", {"t1":0,"t2":0,"t3":0,"t4":0}, {"t1":[],"t2":[],"t3":[],"t4":[]}, 4
        
    if excluded_job_nos is None:
        excluded_job_nos = []
        
    jobs = df.to_dict('records')
    
    # Strip out excluded jobs immediately
    jobs = [j for j in jobs if j.get("Job#") not in excluded_job_nos]
        
    # Size Filtering
    min_sq = sample_squares * 0.40
    if sample_squares < 500:
        max_sq = sample_squares * 2.50
    else:
        max_sq = 10000.0
    jobs = [j for j in jobs if min_sq <= float(j.get("Total Squares") or 0.0) <= max_sq]
    
    tier_counts = {"t1": 0, "t2": 0, "t3": 0, "t4": len(jobs)}
    
    if is_master:
        # Tier 1
        tier1 = [j for j in jobs if all(job_matches_filter(j, k, v) for k, v in filters.items())]
        tier_counts["t1"] = len(tier1)
        
        # Tier 2: Drop Cover Board Type. Keep Spec, Insulation, Material, Thickness, and Attachment.
        t2_filters = {k:v for k,v in filters.items() if k in ["Spec Type", "Insulation Thickness/R-Value", "Roof Material Type", "Roof Material Thickness", "Roof Material Attachment"]}
        tier2 = [j for j in jobs if all(job_matches_filter(j, k, v) for k, v in t2_filters.items())]
        tier_counts["t2"] = len(tier2)
        
        # Tier 3: Drop Cover Board Type AND Roof Material Type and Thickness. Keep Spec, Insulation, and Attachment.
        t3_filters = {k:v for k,v in filters.items() if k in ["Spec Type", "Insulation Thickness/R-Value", "Roof Material Attachment"]}
        tier3 = [j for j in jobs if all(job_matches_filter(j, k, v) for k, v in t3_filters.items())]
        tier_counts["t3"] = len(tier3)
        
        tier_jobs_dict = {"t1": tier1, "t2": tier2, "t3": tier3, "t4": jobs}
        
        if len(tier1) >= min_jobs: return tier1, "Tier 1 (Strict)", tier_counts, tier_jobs_dict, 1
        if len(tier2) >= min_jobs: return tier2, "Tier 2 (Loose)", tier_counts, tier_jobs_dict, 2
        if len(tier3) >= min_jobs: return tier3, "Tier 3 (Broad)", tier_counts, tier_jobs_dict, 3
        
        return jobs, "Tier 4 (Global Fallback)", tier_counts, tier_jobs_dict, 4
    else:
        # Coating Projects
        # Tier 1: Spec Type, Coating Spec, City
        t1_filters = {k:v for k,v in filters.items() if k in ["Spec Type", "Coating Spec", "City"]}
        tier1 = [j for j in jobs if all(job_matches_filter(j, k, v) for k, v in t1_filters.items())]
        tier_counts["t1"] = len(tier1)
        
        # Tier 2: Drop City
        t2_filters = {k:v for k,v in filters.items() if k in ["Spec Type", "Coating Spec"]}
        tier2 = [j for j in jobs if all(job_matches_filter(j, k, v) for k, v in t2_filters.items())]
        tier_counts["t2"] = len(tier2)
        
        # Tier 3: Spec Type only
        t3_filters = {k:v for k,v in filters.items() if k in ["Spec Type"]}
        tier3 = [j for j in jobs if all(job_matches_filter(j, k, v) for k, v in t3_filters.items())]
        tier_counts["t3"] = len(tier3)
        
        tier_jobs_dict = {"t1": tier1, "t2": tier2, "t3": tier3, "t4": jobs}
        
        if len(tier1) >= min_jobs: return tier1, "Tier 1 (Strict)", tier_counts, tier_jobs_dict, 1
        if len(tier2) >= min_jobs: return tier2, "Tier 2 (Loose)", tier_counts, tier_jobs_dict, 2
        if len(tier3) >= min_jobs: return tier3, "Tier 3 (Broad)", tier_counts, tier_jobs_dict, 3
        
        return jobs, "Tier 4 (Global Fallback)", tier_counts, tier_jobs_dict, 4

# Sidebar UI


st.sidebar.title("Configuration")

source_sheet = "TPO-PVC"
sample_squares = st.sidebar.number_input("Target Project Size (Squares)", value=float(DEFAULT_SAMPLE_SQUARES), step=10.0, format="%.2f")
labor_rate = st.sidebar.number_input("Man Day Labor Rate", value=430.00, step=25.0, format="%.2f")
min_jobs = st.sidebar.number_input("Min Comparable Jobs", value=5, step=1)

st.sidebar.markdown("### Filters")
is_master = (source_sheet == "TPO-PVC")
df_active = df_tpopvc_eligible if is_master else df_coating_eligible

def render_selectbox(label, col_name):
    opts = ["All"] + get_filter_options(df_active, col_name)
    return st.sidebar.selectbox(label, opts, index=0)

filters = {}

if is_master:
    spec_type = render_selectbox("Spec Type", "Spec Type")
    ins_r = render_selectbox("Insulation R-Value", "Insulation Thickness/R-Value")
    cb_type = render_selectbox("Cover Board Type", "Cover Board Type")
    mat_type = render_selectbox("Roof Material Type", "Roof Material Type")
    mat_thick = render_selectbox("Roof Material Thickness", "Roof Material Thickness")
    mat_attach = render_selectbox("Roof Material Attachment", "Roof Material Attachment")
    
    filters = {
        "Spec Type": [spec_type] if spec_type else ["All"],
        "Insulation Thickness/R-Value": [ins_r] if ins_r else ["All"],
        "Cover Board Type": [cb_type] if cb_type else ["All"],
        "Roof Material Type": [mat_type] if mat_type else ["All"],
        "Roof Material Thickness": [mat_thick] if mat_thick else ["All"],
        "Roof Material Attachment": [mat_attach] if mat_attach else ["All"],
    }
else:
    coating_spec = render_selectbox("Coating Spec", "Coating Spec")
    
    filters = {
        "Coating Spec": [coating_spec] if coating_spec else ["All"]
    }

# Fallbacks for variables removed from filters UI but still referenced downstream
if not is_master:
    spec_type = "All"
    mat_type = coating_spec
    cb_type = "All"
    mat_attach = "All"

cb_attach = "All"

# Recompute Total Cost (Independent) and $/SQ Ind now that labor_rate is known.
# Formula: Type G Total + Type M Total + (Estimator MD × labor_rate)
for _df in [df_tpopvc, df_coating, df_tpopvc_eligible, df_coating_eligible]:
    if _df is not None and not _df.empty:
        _g  = pd.to_numeric(_df["Type G Total"], errors='coerce').fillna(0.0)
        _m  = pd.to_numeric(_df["Type M Total"], errors='coerce').fillna(0.0)
        _md = pd.to_numeric(_df["Estimator MD"], errors='coerce').fillna(0.0)
        _sq = pd.to_numeric(_df["Total Squares"], errors='coerce').fillna(0.0)
        _tci = _g + _m + (_md * labor_rate)
        _df["Total Cost (Independent)"] = _tci
        _df["$/SQ Ind"] = np.where(_sq > 0, _tci / _sq, 0.0)

# Pass 1: Ghost Run (No exclusions) to find what naturally lands in the active tier
_, _, _, base_tier_jobs_dict, base_active_tier_level = filter_jobs(df_active, filters, min_jobs, is_master, sample_squares, [])

all_jobs_context = {}
if df_active is not None and not df_active.empty:
    # Only pull jobs that would actually show up in the base comparable tables
    valid_jnos = set()
    for t_level in range(1, base_active_tier_level + 1):
        for j in base_tier_jobs_dict.get(f"t{t_level}", []):
            if "Job#" in j:
                valid_jnos.add(j["Job#"])
                
    for _, row in df_active.iterrows():
        jno = row.get("Job#")
        if pd.isna(jno) or jno not in valid_jnos: continue
        sq = pd.to_numeric(row.get("Total Squares"), errors='coerce')
        c_sq = pd.to_numeric(row.get("$/SQ Ind"), errors='coerce')
        sq_str = f"{sq:,.0f} SQ" if not pd.isna(sq) else "? SQ"
        c_sq_str = f"${c_sq:,.2f}/SQ" if not pd.isna(c_sq) else "?/SQ"
        label = f"{jno} | {sq_str} | {c_sq_str}"
        all_jobs_context[label] = jno

# Retrieve exclusions from session state (UI is rendered below)
if "exclude_dropdown" in st.session_state:
    st.session_state["exclude_dropdown"] = [lbl for lbl in st.session_state["exclude_dropdown"] if lbl in all_jobs_context]
selected_labels = st.session_state.get("exclude_dropdown", [])
excluded_job_nos = [all_jobs_context.get(lbl) for lbl in selected_labels]

# Pass 2: Real Run (With exclusions)
matched_jobs, tier_used, tier_counts, tier_jobs_dict, active_tier_level = filter_jobs(df_active, filters, min_jobs, is_master, sample_squares, excluded_job_nos)

# 2. Benchmark Eligibility Filter
eligible_jobs = []
excluded_jobs = []
mat_cols = [c for c in df_active.columns if str(c).startswith("03-")]

if is_master and not df_active.empty:
    for _, row in df_active.iterrows():
        job_no = row.get("Job#", "Unknown")
        reasons = []
        
        comp_date = row.get("Complete")
        rep_date = row.get("Report")
        
        if pd.isnull(comp_date): reasons.append("Missing Complete Date")
        if pd.isnull(rep_date): reasons.append("Missing Report Date")
        if not pd.isnull(comp_date) and not pd.isnull(rep_date):
            if (rep_date - comp_date).days < 30:
                reasons.append("Report Date less than 30 days after Complete Date")
                
        sq = float(row.get("Total Squares") or 0.0)
        labor = float(row.get("Total Labor") or 0.0)
        type_g_orig = float(row.get("Type G Total") or 0.0)
        scaffold = float(row.get("Scaffold") or 0.0)
        type_g = max(0.0, type_g_orig - scaffold)
        
        mats_orig = sum(float(row.get(c) or 0.0) for c in mat_cols)
        special = float(row.get("03-0120 Special (Skylights/Hatches)") or 0.0)
        mats = max(0.0, mats_orig - special)
        
        hist_dc = type_g + labor + mats
        
        if sq <= 0: reasons.append("Total Squares <= 0")
        if labor <= 0: reasons.append("Total Labor <= 0")
        if hist_dc <= 0: reasons.append("Clean Benchmark Direct Cost <= 0")
        
        if reasons:
            excluded_jobs.append({
                "Job #": job_no,
                "Address": row.get("Address", ""),
                "City": row.get("City", ""),
                "Complete Date": comp_date.strftime("%Y-%m-%d") if not pd.isnull(comp_date) else "Invalid/Missing",
                "Report Date": rep_date.strftime("%Y-%m-%d") if not pd.isnull(rep_date) else "Invalid/Missing",
                "Exclusion Reason": ", ".join(reasons)
            })
        else:
            eligible_jobs.append(row.to_dict())
            
    df_active = pd.DataFrame(eligible_jobs) if eligible_jobs else pd.DataFrame(columns=df_active.columns)
    
    if not df_active.empty:
        df_active["Scaffold Excluded"] = get_numeric_series(df_active, "Scaffold", 0.0)
        df_active["Special Items Excluded"] = get_numeric_series(df_active, "03-0120 Special (Skylights/Hatches)", 0.0)
        df_active["Original Material Total"] = df_active[[c for c in mat_cols]].sum(axis=1)
        df_active["Type G Total"] = get_numeric_series(df_active, "Type G Total", 0.0)
        df_active["Clean Benchmark Type G Total"] = (df_active["Type G Total"] - df_active["Scaffold Excluded"]).clip(lower=0)
        df_active["Clean Benchmark Material Total"] = (df_active["Original Material Total"] - df_active["Special Items Excluded"]).clip(lower=0)
        df_active["Total Labor"] = get_numeric_series(df_active, "Total Labor", 0.0)
        df_active["Clean Benchmark Direct Cost"] = df_active["Clean Benchmark Type G Total"] + df_active["Total Labor"] + df_active["Clean Benchmark Material Total"]
        df_active["Original Historical Direct Cost"] = df_active["Type G Total"] + df_active["Total Labor"] + df_active["Original Material Total"]
        df_active["Total Excluded Cost"] = df_active["Original Historical Direct Cost"] - df_active["Clean Benchmark Direct Cost"]
        sq_arr = get_numeric_series(df_active, "Total Squares", 0.0)
        df_active["Clean Benchmark Direct Cost/SQ"] = np.where(sq_arr > 0, df_active["Clean Benchmark Direct Cost"] / sq_arr, 0.0)

# 3. Dynamic Sidebar Filters
with st.sidebar:
    st.markdown("### Tier Coverage")
    max_sq_display = 10000.0 if sample_squares >= 500 else sample_squares * 2.50
    st.sidebar.markdown(f"**Eligible Size Range:** {sample_squares * 0.40:,.2f} - {max_sq_display:,.2f} SQ")
    st.sidebar.markdown(f"* Tier 1 Strict: {tier_counts['t1']} jobs\n* Tier 2 Major Specs: {tier_counts['t2']} jobs\n* Tier 3 Broad System: {tier_counts['t3']} jobs\n* Tier 4 Global Fallback: {tier_counts['t4']} jobs")

def get_rate_median(col_name):
    rates = []
    for j in matched_jobs:
        sq = float(j.get("Total Squares") or 0.0)
        # Type G Total and Type M Total are now the pure benchmark figures,
        # so they no longer need Scaffold or Skylights subtracted out.
        if col_name == "03-0120 Special (Skylights/Hatches)":
            continue
        else:
            val = float(j.get(col_name) or 0.0)
        if val > 0.0 and sq > 0.0:
            rates.append(val / sq)
    return np.median(rates) if rates else 0.0

# Productivity (Volume-Weighted)
valid_labor_jobs = []
for j in matched_jobs:
    sq = float(j.get("Total Squares") or 0.0)
    md = float(j.get("Estimator MD") or 0.0)
    if sq > 0 and md > 0:
        valid_labor_jobs.append(j)

sum_sq = sum(float(j.get("Total Squares") or 0.0) for j in valid_labor_jobs)
sum_md = sum(float(j.get("Estimator MD") or 0.0) for j in valid_labor_jobs)
weighted_sq_md = (sum_sq / sum_md) if sum_md > 0 else 0.0
est_md = (sample_squares / weighted_sq_md) if weighted_sq_md > 0 else 0.0
est_labor_cost = est_md * labor_rate

# Pricing (Median)
cost_per_sq_vals = []
type_g_med_sq = get_rate_median("Type G Total")
mats_med_sq = get_rate_median("Type M Total")

for j in matched_jobs:
    sq = float(j.get("Total Squares") or 0.0)
    type_g = float(j.get("Type G Total") or 0.0)
    labor = float(j.get("Total Labor") or 0.0)
    mats = float(j.get("Type M Total") or 0.0)
    
    hist_direct_cost = type_g + labor + mats
    
    if sq > 0 and hist_direct_cost > 0:
        cps = hist_direct_cost / sq
        if cps > 0:
            cost_per_sq_vals.append(cps)

cps_low = np.percentile(cost_per_sq_vals, 25) if cost_per_sq_vals else 0.0
cps_med = np.median(cost_per_sq_vals) if cost_per_sq_vals else 0.0
cps_high = np.percentile(cost_per_sq_vals, 75) if cost_per_sq_vals else 0.0

expected_type_g = type_g_med_sq * sample_squares
expected_mats = mats_med_sq * sample_squares

# Build Recap Table
recap_data = []

g_col_display_names = {
    "Permit": "01-0100 - General - Permits",
    "Disposal": "01-0110 - General - Disposal",
    "Equipment": "01-0120 - General - Equipment Rental",
    "Fuel": "01-0160 - General - Fuel"
}

g_cols_to_list = [c for c in df_active.columns if str(c).startswith("01-")]
g_cols_to_list.extend(["Permit", "Disposal", "Equipment", "Fuel"])

# 1. Type G Individual Items
# Exclusions: Scaffold, Per Diem, and anything in COST_BREAKDOWN_EXCLUSIONS
G_RECAP_EXCLUSIONS = COST_BREAKDOWN_EXCLUSIONS | {
    "01-0140 - General-Subsistence (Per Diem)",  # unrenamed Per Diem column
    "Scaffold",                                   # renamed Scaffolding column
}
for col in g_cols_to_list:
    if col not in df_active.columns: continue
    if col in G_RECAP_EXCLUSIONS: continue
    rate = get_rate_median(col)
    if rate > 0:
        orig_name = g_col_display_names.get(col, col)
        if " - " in orig_name:
            code = orig_name.split(" - ")[0].strip()
            desc = orig_name.split(" - ", 1)[1].strip()
        elif " " in orig_name:
            code = orig_name.split(" ")[0].strip()
            desc = orig_name.split(" ", 1)[1].strip()
        else:
            code = "01-Misc"
            desc = orig_name
            
        recap_data.append({
            "Cost Code": code, 
            "Description": desc, 
            "Category": "Other / General", 
            "Type": "Type G Item", 
            "Median Rate": rate, 
            "Expected Total": rate * sample_squares
        })
    
# 2. Labor
if est_labor_cost > 0:
    recap_data.append({"Cost Code": "02-0110", "Description": "Labor - Roofing", "Category": "Labor", "Type": "Labor", "Median Rate": est_labor_cost / (sample_squares or 1), "Expected Total": est_labor_cost})

# 3. Materials — skip all columns in COST_BREAKDOWN_EXCLUSIONS
display_mat_cols, mats_suppressed = filter_material_categories_for_selected_system(mat_cols, mat_type)

for col in display_mat_cols:
    if col in COST_BREAKDOWN_EXCLUSIONS: continue
    rate = get_rate_median(col)
    if rate > 0:
        code = col.split(" ")[0]
        desc = col.split(" ", 1)[1] if " " in col else ""
        recap_data.append({"Cost Code": code, "Description": desc, "Category": "Material", "Type": "Material", "Median Rate": rate, "Expected Total": rate * sample_squares})

df_recap = pd.DataFrame(recap_data)

# Precalculate quick reference totals
est_total_cost = expected_type_g + est_labor_cost + expected_mats
sell_30 = est_total_cost / 0.70
sell_40 = est_total_cost / 0.60
sq_val = sample_squares if sample_squares > 0 else 1

view_b_direct_cost = expected_type_g + est_labor_cost + expected_mats

# Build Print-Only Selections Summary Brief
additional_fields_html = ""

def clean_val_for_brief(val):
    if isinstance(val, list):
        cleaned_list = [str(v).strip() for v in val if str(v).strip().lower() != "all"]
        if not cleaned_list:
            return "All"
        return ", ".join(cleaned_list)
    val_str = str(val).strip()
    if val_str.lower() in ["all", "['all']", "[]", "none", "nan", ""]:
        return "All"
    return val_str

source_sheet_val = clean_val_for_brief(source_sheet)
sample_squares_val = f"{sample_squares:,.2f}"
labor_rate_val = f"${labor_rate:,.2f}"
min_jobs_val = clean_val_for_brief(min_jobs)
city_val = "All Region"

if is_master:
    spec_type_val = clean_val_for_brief(spec_type)
    mat_type_str = clean_val_for_brief(mat_type)
    mat_attach_str = clean_val_for_brief(mat_attach)
    cb_type_str = clean_val_for_brief(cb_type)
    ins_r_str = clean_val_for_brief(ins_r)
    
    narrative_text = (
        f"This benchmark evaluates <strong>{spec_type_val}</strong> projects in the "
        f"<strong>{city_val}</strong>, scaled to a target project size of <strong>{sample_squares_val} SQ</strong>. "
        f"Based on a search of the <strong>{source_sheet_val}</strong> database requiring a minimum comparable sample "
        f"of <strong>{min_jobs_val}</strong> jobs, the analysis filters for systems matching "
        f"<strong>{mat_type_str}</strong> roof material (Attachment: <strong>{mat_attach_str}</strong>) and "
        f"<strong>{cb_type_str}</strong> cover board (Insulation R-Value: <strong>{ins_r_str}</strong>). "
        f"All labor requirements are calculated at a base rate of <strong>{labor_rate_val}/MD</strong>."
    )
else:
    coating_spec_str = clean_val_for_brief(coating_spec)
    
    narrative_text = (
        f"This benchmark evaluates projects in the "
        f"<strong>{city_val}</strong>, scaled to a target project size of <strong>{sample_squares_val} SQ</strong>. "
        f"Based on a search of the <strong>{source_sheet_val}</strong> database requiring a minimum comparable sample "
        f"of <strong>{min_jobs_val}</strong> jobs, the analysis filters for systems matching coating specification "
        f"<strong>{coating_spec_str}</strong>. All labor requirements are calculated at a base rate of "
        f"<strong>{labor_rate_val}/MD</strong>."
    )

brief_html = f"""<div class="print-only" style="background-color: #f8fafc; border: 1px solid #cbd5e1; border-radius: 8px; padding: 12px 16px; margin-bottom: 15px; font-family: 'Inter', -apple-system, sans-serif; line-height: 1.4; font-size: 13px; color: #334155;">
<div style="font-size: 14px; font-weight: 700; color: #0f172a; margin-bottom: 6px; border-bottom: 2px solid #3b82f6; padding-bottom: 4px; text-transform: uppercase; letter-spacing: 0.5px;">Historical Estimating Benchmark — Project Brief</div>
<p style="margin: 0; text-align: justify; font-size: 13px;">{narrative_text}</p>
</div>"""
st.markdown(brief_html, unsafe_allow_html=True)

st.markdown("---")

st.header("Current Benchmark Setup")

# Missing Value Warnings
for k, v_list in filters.items():
    for v in v_list:
        if v != "All" and v != "N/A" and v != "[Blank]":
            if k in df_active.columns:
                valid_vals_lower = [normalize_field_value(k, str(x)) for x in df_active[k].dropna() if str(x).strip() != ""]
                v_norm = normalize_field_value(k, str(v))
                if v_norm not in valid_vals_lower:
                    st.warning(f"⚠️ Selected value '{v}' does not currently exist in the historical data for {k}. Results may fall back to a broader tier.")

if tier_used == "Tier 4 (Global Fallback)":
    st.error("INSUFFICIENT HISTORICAL DATA: No close comparable set was found. Results are based on all valid jobs and should not be used for estimating.")

col1, col2, col3 = st.columns(3)
col1.metric("Comparable Jobs Found", len(matched_jobs))
col2.metric("Match Level", tier_used)
col3.metric("Target Roof Size", f"{sample_squares:,.2f}")

st.markdown("---")
st.header("Top Estimate Summary")

if tier_used != "Tier 4 (Global Fallback)":
    t1, t2, t3, t4, t5 = st.columns(5)
    t1.metric("Clean Total Cost", f"${est_total_cost:,.2f}")
    t1.caption("*(Baseline: Excludes Scaffolding, Deck Replacement, and Special Items)*")
    t2.metric("Sell @ 30% Margin", f"${sell_30:,.2f}")
    t3.metric("Sell @ 40% Margin", f"${sell_40:,.2f}")
    t4.metric("Estimated MD", f"{est_md:,.2f}")
    t5.metric("Weighted SQ/MD", f"{weighted_sq_md:,.2f}")
    
    # Inject Sticky Banner at the top of the page
    sticky_html = f"""
    <div id="sticky-banner-container" style="display: none;"></div>
    <style>
    /* Make the parent Streamlit element-container sticky */
    div.element-container:has(#sticky-banner-container) {{
        position: sticky;
        top: 2.875rem;
        z-index: 999;
    }}
    .sticky-banner {{
        background-color: rgba(255, 255, 255, 0.95);
        backdrop-filter: blur(5px);
        padding: 12px 20px;
        border-radius: 8px;
        box-shadow: 0 4px 6px rgba(0,0,0,0.1);
        border: 1px solid #e5e7eb;
        display: flex;
        justify-content: space-between;
        align-items: center;
        margin-bottom: 1rem;
    }}
    .sticky-stat {{
        text-align: center;
        font-family: 'Inter', -apple-system, sans-serif;
    }}
    .sticky-label {{
        font-size: 0.75rem;
        color: #64748b;
        text-transform: uppercase;
        font-weight: 700;
        letter-spacing: 0.5px;
    }}
    .sticky-value {{
        font-size: 1.15rem;
        color: #0f172a;
        font-weight: 800;
    }}
    .sticky-highlight {{
        color: #2563eb;
    }}
    </style>
    <div class="sticky-banner screen-only">
        <div class="sticky-stat"><div class="sticky-label">Clean Total Cost</div><div class="sticky-value">${est_total_cost:,.2f}</div></div>
        <div class="sticky-stat"><div class="sticky-label">Sell @ 30% Margin</div><div class="sticky-value">${sell_30:,.2f}</div></div>
        <div class="sticky-stat"><div class="sticky-label">Sell @ 40% Margin</div><div class="sticky-value sticky-highlight">${sell_40:,.2f}</div></div>
        <div class="sticky-stat"><div class="sticky-label">Estimated MD</div><div class="sticky-value">{est_md:,.2f}</div></div>
        <div class="sticky-stat"><div class="sticky-label">Weighted SQ/MD</div><div class="sticky-value">{weighted_sq_md:,.2f}</div></div>
    </div>
    """
    sticky_header_placeholder.markdown(sticky_html, unsafe_allow_html=True)
    

    
    # Inject printable SaaS Pricing Cards right after chart (hidden on screen)
    cost_per_sq = est_total_cost / (sample_squares or 1)
    sell_30_per_sq = sell_30 / (sample_squares or 1)
    sell_40_per_sq = sell_40 / (sample_squares or 1)
    profit_30 = sell_30 - est_total_cost
    profit_40 = sell_40 - est_total_cost
    
    cards_html = f"""<div class="print-only-flex" style="display: flex; flex-direction: row; justify-content: space-between; gap: 16px; margin-top: 10px; margin-bottom: 20px; font-family: 'Inter', -apple-system, sans-serif; width: 100%;">
<div style="flex: 1; background-color: #f8fafc; border: 1px solid #cbd5e1; border-top: 4px solid #64748b; border-radius: 8px; padding: 14px; text-align: center; box-shadow: 0 1px 3px rgba(0,0,0,0.05);">
<div style="font-size: 11px; font-weight: 700; color: #64748b; text-transform: uppercase; letter-spacing: 0.5px;">Base Cost</div>
<div style="font-size: 24px; font-weight: 800; color: #1e293b; margin: 6px 0 4px 0;">${est_total_cost:,.2f}</div>
<div style="font-size: 12px; font-weight: 600; color: #475569; margin-bottom: 10px;">${cost_per_sq:,.2f} <span style="font-weight: 400; color: #64748b;">/ SQ</span></div>
<div style="border-top: 1px dashed #cbd5e1; padding-top: 8px; text-align: left; font-size: 11px; color: #475569; display: flex; flex-direction: column; gap: 4px;">
<div><strong>Labor:</strong> {est_md:,.2f} Man Days</div>
<div><strong>Productivity:</strong> {weighted_sq_md:,.2f} SQ / MD</div>
</div>
</div>
<div style="flex: 1; background-color: #f0f9ff; border: 1px solid #bae6fd; border-top: 4px solid #0284c7; border-radius: 8px; padding: 14px; text-align: center; box-shadow: 0 1px 3px rgba(0,0,0,0.05);">
<div style="font-size: 11px; font-weight: 700; color: #0369a1; text-transform: uppercase; letter-spacing: 0.5px;">30% Margin Target</div>
<div style="font-size: 24px; font-weight: 800; color: #0284c7; margin: 6px 0 4px 0;">${sell_30:,.2f}</div>
<div style="font-size: 12px; font-weight: 600; color: #0284c7; margin-bottom: 10px;">${sell_30_per_sq:,.2f} <span style="font-weight: 400; color: #0369a1;">/ SQ</span></div>
<div style="border-top: 1px dashed #bae6fd; padding-top: 8px; text-align: left; font-size: 11px; color: #0369a1; display: flex; flex-direction: column; gap: 4px;">
<div><strong>Gross Margin:</strong> 30.0%</div>
<div><strong>Gross Profit:</strong> ${profit_30:,.2f}</div>
</div>
</div>
<div style="flex: 1; background-color: #f0fdf4; border: 1px solid #bbf7d0; border-top: 4px solid #16a34a; border-radius: 8px; padding: 14px; text-align: center; box-shadow: 0 1px 3px rgba(0,0,0,0.05);">
<div style="font-size: 11px; font-weight: 700; color: #15803d; text-transform: uppercase; letter-spacing: 0.5px;">40% Margin Target</div>
<div style="font-size: 24px; font-weight: 800; color: #16a34a; margin: 6px 0 4px 0;">${sell_40:,.2f}</div>
<div style="font-size: 12px; font-weight: 600; color: #16a34a; margin-bottom: 10px;">${sell_40_per_sq:,.2f} <span style="font-weight: 400; color: #15803d;">/ SQ</span></div>
<div style="border-top: 1px dashed #bbf7d0; padding-top: 8px; text-align: left; font-size: 11px; color: #15803d; display: flex; flex-direction: column; gap: 4px;">
<div><strong>Gross Margin:</strong> 40.0%</div>
<div><strong>Gross Profit:</strong> ${profit_40:,.2f}</div>
</div>
</div>
</div>"""
    st.markdown(cards_html, unsafe_allow_html=True)
else:
    st.error("Disabled due to Tier 4 Global Fallback.")



# Allowance Estimators
st.markdown("---")
st.header("Allowance Estimators")
st.markdown("*These are independent budgeting tools and do not affect Clean Benchmark Direct Cost.*")

col_perm, col_disp, col_fuel = st.columns(3)

# PERMIT ESTIMATOR
permit_export_data = {}
with col_perm:
    st.markdown("#### Permit Cost Estimator")
    if not df_permit.empty:
        permit_cities = ["All Cities"] + sorted([c for c in df_permit["City"].dropna().unique() if str(c).strip() != ""])
        permit_counts = df_permit["City"].dropna().astype(str).str.strip().value_counts().to_dict()
        total_permit_count = len(df_permit)
        
        def format_permit_city(city_opt):
            if city_opt == "All Cities":
                return f"All Cities ({total_permit_count})"
            count = permit_counts.get(city_opt, 0)
            return f"{city_opt} ({count})"
            
        p_city = st.selectbox("Permit City", permit_cities, key="permit_city", format_func=format_permit_city)
        
        pdf = df_permit if p_city == "All Cities" else df_permit[df_permit["City"].astype(str).str.strip() == p_city]
        p_count = len(pdf)
        
        if p_count >= 3:
            p_low = np.percentile(pdf["Permit Cost/SQ"], 25)
            p_med = np.median(pdf["Permit Cost/SQ"])
            p_high = np.percentile(pdf["Permit Cost/SQ"], 75)
            
            sq = sample_squares if sample_squares > 0 else 1
            
            # 1. Square-Based Estimate
            tab1, tab2 = st.tabs(["Square-Based", "Contract-Based"])
            
            with tab1:
                st.metric("Estimated Permit Cost", f"${p_med * sq:,.2f}")
                st.markdown(f"**Permit \\$/SQ:** \${p_med:,.2f}/SQ")
                st.markdown(f"**Low / High Range:** \${p_low * sq:,.2f} - \${p_high * sq:,.2f}")
                st.markdown(f"**Data Points Used:** {p_count}")
            
            # 2. Contract-Based Estimate
            pdf_pct = pdf["Permit % of Contract"].dropna()
            p_pct_count = len(pdf_pct)
            
            with tab2:
                if p_pct_count >= 3 and sell_30 > 0:
                    p_pct_low = np.percentile(pdf_pct, 25)
                    p_pct_med = np.median(pdf_pct)
                    p_pct_high = np.percentile(pdf_pct, 75)
                    
                    p_pct_est = (p_pct_med / 100.0) * sell_30
                    p_pct_low_val = (p_pct_low / 100.0) * sell_30
                    p_pct_high_val = (p_pct_high / 100.0) * sell_30
                    
                    st.metric("Estimated Permit Cost", f"${p_pct_est:,.2f}")
                    st.markdown(f"**Permit % of Contract:** {p_pct_med:.3f}%")
                    st.markdown(f"**Low / High Range:** \${p_pct_low_val:,.2f} - \${p_pct_high_val:,.2f}")
                    st.markdown(f"**Data Points Used:** {p_pct_count}")
                    st.markdown(f"*Based on Sell @ 30% Margin: \${sell_30:,.2f}*")
                    
                    permit_export_data = {
                        "Active": True,
                        "City": p_city,
                        "Count": p_count,
                        "Median $/SQ": p_med,
                        "Low": p_low * sq,
                        "Median": p_med * sq,
                        "High": p_high * sq,
                        "Contract_Pct_Med": p_pct_med,
                        "Contract_Low": p_pct_low_val,
                        "Contract_Median": p_pct_est,
                        "Contract_High": p_pct_high_val,
                        "Msg": ""
                    }
                else:
                    st.warning("Insufficient historical contract-value data to estimate by percentage.")
                    permit_export_data = {
                        "Active": True,
                        "City": p_city,
                        "Count": p_count,
                        "Median $/SQ": p_med,
                        "Low": p_low * sq,
                        "Median": p_med * sq,
                        "High": p_high * sq,
                        "Msg": "Insufficient contract-value data"
                    }
            
        else:
            st.error("Insufficient permit data for this city.")
            st.markdown(f"**Data Points Found:** {p_count} (Minimum Required: 3)")
            permit_export_data = {"Active": False, "City": p_city, "Count": p_count, "Msg": "Insufficient data"}
            if p_city != "All Cities" and len(df_permit) >= 3:
                p_med_all = np.median(df_permit["Permit Cost/SQ"])
                st.markdown(f"*All Cities Reference: \${p_med_all * (sample_squares if sample_squares > 0 else 1):,.2f}*")
                
        st.info("Permit estimate is based on historical permit cost per square and is intended for budgeting only. Actual permit costs may vary by jurisdiction, project valuation, review fees, and project-specific requirements.")
    else:
        st.warning("No valid permit data found in historical dataset.")
        permit_export_data = {"Active": False, "City": "N/A", "Count": 0, "Msg": "No valid data"}

# DISPOSAL ESTIMATOR
disposal_export_data = {}
with col_disp:
    st.markdown("#### Disposal Cost Estimator")
    if spec_type not in ["Tear-Off", "Tear Off", "TO"] and not (is_master and spec_type == "All"):
        st.info("Disposal estimator is only available when Spec Type is Tear-Off.")
        disposal_export_data = {"Active": False, "Msg": "Inactive (Spec Type is not Tear-Off)"}
    else:
        if not df_disposal.empty:
            disp_cities = ["All Cities"] + sorted([c for c in df_disposal["City"].dropna().unique() if str(c).strip() != ""])
            disposal_counts = df_disposal["City"].dropna().astype(str).str.strip().value_counts().to_dict()
            total_disposal_count = len(df_disposal)
            
            def format_disposal_city(city_opt):
                if city_opt == "All Cities":
                    return f"All Cities ({total_disposal_count})"
                count = disposal_counts.get(city_opt, 0)
                return f"{city_opt} ({count})"
                
            d_city = st.selectbox("Disposal City", disp_cities, key="disp_city", format_func=format_disposal_city)
            
            user_weight_psf = st.number_input("Estimated Roof Weight (psf)", value=2.0, step=0.5, key="disp_weight")
            sq = sample_squares if sample_squares > 0 else 1
            est_total_tons = (user_weight_psf * sq * 100) / 2000
            
            ddf = df_disposal if d_city == "All Cities" else df_disposal[df_disposal["City"].astype(str).str.strip() == d_city]
            d_count = len(ddf)
            
            if d_count >= 3:
                d_low = np.percentile(ddf["Disposal Cost/SQ"], 25)
                d_med = np.median(ddf["Disposal Cost/SQ"])
                d_high = np.percentile(ddf["Disposal Cost/SQ"], 75)
                
                ddf_tons = ddf[ddf["Total Tons"] > 0]
                t_count = len(ddf_tons)
                has_ton_data = t_count >= 3
                
                tab1, tab2 = st.tabs(["Weight-Based", "Square-Based"])
                
                with tab1:
                    if user_weight_psf > 0 and has_ton_data:
                        t_low = np.percentile(ddf_tons["Disposal Cost/Ton"], 25)
                        t_med = np.median(ddf_tons["Disposal Cost/Ton"])
                        t_high = np.percentile(ddf_tons["Disposal Cost/Ton"], 75)
                        
                        st.metric("Estimated Disposal Cost", f"${t_med * est_total_tons:,.2f}")
                        st.markdown(f"**Disposal \\$/Ton:** \${t_med:,.2f}/Ton")
                        st.markdown(f"**Estimated Tonnage:** {est_total_tons:,.1f} Tons")
                        st.markdown(f"**Data Points Used:** {t_count}")
                    elif user_weight_psf > 0 and not has_ton_data:
                        st.warning("Insufficient historical weight data to estimate by tonnage.")
                    else:
                        st.info("Enter estimated roof weight to calculate tonnage.")
                
                with tab2:
                    st.metric("Estimated Disposal Cost", f"${d_med * sq:,.2f}")
                    st.markdown(f"**Disposal \\$/SQ:** \${d_med:,.2f}/SQ")
                    st.markdown(f"**Low / High Range:** \${d_low * sq:,.2f} - \${d_high * sq:,.2f}")
                    st.markdown(f"**Data Points Used:** {d_count}")
                
                disposal_export_data = {
                    "Active": True,
                    "City": d_city,
                    "Count": d_count,
                    "Median $/SQ": d_med,
                    "Low": d_low * sq,
                    "Median": d_med * sq,
                    "High": d_high * sq,
                    "Msg": ""
                }
                if user_weight_psf > 0 and has_ton_data:
                    disposal_export_data["Estimated Tons"] = est_total_tons
                    disposal_export_data["Median $/Ton"] = t_med
                    disposal_export_data["Weight-Based Total"] = t_med * est_total_tons
            else:
                st.error("Insufficient disposal data for this city.")
                st.markdown(f"**Data Points Found:** {d_count} (Minimum Required: 3)")
                disposal_export_data = {"Active": False, "City": d_city, "Count": d_count, "Msg": "Insufficient data"}
                
                if d_city != "All Cities" and len(df_disposal) >= 3:
                    d_med_all = np.median(df_disposal["Disposal Cost/SQ"])
                    st.markdown(f"*All Cities Reference: \${d_med_all * (sample_squares if sample_squares > 0 else 1):,.2f}*")
                    
            st.info("Disposal estimate is based only on historical Tear-Off projects with recorded disposal costs and is intended for budgeting only. Actual disposal costs may vary based on debris type, dump fees, access, trucking, and hazardous material conditions.")
        else:
            st.warning("No valid disposal data found in historical dataset.")
            disposal_export_data = {"Active": False, "Msg": "No valid data"}

# FUEL ESTIMATOR
fuel_export_data = {}
with col_fuel:
    st.markdown("#### Fuel Cost Estimator")
    if not df_fuel.empty:
        fuel_cities = ["All Cities"] + sorted([c for c in df_fuel["City"].dropna().unique() if str(c).strip() != ""])
        fuel_counts = df_fuel["City"].dropna().astype(str).str.strip().value_counts().to_dict()
        total_fuel_count = len(df_fuel)
        
        def format_fuel_city(city_opt):
            if city_opt == "All Cities":
                return f"All Cities ({total_fuel_count})"
            count = fuel_counts.get(city_opt, 0)
            return f"{city_opt} ({count})"
            
        f_city = st.selectbox("Fuel City", fuel_cities, key="fuel_city", format_func=format_fuel_city)
        
        fdf = df_fuel if f_city == "All Cities" else df_fuel[df_fuel["City"].astype(str).str.strip() == f_city]
        f_count = len(fdf)
        
        if f_count >= 3:
            f_low = np.percentile(fdf["Fuel Cost/MD"], 25)
            f_med = np.median(fdf["Fuel Cost/MD"])
            f_high = np.percentile(fdf["Fuel Cost/MD"], 75)
            
            st.metric("Estimated Fuel Cost", f"${f_med * est_md:,.2f}")
            st.markdown(f"**Fuel \\$/MD:** \${f_med:,.2f}/MD")
            st.markdown(f"**Low / High Range:** \${f_low * est_md:,.2f} - \${f_high * est_md:,.2f}")
            st.markdown(f"**Estimated Man Days:** {est_md:,.2f} MD")
            st.markdown(f"**Data Points Used:** {f_count}")
            
            fuel_export_data = {
                "Active": True,
                "City": f_city,
                "Count": f_count,
                "Median $/MD": f_med,
                "Low": f_low * est_md,
                "Median": f_med * est_md,
                "High": f_high * est_md,
                "Msg": ""
            }
            
            if f_city != "All Cities":
                city_zone = get_zone(f_city)
                if city_zone != "Unzoned":
                    zone_df = df_fuel[df_fuel["Zone"] == city_zone]
                    if len(zone_df) >= 3:
                        z_med = np.median(zone_df["Fuel Cost/MD"])
                        st.markdown(f"*Secondary Metric: {f_city} is in **{city_zone}**. The historical median for this zone is **\${z_med:,.2f}/MD**.*")
                        fuel_export_data["Zone"] = city_zone
                        fuel_export_data["Zone_Med"] = z_med
        else:
            st.error("Insufficient fuel data for this city.")
            st.markdown(f"**Data Points Found:** {f_count} (Minimum Required: 3)")
            fuel_export_data = {"Active": False, "City": f_city, "Count": f_count, "Msg": "Insufficient data"}
            
            if f_city != "All Cities":
                city_zone = get_zone(f_city)
                if city_zone != "Unzoned":
                    zone_df = df_fuel[df_fuel["Zone"] == city_zone]
                    if len(zone_df) >= 3:
                        z_med = np.median(zone_df["Fuel Cost/MD"])
                        st.markdown(f"*Zone Reference: {f_city} is in **{city_zone}**. The historical median for this zone is **\${z_med:,.2f}/MD** (\${z_med * est_md:,.2f}).*")
                        fuel_export_data["Zone"] = city_zone
                        fuel_export_data["Zone_Med"] = z_med
                elif len(df_fuel) >= 3:
                    f_med_all = np.median(df_fuel["Fuel Cost/MD"])
                    st.markdown(f"*All Cities Reference: \${f_med_all * est_md:,.2f}*")
                
        st.info("Fuel estimate is based on historical fuel cost per Man Day. Fuel costs fluctuate significantly with macroeconomic gas prices and distance to the job site.")
    else:
        st.warning("No valid fuel data found in historical dataset.")
        fuel_export_data = {"Active": False, "Msg": "No valid data"}

# Quick Reference Estimate
if tier_used != "Tier 4 (Global Fallback)":
    st.markdown('<div class="print-page-break"></div>', unsafe_allow_html=True)
    st.markdown("---")
    st.header("Quick Reference Estimate")
    
    q1, q2, q3 = st.columns(3)
    q1.metric("Estimated Type G Cost", f"${expected_type_g:,.2f}", f"${type_g_med_sq:,.2f}/SQ")
    q2.metric("Estimated Type L / Labor Cost", f"${est_labor_cost:,.2f}", f"${est_labor_cost/sq_val:,.2f}/SQ")
    q3.metric("Estimated Type M / Material Cost", f"${expected_mats:,.2f}", f"${mats_med_sq:,.2f}/SQ")
    

st.header("Main Benchmark Results")

layout_style = st.radio(
    "Benchmark Layout Style", 
    ["New KPI Tabs (Compact)", "Classic View (Detailed)"], 
    horizontal=True,
    help="Toggle between the new tabbed layout and the classic side-by-side layout."
)

if layout_style == "Classic View (Detailed)":
    viewA, viewB = st.columns(2)
    
    with viewA:
        st.subheader("View A: Historical Cost Benchmark")
        st.markdown("*Historical roof-system benchmark excluding scaffold and specialty items.*")
        st.metric("Historical Low Benchmark Total", f"${cps_low * sample_squares:,.2f}")
        st.metric("Historical Median Benchmark Total", f"${cps_med * sample_squares:,.2f}")
        st.metric("Historical High Benchmark Total", f"${cps_high * sample_squares:,.2f}")
        
        st.markdown("#### Direct Cost/SQ")
        st.markdown(f"**Low:** \${cps_low:,.2f} | **Median:** \${cps_med:,.2f} | **High:** \${cps_high:,.2f}")
    
    with viewB:
        st.subheader("View B: Current-Rate Projection")
        st.markdown("*Historical roof-system productivity and cleaned cost patterns priced with today's labor rate.*")
        if tier_used == "Tier 4 (Global Fallback)":
            st.error("Broad reference only — not reliable for estimating.")
        else:
            st.metric("Current-Rate Projected Direct Cost Total", f"${view_b_direct_cost:,.2f}")
            st.metric("Current-Rate Projected Direct Cost/SQ", f"${(view_b_direct_cost/sample_squares) if sample_squares>0 else 0:,.2f}")
            st.markdown(f"**Expected Type G:** \${expected_type_g:,.2f} | **Expected Labor:** \${est_labor_cost:,.2f} | **Expected Materials:** \${expected_mats:,.2f}")
    
    # Labor Rate Context
    st.markdown("---")
    st.header("Labor Rate Context")
    hist_labor_rates = []
    for j in valid_labor_jobs:
        l = float(j.get("Total Labor", 0))
        md = float(j.get("Estimator MD", 0))
        if md > 0:
            hist_labor_rates.append(l / md)
    median_hist_labor_rate = np.median(hist_labor_rates) if hist_labor_rates else 0.0
    
    st.markdown(f"**Median Historical Labor:** \${median_hist_labor_rate:,.2f}/MD | **Your Selected Rate:** \${labor_rate:,.2f}/MD")
    labor_diff = labor_rate - median_hist_labor_rate
    labor_diff_pct = (labor_diff / median_hist_labor_rate) * 100 if median_hist_labor_rate > 0 else 0
    st.markdown(f"**Difference:** \${labor_diff:,.2f} ({labor_diff_pct:+.2f}%)")
    
    if abs(labor_diff_pct) > 25:
        st.info("💡 **Note:** Your selected labor rate differs significantly from the historical labor rate in this comparable set. Current-rate projection may differ from the historical benchmark primarily because of labor-rate normalization.")

else:
    tab_hist, tab_prod, tab_proj = st.tabs([
        "📚 Historical Raw Data", 
        "👷 Productivity Context (SQ/MD)",
        "📊 Current-Rate Projection"
    ])
    
    with tab_hist:
        c1, c2, c3 = st.columns([1, 1.2, 1.5])
        with c1:
            st.metric("Historical Median Benchmark Total", f"${cps_med * sample_squares:,.2f}")
            st.metric("Median Direct Cost/SQ", f"${cps_med:,.2f}")
        with c2:
            st.markdown("##### Historical Ranges (Cost/SQ)")
            st.markdown(f"**Low (25th):** \${cps_low:,.2f}")
            st.markdown(f"**Median (50th):** \${cps_med:,.2f}")
            st.markdown(f"**High (75th):** \${cps_high:,.2f}")
        with c3:
            st.markdown("##### Data Quality")
            cps_spread = cps_high - cps_low
            cps_spread_pct = (cps_spread / cps_med) * 100 if cps_med > 0 else 0
            
            if cps_spread_pct < 20:
                q_label = "Excellent"
                q_color = "#28a745"
            elif cps_spread_pct < 40:
                q_label = "Good"
                q_color = "#e0a800"
            elif cps_spread_pct < 70:
                q_label = "Fair"
                q_color = "#fd7e14"
            else:
                q_label = "Poor"
                q_color = "#dc3545"
                
            marker_pos = min(cps_spread_pct, 100)
            
            bar_html = f"""
            <div style="padding: 15px 15px 20px 15px; background-color: #f8f9fa; border-radius: 8px; border: 1px solid #e9ecef; margin-top: 5px;">
                <div style="display: flex; justify-content: space-between; font-size: 0.9rem; margin-bottom: 12px;">
                    <span style="color: #495057;">Variance: <strong>{cps_spread_pct:.1f}%</strong></span>
                    <span style="color: {q_color}; font-weight: 600;">{q_label} Confidence</span>
                </div>
                <div style="position: relative; width: 100%; height: 12px; background: linear-gradient(to right, #28a745 20%, #ffc107 40%, #fd7e14 70%, #dc3545 100%); border-radius: 6px;">
                    <div style="position: absolute; top: -6px; bottom: -6px; left: {marker_pos}%; width: 4px; background-color: #212529; border: 2px solid white; border-radius: 3px; transform: translateX(-50%); box-shadow: 0 0 4px rgba(0,0,0,0.4);"></div>
                </div>
                <div style="display: flex; justify-content: space-between; font-size: 0.75rem; color: #6c757d; margin-top: 8px;">
                    <span>0% (Tight)</span>
                    <span>100%+ (Wide)</span>
                </div>
            </div>
            """
            st.markdown(bar_html, unsafe_allow_html=True)
        st.caption("*Historical roof-system benchmark excluding scaffold and specialty items.*")
            
    with tab_prod:
        sqmd_vals = []
        for j in valid_labor_jobs:
            try:
                sqmd = float(j.get("SQ/MD", 0))
                if sqmd > 0: sqmd_vals.append(sqmd)
            except: pass
            
        if sqmd_vals:
            sqmd_low = np.percentile(sqmd_vals, 25)
            sqmd_med = np.median(sqmd_vals)
            sqmd_high = np.percentile(sqmd_vals, 75)
        else:
            sqmd_low = sqmd_med = sqmd_high = 0.0
            
        c1, c2, c3 = st.columns([1.2, 1.2, 1.6])
        with c1:
            st.metric("Weighted Average SQ/MD", f"{weighted_sq_md:,.2f}")
            st.metric("Median SQ/MD", f"{sqmd_med:,.2f}")
        with c2:
            st.markdown("##### Historical Ranges")
            st.markdown(f"**Low (25th):** {sqmd_low:,.2f}")
            st.markdown(f"**Median (50th):** {sqmd_med:,.2f}")
            st.markdown(f"**High (75th):** {sqmd_high:,.2f}")
        with c3:
            st.markdown("##### Productivity Reliability")
            sqmd_spread_val = sqmd_high - sqmd_low
            sqmd_spread_pct = (sqmd_spread_val / sqmd_med) * 100 if sqmd_med > 0 else 0
            
            if sqmd_spread_pct < 20:
                sqmd_quality_label = "Excellent"
                sqmd_q_color = "#28a745"
            elif sqmd_spread_pct < 40:
                sqmd_quality_label = "Good"
                sqmd_q_color = "#e0a800"
            elif sqmd_spread_pct < 70:
                sqmd_quality_label = "Fair"
                sqmd_q_color = "#fd7e14"
            else:
                sqmd_quality_label = "Poor"
                sqmd_q_color = "#dc3545"
                
            sqmd_marker_pos = min(sqmd_spread_pct, 100)
            
            sqmd_bar_html = f"""
            <div style="padding: 15px 15px 20px 15px; background-color: #f8f9fa; border-radius: 8px; border: 1px solid #e9ecef; margin-top: 5px;">
                <div style="display: flex; justify-content: space-between; font-size: 0.9rem; margin-bottom: 12px;">
                    <span style="color: #495057;">Variance: <strong>{sqmd_spread_pct:.1f}%</strong></span>
                    <span style="color: {sqmd_q_color}; font-weight: 600;">{sqmd_quality_label} Confidence</span>
                </div>
                <div style="position: relative; width: 100%; height: 12px; background: linear-gradient(to right, #28a745 20%, #ffc107 40%, #fd7e14 70%, #dc3545 100%); border-radius: 6px;">
                    <div style="position: absolute; top: -6px; bottom: -6px; left: {sqmd_marker_pos}%; width: 4px; background-color: #212529; border: 2px solid white; border-radius: 3px; transform: translateX(-50%); box-shadow: 0 0 4px rgba(0,0,0,0.4);"></div>
                </div>
                <div style="display: flex; justify-content: space-between; font-size: 0.75rem; color: #6c757d; margin-top: 8px;">
                    <span>0% (Tight)</span>
                    <span>100%+ (Wide)</span>
                </div>
            </div>
            """
            st.markdown(sqmd_bar_html, unsafe_allow_html=True)
            
        st.caption("*Historical roof-system productivity metrics (Squares per Man Day).*")
        
    with tab_proj:
        if tier_used == "Tier 4 (Global Fallback)":
            st.error("Broad reference only — not reliable for estimating.")
        else:
            c1, c2 = st.columns([1, 2])
            with c1:
                st.metric("Projected Direct Cost Total", f"${view_b_direct_cost:,.2f}")
                st.metric("Projected Direct Cost/SQ", f"${(view_b_direct_cost/sample_squares) if sample_squares>0 else 0:,.2f}")
            with c2:
                st.markdown("##### Expected Breakdowns")
                sq_safe = sample_squares if sample_squares > 0 else 1
                st.markdown(f"**Type G (Indirects):** \${expected_type_g:,.2f} (*\${expected_type_g / sq_safe:,.2f}/SQ*)")
                st.markdown(f"**Labor (at \${labor_rate:,.0f}/MD):** \${est_labor_cost:,.2f} (*\${est_labor_cost / sq_safe:,.2f}/SQ*)")
                st.markdown(f"**Materials:** \${expected_mats:,.2f} (*\${expected_mats / sq_safe:,.2f}/SQ*)")
            st.caption("*Historical roof-system productivity and cleaned cost patterns priced with today's labor rate.*")
            
            st.markdown("---")
            st.markdown("##### 👷 Labor Rate Context")
            hist_labor_rates = []
            for j in valid_labor_jobs:
                l = float(j.get("Total Labor", 0))
                md = float(j.get("Estimator MD", 0))
                if md > 0:
                    hist_labor_rates.append(l / md)
            median_hist_labor_rate = np.median(hist_labor_rates) if hist_labor_rates else 0.0
            
            labor_diff = labor_rate - median_hist_labor_rate
            labor_diff_pct = (labor_diff / median_hist_labor_rate) * 100 if median_hist_labor_rate > 0 else 0
            
            c3, c4, c5 = st.columns(3)
            c3.metric("Your Selected Rate", f"${labor_rate:,.2f}/MD")
            c4.metric("Median Historical Labor", f"${median_hist_labor_rate:,.2f}/MD")
            c5.metric("Difference", f"${labor_diff:,.2f}", f"{labor_diff_pct:+.2f}%" if median_hist_labor_rate > 0 else None, delta_color="inverse")
            
            if abs(labor_diff_pct) > 25:
                st.info("💡 **Note:** Your selected labor rate differs significantly from the historical labor rate in this comparable set. Current-rate projection may differ from the historical benchmark primarily because of labor-rate normalization.")

st.markdown('<div class="print-page-break"></div>', unsafe_allow_html=True)
st.markdown("---")
st.header("Comparable Jobs")

# Render Contextual Multiselect Directly Above Tables
def clear_exclusions():
    st.session_state["exclude_dropdown"] = []

st.multiselect(
    "Temporarily exclude comparable jobs from this calculation",
    options=list(all_jobs_context.keys()),
    key="exclude_dropdown",
    help="Select jobs that you know had catastrophic anomalies (bad access, terrible crew) to prevent them from skewing the median."
)

if st.session_state.get("exclude_dropdown"):
    excluded_nos = [all_jobs_context.get(lbl, "Unknown") for lbl in st.session_state["exclude_dropdown"]]
    st.warning(f"Currently excluding {len(st.session_state['exclude_dropdown'])} comparable job(s) from the benchmark calculation: {', '.join(excluded_nos)}")
    st.button("Clear exclusions", on_click=clear_exclusions)

if matched_jobs:
    tier_labels = {
        1: "Tier 1 (Strict)",
        2: "Tier 2 (Core Drivers)",
        3: "Tier 3 (Broad Labor)",
        4: "Tier 4 (Global Fallback)"
    }
    
    seen_job_nos = set()
    all_html_tables = []
    
    for t_level in range(1, active_tier_level + 1):
        tier_key = f"t{t_level}"
        raw_tier_jobs = tier_jobs_dict.get(tier_key, [])
        
        # Mutually exclusive filter
        exclusive_jobs = []
        for j in raw_tier_jobs:
            jno = j.get("Job#")
            if jno and jno not in seen_job_nos:
                j_copy = dict(j)
                # Stamp the composite Spec Code on every job (all tiers)
                j_copy["Spec Code"] = build_spec_code(j_copy)
                exclusive_jobs.append(j_copy)
                seen_job_nos.add(jno)
                
        # Calculate medians for this specific tier table
        tier_sq_md_vals = []
        tier_cost_sq_vals = []
        for j in exclusive_jobs:
            try:
                sq_md = float(j.get("SQ/MD", 0))
                if sq_md > 0: tier_sq_md_vals.append(sq_md)
            except (ValueError, TypeError): pass
            
            try:
                c_sq = float(j.get("$/SQ Ind", 0))
                if c_sq > 0: tier_cost_sq_vals.append(c_sq)
            except (ValueError, TypeError): pass
            
        tier_med_sq_md = f"{np.median(tier_sq_md_vals):.2f}" if tier_sq_md_vals else "N/A"
        tier_med_cost_sq = f"${np.median(tier_cost_sq_vals):,.2f}" if tier_cost_sq_vals else "N/A"
        
        col1, col2, col3 = st.columns([2, 1, 1])
        with col1:
            st.markdown(f"### {tier_labels.get(t_level, f'Tier {t_level}')} Matches ({len(exclusive_jobs)} jobs)")
        with col2:
            st.metric("Tier Median $/SQ", tier_med_cost_sq)
        with col3:
            st.metric("Tier Median SQ/MD", tier_med_sq_md)
        
        comp_cols = ["Job#", "Address", "City", "Spec Code", "Total Squares", "Estimator MD", "SQ/MD", 
                       "Type G Total", "Scaffold Excluded", "Total Labor", "Type M Total", "Special Items Excluded", 
                       "Original Historical Direct Cost", "Clean Benchmark Direct Cost", "Clean Benchmark Direct Cost/SQ",
                       "Total Cost (Independent)", "$/SQ Ind"]
                       
        if exclusive_jobs:
            comp_df = pd.DataFrame(exclusive_jobs)
            existing_cols = [c for c in comp_cols if c in comp_df.columns]
            df_comp_display = comp_df[existing_cols].copy()
        else:
            df_comp_display = pd.DataFrame(columns=comp_cols)
            

        # Populate Aerial View 🛰️ link column
        df_comp_display["Aerial View 🛰️"] = df_comp_display.apply(
            lambda r: make_google_maps_link(
                r.get("Address") if "Address" in r.index else None,
                r.get("City") if "City" in r.index else None
            ),
            axis=1
        )
        
        # Reorder columns to place "Aerial View 🛰️" right after "City" or "Address"
        cols = list(df_comp_display.columns)
        if "Aerial View 🛰️" in cols:
            cols.remove("Aerial View 🛰️")
            insert_idx = cols.index("City") + 1 if "City" in cols else (cols.index("Address") + 1 if "Address" in cols else 1)
            cols.insert(insert_idx, "Aerial View 🛰️")
            df_comp_display = df_comp_display[cols]
            
        # Default sort by Total Squares ascending
        if "Total Squares" in df_comp_display.columns:
            df_comp_display = df_comp_display.sort_values(by="Total Squares", ascending=True)
            
        df_comp_display.rename(columns={
            "Total Cost (Independent)": "Total Cost",
            "$/SQ Ind": "$/SQ"
        }, inplace=True)
            
        event = st.dataframe(
            df_comp_display,
            hide_index=True,
            use_container_width=True,
            selection_mode="single-row",
            on_select="rerun",
            key=f"tier_table_{t_level}",
            column_config={
                "Aerial View 🛰️": st.column_config.LinkColumn(
                    "Aerial View 🛰️",
                    help="Click to view satellite map",
                    display_text="View Map 🗺️"
                ),
                "Total Squares": st.column_config.NumberColumn("Total Squares", format="%,.2f"),
                "Estimator MD": st.column_config.NumberColumn("Estimator MD", format="%,.2f"),
                "SQ/MD": st.column_config.NumberColumn("SQ/MD", format="%,.2f"),
                "Type G Total": st.column_config.NumberColumn("Type G Total", format="$%,.2f"),
                "Scaffold Excluded": st.column_config.NumberColumn("Scaffold Excluded", format="$%,.2f"),
                "Total Labor": st.column_config.NumberColumn("Total Labor", format="$%,.2f"),
                "Type M Total": st.column_config.NumberColumn("Type M Total", format="$%,.2f"),
                "Special Items Excluded": st.column_config.NumberColumn("Special Items Excluded", format="$%,.2f"),
                "Original Historical Direct Cost": st.column_config.NumberColumn("Original Historical Direct Cost", format="$%,.2f"),
                "Clean Benchmark Direct Cost": st.column_config.NumberColumn("Clean Benchmark Direct Cost", format="$%,.2f"),
                "Clean Benchmark Direct Cost/SQ": st.column_config.NumberColumn("Clean Benchmark Direct Cost/SQ", format="$%,.2f"),
                "Total Cost": st.column_config.NumberColumn("Total Cost", format="$%,.2f"),
                "$/SQ": st.column_config.NumberColumn("$/SQ", format="$%,.2f")
            }
        )

        if event and len(event.selection.rows) > 0:
            selected_idx = event.selection.rows[0]
            selected_job_row = df_comp_display.iloc[selected_idx]
            job_no = selected_job_row.get("Job#")
            orig_job = next((j for j in exclusive_jobs if j.get("Job#") == job_no), None)

            if orig_job and t_level == 1:
                # --- Tier 1: Cost Code Breakdown ---
                job_sq = pd.to_numeric(orig_job.get("Total Squares"), errors='coerce')
                job_sq = float(job_sq) if not pd.isna(job_sq) else 0.0
                job_md = pd.to_numeric(orig_job.get("Estimator MD"), errors='coerce')
                job_md = float(job_md) if not pd.isna(job_md) else 0.0

                breakdown_rows = []
                pred_g_total = 0.0
                pred_m_total = 0.0

                # Type G rows
                for col in COST_BREAKDOWN_TYPE_G_COLS:
                    if col in COST_BREAKDOWN_EXCLUSIONS:
                        continue
                    raw_val = pd.to_numeric(orig_job.get(col), errors='coerce')
                    raw_val = float(raw_val) if not pd.isna(raw_val) else 0.0
                    if raw_val <= 0.0:
                        continue
                    display_name = DISPLAY_NAME_MAP.get(col, col)
                    per_sq = raw_val / job_sq if job_sq > 0 else 0.0
                    predicted = per_sq * sample_squares
                    pred_g_total += predicted
                    breakdown_rows.append({"Category": display_name, "Type": "G", "$/SQ": per_sq, "Predicted Cost": predicted})

                # Type M rows
                for col in COST_BREAKDOWN_TYPE_M_COLS:
                    if col in COST_BREAKDOWN_EXCLUSIONS:
                        continue
                    raw_val = pd.to_numeric(orig_job.get(col), errors='coerce')
                    raw_val = float(raw_val) if not pd.isna(raw_val) else 0.0
                    if raw_val <= 0.0:
                        continue
                    display_name = DISPLAY_NAME_MAP.get(col, col)
                    per_sq = raw_val / job_sq if job_sq > 0 else 0.0
                    predicted = per_sq * sample_squares
                    pred_m_total += predicted
                    breakdown_rows.append({"Category": display_name, "Type": "M", "$/SQ": per_sq, "Predicted Cost": predicted})

                # Normalized Labor row
                norm_labor = job_md * labor_rate
                labor_per_sq = norm_labor / job_sq if job_sq > 0 else 0.0
                pred_labor_total = labor_per_sq * sample_squares
                breakdown_rows.append({"Category": "02-0110 - Labor - Roofing (Normalized)", "Type": "L", "$/SQ": labor_per_sq, "Predicted Cost": pred_labor_total})

                if breakdown_rows:
                    df_breakdown = pd.DataFrame(breakdown_rows)
                    st.markdown(f"##### 📋 Cost Breakdown — Job {job_no} (Predicted for {sample_squares:,.2f} SQ)")
                    st.dataframe(
                        df_breakdown[["Category", "$/SQ", "Predicted Cost"]],
                        hide_index=True,
                        use_container_width=True,
                        column_config={
                            "Category": st.column_config.TextColumn("Category", width="large"),
                            "$/SQ": st.column_config.NumberColumn("$/SQ", format="$%,.2f"),
                            "Predicted Cost": st.column_config.NumberColumn("Predicted Cost", format="$%,.2f"),
                        }
                    )
                    # Summary metrics
                    pred_total = pred_g_total + pred_m_total + pred_labor_total
                    sc1, sc2, sc3, sc4 = st.columns(4)
                    sc1.metric("Predicted Type G", f"${pred_g_total:,.2f}")
                    sc2.metric("Predicted Type M", f"${pred_m_total:,.2f}")
                    sc3.metric("Predicted Labor", f"${pred_labor_total:,.2f}")
                    sc4.metric("Total Predicted Cost", f"${pred_total:,.2f}")

            elif orig_job and t_level > 1:
                # --- Tier 2/3: Filter Mismatch Detail ---
                st.info(f"**Job {job_no} Mismatches:**")
                mismatch_bullets = []
                for k, v_list in filters.items():
                    if not job_matches_filter(orig_job, k, v_list):
                        expected = ", ".join([str(v) for v in v_list])
                        actual = orig_job.get(k, 'N/A')
                        mismatch_bullets.append(f"* **{k}:** Filter expected [{expected}], Job has [{actual}]")
                if mismatch_bullets:
                    st.markdown("\n".join(mismatch_bullets))

        # Build clean print-only HTML table
        df_print = df_comp_display.copy()
        if "Aerial View 🛰️" in df_print.columns:
            df_print = df_print.drop(columns=["Aerial View 🛰️"])
            
        def format_print_val(col_name, val):
            if pd.isna(val) or val is None:
                return ""
            val_str = str(val).strip()
            if val_str.lower() in ["nan", "none", "n/a", ""]:
                return ""
            try:
                float_val = float(val)
                if col_name in ["Total Squares", "Estimator MD", "SQ/MD"]:
                    return f"{float_val:,.2f}"
                elif col_name in ["Type G Total", "Scaffold Excluded", "Total Labor", "Type M Total", "Special Items Excluded", "Original Historical Direct Cost", "Clean Benchmark Direct Cost", "Clean Benchmark Direct Cost/SQ", "Total Cost", "$/SQ"]:
                    return f"${float_val:,.2f}"
            except (ValueError, TypeError):
                pass
            return val_str

        html_rows = []
        for _, row in df_print.iterrows():
            row_html = "<tr>"
            for col in df_print.columns:
                val_str = format_print_val(col, row[col])
                align = "left" if col in ["Job#", "Address", "City", "Spec Type"] else "right"
                row_html += f'<td style="text-align: {align}; padding: 3px 5px; border-bottom: 1px solid #cbd5e1; white-space: nowrap;">{val_str}</td>'
            row_html += "</tr>"
            html_rows.append(row_html)

        headers_html = "".join([f'<th style="text-align: {"left" if col in ["Job#", "Address", "City", "Spec Type"] else "right"}; padding: 4px 5px; border-bottom: 2px solid #64748b; font-weight: 700; white-space: nowrap;">{col}</th>' for col in df_print.columns])

        print_table_html = f"""<div class="print-only" style="width: 100%; font-family: 'Inter', -apple-system, sans-serif; font-size: 8pt; color: #1e293b; margin-top: 20px;">
<div style="display: flex; justify-content: space-between; align-items: baseline; border-bottom: 2px solid #cbd5e1; margin-bottom: 5px; padding-bottom: 5px;">
    <h4 style="margin: 0; color: #334155; font-size: 14px;">{tier_labels.get(t_level, f'Tier {t_level}')} Matches ({len(exclusive_jobs)} jobs)</h4>
    <div style="display: flex; gap: 20px; font-size: 12px; font-weight: 600; color: #475569;">
        <span>Median $/SQ: {tier_med_cost_sq}</span>
        <span>Median SQ/MD: {tier_med_sq_md}</span>
    </div>
</div>
<table style="width: 100%; border-collapse: collapse;">
<thead>
<tr style="background-color: #f8fafc; border-top: 1px solid #cbd5e1;">{headers_html}</tr>
</thead>
<tbody>
{"".join(html_rows)}
</tbody>
</table>
</div>"""
        all_html_tables.append(print_table_html)
            
    if all_html_tables:
        st.markdown("".join(all_html_tables), unsafe_allow_html=True)

# Export Functionality
st.markdown('<div class="print-page-break"></div>', unsafe_allow_html=True)
st.markdown("---")
if mats_suppressed:
    st.info("Shingle-related material categories were hidden from this recap because the active roof material filter is TPO. Source job totals were not modified.")

st.subheader("Expected Cost Recap")
st.markdown("### Sample Recap")
if not df_recap.empty and tier_used != "Tier 4 (Global Fallback)":
    df_recap_display = df_recap.copy()
    styled_df = df_recap_display.style.hide(axis="index")\
        .format({"Median Rate": "${:,.2f}", "Expected Total": "${:,.2f}"})\
        .set_properties(**{'text-align': 'right'}, subset=['Median Rate', 'Expected Total'])\
        .set_properties(**{'font-weight': 'bold'}, subset=['Expected Total'])\
        .set_table_styles([
            {'selector': 'th', 'props': [('text-align', 'left')]},
            {'selector': 'td', 'props': [('padding', '8px')]}
        ])
    st.table(styled_df)

def to_excel():
    export_jobs = []
    
    tier_labels = {
        1: "Tier 1 (Strict)",
        2: "Tier 2 (Core Drivers)",
        3: "Tier 3 (Broad Labor)",
        4: "Tier 4 (Global Fallback)"
    }
    
    seen_job_nos = set()
    for t_level in range(1, active_tier_level + 1):
        tier_key = f"t{t_level}"
        raw_tier_jobs = tier_jobs_dict.get(tier_key, [])
        for j in raw_tier_jobs:
            jno = j.get("Job#")
            if jno and jno not in seen_job_nos:
                seen_job_nos.add(jno)
                export_jobs.append({
                    "Tier Match": tier_labels.get(t_level, f"Tier {t_level}"),
                    "Job#": j.get("Job#", ""),
                    "Address": j.get("Address", ""),
                    "Total Squares": j.get("Total Squares", 0),
                    "Type G Total": j.get("Type G Total", 0),
                    "Total Labor": j.get("Total Labor", 0),
                    "Type M Total": j.get("Type M Total", 0),
                    "Original Historical Direct Cost": j.get("Original Historical Direct Cost", 0),
                    "Scaffold Excluded": j.get("Scaffold", 0),
                    "Special Items Excluded": j.get("03-0120 Special (Skylights/Hatches)", 0),
                    "Clean Benchmark Direct Cost": j.get("Clean Benchmark Direct Cost", 0),
                    "Clean Benchmark Direct Cost/SQ": j.get("Clean Benchmark Direct Cost/SQ", 0),
                    "Total Cost": j.get("Total Cost (Independent)", 0),
                    "$/SQ": j.get("$/SQ Ind", 0)
                })
        
    df_export_jobs = pd.DataFrame(export_jobs)
    excel_buffer = io.BytesIO()
    with pd.ExcelWriter(excel_buffer, engine='openpyxl') as writer:
        df_recap.to_excel(writer, sheet_name="Recap", index=False)
        df_export_jobs.to_excel(writer, sheet_name="Comparable Jobs", index=False)
        ws_recap = writer.sheets["Recap"]
        ws_recap.append([])
        ws_recap.append(["Note: Permit, Disposal, and Fuel estimators are independent allowance tools and do not affect Clean Benchmark Direct Cost."])
        
        ws_recap.append([])
        ws_recap.append(["Permit Cost Estimator"])
        if permit_export_data.get("Active"):
            ws_recap.append(["City Filter", permit_export_data["City"]])
            ws_recap.append(["Data Points Used", permit_export_data["Count"]])
            ws_recap.append(["Median Permit $/SQ", f"${permit_export_data['Median $/SQ']:,.2f}"])
            ws_recap.append(["Low Permit Estimate (Square-Based)", f"${permit_export_data['Low']:,.2f}"])
            ws_recap.append(["Median Permit Estimate (Square-Based)", f"${permit_export_data['Median']:,.2f}"])
            ws_recap.append(["High Permit Estimate (Square-Based)", f"${permit_export_data['High']:,.2f}"])
            if "Contract_Median" in permit_export_data:
                ws_recap.append(["Median Permit % of Contract", f"{permit_export_data['Contract_Pct_Med']:.3f}%"])
                ws_recap.append(["Low Permit Estimate (Contract-Based)", f"${permit_export_data['Contract_Low']:,.2f}"])
                ws_recap.append(["Median Permit Estimate (Contract-Based)", f"${permit_export_data['Contract_Median']:,.2f}"])
                ws_recap.append(["High Permit Estimate (Contract-Based)", f"${permit_export_data['Contract_High']:,.2f}"])
            else:
                ws_recap.append(["Contract-Based Estimate", "Insufficient contract-value data"])
        else:
            ws_recap.append(["Status", permit_export_data.get("Msg", "Inactive")])
            
        ws_recap.append([])
        ws_recap.append(["Disposal Cost Estimator"])
        if disposal_export_data.get("Active"):
            ws_recap.append(["City Filter", disposal_export_data["City"]])
            ws_recap.append(["Data Points Used", disposal_export_data["Count"]])
            ws_recap.append(["Median Disposal $/SQ", f"${disposal_export_data['Median $/SQ']:,.2f}"])
            ws_recap.append(["Low Disposal Estimate", f"${disposal_export_data['Low']:,.2f}"])
            ws_recap.append(["Median Disposal Estimate", f"${disposal_export_data['Median']:,.2f}"])
            ws_recap.append(["High Disposal Estimate", f"${disposal_export_data['High']:,.2f}"])
        else:
            ws_recap.append(["Status", disposal_export_data.get("Msg", "Inactive")])
            
        ws_recap.append([])
        ws_recap.append(["Fuel Cost Estimator"])
        if fuel_export_data.get("Active"):
            ws_recap.append(["City Filter", fuel_export_data["City"]])
            ws_recap.append(["Data Points Used", fuel_export_data["Count"]])
            ws_recap.append(["Median Fuel $/MD", f"${fuel_export_data['Median $/MD']:,.2f}"])
            ws_recap.append(["Low Fuel Estimate", f"${fuel_export_data['Low']:,.2f}"])
            ws_recap.append(["Median Fuel Estimate", f"${fuel_export_data['Median']:,.2f}"])
            ws_recap.append(["High Fuel Estimate", f"${fuel_export_data['High']:,.2f}"])
            if "Zone" in fuel_export_data:
                ws_recap.append(["Secondary Zone Reference", fuel_export_data["Zone"]])
                ws_recap.append(["Zone Median $/MD", f"${fuel_export_data['Zone_Med']:,.2f}"])
        else:
            ws_recap.append(["Status", fuel_export_data.get("Msg", "Inactive")])
            if "Zone" in fuel_export_data:
                ws_recap.append(["Zone Fallback Used", fuel_export_data["Zone"]])
                ws_recap.append(["Zone Median $/MD", f"${fuel_export_data['Zone_Med']:,.2f}"])

        for _ in range(3):
            ws_recap.append([])
        ws_recap.append(["--- ESTIMATING MODEL ---"])
        ws_recap.append(["Estimating Model Output Generated by Historical Benchmark Tool"])
        ws_recap.append(["Note: Clean benchmark excludes Scaffold and 03-0120 Special (Skylights/Hatches)."])
    return excel_buffer.getvalue()

if st.button("Download Recap Excel"):
    excel_data = to_excel()
    st.download_button(label="Click here to download", data=excel_data, file_name="Estimated_Recap.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


st.markdown("---")
st.info("**Note:** This tool provides a historical baseline based on past performance. It is a sanity-check mechanism and does not account for unique site conditions, extreme chop/complexity, or current supply chain volatility.")
