import streamlit as st
import pandas as pd
import numpy as np
import altair as alt
from pathlib import Path
import os
import io

st.set_page_config(page_title="Job Simulation", layout="wide")

# Global CSS
st.markdown("""
<style>
/* Premium Metric aesthetics */
.premium-metric { text-align: center; background-color: #f8f9fa; padding: 1.5rem; border-radius: 8px; border: 1px solid #e0e0e0; }
.premium-title { color: #555; margin-bottom: 0; font-size: 1.1rem; }
.premium-value { color: #0F52BA; margin-top: 0; font-size: 2.5rem; font-weight: bold; line-height: 1.2; }
.premium-sub { color: #888; margin-top: 0; font-size: 1rem; }

/* Step Headers */
.step-header { color: #0F52BA; margin-bottom: 1.5rem; border-bottom: 2px solid #0F52BA; padding-bottom: 0.5rem; }

/* Print styling */
@media print {
    section[data-testid="stSidebar"], div.stButton, header, footer {
        display: none !important;
    }
    section.main {
        width: 100% !important;
        padding: 0 !important;
        margin: 0 !important;
    }
}
</style>
""", unsafe_allow_html=True)

PROJECT_ROOT = Path(__file__).resolve().parent.parent
LOCAL_FILE = str(PROJECT_ROOT / "data" / "Demo_Master_List.xlsx")

@st.cache_data(ttl=60)
def load_sim_data():
    if not os.path.exists(LOCAL_FILE):
        return None
    try:
        df = pd.read_excel(LOCAL_FILE, sheet_name="FULL LIST")
    except Exception as e:
        st.error(f"Error reading file: {e}")
        return None
        
    # Clean string columns
    text_cols = ["City", "Address", "Spec Type", "Insulation R-Value", "Insulation Attachment", "Cover Board Type", "Cover Board Thickness", 
                 "Cover Board Attachment", "Roof Material Type", "Roof Material Thickness", "Roof Material Attachment"]
    for col in text_cols:
        if col in df.columns:
            df[col] = df[col].fillna("")
            df[col] = df[col].astype(str).str.strip()
            df[col] = df[col].replace(['nan', 'None', 'NaN', '<NA>'], '')
            
    # Clean numeric columns
    num_cols = ["Total Squares", "Type M - Benchmark", "Man Days", "Permit", "Disposal", "Equipment", "Fuel", "Outside Services"]
    
    col_map = {
        "01-0100 - General - Permits": "Permit",
        "01-0110 - General - Disposal": "Disposal",
        "01-0120 - General - Equipment Rental": "Equipment",
        "01-0160 - General - Fuel": "Fuel",
        "01-0200 - General-Outside Services": "Outside Services"
    }
    df = df.rename(columns=col_map)
    
    for col in num_cols:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0.0)
            
    # Parse dates
    date_cols = ["Complete", "Report"]
    for col in date_cols:
        if col in df.columns:
            df[col] = pd.to_datetime(df[col], errors='coerce')
            
    return df

df_all = load_sim_data()

st.title("Job Simulation")

if df_all is None:
    st.error("Dashboard data not found. Please run the Historical Dashboard and click 'Update Dashboard Data'.")
    st.stop()

# Get unique options, removing empties
def get_options(col_name):
    if col_name not in df_all.columns:
        return ["Any"]
    opts = df_all[col_name].unique()
    valid_opts = [str(o) for o in opts if o and str(o).lower() not in ('nan', 'none', '')]
    
    if col_name == "Spec Type":
        valid_opts = [o for o in valid_opts if o.lower() != "master list"]
        return [""] + sorted(valid_opts)
        
    return ["Any", "(Blank)"] + sorted(valid_opts)


# Initialize all calculation variables so they exist even if filters aren't selected yet
city_filter = ""
spec_type = ""
tear_off = False
deck_type = ""
to_material = ""
to_weight = ""
weighted_avg_mat_cost = 0.0
weighted_avg_lab_cost = 0.0
final_gen_cost = 0.0
tot_mat_cost = 0.0
total_lab_cost = 0.0
base_cost = 0.0
gross_profit = 0.0
final_sell = 0.0
gp_pct = 0.0
final_sell_sq = 0.0
tot_overhead = 0.0
tot_profit = 0.0
est_md = 0.0
proj_sq_md = 0.0

# --- STEP 1: JOB INFO ---
with st.container(border=True):
    st.markdown("<h2 class='step-header'>Job Info</h2>", unsafe_allow_html=True)
    st.markdown('<div class="hide-me-print"></div>', unsafe_allow_html=True)
    col1, col2 = st.columns(2)
    total_squares = col1.number_input("Total Squares", value=500.0, step=10.0, min_value=0.1)
    
    spec_type = col2.selectbox("Spec Type", get_options("Spec Type"))
    
    col4, col5, col6 = st.columns(3)
    insulation = col4.selectbox("Insulation", get_options("Insulation R-Value"))
    cb_type = col5.selectbox("Cover Board", get_options("Cover Board Type"))
    mat_type = col6.selectbox("Roof Material", get_options("Roof Material Type"))

    col10, col11, col12 = st.columns(3)
    cb_thick = col11.selectbox("Cover Board Thickness", get_options("Cover Board Thickness"))
    mat_thick = col12.selectbox("Roof Material Thickness", get_options("Roof Material Thickness"))

    col7, col8, col9 = st.columns(3)
    ins_att = col7.selectbox("Insulation Attachment", get_options("Insulation Attachment"))
    cb_att = col8.selectbox("Cover Board Attachment", get_options("Cover Board Attachment"))
    mat_att = col9.selectbox("Roof Material Attachment", get_options("Roof Material Attachment"))

# --- STEP 2: MATERIALS SIMULATION ---
with st.container(border=True):
    st.markdown("<h2 class='step-header'>Materials Simulation</h2>", unsafe_allow_html=True)
    st.markdown('<div class="hide-me-print"></div>', unsafe_allow_html=True)

if not spec_type or spec_type == "Any":
    st.info("Please select a Spec Type to run the Materials simulation.")
else:
    # 1. Exact Spec Filtering
    mask = (df_all["Spec Type"] == spec_type)
    
    def apply_filter(current_mask, col_name, selected_val):
        if selected_val == "Any":
            return current_mask
        elif selected_val == "(Blank)":
            return current_mask & (df_all[col_name] == "")
        else:
            return current_mask & (df_all[col_name] == selected_val)
            
    mask = apply_filter(mask, "Insulation R-Value", insulation)
    mask = apply_filter(mask, "Insulation Attachment", ins_att)
    mask = apply_filter(mask, "Cover Board Type", cb_type)
    mask = apply_filter(mask, "Cover Board Thickness", cb_thick)
    mask = apply_filter(mask, "Cover Board Attachment", cb_att)
    mask = apply_filter(mask, "Roof Material Type", mat_type)
    mask = apply_filter(mask, "Roof Material Thickness", mat_thick)
    mask = apply_filter(mask, "Roof Material Attachment", mat_att)
    
    df_filtered = df_all[mask].copy()
    
    if df_filtered.empty:
        st.warning("No historical jobs match this exact specification.")
    else:
        mat_col = "Type M - Benchmark"
        if mat_col not in df_filtered.columns:
            st.error(f"Required column '{mat_col}' not found in data.")
            st.stop()
        
        weighted_avg_mat_cost = 0.0

        # Calculate unweighted Cost Per SQ for each job
        df_filtered["Mat_Cost_Per_SQ"] = np.where(
            df_filtered["Total Squares"] > 0, 
            df_filtered[mat_col] / df_filtered["Total Squares"], 
            0.0
        )
        
        # 2. Relevance Weighting Logic
        # 1.0 weight if within +/- 50%, 0.5 otherwise
        def calc_weight(sq):
            if sq <= 0: return 0.0
            if (total_squares * 0.5) <= sq <= (total_squares * 1.5):
                return 1.0
            return 0.5
            
        df_filtered["Relevance_Weight"] = df_filtered["Total Squares"].apply(calc_weight)
        
        # Final Weight = Job Squares * Relevance Multiplier (per skill guidelines)
        df_filtered["Final_Weight"] = df_filtered["Total Squares"] * df_filtered["Relevance_Weight"]
        
        # 30-Day Rule Eligibility
        df_filtered["Days_Diff"] = (df_filtered["Report"] - df_filtered["Complete"]).dt.days
        df_filtered["Eligible"] = df_filtered["Days_Diff"] >= 30
        
        ineligible_mats = df_filtered[~df_filtered["Eligible"]]
        
        # Filter out $0 materials for accurate calculations, AND enforce Eligibility
        valid_mats = df_filtered[df_filtered["Eligible"] & (df_filtered["Mat_Cost_Per_SQ"] > 0)].copy()
        
        if valid_mats.empty:
            st.warning("Found matching jobs, but none are eligible or have valid material costs.")
        else:
            total_weighted_cost = (valid_mats["Mat_Cost_Per_SQ"] * valid_mats["Final_Weight"]).sum()
            total_weights = valid_mats["Final_Weight"].sum()
            
            weighted_avg_mat_cost = total_weighted_cost / total_weights if total_weights > 0 else 0.0
            
            st.markdown(
                f"""
                <div class='premium-metric'>
                    <h3 class='premium-title'>Projected Material Cost</h3>
                    <h1 class='premium-value'>${weighted_avg_mat_cost * total_squares:,.2f}</h1>
                    <p style='font-size: 1.2rem; color: #555; margin-bottom: 5px;'>${weighted_avg_mat_cost:,.2f} / SQ</p>
                    <p class='premium-sub'>Calculated from {len(valid_mats)} eligible comparable jobs</p>
                </div>
                <br>
                """, unsafe_allow_html=True
            )
            
            # --- Visual Chart ---
            st.markdown("##### Historical Data Points (Relevance Weighting)")
            
            # Label for legend
            valid_mats["Weight_Category"] = valid_mats["Relevance_Weight"].map({
                1.0: f"High Relevance ({(total_squares*0.5):.0f} - {(total_squares*1.5):.0f} SQ)", 
                0.5: "Baseline Relevance (Outside Range)"
            })
            
            chart = alt.Chart(valid_mats).mark_circle(size=80).encode(
                x=alt.X('Total Squares:Q', title='Job Size (Squares)'),
                y=alt.Y('Mat_Cost_Per_SQ:Q', title='Material Cost per SQ ($)', scale=alt.Scale(zero=False)),
                color=alt.Color('Weight_Category:N', scale=alt.Scale(
                    domain=[f"High Relevance ({(total_squares*0.5):.0f} - {(total_squares*1.5):.0f} SQ)", "Baseline Relevance (Outside Range)"], 
                    range=['#1f77b4', '#aec7e8']
                )),
                tooltip=['Job#', 'Total Squares', 'Mat_Cost_Per_SQ']
            ).interactive()
            
            # Vertical rule indicating the simulated job size
            rule = alt.Chart(pd.DataFrame({'x': [total_squares]})).mark_rule(color='red', strokeDash=[5, 5]).encode(x='x:Q')
            
            st.altair_chart(chart + rule, use_container_width=True)
            
            with st.expander(f"View Matching Jobs ({len(valid_mats)})"):
                display_df = valid_mats[['Job#', 'Address', 'City', 'Total Squares', 'Mat_Cost_Per_SQ', 'Relevance_Weight']].copy()
                display_df.rename(columns={'Mat_Cost_Per_SQ': 'Cost/SQ'}, inplace=True)
                
                # Sort by highest relevance weight, then by closest size
                display_df['Size_Diff'] = (display_df['Total Squares'] - total_squares).abs()
                display_df = display_df.sort_values(by=['Relevance_Weight', 'Size_Diff'], ascending=[False, True]).drop(columns=['Size_Diff'])
                
                st.dataframe(display_df.style.format({'Total Squares': '{:,.1f}', 'Cost/SQ': '${:,.2f}', 'Relevance_Weight': '{:.1f}'}))



# --- STEP 3: LABOR SIMULATION ---
with st.container(border=True):
    st.markdown("<h2 class='step-header'>Labor Simulation</h2>", unsafe_allow_html=True)
    st.markdown('<div class="hide-me-print"></div>', unsafe_allow_html=True)
    
    total_lab_cost = 0.0
    est_md = 0.0  # Initialize globally for use in Step 4
    
    with st.container(border=True):
        st.markdown("#### Estimator Controls")
        col_lab1, col_lab2 = st.columns(2, vertical_alignment="center")
        md_rate = col_lab1.number_input("Man Day Rate ($)", value=430.00, step=5.00, min_value=0.0)
        diff_slider = col_lab2.slider("Difficulty Factor (Production Rate Adjustment)", min_value=-25, max_value=25, value=0, step=5, format="%d%%", help="Slide left (-%) for slower production (harder job). Slide right (+%) for faster production (easier job).")


if not spec_type or spec_type == "Any":
    st.info("Please select a Spec Type to run the Labor simulation.")
else:
    # Helper to classify Cover Board
    def get_cb_class(cb_str):
        cb_str = str(cb_str).lower()
        if cb_str in ["cgf", "hd polyiso", "fan fold", "hd"]: return "Light"
        if cb_str in ["densdeck", "gypsum", "dens deck", "securerock"]: return "Heavy"
        return "Other"
        
    def get_cb_thickness(thick_str):
        if "1/2" in str(thick_str) or ".5" in str(thick_str): return "1/2"
        return "1/4"

    def get_efficiency(sq):
        if sq < 100: return 0.20
        if sq < 500: return 0.40
        if sq < 1000: return 0.70
        return 1.00

    sim_cb_class = get_cb_class(cb_type)
    sim_cb_thick = get_cb_thickness(cb_thick)
    sim_eff = get_efficiency(total_squares)
    
    # 1. Base Filter (Spec Type)
    lab_mask = (df_all["Spec Type"] == spec_type)
    
    # 2. Attachment Equivalency Filtering
    # Insulation
    if ins_att != "Any" and ins_att != "(Blank)":
        lab_mask = lab_mask & (df_all["Insulation Attachment"] == ins_att)
    elif ins_att == "(Blank)":
        lab_mask = lab_mask & (df_all["Insulation Attachment"] == "")
        
    # Cover Board
    if cb_att != "Any" and cb_att != "(Blank)":
        lab_mask = lab_mask & (df_all["Cover Board Attachment"] == cb_att)
    elif cb_att == "(Blank)":
        lab_mask = lab_mask & (df_all["Cover Board Attachment"] == "")
        
    # Roof Material (Mech == Rhino equivalent)
    if mat_att != "Any" and mat_att != "(Blank)":
        is_sim_mech = "mechanically attached" in mat_att.lower()
        is_sim_rhino = "rhinobond" in mat_att.lower()
        
        if is_sim_mech or is_sim_rhino:
            # Allow both
            lab_mask = lab_mask & (df_all["Roof Material Attachment"].str.lower().str.contains("mechanically attached|rhinobond", na=False))
        else:
            lab_mask = lab_mask & (df_all["Roof Material Attachment"] == mat_att)
    elif mat_att == "(Blank)":
        lab_mask = lab_mask & (df_all["Roof Material Attachment"] == "")

    df_lab = df_all[lab_mask].copy()
    
    if df_lab.empty:
        st.warning("No historical jobs match these attachment requirements.")
    else:
        if "Man Days" not in df_lab.columns:
            st.error("Required column 'Man Days' not found.")
            st.stop()
            
        # Calculate raw SQ/MD
        df_lab["Raw_SQ_MD"] = np.where(df_lab["Man Days"] > 0, df_lab["Total Squares"] / df_lab["Man Days"], 0)
        df_lab = df_lab[df_lab["Raw_SQ_MD"] > 0]
        
        if df_lab.empty:
            st.warning("No valid labor data found for matching jobs.")
        else:
            # Calculate adjustments
            def calculate_adjusted_sq_md(row):
                sq_md = row["Raw_SQ_MD"]
                hist_cb_class = get_cb_class(row["Cover Board Type"])
                hist_mat_att = str(row["Roof Material Attachment"]).lower()
                
                # A. Cover Board Adjustment
                cb_adj = 1.0
                if sim_cb_class == "Heavy" and hist_cb_class == "Light":
                    cb_adj = 0.75 if sim_cb_thick == "1/2" else 0.85
                elif sim_cb_class == "Light" and hist_cb_class == "Heavy":
                    hist_thick = get_cb_thickness(row["Cover Board Thickness"])
                    cb_adj = 1.25 if hist_thick == "1/2" else 1.15
                
                sq_md *= cb_adj
                
                # B. Rhino vs Mech Adjustment
                mat_adj = 1.0
                if "mechanically attached" in str(mat_att).lower() and "rhinobond" in hist_mat_att:
                    mat_adj = 0.95 # history was faster, reduce it
                elif "rhinobond" in str(mat_att).lower() and "mechanically attached" in hist_mat_att:
                    mat_adj = 1.05 # history was slower, boost it
                    
                sq_md *= mat_adj
                
                # C. Job Size Efficiency Adjustment
                hist_eff = get_efficiency(row["Total Squares"])
                size_adj = sim_eff / hist_eff
                
                sq_md *= size_adj
                
                return sq_md, cb_adj, mat_adj, size_adj
                
            adj_results = df_lab.apply(calculate_adjusted_sq_md, axis=1)
            df_lab["Simulated_SQ_MD"] = [r[0] for r in adj_results]
            df_lab["CB_Adj"] = [r[1] for r in adj_results]
            df_lab["Mat_Adj"] = [r[2] for r in adj_results]
            df_lab["Size_Adj"] = [r[3] for r in adj_results]
            
            # Relevance Weighting for Thickness
            def calc_lab_weight(row):
                w = 1.0
                if mat_thick != "Any" and mat_thick != "(Blank)":
                    if row["Roof Material Thickness"] != mat_thick:
                        w = 0.5
                return w
                
            df_lab["Relevance_Weight"] = df_lab.apply(calc_lab_weight, axis=1)
            df_lab["Final_Weight"] = df_lab["Total Squares"] * df_lab["Relevance_Weight"]
            
            # 30-Day Rule Eligibility
            df_lab["Days_Diff"] = (df_lab["Report"] - df_lab["Complete"]).dt.days
            df_lab["Eligible"] = df_lab["Days_Diff"] >= 30
            
            inel_lab = df_lab[~df_lab["Eligible"]]
            val_lab = df_lab[df_lab["Eligible"]].copy()
            
            if val_lab.empty:
                st.warning("Found matching jobs, but none are eligible under the 30-Day Rule.")
            else:
                total_weighted_sq_md = (val_lab["Simulated_SQ_MD"] * val_lab["Final_Weight"]).sum()
                total_weights = val_lab["Final_Weight"].sum()
                
                base_proj_sq_md = total_weighted_sq_md / total_weights if total_weights > 0 else 0.0
                proj_sq_md = base_proj_sq_md * (1.0 + (diff_slider / 100.0))
                
                est_md = total_squares / proj_sq_md if proj_sq_md > 0 else 0.0
                total_lab_cost = est_md * md_rate
                cost_per_sq = total_lab_cost / total_squares if total_squares > 0 else 0.0
                
                mc1, mc2, mc3 = st.columns(3)
                base_str = f"Base: {base_proj_sq_md:,.2f}" if diff_slider != 0 else "Base Rate"
                
                mc1.markdown(f"""
                <div class='premium-metric'>
                    <h3 class='premium-title'>Projected SQ / MD</h3>
                    <h1 class='premium-value'>{proj_sq_md:,.2f}</h1>
                    <p class='premium-sub'>{base_str}</p>
                </div>
                """, unsafe_allow_html=True)
                
                mc2.markdown(f"""
                <div class='premium-metric'>
                    <h3 class='premium-title'>Estimated Man Days</h3>
                    <h1 class='premium-value'>{est_md:,.2f}</h1>
                    <p class='premium-sub'>Total SQs / Rate</p>
                </div>
                """, unsafe_allow_html=True)
                
                mc3.markdown(f"""
                <div class='premium-metric'>
                    <h3 class='premium-title'>Total Labor Cost</h3>
                    <h1 class='premium-value'>${total_lab_cost:,.2f}</h1>
                    <p class='premium-sub'>${cost_per_sq:,.2f} / SQ</p>
                </div>
                """, unsafe_allow_html=True)
                
                st.markdown("<br>", unsafe_allow_html=True)
                
                # Chart
                st.markdown("##### Labor Efficiency Curve (Dynamic Scaling)")
                val_lab["Weight_Category"] = val_lab["Relevance_Weight"].map({1.0: "High Relevance (Matched Thick)", 0.5: "Baseline (Differing Thick)"})
                
                chart_lab = alt.Chart(val_lab).mark_circle(size=80).encode(
                    x=alt.X('Total Squares:Q', title='Historical Job Size (Squares)'),
                    y=alt.Y('Simulated_SQ_MD:Q', title='Simulated Production Rate (SQ/MD)', scale=alt.Scale(zero=False)),
                    color=alt.Color('Weight_Category:N', scale=alt.Scale(domain=["High Relevance (Matched Thick)", "Baseline (Differing Thick)"], range=['#2ca02c', '#98df8a'])),
                    tooltip=['Job#', 'Total Squares', 'Raw_SQ_MD', 'Simulated_SQ_MD']
                ).interactive()
                rule_lab = alt.Chart(pd.DataFrame({'x': [total_squares]})).mark_rule(color='red', strokeDash=[5, 5]).encode(x='x:Q')
                
                st.altair_chart(chart_lab + rule_lab, use_container_width=True)
                
                with st.expander(f"View Matching Labor Jobs ({len(val_lab)})"):
                    disp_lab = val_lab[['Job#', 'Total Squares', 'Raw_SQ_MD', 'CB_Adj', 'Mat_Adj', 'Size_Adj', 'Simulated_SQ_MD', 'Relevance_Weight']].copy()
                    
                    # Sort by highest relevance weight, then by closest size
                    disp_lab['Size_Diff'] = (disp_lab['Total Squares'] - total_squares).abs()
                    disp_lab = disp_lab.sort_values(by=['Relevance_Weight', 'Size_Diff'], ascending=[False, True]).drop(columns=['Size_Diff'])
                    
                    # Format as percentage strings or factors
                    disp_lab['CB_Adj'] = disp_lab['CB_Adj'].apply(lambda x: f"{x:.2f}x")
                    disp_lab['Mat_Adj'] = disp_lab['Mat_Adj'].apply(lambda x: f"{x:.2f}x")
                    disp_lab['Size_Adj'] = disp_lab['Size_Adj'].apply(lambda x: f"{x:.2f}x")
                    
                    st.dataframe(disp_lab.style.format({
                        'Total Squares': '{:,.1f}',
                        'Raw_SQ_MD': '{:,.2f}',
                        'Simulated_SQ_MD': '{:,.2f}',
                        'Relevance_Weight': '{:.1f}'
                    }))

# --- STEP 4: GENERAL COSTS ---
with st.container(border=True):
    st.markdown("<h2 class='step-header'>General Costs</h2>", unsafe_allow_html=True)
    st.markdown('<div class="hide-me-print"></div>', unsafe_allow_html=True)

    final_gen_cost = 0.0
    if not spec_type or spec_type == "Any":
        st.info("Please select a Spec Type to run the General Costs simulation.")
    else:
        # 1. Base Filter (Spec Type) for General Costs (using 30-Day rule)
        gen_mask = (df_all["Spec Type"] == spec_type)
        df_gen = df_all[gen_mask].copy()
        df_gen["Days_Diff"] = (df_gen["Report"] - df_gen["Complete"]).dt.days
        df_gen["Eligible"] = df_gen["Days_Diff"] >= 30
        val_gen = df_gen[df_gen["Eligible"]].copy()
    
        if val_gen.empty:
            st.warning("No historical jobs match these requirements.")
        else:
            # Pre-calculate Relevance Weights for the gen subset so we can do weighted averages
            def calc_gen_weight(sq):
                if sq <= 0: return 0.0
                if (total_squares * 0.5) <= sq <= (total_squares * 1.5): return 1.0
                return 0.5
            val_gen["Relevance_Weight"] = val_gen["Total Squares"].apply(calc_gen_weight)
            val_gen["Final_Weight"] = val_gen["Total Squares"] * val_gen["Relevance_Weight"]
            total_gen_weights = val_gen["Final_Weight"].sum()
        
            # UI Header
            st.markdown("---")
            h1, h2, h3, h4 = st.columns([3, 4, 2, 2], vertical_alignment="center")
            h1.markdown("**Cost Code**")
            h2.markdown("**Parameters**")
            h3.markdown("<div style='text-align: right;'><b>Total Cost</b></div>", unsafe_allow_html=True)
            h4.markdown("<div style='text-align: right;'><b>Unit Cost</b></div>", unsafe_allow_html=True)
        
            with st.container(border=True):
                st.markdown("#### System Calculated Allowances")
            
                # --- PERMIT ---
                permit_mask = (val_gen["Permit"] > 0) & (val_gen["Total Squares"] > 0)
                valid_permits = val_gen[permit_mask].copy()
                valid_permits["Permit_SQ"] = valid_permits["Permit"] / valid_permits["Total Squares"]
            
                regions = {
                    "SF / Peninsula": ['San Francisco', 'Daly City', 'Brisbane', 'South San Francisco', 'Burlingame', 'San Mateo', 'Foster City', 'Belmont', 'San Carlos', 'Redwood City', 'Menlo Park', 'Palo Alto'],
                    "South Bay": ['Mountain View', 'Sunnyvale', 'Santa Clara', 'San Jose', 'Campbell', 'Los Gatos', 'Milpitas'],
                    "East Bay": ['Oakland', 'Berkeley', 'Emeryville', 'Albany', 'Alameda', 'San Leandro', 'Hayward', 'Union City', 'Fremont', 'Newark', 'Richmond', 'Concord', 'Pleasant Hill', 'Walnut Creek', 'Lafayette', 'Danville', 'Dublin', 'Pleasanton', 'Livermore'],
                    "North Bay": ['San Rafael', 'Corte Madera', 'Santa Rosa', 'Willits'],
                    "Central Valley": ['Sacramento', 'Stockton', 'Tracy', 'Salida', 'Oakdale', 'Turlock', 'Bakersfield', 'Angels Camp']
                }
                db_cities = sorted([str(c).strip() for c in val_gen["City"].dropna().unique() if str(c).strip() != ""])
                permit_opts = ["All Cities", "Region: SF / Peninsula", "Region: South Bay", "Region: East Bay", "Region: North Bay", "Region: Central Valley"] + [f"City: {c}" for c in db_cities]
            
                p1, p2, p3, p4 = st.columns([3, 4, 2, 2], vertical_alignment="center")
                p1.markdown("**01-0100 - General - Permits**")
                p_sel = p2.selectbox("Permit Location", permit_opts, label_visibility="collapsed")
            
                p_cost_sq = 0.0
                if not valid_permits.empty:
                    if p_sel == "All Cities":
                        p_cost_sq = valid_permits["Permit_SQ"].median()
                    elif p_sel.startswith("Region: "):
                        r_name = p_sel.split(": ")[1]
                        r_cities = regions.get(r_name, [])
                        r_df = valid_permits[valid_permits["City"].isin(r_cities)]
                        p_cost_sq = r_df["Permit_SQ"].median() if not r_df.empty else valid_permits["Permit_SQ"].median()
                    else:
                        c_name = p_sel.split(": ")[1]
                        c_df = valid_permits[valid_permits["City"] == c_name]
                        p_cost_sq = c_df["Permit_SQ"].median() if len(c_df) >= 3 else valid_permits["Permit_SQ"].median()
                if np.isnan(p_cost_sq): p_cost_sq = 0.0
                tot_permit = p_cost_sq * total_squares
            
                p3.markdown(f"<div style='text-align: right; font-weight: bold;'>${tot_permit:,.2f}</div>", unsafe_allow_html=True)
                p4.markdown(f"<div style='text-align: right; color: #555;'>${p_cost_sq:,.2f} / SQ</div>", unsafe_allow_html=True)
            
                # --- DISPOSAL ---
                disp_mask = (val_gen["Disposal"] > 0) & (val_gen["Total Squares"] > 0)
                valid_disp = val_gen[disp_mask].copy()
                valid_disp["Disp_SQ"] = valid_disp["Disposal"] / valid_disp["Total Squares"]
            
                d1, d2, d3, d4 = st.columns([3, 4, 2, 2], vertical_alignment="center")
                d1.markdown("**01-0110 - General - Disposal**")
            
                tot_disp = 0.0
                d_cost_sq = 0.0
                if spec_type in ["Overlay", "Coating"]:
                    d2.markdown("*N/A (No Tear-Off)*")
                    d3.markdown(f"<div style='text-align: right; font-weight: bold;'>$0.00</div>", unsafe_allow_html=True)
                    d4.markdown(f"<div style='text-align: right; color: #555;'>$0.00 / SQ</div>", unsafe_allow_html=True)
                else:
                    d_sel = d2.selectbox("Disposal Location", permit_opts, label_visibility="collapsed")
                    if not valid_disp.empty:
                        if d_sel == "All Cities":
                            d_cost_sq = valid_disp["Disp_SQ"].median()
                        elif d_sel.startswith("Region: "):
                            r_name = d_sel.split(": ")[1]
                            r_cities = regions.get(r_name, [])
                            r_df = valid_disp[valid_disp["City"].isin(r_cities)]
                            d_cost_sq = r_df["Disp_SQ"].median() if not r_df.empty else valid_disp["Disp_SQ"].median()
                        else:
                            c_name = d_sel.split(": ")[1]
                            c_df = valid_disp[valid_disp["City"] == c_name]
                            d_cost_sq = c_df["Disp_SQ"].median() if len(c_df) >= 3 else valid_disp["Disp_SQ"].median()
                    if np.isnan(d_cost_sq): d_cost_sq = 0.0
                    tot_disp = d_cost_sq * total_squares
                    d3.markdown(f"<div style='text-align: right; font-weight: bold;'>${tot_disp:,.2f}</div>", unsafe_allow_html=True)
                    d4.markdown(f"<div style='text-align: right; color: #555;'>${d_cost_sq:,.2f} / SQ</div>", unsafe_allow_html=True)
                
                # --- STD EQUIPMENT ---
                eq_mask = (val_gen["Equipment"] > 0) & (val_gen["Total Squares"] > 0)
                valid_eq = val_gen[eq_mask].copy()
                eq_cost_sq = 0.0
                if not valid_eq.empty:
                    eq_cost_sq = (valid_eq["Equipment"] / valid_eq["Total Squares"]).median()
                if np.isnan(eq_cost_sq): eq_cost_sq = 0.0
                tot_equip = eq_cost_sq * total_squares
            
                e1, e2, e3, e4 = st.columns([3, 4, 2, 2], vertical_alignment="center")
                e1.markdown("**01-0120 - General - Std. Equip**")
                e2.markdown(f"*System Calculated ({len(valid_eq)} jobs)*")
                e3.markdown(f"<div style='text-align: right; font-weight: bold;'>${tot_equip:,.2f}</div>", unsafe_allow_html=True)
                e4.markdown(f"<div style='text-align: right; color: #555;'>${eq_cost_sq:,.2f} / SQ</div>", unsafe_allow_html=True)
            
                # --- FUEL ---
                fuel_mask = (val_gen["Fuel"] > 0) & (val_gen["Man Days"] > 0)
                valid_fuel = val_gen[fuel_mask].copy()
                f_cost_md = 0.0
                if not valid_fuel.empty:
                    valid_fuel["Fuel_MD"] = valid_fuel["Fuel"] / valid_fuel["Man Days"]
                    f_cost_md = valid_fuel["Fuel_MD"].median()
                if np.isnan(f_cost_md): f_cost_md = 0.0
                tot_fuel = f_cost_md * est_md
            
                f1, f2, f3, f4 = st.columns([3, 4, 2, 2], vertical_alignment="center")
                f1.markdown("**01-0160 - General - Fuel**")
                f2.markdown(f"*System Calculated ({len(valid_fuel)} jobs) * {est_md:,.2f} MDs*")
                f3.markdown(f"<div style='text-align: right; font-weight: bold;'>${tot_fuel:,.2f}</div>", unsafe_allow_html=True)
                f4.markdown(f"<div style='text-align: right; color: #555;'>${f_cost_md:,.2f} / MD</div>", unsafe_allow_html=True)
                
                # --- OUTSIDE SERVICES ---
                os_mask = (val_gen["Outside Services"] > 0) & (val_gen["Total Squares"] > 0)
                valid_os = val_gen[os_mask].copy()
                os_cost_sq = 0.0
                if not valid_os.empty:
                    os_cost_sq = (valid_os["Outside Services"] / valid_os["Total Squares"]).median()
                if np.isnan(os_cost_sq): os_cost_sq = 0.0
                tot_os = os_cost_sq * total_squares
            
                os1, os2, os3, os4 = st.columns([3, 4, 2, 2], vertical_alignment="center")
                os1.markdown("**01-0200 - General-Outside Services**")
                os2.markdown(f"*System Calculated ({len(valid_os)} jobs)*")
                os3.markdown(f"<div style='text-align: right; font-weight: bold;'>${tot_os:,.2f}</div>", unsafe_allow_html=True)
                os4.markdown(f"<div style='text-align: right; color: #555;'>${os_cost_sq:,.2f} / SQ</div>", unsafe_allow_html=True)

            st.markdown("<br>", unsafe_allow_html=True)
            with st.container(border=True):
                st.markdown("#### Estimator Allowances")
            
                # --- HEAVY EQUIPMENT ---
                he1, he2, he3, he4 = st.columns([3, 4, 2, 2], vertical_alignment="center")
                he1.markdown("**01-0120 - General - Heavy Equip**")
                he2.markdown("<div style='color: #888;'><em>Enter Lump Sum Allowance -></em></div>", unsafe_allow_html=True)
                tot_heavy = he3.number_input("Heavy Equip Allowance ($)", min_value=0.0, value=0.0, step=100.0, label_visibility="collapsed", key="he_num")
                he4.markdown(f"<div style='text-align: right; color: #555;'>-</div>", unsafe_allow_html=True)
            
                # --- PER DIEM ---
                pd1, pd2, pd3, pd4 = st.columns([3, 4, 2, 2], vertical_alignment="center")
                pd1.markdown("**01-0140 - General - Per Diem**")
                req_perdiem = pd2.toggle("Add Out of Town (Per Diem)?", key="tgl_pd")
                tot_perdiem = 0.0
                pd_rate = 0.0
                if req_perdiem:
                    pd_rate = pd2.number_input("Per Diem Rate ($/MD)", min_value=0.0, value=100.0, step=10.0, label_visibility="collapsed")
                    tot_perdiem = pd_rate * est_md
                pd3.markdown(f"<div style='text-align: right; font-weight: bold;'>${tot_perdiem:,.2f}</div>", unsafe_allow_html=True)
                if req_perdiem:
                    pd4.markdown(f"<div style='text-align: right; color: #555;'>${pd_rate:,.2f} / MD</div>", unsafe_allow_html=True)
                else:
                    pd4.markdown(f"<div style='text-align: right; color: #555;'>-</div>", unsafe_allow_html=True)
                
                # --- WARRANTY ---
                w1, w2, w3, w4 = st.columns([3, 4, 2, 2], vertical_alignment="center")
                w1.markdown("**01-0190 - General-Mfg Warranty**")
                req_warranty = w2.toggle("Require Manufacturer Warranty?", key="tgl_warr")
                tot_warr = 0.0
                warr_rate = 0.0
                if req_warranty:
                    warr_rate = w2.number_input("Warranty Rate ($/SQ)", min_value=0.0, value=10.0, step=1.0, label_visibility="collapsed")
                tot_warr = warr_rate * total_squares
                w3.markdown(f"<div style='text-align: right; font-weight: bold;'>${tot_warr:,.2f}</div>", unsafe_allow_html=True)
                if req_warranty:
                    w4.markdown(f"<div style='text-align: right; color: #555;'>${warr_rate:,.2f} / SQ</div>", unsafe_allow_html=True)
                else:
                    w4.markdown(f"<div style='text-align: right; color: #555;'>-</div>", unsafe_allow_html=True)
                
                # --- MISC ---
                m1, m2, m3, m4 = st.columns([3, 4, 2, 2], vertical_alignment="center")
                m1.markdown("**Misc Allowances**")
                m2.markdown("<div style='color: #888;'><em>Enter Lump Sum Allowance -></em></div>", unsafe_allow_html=True)
                tot_misc = m3.number_input("Misc Allowance ($)", min_value=0.0, value=0.0, step=100.0, label_visibility="collapsed", key="misc_num")
                m4.markdown(f"<div style='text-align: right; color: #555;'>-</div>", unsafe_allow_html=True)
            
        final_gen_cost = tot_permit + tot_disp + tot_fuel + tot_equip + tot_os + tot_heavy + tot_perdiem + tot_warr + tot_misc
        gen_cost_sq = final_gen_cost / total_squares if total_squares > 0 else 0.0
        
        st.markdown("<br>", unsafe_allow_html=True)
        fc1, fc2, fc3, fc4 = st.columns([3, 4, 2, 2], vertical_alignment="center")
        fc2.markdown("### Total General Costs")
        fc3.markdown(f"<div style='text-align: right;'><h3 style='color: #0F52BA; margin: 0;'>${final_gen_cost:,.2f}</h3></div>", unsafe_allow_html=True)
        fc4.markdown(f"<div style='text-align: right;'><h4 style='color: #555; margin: 0;'>${gen_cost_sq:,.2f} / SQ</h4></div>", unsafe_allow_html=True)

    # --- STEP 5: MARGIN & PROFIT ---
    with st.container(border=True):
        st.markdown("<h2 class='step-header'>Margin & Profit</h2>", unsafe_allow_html=True)
        st.markdown('<div class="hide-me-print"></div>', unsafe_allow_html=True)

        tot_mat_cost = weighted_avg_mat_cost * total_squares
        base_cost = tot_mat_cost + total_lab_cost + final_gen_cost
    base_cost_sq = base_cost / total_squares if total_squares > 0 else 0.0

    st.markdown(
        f"""
        <div style='text-align: center; margin-bottom: 2rem;'>
            <h3 style='margin-bottom: 0;'>Total Base Cost</h3>
            <h2 style='color: #555; margin-top: 0;'>${base_cost:,.2f}</h2>
            <p style='color: #888;'>(${base_cost_sq:,.2f} / SQ)</p>
        </div>
        """, unsafe_allow_html=True
    )

    with st.container(border=True):
        st.markdown("#### Markup Application")
        mc1, mc2 = st.columns(2, vertical_alignment="center")
    
        oh_pct = mc1.slider("Overhead Markup %", min_value=0.0, max_value=100.0, value=33.0, step=1.0)
        oh_amt = base_cost * (oh_pct / 100.0)
        subtotal_w_oh = base_cost + oh_amt
        mc1.markdown(f"<div style='text-align: right; color: #555;'>+ ${oh_amt:,.2f} (Overhead)</div>", unsafe_allow_html=True)
    
        prof_pct = mc2.slider("Profit Markup %", min_value=0.0, max_value=100.0, value=15.0, step=1.0)
    prof_amt = subtotal_w_oh * (prof_pct / 100.0)
    mc2.markdown(f"<div style='text-align: right; color: #555;'>+ ${prof_amt:,.2f} (Profit)</div>", unsafe_allow_html=True)

    final_sell = subtotal_w_oh + prof_amt
    final_sell_sq = final_sell / total_squares if total_squares > 0 else 0.0
    gross_profit = final_sell - base_cost
gp_pct = (gross_profit / final_sell * 100.0) if final_sell > 0 else 0.0

st.markdown("<br>", unsafe_allow_html=True)

st.markdown(
    f"""
    <div style='background-color: #f0f8ff; padding: 2rem; border-radius: 10px; border: 2px solid #0F52BA; text-align: center; margin-bottom: 2rem;'>
        <h2 style='color: #0F52BA; margin-bottom: 0;'>Final Sell Price</h2>
        <h1 style='color: #0F52BA; font-size: 3.5rem; margin-top: 0;'>${final_sell:,.2f}</h1>
        <h4 style='color: #555; margin-top: -10px;'>${final_sell_sq:,.2f} / SQ</h4>
        <hr style='border: 1px solid #cce5ff; margin: 1.5rem 0;'>
        <div style='display: flex; justify-content: space-around;'>
            <div>
                <p style='color: #555; font-weight: bold; margin-bottom: 0;'>Gross Profit ($)</p>
                <h2 style='color: #2e8b57; margin-top: 0;'>${gross_profit:,.2f}</h2>
            </div>
            <div>
                <p style='color: #555; font-weight: bold; margin-bottom: 0;'>Gross Profit (%)</p>
                <h2 style='color: #2e8b57; margin-top: 0;'>{gp_pct:,.1f}%</h2>
            </div>
        </div>
    </div>
    """, unsafe_allow_html=True
)



# --- STEP 6: Summary & Export ---
with st.container(border=True):
    st.markdown("<h2 class='step-header'>Summary & Export</h2>", unsafe_allow_html=True)
    
    # Remove the broken print CSS - replaced with proper PDF generation
    
    # 1. Job Profile Summary
    st.markdown("### Job Profile Summary")
    
    filters = []
    if spec_type and spec_type not in ["Any", "(Blank)"]:
        filters.append(f"**Spec:** {spec_type}")
    filters.append(f"**Size:** {total_squares:,.0f} SQ")
    
    if insulation and insulation not in ["Any", "(Blank)"]:
        att_str = f" ({ins_att})" if ins_att not in ["Any", "(Blank)"] else ""
        filters.append(f"**Insulation:** {insulation}{att_str}")
        
    if cb_type and cb_type not in ["Any", "(Blank)"]:
        thick_str = f" {cb_thick}" if cb_thick not in ["Any", "(Blank)"] else ""
        att_str = f" ({cb_att})" if cb_att not in ["Any", "(Blank)"] else ""
        filters.append(f"**Cover Board:** {cb_type}{thick_str}{att_str}")
        
    if mat_type and mat_type not in ["Any", "(Blank)"]:
        thick_str = f" {mat_thick}" if mat_thick not in ["Any", "(Blank)"] else ""
        att_str = f" ({mat_att})" if mat_att not in ["Any", "(Blank)"] else ""
        filters.append(f"**Roof System:** {mat_type}{thick_str}{att_str}")
        
    st.markdown(" | ".join(filters))
    
    st.markdown("<hr>", unsafe_allow_html=True)
    
    # 2. Executive Summary Metrics
    st.markdown("### Executive Summary")
    c1, c2, c3 = st.columns(3)
    
    # Calculate $/SQ for base cost
    base_cost_sq = base_cost / total_squares if total_squares > 0 else 0.0
    
    with c1:
        st.markdown(
            f'''
            <div class='premium-metric'>
                <h3 class='premium-title'>Total Cost</h3>
                <h1 class='premium-value' style='color: #888;'>${base_cost:,.2f}</h1>
                <p class='premium-sub'>${base_cost_sq:,.2f} / SQ</p>
            </div>
            ''', unsafe_allow_html=True
        )
    with c2:
        st.markdown(
            f'''
            <div class='premium-metric'>
                <h3 class='premium-title'>Total Gross Profit</h3>
                <h1 class='premium-value' style='color: #2e8b57;'>${gross_profit:,.2f}</h1>
                <p class='premium-sub'>{gp_pct:,.1f}% Margin</p>
            </div>
            ''', unsafe_allow_html=True
        )
    with c3:
        st.markdown(
            f'''
            <div class='premium-metric'>
                <h3 class='premium-title'>Final Sell Price</h3>
                <h1 class='premium-value' style='color: #0F52BA;'>${final_sell:,.2f}</h1>
                <p class='premium-sub'>${final_sell_sq:,.2f} / SQ</p>
            </div>
            ''', unsafe_allow_html=True
        )
        
    st.markdown(
        f'''
        <div style="display: flex; justify-content: center; gap: 4rem; margin-top: 1rem; color: #555; font-size: 1.1rem;">
            <div><strong>Labor Productivity:</strong> {proj_sq_md:,.2f} SQ / MD</div>
            <div><strong>Total Man Days:</strong> {est_md:,.1f}</div>
        </div>
        ''', unsafe_allow_html=True
    )
        
    st.markdown("<br>", unsafe_allow_html=True)
    
    # 3. Cost Breakdown Chart and Table
    col_chart, col_table = st.columns([1, 1])
    
    bd_data = pd.DataFrame({
        "Category": ["Materials", "Labor", "General Costs", "Overhead", "Profit"],
        "Amount": [tot_mat_cost, total_lab_cost, final_gen_cost, tot_overhead, tot_profit]
    })
    bd_data = bd_data[bd_data["Amount"] > 0]
    
    # Calculate $/SQ column for the table
    bd_data["$/SQ"] = np.where(total_squares > 0, bd_data["Amount"] / total_squares, 0.0)
    
    with col_chart:
        st.markdown("#### Cost Distribution")
        chart = alt.Chart(bd_data).mark_arc(innerRadius=50).encode(
            theta=alt.Theta(field="Amount", type="quantitative"),
            color=alt.Color(field="Category", type="nominal", scale=alt.Scale(scheme='blues')),
            tooltip=['Category', alt.Tooltip('Amount', format='$,.2f')]
        ).interactive()
        st.altair_chart(chart, use_container_width=True)
        
    with col_table:
        st.markdown("#### Financial Breakdown")
        disp_bd = bd_data.copy()
        disp_bd['Amount'] = disp_bd['Amount'].apply(lambda x: f"${x:,.2f}")
        disp_bd['$/SQ'] = disp_bd['$/SQ'].apply(lambda x: f"${x:,.2f}")
        disp_bd.loc[len(disp_bd)] = ["TOTAL COST", f"${base_cost:,.2f}", f"${base_cost_sq:,.2f}"]
        disp_bd.loc[len(disp_bd)] = ["TOTAL GROSS PROFIT", f"${gross_profit:,.2f}", f"${gross_profit / total_squares if total_squares > 0 else 0:,.2f}"]
        disp_bd.loc[len(disp_bd)] = ["TOTAL SELL PRICE", f"${final_sell:,.2f}", f"${final_sell_sq:,.2f}"]
        st.dataframe(disp_bd, hide_index=True, use_container_width=True)
        
    st.markdown("<hr>", unsafe_allow_html=True)
    
    # 4. Export Buttons
    st.markdown("### Export Estimate")
    
    # --- PDF Generation Function ---
    def generate_pdf():
        from fpdf import FPDF
        
        pdf = FPDF(orientation='P', unit='mm', format='Letter')
        pdf.set_auto_page_break(auto=True, margin=15)
        pdf.add_page()
        
        # --- Title ---
        pdf.set_font("Helvetica", "B", 20)
        pdf.set_text_color(15, 82, 186)  # #0F52BA
        pdf.cell(0, 12, "Roofing Estimate Summary", ln=True, align="C")
        pdf.ln(2)
        
        # Thin divider line
        pdf.set_draw_color(15, 82, 186)
        pdf.set_line_width(0.8)
        pdf.line(15, pdf.get_y(), 195, pdf.get_y())
        pdf.ln(6)
        
        # --- Job Profile ---
        pdf.set_font("Helvetica", "B", 12)
        pdf.set_text_color(51, 51, 51)
        pdf.cell(0, 8, "Job Profile", ln=True)
        pdf.ln(1)
        
        pdf.set_font("Helvetica", "", 10)
        pdf.set_text_color(80, 80, 80)
        
        profile_items = []
        if spec_type and spec_type not in ["Any", "(Blank)"]:
            profile_items.append(f"Spec: {spec_type}")
        profile_items.append(f"Size: {total_squares:,.0f} SQ")
        if insulation and insulation not in ["Any", "(Blank)"]:
            att_s = f" ({ins_att})" if ins_att not in ["Any", "(Blank)"] else ""
            profile_items.append(f"Insulation: {insulation}{att_s}")
        if cb_type and cb_type not in ["Any", "(Blank)"]:
            t_s = f" {cb_thick}" if cb_thick not in ["Any", "(Blank)"] else ""
            a_s = f" ({cb_att})" if cb_att not in ["Any", "(Blank)"] else ""
            profile_items.append(f"Cover Board: {cb_type}{t_s}{a_s}")
        if mat_type and mat_type not in ["Any", "(Blank)"]:
            t_s = f" {mat_thick}" if mat_thick not in ["Any", "(Blank)"] else ""
            a_s = f" ({mat_att})" if mat_att not in ["Any", "(Blank)"] else ""
            profile_items.append(f"Roof System: {mat_type}{t_s}{a_s}")
        
        pdf.cell(0, 6, "  |  ".join(profile_items), ln=True)
        pdf.ln(6)
        
        # --- Executive Summary KPIs ---
        pdf.set_font("Helvetica", "B", 12)
        pdf.set_text_color(51, 51, 51)
        pdf.cell(0, 8, "Executive Summary", ln=True)
        pdf.ln(2)
        
        # 3-column KPI boxes
        box_w = 56
        box_h = 28
        gap = 6
        x_start = 15
        y_start = pdf.get_y()
        
        kpi_data = [
            ("Total Cost", f"${base_cost:,.2f}", f"${base_cost_sq:,.2f} / SQ", (136, 136, 136)),
            ("Total Gross Profit", f"${gross_profit:,.2f}", f"{gp_pct:,.1f}% Margin", (46, 139, 87)),
            ("Final Sell Price", f"${final_sell:,.2f}", f"${final_sell_sq:,.2f} / SQ", (15, 82, 186)),
        ]
        
        for i, (title, value, sub, color) in enumerate(kpi_data):
            x = x_start + i * (box_w + gap)
            # Light grey background box
            pdf.set_fill_color(248, 249, 250)
            pdf.set_draw_color(220, 220, 220)
            pdf.rect(x, y_start, box_w, box_h, style="DF")
            
            # Title
            pdf.set_xy(x, y_start + 2)
            pdf.set_font("Helvetica", "", 8)
            pdf.set_text_color(100, 100, 100)
            pdf.cell(box_w, 5, title, align="C")
            
            # Value
            pdf.set_xy(x, y_start + 9)
            pdf.set_font("Helvetica", "B", 16)
            pdf.set_text_color(*color)
            pdf.cell(box_w, 8, value, align="C")
            
            # Sub
            pdf.set_xy(x, y_start + 20)
            pdf.set_font("Helvetica", "", 8)
            pdf.set_text_color(140, 140, 140)
            pdf.cell(box_w, 5, sub, align="C")
        
        pdf.set_y(y_start + box_h + 4)
        
        # Labor KPIs row
        pdf.set_font("Helvetica", "", 9)
        pdf.set_text_color(100, 100, 100)
        labor_line = f"Labor Productivity: {proj_sq_md:,.2f} SQ / MD     |     Total Man Days: {est_md:,.1f}"
        pdf.cell(0, 6, labor_line, ln=True, align="C")
        pdf.ln(8)
        
        # --- Financial Breakdown Table ---
        pdf.set_font("Helvetica", "B", 12)
        pdf.set_text_color(51, 51, 51)
        pdf.cell(0, 8, "Financial Breakdown", ln=True)
        pdf.ln(2)
        
        # Table header
        col_widths = [75, 50, 50]
        headers = ["Category", "Amount", "$/SQ"]
        
        pdf.set_fill_color(15, 82, 186)
        pdf.set_text_color(255, 255, 255)
        pdf.set_font("Helvetica", "B", 10)
        for j, h in enumerate(headers):
            pdf.cell(col_widths[j], 8, h, border=1, fill=True, align="C" if j > 0 else "L")
        pdf.ln()
        
        # Build table rows from bd_data + summary rows
        table_rows = []
        for _, row in bd_data.iterrows():
            table_rows.append((row["Category"], f"${row['Amount']:,.2f}", f"${row['$/SQ']:,.2f}", False))
        
        gp_sq = gross_profit / total_squares if total_squares > 0 else 0
        table_rows.append(("TOTAL COST", f"${base_cost:,.2f}", f"${base_cost_sq:,.2f}", True))
        table_rows.append(("TOTAL GROSS PROFIT", f"${gross_profit:,.2f}", f"${gp_sq:,.2f}", True))
        table_rows.append(("TOTAL SELL PRICE", f"${final_sell:,.2f}", f"${final_sell_sq:,.2f}", True))
        
        for cat, amt, sq, is_bold in table_rows:
            if is_bold:
                pdf.set_font("Helvetica", "B", 10)
                pdf.set_fill_color(240, 242, 246)
                pdf.set_text_color(30, 30, 30)
                fill = True
            else:
                pdf.set_font("Helvetica", "", 10)
                pdf.set_text_color(60, 60, 60)
                fill = False
            
            pdf.cell(col_widths[0], 7, f"  {cat}", border=1, fill=fill)
            pdf.cell(col_widths[1], 7, amt, border=1, align="R", fill=fill)
            pdf.cell(col_widths[2], 7, sq, border=1, align="R", fill=fill)
            pdf.ln()
        
        pdf.ln(10)
        
        # --- Footer ---
        pdf.set_font("Helvetica", "I", 8)
        pdf.set_text_color(160, 160, 160)
        from datetime import datetime
        pdf.cell(0, 5, f"Generated {datetime.now().strftime('%B %d, %Y at %I:%M %p')}  |  Historical Roofing Benchmark Dashboard", align="C")
        
        return bytes(pdf.output())
    
    col_btn1, col_btn2 = st.columns(2)
    with col_btn1:
        pdf_bytes = generate_pdf()
        st.download_button(
            label="📄 Download PDF",
            data=pdf_bytes,
            file_name="Estimate_Summary.pdf",
            mime="application/pdf",
            type="primary"
        )
    with col_btn2:
        output = io.BytesIO()
        writer = pd.ExcelWriter(output, engine='openpyxl')
        export_df = bd_data.copy()
        export_df.loc[len(export_df)] = ["TOTAL COST", base_cost, base_cost_sq]
        export_df.loc[len(export_df)] = ["TOTAL GROSS PROFIT", gross_profit, (gross_profit / total_squares if total_squares > 0 else 0)]
        export_df.loc[len(export_df)] = ["TOTAL SELL PRICE", final_sell, final_sell_sq]
        export_df.to_excel(writer, sheet_name='Summary', index=False)
        worksheet = writer.sheets['Summary']
        worksheet.column_dimensions['A'].width = 25
        worksheet.column_dimensions['B'].width = 15
        worksheet.column_dimensions['C'].width = 15
        for row in worksheet.iter_rows(min_row=2, max_row=worksheet.max_row, min_col=2, max_col=3):
            for cell in row:
                cell.number_format = '$#,##0.00'
        writer.close()
        processed_data = output.getvalue()
        
        st.download_button(
            label="📊 Download Excel (.xlsx)",
            data=processed_data,
            file_name="Estimate_Summary.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary"
        )

