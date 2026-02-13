"""
GSC Attainment Report Automator - Streamlit App
Step-by-step wizard for generating manager reports and Outlook email drafts.
"""

import os
import re
import pandas as pd
import streamlit as st

from generate_manager_reports import (
    extract_manager_name, extract_manager_id, sanitize_filename,
    generate_all_reports, get_all_regions, get_fiscal_year, REPORT_COLUMNS,
)
from create_email_drafts import (
    load_email_mapping, clean_display_name,
)

# ── Page config ───────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="GSC Attainment Report Automator",
    page_icon="📊",
    layout="wide",
)

# Additional imports for .zip download
import zipfile
import io
import json
import tempfile
from datetime import datetime

# ── Custom CSS for Microsoft-style look ──────────────────────────────────────
st.markdown("""
<style>
    /* Main title - Microsoft Fluent Design */
    h1 {
        background-color: #0078D4 !important;
        color: white !important;
        font-size: 1.75rem !important;
        font-weight: 600 !important;
        padding: 1.25rem 1.5rem !important;
        border-radius: 4px !important;
        margin-bottom: 1.5rem !important;
        box-shadow: 0 1.6px 3.6px 0 rgba(0,0,0,.132), 0 0.3px 0.9px 0 rgba(0,0,0,.108);
        letter-spacing: 0.3px;
    }

    /* Step headers - Microsoft style with dark mode support */
    h2 {
        font-size: 1rem !important;
        font-weight: 600 !important;
        color: var(--text-color) !important;
        border-left: 3px solid #0078D4;
        padding-left: 12px;
        margin-top: 0.75rem !important;
        margin-bottom: 0.75rem !important;
        letter-spacing: 0.2px;
    }

    /* Subheaders with dark mode support */
    h3 {
        font-size: 0.875rem !important;
        font-weight: 600 !important;
        color: var(--text-color) !important;
    }

    /* Dark mode text colors */
    @media (prefers-color-scheme: dark) {
        h2, h3 {
            color: #FFFFFF !important;
        }
    }

    /* Light mode text colors (Streamlit default) */
    [data-theme="light"] h2,
    [data-theme="light"] h3 {
        color: #201F1E !important;
    }

    /* Dark mode text colors (Streamlit dark theme) */
    [data-theme="dark"] h2,
    [data-theme="dark"] h3 {
        color: #FFFFFF !important;
    }

    /* Multiselect tag colors - Microsoft blue */
    span[data-baseweb="tag"] {
        background-color: #DEECF9 !important;
        color: #0078D4 !important;
        border: 1px solid #0078D4 !important;
    }
    span[data-baseweb="tag"] span[role="presentation"] {
        color: #0078D4 !important;
    }

    /* Primary buttons - Microsoft blue */
    .stButton > button[kind="primary"] {
        background-color: #0078D4 !important;
        border: none !important;
        border-radius: 2px !important;
        padding: 0.5rem 1.25rem !important;
        font-weight: 600 !important;
        transition: background-color 0.1s ease !important;
    }
    .stButton > button[kind="primary"]:hover {
        background-color: #106EBE !important;
        box-shadow: 0 3.2px 7.2px 0 rgba(0,0,0,.132), 0 0.6px 1.8px 0 rgba(0,0,0,.108) !important;
    }

    /* Secondary buttons with dark mode support */
    .stButton > button:not([kind="primary"]) {
        border-radius: 2px !important;
        border: 1px solid #8A8886 !important;
        transition: background-color 0.1s ease !important;
    }

    /* Light mode secondary buttons */
    [data-theme="light"] .stButton > button:not([kind="primary"]) {
        background-color: white !important;
        color: #201F1E !important;
    }
    [data-theme="light"] .stButton > button:not([kind="primary"]):hover {
        background-color: #F3F2F1 !important;
        border-color: #323130 !important;
    }

    /* Dark mode secondary buttons */
    [data-theme="dark"] .stButton > button:not([kind="primary"]) {
        background-color: #2D2D2D !important;
        color: #FFFFFF !important;
        border-color: #5A5A5A !important;
    }
    [data-theme="dark"] .stButton > button:not([kind="primary"]):hover {
        background-color: #3D3D3D !important;
        border-color: #707070 !important;
    }

    /* File uploader area */
    [data-testid="stFileUploader"] {
        border-radius: 2px;
        border: 1px solid #EDEBE9;
    }

    /* Metrics - Microsoft card style */
    [data-testid="stMetric"] {
        background: white;
        border: 1px solid #EDEBE9;
        border-radius: 2px;
        padding: 12px 16px;
        box-shadow: 0 1.6px 3.6px 0 rgba(0,0,0,.132), 0 0.3px 0.9px 0 rgba(0,0,0,.108);
    }

    /* Dividers */
    hr {
        border-color: #EDEBE9 !important;
        margin-top: 1.5rem !important;
        margin-bottom: 1rem !important;
    }

    /* Expander */
    [data-testid="stExpander"] {
        border: 1px solid #EDEBE9;
        border-radius: 2px;
    }

    /* Progress bar - Microsoft blue */
    .stProgress > div > div {
        background-color: #0078D4 !important;
    }

    /* Success/Info/Warning boxes - Microsoft style */
    .stAlert {
        border-radius: 2px;
        border-left: 3px solid;
    }

    /* Text inputs and selects */
    input, select, textarea {
        border-radius: 2px !important;
    }

    /* General app background */
    [data-theme="light"] .main {
        background-color: #FAF9F8;
    }

    /* Dark mode adjustments */
    [data-theme="dark"] .main {
        background-color: #1E1E1E;
    }

    /* Captions and small text for dark mode */
    [data-theme="dark"] .stCaption,
    [data-theme="dark"] [data-testid="stCaption"] {
        color: #B3B3B3 !important;
    }
</style>
""", unsafe_allow_html=True)

st.title("GSC Attainment Report Automator")

# ── Reset button (top-right) ─────────────────────────────────────────────────
_, reset_col = st.columns([8, 1])
with reset_col:
    if st.button("🔄 Reset", use_container_width=True):
        # Increment reset counter to force new file_uploader widget keys
        reset_count = st.session_state.get("reset_count", 0) + 1
        for key in list(st.session_state.keys()):
            del st.session_state[key]
        st.session_state["reset_count"] = reset_count
        st.rerun()


# ── Helper functions ──────────────────────────────────────────────────────────

def validate_attainment_file(uploaded_file):
    """Validate and load the attainment Excel file. Returns (df, error_msg)."""
    try:
        df = pd.read_excel(uploaded_file, sheet_name="in")
    except Exception as e:
        return None, f"Cannot read sheet 'in': {e}"

    if "Level_1_Manager" not in df.columns:
        return None, "Missing required column: Level_1_Manager"

    # Rename Plan_Period column if present
    if "Plan_Period;MBO_Description" in df.columns:
        df.rename(columns={"Plan_Period;MBO_Description": "Plan_Period"}, inplace=True)

    return df, None


def validate_sales_comp_file(uploaded_file):
    """Validate and load the Sales Compensation Report. Returns (email_map, error_msg)."""
    try:
        df = pd.read_excel(uploaded_file, sheet_name="Sheet1", header=3)
    except Exception as e:
        return None, f"Cannot read sheet 'Sheet1' (header=3): {e}"

    if "Employee ID" not in df.columns or "Email - Work" not in df.columns:
        return None, "Missing required columns: 'Employee ID' and/or 'Email - Work'"

    email_map = {}
    for _, row in df.iterrows():
        emp_id = row.get("Employee ID")
        email = row.get("Email - Work")
        if pd.notna(emp_id) and pd.notna(email):
            clean_id = str(emp_id).lstrip("0") or "0"
            email_map[clean_id] = str(email).strip()

    return email_map, None




# ── Initialize session state ──────────────────────────────────────────────────
for key, default in [
    ("attainment_df", None),
    ("email_map", None),
    ("reports_generated", False),
    ("report_results", None),
    ("available_regions", None),
    ("fiscal_year", None),
    ("temp_dir", None),
    ("reset_count", 0),
]:
    if key not in st.session_state:
        st.session_state[key] = default

rc = st.session_state.reset_count  # used in widget keys to force re-creation


# ══════════════════════════════════════════════════════════════════════════════
# STEP 1: File Upload
# ══════════════════════════════════════════════════════════════════════════════
st.header("Step 1 — Upload Data Files")

col1, col2 = st.columns(2)

with col1:
    st.subheader("Global Attainment Report")
    attainment_file = st.file_uploader(
        "Upload the Global Attainment Club Excel file",
        type=["xlsx"],
        key=f"attainment_uploader_{rc}",
    )

with col2:
    st.subheader("Sales Compensation Report")
    sales_comp_file = st.file_uploader(
        "Upload the Sales Compensation Report Excel file",
        type=["xlsx"],
        key=f"sales_comp_uploader_{rc}",
    )

# Track if this is a new file upload (reset when file is uploaded)
if attainment_file is not None:
    # Check if this is a new upload by comparing file name or forcing revalidation
    if st.session_state.attainment_df is None:
        with st.spinner("Validating Attainment file..."):
            df, err = validate_attainment_file(attainment_file)
            if err:
                st.error(f"Attainment file error: {err}")
            else:
                st.session_state.attainment_df = df
                st.session_state.available_regions = get_all_regions(df)
                st.session_state.fiscal_year = get_fiscal_year(df)
                # Reset report generation state on new file
                st.session_state.reports_generated = False
                st.session_state.report_results = None

if st.session_state.attainment_df is not None:
    df = st.session_state.attainment_df
    n_managers = df["Level_1_Manager"].dropna().nunique()
    fiscal_year = st.session_state.fiscal_year or "FY26"
    n_regions = len(st.session_state.available_regions) if st.session_state.available_regions else 0
    st.success(
        f"Attainment file loaded: **{len(df):,}** rows, **{n_managers}** managers, "
        f"**{n_regions}** regions, **{fiscal_year}**"
    )

# Validate sales comp file
if sales_comp_file is not None:
    if st.session_state.email_map is None:
        with st.spinner("Validating Sales Compensation file..."):
            email_map, err = validate_sales_comp_file(sales_comp_file)
            if err:
                st.error(f"Sales Comp file error: {err}")
            else:
                st.session_state.email_map = email_map
                # Reset report generation state on new file
                st.session_state.reports_generated = False
                st.session_state.report_results = None

if st.session_state.email_map is not None:
    st.success(f"Sales Comp file loaded: **{len(st.session_state.email_map):,}** email records")


# ══════════════════════════════════════════════════════════════════════════════
# STEP 2: Generate Reports (with Region filter)
# ══════════════════════════════════════════════════════════════════════════════
files_ready = (
    st.session_state.attainment_df is not None
    and st.session_state.email_map is not None
)

if files_ready:
    st.divider()
    st.header("Step 2 — Generate Manager Reports")

    if st.session_state.reports_generated:
        results = st.session_state.report_results
        st.success(
            f"Reports generated: **{results['total']}** managers across "
            f"**{len(results['region_counts'])}** regions"
        )
        with st.expander("Region distribution"):
            for region, count in sorted(results["region_counts"].items()):
                st.write(f"**{region}**: {count} reports")
    else:
        # Region selection before generation
        available_regions = st.session_state.available_regions

        if not available_regions:
            st.error("""
⚠️ **No regions found in the attainment data.**

Please ensure:
- The Excel file has a 'Region' column
- The 'Level_1_Manager' column contains manager names
- At least one manager has a valid Region value
            """)
            st.stop()

        st.write(f"**{len(available_regions)}** regions available. Select which regions to generate reports for:")
        gen_regions = st.multiselect(
            "Regions to generate:",
            options=available_regions,
            default=available_regions,
            key="gen_region_filter",
        )

        if gen_regions:
            if st.button("Generate Reports", type="primary"):
                # Create temporary directory for report generation
                temp_dir = tempfile.mkdtemp()
                st.session_state.temp_dir = temp_dir

                progress_bar = st.progress(0)
                status_text = st.empty()

                def on_progress(current, total, message):
                    progress_bar.progress(current / total)
                    status_text.text(f"[{current}/{total}] {message}")

                with st.spinner("Generating reports..."):
                    results = generate_all_reports(
                        st.session_state.attainment_df,
                        temp_dir,
                        progress_callback=on_progress,
                        selected_regions=gen_regions,
                        fiscal_year=st.session_state.fiscal_year,
                    )

                st.session_state.reports_generated = True
                st.session_state.report_results = results

                progress_bar.progress(1.0)
                status_text.text("Complete!")
                st.rerun()
        else:
            st.warning("Select at least one region to generate reports.")

    # ── Download Section (After Reports Generated) ──
    if st.session_state.reports_generated:
        st.divider()
        st.subheader("📦 Download Reports")

        results = st.session_state.report_results
        fiscal_year = st.session_state.fiscal_year or "FY26"
        email_map = st.session_state.email_map

        # Create .zip file in memory
        zip_buffer = io.BytesIO()
        with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zip_file:
            # Add all Excel files
            for mgr_info in results["managers"]:
                # mgr_info: (full_name, region, safe_name, filepath)
                filepath = mgr_info[3]
                region = mgr_info[1]
                arcname = os.path.join(region, os.path.basename(filepath))
                zip_file.write(filepath, arcname)

            # Build metadata JSON
            metadata = {
                "fiscal_year": fiscal_year,
                "generated_date": datetime.today().strftime("%Y-%m-%d"),
                "total_reports": results["total"],
                "managers": [
                    {
                        "name": clean_display_name(m[0]),
                        "safe_name": m[2],
                        "region": m[1],
                        "email": email_map.get(extract_manager_id(m[0])),
                        "filepath": os.path.join(m[1], os.path.basename(m[3]))
                    }
                    for m in results["managers"]
                ]
            }
            zip_file.writestr("manager_metadata.json", json.dumps(metadata, indent=2))

        # Download buttons
        col1, col2 = st.columns(2)

        with col1:
            st.download_button(
                label="📦 Download Reports (.zip)",
                data=zip_buffer.getvalue(),
                file_name=f"Manager_Reports_{fiscal_year}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.zip",
                mime="application/zip",
                type="primary",
                use_container_width=True
            )

        with col2:
            # Check if EmailManager.exe exists
            exe_path = os.path.join(os.path.dirname(__file__), "dist", "EmailManager.exe")
            if os.path.exists(exe_path):
                # Wrap .exe in .zip to avoid browser download blocking
                exe_zip_buffer = io.BytesIO()
                with zipfile.ZipFile(exe_zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zf:
                    zf.write(exe_path, "EmailManager.exe")
                st.download_button(
                    label="📧 Download Email Manager (.zip)",
                    data=exe_zip_buffer.getvalue(),
                    file_name="EmailManager.zip",
                    mime="application/zip",
                    use_container_width=True
                )
            else:
                st.button(
                    "📧 Email Manager (not built)",
                    disabled=True,
                    use_container_width=True,
                    help="Run 'pyinstaller email_manager.spec' to build EmailManager.exe"
                )

        st.info("""
**✅ Reports generated successfully!**

**📥 Next Steps:**
1. **Download Reports (.zip)**: Contains all manager reports + metadata
2. **Download Email Manager**: Windows app for sending emails via Outlook
3. Run EmailManager.exe → Load .zip → Send emails
        """)
