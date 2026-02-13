"""
GSC Attainment Report Automator - Streamlit App
Step-by-step wizard for generating manager reports and Outlook email drafts.
"""

import os
import re
import tkinter as tk
from tkinter import filedialog
import pandas as pd
import streamlit as st

from generate_manager_reports import (
    extract_manager_name, extract_manager_id, sanitize_filename,
    generate_all_reports, get_all_regions, get_fiscal_year, REPORT_COLUMNS,
)
from create_email_drafts import (
    load_email_mapping, clean_display_name,
    create_drafts_batch, get_drafts_from_folder, send_drafts_batch,
)

# ── Page config ───────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="GSC Attainment Report Automator",
    page_icon="📊",
    layout="wide",
)

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


# ── Helper functions ──────────────────────────────────────────────────────────

def select_folder_dialog():
    """Open a native Windows folder picker dialog. Returns path or empty string."""
    root = tk.Tk()
    root.withdraw()
    root.wm_attributes("-topmost", 1)
    folder = filedialog.askdirectory(
        title="Select Output Folder",
        mustexist=False,
    )
    root.destroy()
    return folder


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


def build_manager_list(attainment_df, email_map, output_dir):
    """
    Build list of manager info from generated report files.
    Returns list of dicts with keys: name, region, email, filepath, safe_name
    """
    managers = []
    for root, _, filenames in os.walk(output_dir):
        for fname in filenames:
            if fname.startswith("FY26_Attainment_") and fname.endswith(".xlsx"):
                filepath = os.path.join(root, fname)
                region = os.path.basename(root)

                # Extract manager name from filename
                parts = fname[len("FY26_Attainment_"):-5]  # remove prefix and .xlsx
                last_underscore = parts.rfind("_")
                safe_name = parts[:last_underscore] if last_underscore > 0 else parts

                # Look up full name and email via attainment data
                mgr_full_name = None
                mgr_email = None
                for mgr in attainment_df["Level_1_Manager"].dropna().unique():
                    clean = extract_manager_name(mgr)
                    if sanitize_filename(clean) == safe_name:
                        mgr_full_name = mgr
                        emp_id = extract_manager_id(mgr)
                        if emp_id:
                            clean_id = emp_id.lstrip("0") or "0"
                            mgr_email = email_map.get(clean_id)
                        break

                display_name = clean_display_name(mgr_full_name) if mgr_full_name else safe_name

                managers.append({
                    "name": display_name,
                    "region": region,
                    "email": mgr_email,
                    "filepath": filepath,
                    "safe_name": safe_name,
                })

    managers.sort(key=lambda m: (m["region"], m["name"]))
    return managers


# ── Initialize session state ──────────────────────────────────────────────────
for key, default in [
    ("attainment_df", None),
    ("email_map", None),
    ("output_dir", ""),
    ("reports_generated", False),
    ("report_results", None),
    ("manager_list", None),
    ("available_regions", None),
    ("fiscal_year", None),
]:
    if key not in st.session_state:
        st.session_state[key] = default


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
        key="attainment_uploader",
    )

with col2:
    st.subheader("Sales Compensation Report")
    sales_comp_file = st.file_uploader(
        "Upload the Sales Compensation Report Excel file",
        type=["xlsx"],
        key="sales_comp_uploader",
    )

# Validate attainment file
if attainment_file is not None and st.session_state.attainment_df is None:
    with st.spinner("Validating Attainment file..."):
        df, err = validate_attainment_file(attainment_file)
        if err:
            st.error(f"Attainment file error: {err}")
        else:
            st.session_state.attainment_df = df
            st.session_state.available_regions = get_all_regions(df)
            st.session_state.fiscal_year = get_fiscal_year(df)

if st.session_state.attainment_df is not None:
    df = st.session_state.attainment_df
    n_managers = df["Level_1_Manager"].dropna().nunique()
    fiscal_year = st.session_state.fiscal_year or "FY26"
    st.success(
        f"Attainment file loaded: **{len(df):,}** rows, **{n_managers}** managers, "
        f"**{fiscal_year}**"
    )

# Validate sales comp file
if sales_comp_file is not None and st.session_state.email_map is None:
    with st.spinner("Validating Sales Compensation file..."):
        email_map, err = validate_sales_comp_file(sales_comp_file)
        if err:
            st.error(f"Sales Comp file error: {err}")
        else:
            st.session_state.email_map = email_map

if st.session_state.email_map is not None:
    st.success(f"Sales Comp file loaded: **{len(st.session_state.email_map):,}** email records")


# ══════════════════════════════════════════════════════════════════════════════
# STEP 2: Output Folder
# ══════════════════════════════════════════════════════════════════════════════
files_ready = (
    st.session_state.attainment_df is not None
    and st.session_state.email_map is not None
)

if files_ready:
    st.divider()
    st.header("Step 2 — Choose Output Folder")

    st.caption(
        'Select a parent folder. Reports will be saved under a '
        '"Manager report" subfolder with Region subfolders inside.'
    )

    col_input, col_btn = st.columns([5, 1])

    with col_input:
        default_dir = r"C:\Attainment Reports"
        output_dir = st.text_input(
            "Parent folder path:",
            value=st.session_state.output_dir or default_dir,
            label_visibility="collapsed",
            placeholder="Enter parent folder path or use Browse button...",
        )

    with col_btn:
        if st.button("Browse...", use_container_width=True):
            selected = select_folder_dialog()
            if selected:
                output_dir = selected
                st.session_state.output_dir = selected
                st.rerun()

    st.session_state.output_dir = output_dir

    if output_dir:
        report_dir = os.path.join(output_dir, "Manager report")
        if os.path.isdir(report_dir):
            st.info(f"Output: `{report_dir}` (already exists)")
        elif os.path.isdir(output_dir):
            st.info(f"Output: `{report_dir}` (will be created)")
        else:
            st.warning(f"Output: `{report_dir}` (folders will be created automatically)")


# ══════════════════════════════════════════════════════════════════════════════
# STEP 3: Generate Reports (with Region filter)
# ══════════════════════════════════════════════════════════════════════════════
if files_ready and st.session_state.output_dir:
    st.divider()
    st.header("Step 3 — Generate Manager Reports")

    # Actual output directory: parent / Manager report
    actual_output_dir = os.path.join(st.session_state.output_dir, "Manager report")

    if st.session_state.reports_generated:
        results = st.session_state.report_results
        st.success(
            f"Reports generated: **{results['total']}** managers across "
            f"**{len(results['region_counts'])}** regions"
        )
        with st.expander("Region distribution"):
            for region, count in sorted(results["region_counts"].items()):
                st.write(f"**{region}**: {count} reports")
        st.caption(f"Saved to: `{actual_output_dir}`")
    else:
        # Region selection before generation
        available_regions = st.session_state.available_regions or []

        st.write("Select regions to generate reports for:")
        gen_regions = st.multiselect(
            "Regions to generate:",
            options=available_regions,
            default=available_regions,
            key="gen_region_filter",
        )

        if gen_regions:
            st.caption(
                f"Reports will be saved to `{actual_output_dir}` "
                f"with Region subfolders."
            )

            if st.button("Generate Reports", type="primary"):
                progress_bar = st.progress(0)
                status_text = st.empty()

                def on_progress(current, total, message):
                    progress_bar.progress(current / total)
                    status_text.text(f"[{current}/{total}] {message}")

                with st.spinner("Generating reports..."):
                    results = generate_all_reports(
                        st.session_state.attainment_df,
                        actual_output_dir,
                        progress_callback=on_progress,
                        selected_regions=gen_regions,
                        fiscal_year=st.session_state.fiscal_year,
                    )

                st.session_state.reports_generated = True
                st.session_state.report_results = results

                # Build manager list for email step
                st.session_state.manager_list = build_manager_list(
                    st.session_state.attainment_df,
                    st.session_state.email_map,
                    actual_output_dir,
                )

                progress_bar.progress(1.0)
                status_text.text("Complete!")
                st.rerun()
        else:
            st.warning("Select at least one region to generate reports.")


# ══════════════════════════════════════════════════════════════════════════════
# STEP 4 & 5: Email Draft Selection and Generation
# ══════════════════════════════════════════════════════════════════════════════
if st.session_state.reports_generated and st.session_state.manager_list:
    st.divider()
    st.header("Step 4 — Select Email Recipients")

    manager_list = st.session_state.manager_list
    all_regions = sorted(set(m["region"] for m in manager_list))

    # Region filter
    selected_regions = st.multiselect(
        "Filter by Region:",
        options=all_regions,
        default=all_regions,
        key="email_region_filter",
    )

    # Filter managers by selected regions
    filtered_managers = [m for m in manager_list if m["region"] in selected_regions]

    # Manager filter
    manager_options = [
        f"{m['name']} ({m['region']})" for m in filtered_managers
    ]

    selected_manager_labels = st.multiselect(
        f"Select Managers ({len(filtered_managers)} available):",
        options=manager_options,
        default=manager_options,
        key="manager_filter",
    )

    # Resolve selected managers back to manager dicts
    selected_managers = [
        m for m, label in zip(filtered_managers, manager_options)
        if label in selected_manager_labels
    ]

    # Matching summary
    with_email = [m for m in selected_managers if m["email"]]
    without_email = [m for m in selected_managers if not m["email"]]

    col1, col2, col3 = st.columns(3)
    col1.metric("Selected", len(selected_managers))
    col2.metric("Email Matched", len(with_email))
    col3.metric("No Email", len(without_email))

    if without_email:
        with st.expander(f"Managers without email ({len(without_email)})"):
            for m in without_email:
                st.write(f"- {m['name']} ({m['region']})")

    # ── Step 5: Generate Email Drafts ──
    st.divider()
    st.header("Step 5 — Generate Outlook Email Drafts")

    if not with_email:
        st.warning("No managers with matching email addresses selected.")
    else:
        st.write(
            f"**{len(with_email)}** email drafts will be created in your "
            f"Outlook **Drafts / Manager Report** folder."
        )

        if st.button("Generate Email Drafts", type="primary"):
            progress_bar = st.progress(0)
            status_text = st.empty()

            def on_email_progress(current, total, message):
                progress_bar.progress(current / total)
                status_text.text(f"[{current}/{total}] {message}")

            matched_list = [
                (m["filepath"], m["name"], m["email"]) for m in with_email
            ]

            try:
                results = create_drafts_batch(
                    matched_list,
                    target_folder_name="Manager Report",
                    progress_callback=on_email_progress,
                    fiscal_year=st.session_state.fiscal_year or "FY26",
                )

                progress_bar.progress(1.0)
                status_text.text("Complete!")

                st.success(
                    f"Done! **{results['created']}** drafts created, "
                    f"**{results['failed']}** failed."
                )

                if results["failures_detail"]:
                    with st.expander("Failed drafts"):
                        for name, email, err in results["failures_detail"]:
                            st.write(f"- **{name}** ({email}): {err}")

                st.info("Check your Outlook > Drafts > Manager Report folder.")

            except ImportError:
                st.error("pywin32 is not installed. Run: `pip install pywin32`")
            except Exception as e:
                st.error(f"Could not connect to Outlook: {e}")


# ══════════════════════════════════════════════════════════════════════════════
# STEP 6: Send Email Drafts
# ══════════════════════════════════════════════════════════════════════════════
if st.session_state.reports_generated:
    st.divider()
    st.header("Step 6 — Send Email Drafts")

    st.caption(
        "⚠️ Warning: Sending emails cannot be undone. "
        "Make sure to review all drafts in Outlook before sending."
    )

    col_refresh, col_spacer = st.columns([1, 3])

    with col_refresh:
        if st.button("🔄 Load Drafts from Outlook"):
            try:
                import win32com.client
                outlook = win32com.client.Dispatch("Outlook.Application")
                folder, draft_items = get_drafts_from_folder(outlook, "Manager Report")

                if folder is None:
                    st.warning("No 'Manager Report' folder found in Outlook Drafts.")
                    st.session_state["draft_items"] = []
                else:
                    st.session_state["draft_items"] = draft_items
                    st.session_state["outlook"] = outlook
                    st.session_state["draft_folder"] = folder
                    st.success(f"Loaded {len(draft_items)} drafts from Outlook.")

            except ImportError:
                st.error("pywin32 is not installed. Run: `pip install pywin32`")
            except Exception as e:
                st.error(f"Could not connect to Outlook: {e}")

    if "draft_items" in st.session_state and st.session_state["draft_items"]:
        draft_items = st.session_state["draft_items"]

        st.write(f"**{len(draft_items)}** draft emails found in Outlook Drafts/Manager Report folder.")

        # Create selection options
        draft_options = [
            f"{d['subject']} → {d['to']}" for d in draft_items
        ]

        # Select All / Deselect All buttons
        col_select_all, col_deselect_all, col_spacer2 = st.columns([1, 1, 2])

        with col_select_all:
            if st.button("✅ Select All", use_container_width=True):
                st.session_state["send_draft_filter_default"] = draft_options
                st.rerun()

        with col_deselect_all:
            if st.button("❌ Deselect All", use_container_width=True):
                st.session_state["send_draft_filter_default"] = []
                st.rerun()

        # Get default value from session state or empty list
        default_selection = st.session_state.get("send_draft_filter_default", [])

        selected_labels = st.multiselect(
            f"Select drafts to send ({len(draft_items)} available):",
            options=draft_options,
            default=default_selection,
            key="send_draft_filter",
        )

        selected_count = len(selected_labels)

        if selected_count > 0:
            st.warning(
                f"⚠️ You are about to send **{selected_count}** email(s). "
                f"This action cannot be undone."
            )

            # Map selected labels back to indices
            selected_indices = [
                draft_items[i]["index"]
                for i, label in enumerate(draft_options)
                if label in selected_labels
            ]

            if st.button(f"✉️ Send {selected_count} Selected Email(s)", type="primary"):
                progress_bar = st.progress(0)
                status_text = st.empty()

                def on_send_progress(current, total, subject):
                    progress_bar.progress(current / total)
                    status_text.text(f"[{current}/{total}] {subject}")

                try:
                    results = send_drafts_batch(
                        st.session_state["outlook"],
                        st.session_state["draft_folder"],
                        selected_indices,
                        progress_callback=on_send_progress,
                    )

                    progress_bar.progress(1.0)
                    status_text.text("Complete!")

                    if results["sent"] > 0:
                        st.success(
                            f"✅ Successfully sent **{results['sent']}** email(s)!"
                        )

                    if results["failed"] > 0:
                        st.error(f"❌ Failed to send **{results['failed']}** email(s).")

                        if results["failures_detail"]:
                            with st.expander("Failed emails"):
                                for subject, err in results["failures_detail"]:
                                    st.write(f"- **{subject}**: {err}")

                    # Clear draft items to force reload
                    st.session_state["draft_items"] = []
                    st.info("Click 'Load Drafts from Outlook' to refresh the list.")

                except Exception as e:
                    st.error(f"Error sending emails: {e}")
        else:
            st.info("Select at least one draft to send.")
    elif "draft_items" in st.session_state:
        st.info("No drafts found. Click 'Load Drafts from Outlook' to check again.")
