# GSC Attainment Report Automator — Development Reference

## Table of Contents

- [Project Overview](#project-overview)
- [Architecture](#architecture)
- [File Structure](#file-structure)
- [Data Flow](#data-flow)
- [Module Reference](#module-reference)
  - [app.py — Streamlit UI](#apppy--streamlit-ui)
  - [generate_manager_reports.py — Report Generator](#generate_manager_reportspy--report-generator)
  - [create_email_drafts.py — Outlook Email Draft Creator](#create_email_draftspy--outlook-email-draft-creator)
- [Input Data Specifications](#input-data-specifications)
- [Output Specifications](#output-specifications)
- [Key Design Decisions](#key-design-decisions)
- [Known Issues and Fixes](#known-issues-and-fixes)
- [Dependencies](#dependencies)
- [Deployment and Sharing](#deployment-and-sharing)

---

## Project Overview

Automates the end-to-end workflow for generating per-manager attainment reports and creating Outlook email drafts with those reports attached.

**Two interfaces are available:**
1. **Streamlit Web UI** (`app.py`) — 5-step wizard with file upload, folder selection, region/manager filtering, and Outlook draft generation.
2. **CLI scripts** — `generate_manager_reports.py` and `create_email_drafts.py` can be run independently from the command line.

**GitHub Repository:** `https://github.com/LI-JBLEE/AttnRptAutomation`

---

## Architecture

```
┌─────────────────────────────────────────────────────────────┐
│  app.py (Streamlit UI — 5-Step Wizard)                      │
│  ┌────────────┐ ┌────────────┐ ┌──────────────────────────┐ │
│  │ Step 1-2   │→│  Step 3    │→│    Step 4-5              │ │
│  │ File Upload│ │ Generate   │ │ Region/Manager Select    │ │
│  │ Folder Pick│ │ Reports    │ │ → Outlook Email Drafts   │ │
│  └────────────┘ └────────────┘ └──────────────────────────┘ │
│       ↓              ↓              ↓                       │
│  generate_manager_reports.py   create_email_drafts.py       │
│  (core functions imported)     (core functions imported)    │
└─────────────────────────────────────────────────────────────┘
```

Both CLI scripts expose reusable functions that `app.py` imports directly. The `main()` functions in each CLI script are preserved for standalone usage.

---

## File Structure

```
Attainment report automation/
├── app.py                          # Streamlit UI (5-step wizard)
├── generate_manager_reports.py     # Report generation engine
├── create_email_drafts.py          # Outlook email draft creator
├── .gitignore                      # Excludes .xlsx, .csv, Manager report/, etc.
├── DEVELOPMENT.md                  # This file
│
├── FY26_Global_Attainment_Club.xlsx          # Input: attainment data (not in git)
├── Sales Compensation Report (Daily) *.xlsx  # Input: email mapping (not in git)
│
└── Manager report/                 # Generated output (not in git)
    ├── APAC/
    │   └── FY26_Attainment_{Name}_{YYYYMMDD}.xlsx
    ├── CHINA/
    ├── EMEA/
    ├── LATAM/
    └── NAMER/
```

---

## Data Flow

```
Step 1: Upload Files
  FY26_Global_Attainment_Club.xlsx ──→ attainment_df (pandas DataFrame)
  Sales Compensation Report *.xlsx ──→ email_map {emp_id: email}

Step 2: Choose Output Folder
  User selects parent folder ──→ reports saved to {parent}/Manager report/{Region}/

Step 3: Generate Reports
  attainment_df ──→ generate_all_reports() ──→ 688 Excel files in Region subfolders

Step 4: Select Recipients
  Scan generated files ──→ match to emails via Employee ID ──→ user filters by Region/Manager

Step 5: Create Email Drafts
  selected managers ──→ create_drafts_batch() ──→ Outlook Drafts/Manager Report folder
```

---

## Module Reference

### app.py — Streamlit UI

**Run:** `streamlit run app.py`

**5-Step Wizard:**

| Step | Description | Key Components |
|------|-------------|----------------|
| 1 | Upload attainment + sales comp Excel files | `st.file_uploader`, validation functions |
| 2 | Choose parent output folder | `st.text_input` + tkinter `filedialog.askdirectory()` |
| 3 | Select regions, generate reports | `st.multiselect`, `st.progress`, calls `generate_all_reports()` |
| 4 | Filter by Region + Manager for email recipients | `st.multiselect`, metrics display |
| 5 | Generate Outlook email drafts | Calls `create_drafts_batch()`, progress bar |

**Key Functions:**

| Function | Purpose |
|----------|---------|
| `select_folder_dialog()` | Opens native Windows folder picker via tkinter |
| `validate_attainment_file(uploaded_file)` | Validates "in" sheet and `Level_1_Manager` column; renames `Plan_Period;MBO_Description` → `Plan_Period` |
| `validate_sales_comp_file(uploaded_file)` | Validates "Sheet1" (header=3), builds `{emp_id: email}` map |
| `build_manager_list(attainment_df, email_map, output_dir)` | Scans generated report files, matches to emails, returns list of manager dicts |

**Session State Keys:**

| Key | Type | Description |
|-----|------|-------------|
| `attainment_df` | DataFrame | Uploaded attainment data |
| `email_map` | dict | `{emp_id_str: email}` mapping |
| `output_dir` | str | User-selected parent folder path |
| `reports_generated` | bool | Whether Step 3 is complete |
| `report_results` | dict | `{total, region_counts, managers}` |
| `manager_list` | list | Manager dicts for Steps 4-5 |
| `available_regions` | list | Sorted region strings from attainment data |

**Custom CSS:** Gradient title, blue left-border step headers (1.15rem), light blue multiselect tags (`#dbeafe`), rounded buttons/cards, styled metrics.

---

### generate_manager_reports.py — Report Generator

**CLI Usage:** `python generate_manager_reports.py`

Generates one Excel report per L1 manager with recursive hierarchy structure, outline grouping, color-coded attainment values, and region-based subfolders.

**Key Exported Functions (used by app.py):**

| Function | Signature | Returns |
|----------|-----------|---------|
| `extract_manager_name(full_name)` | `str → str` | Clean name from `"First Last (ID)"` format |
| `extract_manager_id(full_name)` | `str → str\|None` | Employee ID from `"First Last (ID)"` format |
| `sanitize_filename(name)` | `str → str` | Windows-safe filename; strips parenthesized content (e.g. Chinese aliases) |
| `get_all_regions(source_df)` | `DataFrame → list[str]` | Sorted unique region names |
| `generate_all_reports(source_df, output_dir, progress_callback=None, selected_regions=None)` | See below | `{total, region_counts, managers}` |

**`generate_all_reports()` Details:**

```python
def generate_all_reports(source_df, output_dir, progress_callback=None,
                         selected_regions=None):
    """
    Args:
        source_df: DataFrame with Plan_Period column already renamed
        output_dir: e.g. "C:/Reports/Manager report"
        progress_callback: callable(current, total, message) for UI updates
        selected_regions: list of region strings to filter by (None = all)

    Returns:
        {
            "total": int,             # number of managers processed
            "region_counts": dict,    # {"APAC": 85, "EMEA": 210, ...}
            "managers": list          # [(full_name, region, safe_name, filepath), ...]
        }
    """
```

**Internal Functions (not exported but important for understanding):**

| Function | Purpose |
|----------|---------|
| `build_id_mappings(df)` | Builds `{emp_id: PersonName}`, set of L1 manager IDs, `{emp_id: L1ManagerString}` |
| `build_manager_region_map(df, person_id_to_name)` | Maps `L1_Manager_string → Region` using ID matching, falls back to mode of direct reports |
| `build_hierarchy_data(df, manager, ...)` | Recursively builds hierarchical data list with depth tracking, section headers, and circular-reference guards |
| `write_report(wb, manager, hierarchy_data, columns)` | Writes formatted Excel report: title bar, headers, color-coded data, outline grouping, freeze panes |

**Report Excel Format:**
- Row 1: Title bar (dark navy, merged)
- Row 2: Subtitle with date and manager name
- Row 3: Blank spacer
- Row 4: Column headers (frozen, auto-filtered)
- Row 5+: Data rows with hierarchy indentation, section headers, outline grouping
- Color coding: green (>=100%), yellow (80-99%), red (<80%) for attainment columns
- Alternating row colors by hierarchy depth level

**File Cleanup Behavior:**
- Only deletes `FY26_Attainment_*.xlsx` files in region subfolders being regenerated
- Does NOT delete the entire output directory (previous bug fix)

---

### create_email_drafts.py — Outlook Email Draft Creator

**CLI Usage:** `python create_email_drafts.py "Manager report/APAC" [--dry-run]`

**Key Exported Functions (used by app.py):**

| Function | Signature | Returns |
|----------|-----------|---------|
| `clean_display_name(full_name)` | `str → str` | Name without ID or parenthesized aliases |
| `load_email_mapping(sales_comp_file)` | `path → dict` | `{emp_id_str: email}` from Sales Comp report |
| `create_drafts_batch(matched_list, target_folder_name, progress_callback)` | See below | `{created, failed, failures_detail}` |

**`create_drafts_batch()` Details:**

```python
def create_drafts_batch(matched_list, target_folder_name="Manager Report",
                        progress_callback=None):
    """
    Args:
        matched_list: [(filepath, clean_name, email), ...]
        target_folder_name: Outlook Drafts subfolder name
        progress_callback: callable(current, total, message)

    Returns:
        {
            "created": int,
            "failed": int,
            "failures_detail": [(name, email, error_str), ...]
        }
    """
```

**Outlook COM Integration:**
- Uses `win32com.client.Dispatch("Outlook.Application")` for Outlook automation
- `GetDefaultFolder(16)` = olFolderDrafts
- Creates "Manager Report" subfolder under Drafts via `get_or_create_drafts_subfolder()`
- `mail.Save()` saves as draft (does NOT send)
- `mail.Move(target_folder)` moves draft into the subfolder

**Email Template:**
- Subject: `"FY26 Attainment Report - {manager_name}"`
- Body: HTML format (Calibri font, professional styling)
- Attachment: the manager's Excel report file

---

## Input Data Specifications

### Global Attainment Report (`FY26_Global_Attainment_Club.xlsx`)

| Property | Value |
|----------|-------|
| Sheet name | `in` |
| Key columns | `Level_1_Manager`, `Level_2_Manager`, `Person Name`, `LI_EMP_ID`, `Region` |
| Special column | `Plan_Period;MBO_Description` (renamed to `Plan_Period` at load time) |
| Name format | `"First Last (EmpID)"` — e.g. `"John Smith (12345)"` |
| Typical size | ~6,595 rows, 64 columns, 688 unique L1 managers |
| Regions | APAC, CHINA, EMEA, LATAM, NAMER |

### Sales Compensation Report

| Property | Value |
|----------|-------|
| Sheet name | `Sheet1` |
| Header row | Row 3 (0-indexed), i.e. `header=3` in pandas |
| Key columns | `Employee ID`, `Email - Work` |
| Employee ID format | Zero-padded string, e.g. `"000081"` |
| ID normalization | Leading zeros are stripped to match attainment IDs: `"000081"` → `"81"` |

---

## Output Specifications

### Report Files

- **Path pattern:** `{parent_folder}/Manager report/{Region}/FY26_Attainment_{SafeName}_{YYYYMMDD}.xlsx`
- **Filename sanitization:** Removes parenthesized content (e.g. Chinese name aliases like `"(黄策)"`), replaces illegal Windows filename characters with `_`
- **One file per L1 manager**, containing all direct and indirect reports in a recursive hierarchy

### Outlook Email Drafts

- **Location:** Outlook Drafts folder → "Manager Report" subfolder
- **One draft per selected manager**, with the corresponding report file attached
- **Not sent** — user reviews and sends manually from Outlook

---

## Key Design Decisions

### ID-Based Matching
Employee IDs (extracted from `"Name (ID)"` format) are used for matching across all data sources because `Person Name` and `Level_1_Manager` strings can differ (e.g. name aliases, parenthesized Chinese names).

### Leading Zero Normalization
The Sales Comp report uses zero-padded IDs (`"000081"`) while attainment data uses plain IDs (`"81"`). All IDs are normalized by stripping leading zeros: `str(emp_id).lstrip("0") or "0"`.

### Parenthesized Alias Handling
Some manager names contain parenthesized aliases (e.g. Chinese names like `"Ce Huang (黄策)"`). These are:
- Stripped from filenames via `sanitize_filename()`
- Stripped from email display names via `clean_display_name()`
- Stripped from report titles via regex in `write_report()`

### Safe File Cleanup
Report regeneration only deletes existing `FY26_Attainment_*.xlsx` files in the specific region subfolders being regenerated. This prevents accidentally destroying user data in the parent folder (a previous bug caused by `shutil.rmtree()`).

### Manager Report Subfolder
The user selects a parent folder, and reports are automatically saved under a `Manager report` subfolder with Region subfolders inside. This keeps generated reports organized and separate from other files in the parent folder.

### Region-Based Filtering
Both report generation (Step 3) and email draft creation (Step 4) support region filtering. Users can select specific regions to process, reducing generation time and allowing targeted distribution.

### Folder Picker
A native Windows folder picker dialog (via tkinter `filedialog.askdirectory()`) is integrated into the Streamlit UI alongside a manual text input field.

---

## Known Issues and Fixes

| Issue | Root Cause | Fix Applied |
|-------|-----------|-------------|
| `PermissionError` when generating reports | `shutil.rmtree()` deleted the entire user-selected folder, including existing business data and OneDrive-locked files | Removed `shutil.rmtree()`, now only deletes `FY26_Attainment_*.xlsx` files in targeted region subfolders |
| Email match failures (~2 of 688) | Some managers exist in attainment data but not in the Sales Comp report | Displayed as warnings in UI; does not block other drafts |
| `pip install streamlit` fails silently | pip not on PATH in some environments | Use `python -m pip install streamlit` instead |
| Name mismatch between Person Name and Level_1_Manager | Same person may have different name strings in different columns | ID-based matching resolves this; names are matched by employee ID, not string comparison |
| Chinese name aliases in filenames/emails | Names like `"Ce Huang (黄策)"` cause issues in filenames and email subjects | `sanitize_filename()` and `clean_display_name()` strip all parenthesized content |

---

## Dependencies

```
pip install streamlit pandas openpyxl pywin32
```

| Package | Version | Purpose |
|---------|---------|---------|
| `streamlit` | latest | Web UI framework |
| `pandas` | latest | Excel file loading and data manipulation |
| `openpyxl` | latest | Excel file writing with formatting |
| `pywin32` | latest | Outlook COM automation (`win32com.client`) |
| `tkinter` | built-in | Native Windows folder picker dialog |

**Platform requirement:** Windows only (Outlook COM requires desktop Outlook installed)

---

## Deployment and Sharing

### Local Distribution
1. Share the 3 Python files: `app.py`, `generate_manager_reports.py`, `create_email_drafts.py`
2. Each user installs dependencies: `pip install streamlit pandas openpyxl pywin32`
3. Each user runs locally: `streamlit run app.py`
4. All processing (file I/O, Outlook draft creation) happens on the user's local machine

### Limitations
- **Cannot be hosted as a web app** (e.g. GitHub Pages, Streamlit Cloud) because:
  - Requires local file system access for saving reports
  - Requires Outlook COM (Windows desktop Outlook) for email drafts
  - Data files are processed locally, not uploaded to a server
- Each user must have Microsoft Outlook installed and running on their Windows machine

### Git Configuration
The `.gitignore` excludes all data files (`.xlsx`, `.csv`), generated reports (`Manager report/`), Python cache, IDE files, and Claude Code metadata. Only the Python source files are tracked.
