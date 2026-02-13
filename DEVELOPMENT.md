# GSC Attainment Report Automator — Development Reference

## Table of Contents

- [Project Overview](#project-overview)
- [Architecture](#architecture)
- [File Structure](#file-structure)
- [Data Flow](#data-flow)
- [Module Reference](#module-reference)
  - [app.py — Streamlit Web App](#apppy--streamlit-web-app)
  - [email_manager.py — EmailManager GUI](#email_managerpy--emailmanager-gui)
  - [generate_manager_reports.py — Report Generator](#generate_manager_reportspy--report-generator)
  - [create_email_drafts.py — CLI Email Draft Creator](#create_email_draftspy--cli-email-draft-creator)
- [Input Data Specifications](#input-data-specifications)
- [Output Specifications](#output-specifications)
- [Key Design Decisions](#key-design-decisions)
- [Known Issues and Fixes](#known-issues-and-fixes)
- [Dependencies](#dependencies)
- [Build and Deployment](#build-and-deployment)

---

## Project Overview

Automates the end-to-end workflow for generating per-manager attainment reports and distributing them via Outlook email.

**Two-component architecture:**
1. **Streamlit Web App** (`app.py`) — Cloud-hosted report generation with .zip download
2. **EmailManager.exe** (`email_manager.py`) — Standalone Windows desktop app for Outlook email operations

**GitHub Repository:** `https://github.com/LI-JBLEE/AttnRptAutomation`
**Live Web App:** `https://manager-attn-report.streamlit.app/`

---

## Architecture

```
┌─────────────────────────────────────────────────────────────┐
│  Streamlit Web App (app.py)                                  │
│  ┌────────────┐ ┌────────────┐ ┌──────────────────────────┐ │
│  │  Step 1    │→│  Step 2    │→│  Download                │ │
│  │ File Upload│ │ Generate   │ │ Reports.zip              │ │
│  │ (2 files)  │ │ Reports    │ │ + EmailManager.zip       │ │
│  └────────────┘ └────────────┘ └──────────────────────────┘ │
│       ↓              ↓                                       │
│  generate_manager_reports.py   create_email_drafts.py       │
│  (core functions imported)     (validation functions only)  │
└─────────────────────────────────────────────────────────────┘
                         │
                    .zip download
                         ↓
┌─────────────────────────────────────────────────────────────┐
│  EmailManager.exe (email_manager.py — standalone)           │
│  ┌────────────┐ ┌────────────┐ ┌──────────────────────────┐ │
│  │  Step 1    │→│  Step 2    │→│  Step 3                  │ │
│  │ Load .zip  │ │ Select     │ │ Create Drafts tab:       │ │
│  │ or folder  │ │ Recipients │ │  - Edit subject/body     │ │
│  │            │ │ (regions + │ │  - Create Outlook drafts │ │
│  │            │ │  managers) │ │ Send Drafts tab:         │ │
│  │            │ │            │ │  - Load/send from Outlook│ │
│  └────────────┘ └────────────┘ └──────────────────────────┘ │
│       ↓                              ↓                       │
│  Embedded functions              win32com.client             │
│  (self-contained .exe)           (Outlook COM)               │
└─────────────────────────────────────────────────────────────┘
```

The web app generates reports and bundles them into a .zip with metadata. The EmailManager.exe loads that .zip and handles all Outlook operations locally.

---

## File Structure

```
AttnRptAutomation/
├── app.py                          # Streamlit web app (Steps 1-2 + download)
├── email_manager.py                # EmailManager GUI (Tkinter, standalone .exe source)
├── email_manager.spec              # PyInstaller build configuration
├── generate_manager_reports.py     # Report generation engine
├── create_email_drafts.py          # CLI email draft creator (legacy)
├── requirements.txt                # Python dependencies
├── .streamlit/
│   └── config.toml                 # Streamlit theme (primaryColor: #0078D4)
├── dist/
│   └── EmailManager.exe            # Built executable (~13MB)
├── build/                          # PyInstaller build artifacts (not in git)
├── DEVELOPMENT.md                  # This file
├── README.md                       # Project overview (Korean)
├── USER_MANUAL.md                  # User manual (English)
└── .gitignore                      # Excludes .xlsx, .csv, build/, etc.
```

---

## Data Flow

```
┌─ Web App (Streamlit Cloud) ─────────────────────────────────┐
│                                                              │
│  Step 1: Upload Files                                        │
│    Attainment Report (.xlsx) ──→ attainment_df (DataFrame)   │
│    Sales Comp Report (.xlsx) ──→ email_map {emp_id: email}   │
│                                                              │
│  Step 2: Generate Reports                                    │
│    attainment_df ──→ generate_all_reports() ──→ Excel files  │
│    (generated in temp directory, in-memory)                   │
│                                                              │
│  Download: Bundle into .zip                                  │
│    Excel files + manager_metadata.json ──→ Reports.zip       │
│    EmailManager.exe ──→ EmailManager.zip                     │
└──────────────────────────────────────────────────────────────┘
                              │
                         .zip files
                              ↓
┌─ EmailManager.exe (local Windows PC) ───────────────────────┐
│                                                              │
│  Step 1: Load .zip                                           │
│    Reports.zip ──→ extract to temp dir ──→ load metadata     │
│                                                              │
│  Step 2: Select Recipients                                   │
│    Region filter + manager list ──→ selected managers        │
│                                                              │
│  Step 3: Email Operations                                    │
│    Create Drafts tab:                                        │
│      Edit subject/body templates ──→ create_drafts_batch()   │
│      ──→ Outlook Drafts/Manager Report folder                │
│    Send Drafts tab:                                          │
│      Load from Outlook ──→ select ──→ send_drafts_batch()    │
└──────────────────────────────────────────────────────────────┘
```

---

## Module Reference

### app.py — Streamlit Web App

**Run locally:** `python -m streamlit run app.py`
**Deployed at:** `https://manager-attn-report.streamlit.app/`

**2-Step Wizard + Download:**

| Step | Description | Key Components |
|------|-------------|----------------|
| 1 | Upload attainment + sales comp Excel files | `st.file_uploader` with dynamic keys for reset, validation functions |
| 2 | Select regions, generate reports | `st.multiselect`, `st.progress`, calls `generate_all_reports()` |
| Download | .zip of reports + EmailManager.zip | In-memory zip creation, `st.download_button` |

**Key Functions:**

| Function | Purpose |
|----------|---------|
| `validate_attainment_file(uploaded_file)` | Validates "in" sheet and `Level_1_Manager` column; renames `Plan_Period;MBO_Description` → `Plan_Period` |
| `validate_sales_comp_file(uploaded_file)` | Validates "Sheet1" (header=3), builds `{emp_id: email}` map |

**Session State Keys:**

| Key | Type | Description |
|-----|------|-------------|
| `attainment_df` | DataFrame | Uploaded attainment data |
| `email_map` | dict | `{emp_id_str: email}` mapping |
| `reports_generated` | bool | Whether Step 2 is complete |
| `report_results` | dict | `{total, region_counts, managers}` |
| `available_regions` | list | Sorted region strings from attainment data |
| `fiscal_year` | str | Detected fiscal year (e.g. "FY26") |
| `temp_dir` | str | Temp directory for generated reports |
| `reset_count` | int | Incremented on reset to force new widget keys |

**Reset Mechanism:**
The Reset button increments `reset_count` and clears all session state. File uploader widgets use `key=f"attainment_uploader_{rc}"` so they are recreated as new widgets, clearing previously uploaded files.

**Theme:** `.streamlit/config.toml` sets `primaryColor = "#0078D4"` (Microsoft blue).

---

### email_manager.py — EmailManager GUI

**Run:** `python email_manager.py` or `dist/EmailManager.exe`

Self-contained Tkinter GUI that embeds all email-related functions (no imports from other project files). Built as a standalone .exe via PyInstaller.

**Class: `EmailManagerApp`**

| Method | Purpose |
|--------|---------|
| `_create_widgets()` | Builds full GUI: load, recipients, email template editor, send tabs |
| `_load_zip_file()` | Extracts .zip, loads `manager_metadata.json`, sets fiscal year |
| `_load_folder()` | Scans folder for report files, prompts for Sales Comp file |
| `_update_region_checkboxes()` | Creates region filter checkboxes from loaded data |
| `_update_manager_list()` | Filters and displays managers based on region selection |
| `_get_default_subject()` | Returns `"{fiscal_year} Attainment Report - {manager_name}"` |
| `_get_default_body()` | Returns default plain text email body with placeholders |
| `_reset_template()` | Restores subject and body fields to defaults |
| `_create_drafts()` | Reads template from UI, runs `create_drafts_batch()` in thread |
| `_load_outlook_drafts()` | Loads drafts from Outlook Drafts/Manager Report folder |
| `_send_drafts()` | Sends selected drafts via `send_drafts_batch()` in thread |

**Email Template Editor (Create Drafts tab):**
- Subject field (`ttk.Entry`) — editable, pre-filled with default template
- Body field (`tk.Text`, 8 lines) — editable, pre-filled with default plain text
- Placeholders: `{manager_name}`, `{fiscal_year}` — resolved at draft creation time
- "Reset to Default" button restores default templates

**Embedded Functions (not imported, copied for standalone .exe):**

| Function | Purpose |
|----------|---------|
| `clean_display_name(full_name)` | Removes IDs and parenthesized aliases from names |
| `plain_text_to_html(text)` | Converts plain text to HTML with Calibri font styling |
| `get_email_subject(fiscal_year)` | Default subject template |
| `get_email_html(fiscal_year)` | Default HTML body template |
| `create_draft(outlook, ...)` | Creates single Outlook draft; accepts custom `subject_template` and `body_text` |
| `create_drafts_batch(matched_list, ...)` | Creates drafts for multiple managers; passes custom templates through |
| `get_drafts_from_folder(outlook, folder_name)` | Retrieves drafts from Outlook subfolder using EntryID |
| `send_drafts_batch(draft_items, ...)` | Sends drafts; creates own COM object per thread |

**Outlook COM Integration:**
- Uses `win32com.client.Dispatch("Outlook.Application")`
- `GetDefaultFolder(16)` = olFolderDrafts
- Creates "Manager Report" subfolder under Drafts
- Draft creation and sending run in background threads
- Each thread creates its own COM object (COM objects cannot be marshalled across threads)
- Drafts are referenced by `EntryID` (persistent identifier) for reliable retrieval

---

### generate_manager_reports.py — Report Generator

**CLI Usage:** `python generate_manager_reports.py`

Generates one Excel report per L1 manager with recursive hierarchy structure, outline grouping, color-coded attainment values, and region-based subfolders.

**Key Exported Functions (used by app.py):**

| Function | Signature | Returns |
|----------|-----------|---------|
| `extract_manager_name(full_name)` | `str → str` | Clean name from `"First Last (ID)"` format |
| `extract_manager_id(full_name)` | `str → str\|None` | Employee ID from `"First Last (ID)"` format |
| `sanitize_filename(name)` | `str → str` | Windows-safe filename; strips parenthesized content |
| `get_all_regions(source_df)` | `DataFrame → list[str]` | Sorted unique region names |
| `get_fiscal_year(source_df)` | `DataFrame → str` | Fiscal year string (e.g. "FY26") from data |
| `generate_all_reports(...)` | See below | `{total, region_counts, managers}` |

**`generate_all_reports()` Details:**

```python
def generate_all_reports(source_df, output_dir, progress_callback=None,
                         selected_regions=None, fiscal_year=None):
    """
    Returns:
        {
            "total": int,
            "region_counts": dict,    # {"APAC": 85, "EMEA": 210, ...}
            "managers": list          # [(full_name, region, safe_name, filepath), ...]
        }
    """
```

**Report Excel Format:**
- Row 1: Title bar (dark navy, merged)
- Row 2: Subtitle with date and manager name
- Row 3: Blank spacer
- Row 4: Column headers (frozen, auto-filtered)
- Row 5+: Data rows with hierarchy indentation, section headers, outline grouping
- Color coding: green (>=100%), yellow (80-99%), red (<80%) for attainment columns
- Alternating row colors by hierarchy depth level

---

### create_email_drafts.py — CLI Email Draft Creator

**CLI Usage:** `python create_email_drafts.py "Manager report/APAC" [--dry-run]`

Legacy standalone script for creating Outlook drafts from the command line. Functions are also imported by `app.py` for validation (`load_email_mapping`, `clean_display_name`).

Note: `email_manager.py` embeds its own copies of the email functions for standalone .exe operation — it does not import from this file.

---

## Input Data Specifications

### Global Attainment Report

| Property | Value |
|----------|-------|
| Sheet name | `in` |
| Key columns | `Level_1_Manager`, `Level_2_Manager`, `Person Name`, `LI_EMP_ID`, `Region` |
| Special column | `Plan_Period;MBO_Description` (renamed to `Plan_Period` at load time) |
| Name format | `"First Last (EmpID)"` — e.g. `"John Smith (12345)"` |
| Fiscal Year detection | From `Fiscal Year` column values |
| Regions | APAC, CHINA, EMEA, LATAM, NAMER |

### Sales Compensation Report

| Property | Value |
|----------|-------|
| Sheet name | `Sheet1` |
| Header row | Row 4 (1-indexed), i.e. `header=3` in pandas |
| Key columns | `Employee ID`, `Email - Work` |
| Employee ID format | Zero-padded string, e.g. `"000081"` |
| ID normalization | Leading zeros stripped: `"000081"` → `"81"` |

---

## Output Specifications

### Report .zip File

- **Filename:** `Manager_Reports_{FY}_{YYYYMMDD_HHMMSS}.zip`
- **Contents:**
  - `{Region}/FY26_Attainment_{SafeName}_{YYYYMMDD}.xlsx` — per-manager Excel reports
  - `manager_metadata.json` — manager name, email, region, filepath mappings

### manager_metadata.json

```json
{
  "fiscal_year": "FY26",
  "generated_date": "2026-02-13",
  "total_reports": 688,
  "managers": [
    {
      "name": "John Smith",
      "safe_name": "John_Smith",
      "region": "NAMER",
      "email": "john.smith@company.com",
      "filepath": "NAMER/FY26_Attainment_John_Smith_20260213.xlsx"
    }
  ]
}
```

### Outlook Email Drafts

- **Location:** Outlook Drafts → "Manager Report" subfolder
- **Subject:** Configurable template, default: `{fiscal_year} Attainment Report - {manager_name}`
- **Body:** Configurable plain text template, converted to HTML (Calibri font)
- **Attachment:** Manager's Excel report file
- **Template variables:** `{manager_name}`, `{fiscal_year}` — resolved at draft creation time

---

## Key Design Decisions

### Two-Component Architecture
The web app handles report generation (cross-platform, cloud-hosted) while email operations are handled by a local Windows .exe. This separation allows:
- Report generation from any device with a browser
- Email sending only from Windows PCs with Outlook (COM requirement)
- No email addresses uploaded to the cloud (security)

### Standalone .exe (EmailManager)
`email_manager.py` embeds all email-related functions rather than importing from `create_email_drafts.py`. This allows PyInstaller to produce a single self-contained .exe without requiring other project files.

### Customizable Email Templates
Email subject and body are editable in the EmailManager GUI. Templates use `{manager_name}` and `{fiscal_year}` placeholders that are resolved at draft creation time. Plain text body input is converted to HTML via `plain_text_to_html()`.

### ID-Based Matching
Employee IDs (extracted from `"Name (ID)"` format) are used for matching across all data sources because name strings can differ (aliases, parenthesized Chinese names).

### Leading Zero Normalization
Sales Comp uses zero-padded IDs (`"000081"`), attainment data uses plain IDs (`"81"`). All IDs are normalized by `str(emp_id).lstrip("0") or "0"`.

### Parenthesized Alias Handling
Names like `"Ce Huang (黄策)"` are stripped from filenames (`sanitize_filename()`), email display names (`clean_display_name()`), and report titles.

### COM Threading
Outlook COM objects must be created and used in the same thread. `create_drafts_batch()` and `send_drafts_batch()` each create their own COM instances. Draft items are referenced by `EntryID` (persistent Outlook identifier) rather than COM object references.

### Reset Mechanism (Web App)
Streamlit's `file_uploader` retains files when the widget key stays the same. The Reset button increments a `reset_count` counter used in widget keys (`f"attainment_uploader_{rc}"`) to force widget re-creation.

### .zip Download for .exe
`EmailManager.exe` is wrapped in a .zip file for download to avoid browser .exe download blocking. Windows SmartScreen warning still appears on first run (requires code signing to eliminate).

---

## Known Issues and Fixes

| Issue | Root Cause | Fix Applied |
|-------|-----------|-------------|
| `PermissionError` when generating reports | `shutil.rmtree()` deleted entire user folder | Only deletes `FY*_Attainment_*.xlsx` files in targeted region subfolders |
| Email match failures (~2 of 688) | Managers in attainment data but not in Sales Comp | Displayed as warnings; does not block other drafts |
| Reset button not clearing uploaded files | Streamlit `file_uploader` retains files when key stays the same | Dynamic widget keys using `reset_count` counter |
| COM "marshalled for different thread" error | Outlook COM objects shared across threads | Each thread creates its own COM object |
| "Inline response" send failure | Draft is in reply/forward mode | User opens draft in new window, saves, retries |
| Browser blocks .exe download | Browsers block direct .exe downloads | Wrapped in .zip file |
| SmartScreen warning on .exe | Unsigned executable | User instructions to bypass; code signing for future |

---

## Dependencies

### Web App (Streamlit Cloud)

```
pip install streamlit pandas openpyxl
```

### EmailManager.exe (Windows only)

```
pip install pywin32
```

### Full Development Environment

```
pip install -r requirements.txt
pip install pyinstaller  # for building .exe
```

| Package | Purpose |
|---------|---------|
| `streamlit` | Web UI framework |
| `pandas` | Excel file loading and data manipulation |
| `openpyxl` | Excel file writing with formatting |
| `pywin32` | Outlook COM automation (`win32com.client`) |
| `pyinstaller` | Building EmailManager.exe |
| `tkinter` | EmailManager GUI (built-in) |

---

## Build and Deployment

### Building EmailManager.exe

```bash
pip install pyinstaller
pyinstaller email_manager.spec
# Output: dist/EmailManager.exe (~13MB)
```

The `.spec` file excludes heavy packages (pandas, numpy, etc.) not used by EmailManager to reduce file size.

### Deploying Web App

The Streamlit app is deployed on Streamlit Cloud, reading directly from the GitHub repository. All code changes must be pushed to `main` branch to reflect on the live app.

```bash
git push origin main
# Streamlit Cloud auto-deploys from main
# If not auto-deployed, manually reboot from Streamlit Cloud dashboard
```

### Local Development

```bash
# Web App
python -m streamlit run app.py

# EmailManager GUI
python email_manager.py
```

### Git Configuration

The `.gitignore` excludes data files (`.xlsx`, `.csv`), build artifacts (`build/`), Python cache, and IDE files. `dist/EmailManager.exe` is tracked in git for distribution via the web app.
