"""
FY26 Attainment Report - Outlook Email Draft Creator
Creates Outlook drafts with attainment reports attached for managers
in a specified folder. Does NOT send emails.
"""

import argparse
import os
import re
import sys
import pandas as pd

# Reuse utilities from the report generator
from generate_manager_reports import (
    extract_manager_name, extract_manager_id, sanitize_filename
)


def clean_display_name(full_name):
    """Clean manager name for email display: remove ID and any parenthesized content."""
    name = extract_manager_name(full_name)  # removes trailing "(ID)"
    name = re.sub(r'\s*\([^)]*\)', '', name)  # removes remaining parenthesized aliases
    return name.strip()

# ── Configuration ──────────────────────────────────────────────────────────────
BASE_DIR = r"c:\Claude\Attainment report automation"
ATTAINMENT_FILE = os.path.join(BASE_DIR, "FY26_Global_Attainment_Club.xlsx")
SALES_COMP_FILE = os.path.join(
    BASE_DIR, "Sales Compensation Report (Daily) 2026-02-12 08_03 PST.xlsx"
)

EMAIL_SUBJECT = "FY26 Attainment Report - {manager_name}"
EMAIL_HTML = """\
<html>
<body style="font-family: Calibri, Arial, sans-serif; font-size: 11pt; color: #333;">
<p>Hi {manager_name},</p>

<p>Please find attached your <b>FY26 Attainment Report</b>.</p>

<p>This report includes attainment data for your team, organized by hierarchy
with quarterly, half-year, and annual breakdowns.</p>

<p>If you have any questions about the data, please reach out to the
Sales Compensation team.</p>

<p>Best regards,<br>
Sales Compensation</p>
</body>
</html>
"""


def load_email_mapping(sales_comp_file):
    """Load {emp_id_str: email} mapping from Sales Compensation Report."""
    print(f"Loading email data from: {os.path.basename(sales_comp_file)}")
    # Header is on row 3 (0-indexed); Employee ID is zero-padded string ("000081")
    df = pd.read_excel(sales_comp_file, sheet_name="Sheet1", header=3)

    email_map = {}
    for _, row in df.iterrows():
        emp_id = row.get("Employee ID")
        email = row.get("Email - Work")
        if pd.notna(emp_id) and pd.notna(email):
            # Strip leading zeros to match attainment data IDs (e.g. "000081" -> "81")
            clean_id = str(emp_id).lstrip("0") or "0"
            email_map[clean_id] = str(email).strip()

    print(f"  Loaded {len(email_map)} employee email records")
    return email_map


def build_manager_name_to_id(attainment_file):
    """
    Build mapping from sanitized manager name -> (full_l1_manager_string, emp_id)
    using Level_1_Manager column from attainment data.
    """
    print(f"Loading attainment data from: {os.path.basename(attainment_file)}")
    df = pd.read_excel(attainment_file, sheet_name="in")

    name_to_info = {}
    for mgr in df["Level_1_Manager"].dropna().unique():
        emp_id = extract_manager_id(mgr)
        clean_name = extract_manager_name(mgr)
        safe_name = sanitize_filename(clean_name)
        if emp_id:
            # Strip leading zeros to normalize ID (e.g. "000475" -> "475")
            clean_id = emp_id.lstrip("0") or "0"
            name_to_info[safe_name] = (mgr, clean_id)

    print(f"  Found {len(name_to_info)} unique managers with IDs")
    return name_to_info


def scan_report_files(folder_path):
    """
    Scan folder (and subfolders) for FY26_Attainment_*.xlsx files.
    Returns list of (filepath, manager_name_from_filename).
    """
    files = []
    for root, _, filenames in os.walk(folder_path):
        for fname in filenames:
            if fname.startswith("FY26_Attainment_") and fname.endswith(".xlsx"):
                # Extract manager name from filename:
                # FY26_Attainment_{ManagerName}_{YYYYMMDD}.xlsx
                parts = fname[len("FY26_Attainment_"):]  # remove prefix
                # Remove .xlsx and the last _YYYYMMDD part
                parts = parts[:-5]  # remove .xlsx
                last_underscore = parts.rfind("_")
                if last_underscore > 0:
                    mgr_name = parts[:last_underscore]
                else:
                    mgr_name = parts
                files.append((os.path.join(root, fname), mgr_name))
    return files


def get_or_create_drafts_subfolder(outlook, folder_name="Manager Report"):
    """
    Get or create a subfolder under the Outlook Drafts folder.
    Returns the target folder MAPI object.
    """
    ns = outlook.GetNamespace("MAPI")
    drafts = ns.GetDefaultFolder(16)  # olFolderDrafts

    # Check if subfolder already exists
    for i in range(drafts.Folders.Count):
        folder = drafts.Folders.Item(i + 1)  # 1-indexed
        if folder.Name == folder_name:
            return folder

    # Create subfolder
    return drafts.Folders.Add(folder_name)


def create_draft(outlook, to_email, manager_name, attachment_path,
                 target_folder=None):
    """Create a single Outlook draft email with the report attached."""
    mail = outlook.CreateItem(0)  # olMailItem
    mail.To = to_email
    mail.Subject = EMAIL_SUBJECT.format(manager_name=manager_name)
    mail.HTMLBody = EMAIL_HTML.format(manager_name=manager_name)
    mail.Attachments.Add(os.path.abspath(attachment_path))
    mail.Save()  # Save as Draft — does NOT send

    # Move to target subfolder if specified
    if target_folder is not None:
        mail.Move(target_folder)


def create_drafts_batch(matched_list, target_folder_name="Manager Report",
                        progress_callback=None):
    """
    Create Outlook drafts for a list of matched managers.

    Args:
        matched_list: list of (filepath, clean_name, email)
        target_folder_name: Outlook Drafts subfolder name
        progress_callback: optional callable(current, total, message)

    Returns:
        dict with keys: created, failed, failures_detail
    """
    import win32com.client
    outlook = win32com.client.Dispatch("Outlook.Application")
    target_folder = get_or_create_drafts_subfolder(outlook, target_folder_name)

    total = len(matched_list)
    created = 0
    failed = 0
    failures_detail = []

    for i, (filepath, clean_name, email) in enumerate(matched_list, 1):
        try:
            create_draft(outlook, email, clean_name, filepath, target_folder)
            created += 1
        except Exception as e:
            failed += 1
            failures_detail.append((clean_name, email, str(e)))

        if progress_callback:
            progress_callback(i, total, clean_name)

    return {
        "created": created,
        "failed": failed,
        "failures_detail": failures_detail,
    }


def main():
    parser = argparse.ArgumentParser(
        description="Create Outlook email drafts for manager attainment reports."
    )
    parser.add_argument(
        "folder",
        help='Folder containing report files (e.g. "Manager report/APAC")'
    )
    parser.add_argument(
        "--dry-run",
        action="store_true",
        default=False,
        help="Show mapping results without creating drafts (default: create drafts)"
    )
    args = parser.parse_args()

    # Resolve folder path
    folder = args.folder
    if not os.path.isabs(folder):
        folder = os.path.join(BASE_DIR, folder)

    if not os.path.isdir(folder):
        print(f"ERROR: Folder not found: {folder}")
        sys.exit(1)

    # Step 1: Load email mapping
    email_map = load_email_mapping(SALES_COMP_FILE)

    # Step 2: Build manager name -> ID mapping
    name_to_info = build_manager_name_to_id(ATTAINMENT_FILE)

    # Step 3: Scan report files
    print(f"\nScanning folder: {folder}")
    report_files = scan_report_files(folder)
    print(f"  Found {len(report_files)} report files")

    if not report_files:
        print("No report files found. Exiting.")
        sys.exit(0)

    # Step 4: Match files to emails
    matched = []
    no_id = []
    no_email = []

    for filepath, mgr_name_from_file in report_files:
        info = name_to_info.get(mgr_name_from_file)
        if not info:
            no_id.append((filepath, mgr_name_from_file))
            continue

        full_name, emp_id = info
        email = email_map.get(emp_id)
        if not email:
            no_email.append((filepath, mgr_name_from_file, emp_id))
            continue

        clean_name = clean_display_name(full_name)
        matched.append((filepath, clean_name, email))

    # Step 5: Print summary
    print(f"\n{'='*60}")
    print(f"MATCHING RESULTS")
    print(f"{'='*60}")
    print(f"  Total files:      {len(report_files)}")
    print(f"  Matched:          {len(matched)}")
    print(f"  No ID match:      {len(no_id)}")
    print(f"  No email found:   {len(no_email)}")

    if no_id:
        print(f"\n  Manager names with no ID match:")
        for fp, name in sorted(no_id, key=lambda x: x[1]):
            print(f"    - {name}  ({os.path.basename(fp)})")

    if no_email:
        print(f"\n  Managers with ID but no email:")
        for fp, name, eid in sorted(no_email, key=lambda x: x[1]):
            print(f"    - {name} (ID: {eid})  ({os.path.basename(fp)})")

    if args.dry_run:
        print(f"\n{'='*60}")
        print("DRY RUN — No drafts created.")
        print(f"{'='*60}")
        if matched:
            print(f"\nSample matches (first 10):")
            for fp, name, email in matched[:10]:
                print(f"  {name:30s} -> {email}")
        return

    # Step 6: Create Outlook drafts
    print(f"\nCreating Outlook drafts...")
    try:
        import win32com.client
        outlook = win32com.client.Dispatch("Outlook.Application")
    except ImportError:
        print("ERROR: pywin32 not installed. Run: pip install pywin32")
        sys.exit(1)
    except Exception as e:
        print(f"ERROR: Could not connect to Outlook: {e}")
        sys.exit(1)

    created = 0
    failed = 0
    for filepath, clean_name, email in matched:
        try:
            create_draft(outlook, email, clean_name, filepath)
            created += 1
            if created % 20 == 0 or created == len(matched):
                print(f"  [{created}/{len(matched)}] Draft created for {clean_name}")
        except Exception as e:
            failed += 1
            print(f"  FAILED: {clean_name} ({email}) — {e}")

    print(f"\n{'='*60}")
    print(f"DONE!")
    print(f"  Drafts created:  {created}")
    print(f"  Failed:          {failed}")
    print(f"  Skipped:         {len(no_id) + len(no_email)}")
    print(f"{'='*60}")
    print(f"\nCheck your Outlook Drafts folder to review and send.")


if __name__ == "__main__":
    main()
