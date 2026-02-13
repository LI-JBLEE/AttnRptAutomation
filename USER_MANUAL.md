# GSC Attainment Report Automator - User Manual

## Overview

The GSC Attainment Report Automator generates per-manager Attainment Reports and distributes them via Outlook email.

**Web App URL**: https://manager-attn-report.streamlit.app/

The tool consists of two components:

| Component | Purpose | Environment |
|-----------|---------|-------------|
| Web App (Streamlit) | Generate reports and download | Any web browser (PC/mobile) |
| EmailManager.exe | Create Outlook drafts and send emails | Windows with Outlook installed |

---

## Part 1: Web App - Report Generation

### Access

Open the following URL in your web browser:

**https://manager-attn-report.streamlit.app/**

### Step 1 — Upload Data Files

Upload two Excel files:

#### Global Attainment Report (Left)
- Excel file (.xlsx) containing attainment data
- Required sheet: `in`
- Required columns: `Level_1_Manager`, `Region`
- After upload, the app validates the file and displays row count, manager count, region count, and fiscal year

#### Sales Compensation Report (Right)
- Excel file (.xlsx) containing employee email mappings
- Required sheet: `Sheet1`
- Required columns: `Employee ID`, `Email - Work`
- Header row must be on row 4 (rows 1-3 are title/blank rows)

> Both files must be uploaded before Step 2 appears.

### Step 2 — Generate Manager Reports

1. **Select Regions**: Choose which regions to generate reports for
   - Default: all regions selected
   - You can deselect regions to generate a subset of reports
2. Click **Generate Reports**
   - A progress bar and status text show the generation progress
   - On completion, the total number of reports and regional distribution are displayed

### Download

After report generation, two download buttons appear:

#### Download Reports (.zip)
- A .zip file containing per-manager Excel reports organized by region folders
- Filename format: `Manager_Reports_FY26_20260213_153045.zip`
- Contents:
  - Region folders > Manager Excel reports
  - `manager_metadata.json`: Manager name, email, region, and file path mappings

#### Download Email Manager (.zip)
- A .zip file containing EmailManager.exe
- Extract the .zip and run the executable

### Reset

Click the **Reset** button in the top-right corner to clear all uploaded files and reset the app to its initial state. Use this when you need to start over with different files.

---

## Part 2: EmailManager.exe - Email Distribution

### Requirements

- Windows 10 / 11
- Microsoft Outlook (installed and configured with a mail account)

### How to Run

1. Extract `EmailManager.zip` downloaded from the web app
2. Double-click `EmailManager.exe` to launch

> **If Windows SmartScreen warning appears:**
> 1. Click "More info"
> 2. Click "Run anyway"
>
> This app runs locally on your PC and does not send data to any external servers.

### Step 1 — Load Reports

Load the .zip file downloaded from the web app:

1. Click **Load .zip File**
2. Select the `Manager_Reports_FY26_YYYYMMDD_HHMMSS.zip` file
3. Once loaded, the manager count and fiscal year are displayed

> Alternatively, click **Load Folder** to load reports from a local folder.
> In this case, you will need to separately select the Sales Compensation Report file.

### Step 2 — Select Recipients

Select which managers should receive emails:

#### Region Filter
- Use region checkboxes to filter managers
- Unchecking a region removes its managers from the list

#### Manager List
- Icon next to each manager: ✓ = email address matched, ✗ = no email found
- Click to select/deselect individual managers
- **Select All**: Select all managers in the list
- **Deselect All**: Clear all selections

> Managers without email addresses (✗) are automatically excluded when creating drafts.

### Step 3 — Email Operations

This section has two tabs:

#### Tab 1: Create Drafts

1. With managers selected in Step 2, click **Create Outlook Drafts**
2. Click "Yes" in the confirmation dialog
3. Drafts are created in Outlook's **Drafts > Manager Report** folder
4. Each draft includes the corresponding manager's Attainment Report as an attachment

> After creating drafts, you can review and edit them in Outlook before sending.

#### Tab 2: Send Drafts

1. Click **Load Drafts** to load the draft list from Outlook
2. Select the drafts you want to send (use Select All / Deselect All as needed)
3. Click **Send Selected**
4. Click "Yes" in the confirmation dialog
5. Selected drafts are sent sequentially

> Sending cannot be undone. Always review draft contents before sending.

### Activity Log

The Activity Log at the bottom of the window shows progress and results for all operations. Detailed error messages are displayed here when issues occur.

---

## End-to-End Workflow

```
1. Open Web App (https://manager-attn-report.streamlit.app/)
   |
2. Upload Attainment Report + Sales Comp Report
   |
3. Select Regions -> Generate Reports
   |
4. Download .zip file + EmailManager.zip
   |
5. Run EmailManager.exe -> Load .zip file
   |
6. Select managers -> Create Outlook Drafts
   |
7. Review drafts in Outlook
   |
8. Send Drafts tab -> Load Drafts -> Send Selected
   |
9. Emails sent!
```

---

## Troubleshooting

### Web App

| Issue | Solution |
|-------|----------|
| Error after file upload | Verify sheet and column names (Attainment: `in` sheet, Sales Comp: `Sheet1` sheet) |
| No regions displayed | Ensure the Attainment file has a `Region` column |
| Need to re-upload files | Click the **Reset** button in the top-right corner |

### EmailManager.exe

| Issue | Solution |
|-------|----------|
| SmartScreen warning on launch | Click "More info" then "Run anyway" |
| .zip load failure | Ensure the .zip was downloaded from the web app and is not corrupted |
| Outlook connection error | Verify Outlook is running and a mail account is configured |
| "Manager Report" folder not found | Run Create Drafts once — the folder is created automatically |
| Send failure (inline response error) | Open the draft in a new Outlook window, save it, then retry |
| Send Selected button disabled | Click Load Drafts first to load the draft list |

---

## Email Format

Each generated email follows this format:

- **Subject**: `FY26 Attainment Report - [Manager Name]`
- **Body**:
  > Hi [Manager Name],
  >
  > Please find attached your FY26 Attainment Report.
  >
  > This report includes attainment data for your team, organized by hierarchy
  > with quarterly, half-year, and annual breakdowns.
  >
  > If you have any questions about the data, please reach out to the
  > Sales Compensation team.
  >
  > Best regards,
  > Sales Compensation
- **Attachment**: The manager's Attainment Report Excel file

---

## FAQ

**Q: Can I use the generated reports without sending emails?**
A: Yes. Download the .zip file and use the Excel reports directly — they are organized by region folders.

**Q: Can I send emails to only specific regions?**
A: Yes. In EmailManager, use the region checkboxes to filter, then select only the managers you want to email.

**Q: Can I edit drafts after creating them?**
A: Yes. Open the drafts in Outlook's Drafts > Manager Report folder and edit them directly. Then use the Send Drafts tab in EmailManager to send them.

**Q: Will it resend emails that were already sent?**
A: No. Already-sent emails are automatically detected and skipped.

**Q: Can I use this on Mac?**
A: The Web App (report generation) works on any OS. EmailManager.exe requires Windows with Outlook installed.

---

## Contact

For issues and feature requests: https://github.com/LI-JBLEE/AttnRptAutomation/issues
