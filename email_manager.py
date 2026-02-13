"""
GSC Email Manager - Standalone Email Operations Tool
Loads report .zip files from web app and manages Outlook email operations.

This is a self-contained executable - all dependencies are embedded.
"""

import os
import sys
import re
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import threading
import zipfile
import json
import tempfile
from datetime import datetime
import pandas as pd


# ══════════════════════════════════════════════════════════════════════════════
# Embedded functions from create_email_drafts.py (for standalone .exe)
# ══════════════════════════════════════════════════════════════════════════════

def clean_display_name(full_name):
    """Clean manager name for email display: remove ID and any parenthesized content."""
    # Remove trailing (ID)
    name = re.sub(r'\s*\(\d+\)\s*$', '', full_name).strip()
    # Remove remaining parenthesized aliases
    name = re.sub(r'\s*\([^)]*\)', '', name).strip()
    return name


def get_email_subject(fiscal_year="FY26"):
    """Generate email subject with fiscal year."""
    return f"{fiscal_year} Attainment Report - {{manager_name}}"


def get_email_html(fiscal_year="FY26"):
    """Generate email HTML body with fiscal year."""
    return f"""\
<html>
<body style="font-family: Calibri, Arial, sans-serif; font-size: 11pt; color: #333;">
<p>Hi {{manager_name}},</p>

<p>Please find attached your <b>{fiscal_year} Attainment Report</b>.</p>

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
    df = pd.read_excel(sales_comp_file, sheet_name="Sheet1", header=3)

    email_map = {}
    for _, row in df.iterrows():
        emp_id = row.get("Employee ID")
        email = row.get("Email - Work")
        if pd.notna(emp_id) and pd.notna(email):
            clean_id = str(emp_id).lstrip("0") or "0"
            email_map[clean_id] = str(email).strip()

    return email_map


def get_or_create_drafts_subfolder(outlook, folder_name="Manager Report"):
    """Get or create a subfolder under the Outlook Drafts folder."""
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
                 target_folder=None, fiscal_year="FY26"):
    """Create a single Outlook draft email with the report attached."""
    mail = outlook.CreateItem(0)  # olMailItem
    mail.To = to_email
    subject_template = get_email_subject(fiscal_year)
    html_template = get_email_html(fiscal_year)
    mail.Subject = subject_template.format(manager_name=manager_name)
    mail.HTMLBody = html_template.format(manager_name=manager_name)
    mail.Attachments.Add(os.path.abspath(attachment_path))
    mail.Save()  # Save as Draft — does NOT send

    # Move to target subfolder if specified
    if target_folder is not None:
        mail.Move(target_folder)


def create_drafts_batch(matched_list, target_folder_name="Manager Report",
                        progress_callback=None, fiscal_year="FY26"):
    """Create Outlook drafts for a list of matched managers."""
    import win32com.client
    outlook = win32com.client.Dispatch("Outlook.Application")
    target_folder = get_or_create_drafts_subfolder(outlook, target_folder_name)

    total = len(matched_list)
    created = 0
    failed = 0
    failures_detail = []

    for i, (filepath, clean_name, email) in enumerate(matched_list, 1):
        try:
            create_draft(outlook, email, clean_name, filepath, target_folder, fiscal_year)
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


def get_drafts_from_folder(outlook, folder_name="Manager Report"):
    """Get all draft emails from Outlook Drafts/folder_name subfolder."""
    ns = outlook.GetNamespace("MAPI")
    drafts = ns.GetDefaultFolder(16)  # olFolderDrafts

    # Find the subfolder
    for i in range(drafts.Folders.Count):
        folder = drafts.Folders.Item(i + 1)  # 1-indexed
        if folder.Name == folder_name:
            draft_items = []
            for j in range(folder.Items.Count):
                item = folder.Items.Item(j + 1)
                draft_items.append({
                    "subject": item.Subject,
                    "to": item.To,
                    "index": j + 1,
                    "item": item,
                })
            return folder, draft_items

    return None, []


def send_drafts_batch(outlook, folder, selected_indices, progress_callback=None):
    """Send selected draft emails from a folder."""
    sent = 0
    failed = 0
    failures_detail = []
    total = len(selected_indices)

    for i, idx in enumerate(selected_indices, 1):
        try:
            item = folder.Items.Item(idx)
            subject = item.Subject
            item.Send()  # Send the email immediately
            sent += 1

            if progress_callback:
                progress_callback(i, total, subject)
        except Exception as e:
            failed += 1
            failures_detail.append((subject if 'subject' in locals() else f"Index {idx}", str(e)))

            if progress_callback:
                progress_callback(i, total, f"FAILED: {subject if 'subject' in locals() else f'Index {idx}'}")

    return {
        "sent": sent,
        "failed": failed,
        "failures_detail": failures_detail,
    }


# ══════════════════════════════════════════════════════════════════════════════
# EmailManager GUI Application
# ══════════════════════════════════════════════════════════════════════════════


class EmailManagerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("GSC Email Manager")
        self.root.geometry("900x750")
        self.root.resizable(True, True)

        # Data storage
        self.metadata = None
        self.temp_dir = None
        self.selected_managers = []
        self.fiscal_year = "FY26"

        # Outlook COM objects
        self.outlook = None
        self.draft_folder = None
        self.draft_items = []

        self._create_widgets()

    def _create_widgets(self):
        """Create all GUI widgets."""
        # Main container with padding
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(0, weight=1)

        current_row = 0

        # ── Title ──
        title_label = ttk.Label(
            main_frame,
            text="GSC Email Manager",
            font=("Segoe UI", 16, "bold"),
            foreground="#0078D4"
        )
        title_label.grid(row=current_row, column=0, pady=(0, 20), sticky=tk.W)
        current_row += 1

        # ── Step 1: Load Reports ──
        self._add_section_header(main_frame, current_row, "Step 1 — Load Reports")
        current_row += 1

        load_frame = ttk.Frame(main_frame)
        load_frame.grid(row=current_row, column=0, sticky=(tk.W, tk.E), pady=5)
        load_frame.columnconfigure(1, weight=1)

        ttk.Label(load_frame, text="Load from:").grid(row=0, column=0, sticky=tk.W, padx=(0, 10))

        self.load_path_var = tk.StringVar()
        ttk.Entry(load_frame, textvariable=self.load_path_var, state="readonly").grid(
            row=0, column=1, sticky=(tk.W, tk.E), padx=(0, 10)
        )

        ttk.Button(load_frame, text="📂 Load .zip File", command=self._load_zip_file).grid(
            row=0, column=2, padx=(0, 5)
        )

        ttk.Button(load_frame, text="📁 Load Folder", command=self._load_folder).grid(
            row=0, column=3
        )
        current_row += 1

        # Load status
        self.load_status_var = tk.StringVar(value="No reports loaded")
        ttk.Label(main_frame, textvariable=self.load_status_var, foreground="gray").grid(
            row=current_row, column=0, sticky=tk.W, pady=(5, 15)
        )
        current_row += 1

        # ── Step 2: Filter & Select Recipients ──
        self._add_section_header(main_frame, current_row, "Step 2 — Select Recipients")
        current_row += 1

        # Region filter frame
        self.region_filter_frame = ttk.LabelFrame(main_frame, text="Filter by Region", padding="10")
        self.region_filter_frame.grid(row=current_row, column=0, sticky=(tk.W, tk.E), pady=5)
        self.region_checkboxes = {}
        current_row += 1

        # Manager list frame
        manager_frame = ttk.Frame(main_frame)
        manager_frame.grid(row=current_row, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), pady=5)
        manager_frame.columnconfigure(0, weight=1)
        manager_frame.rowconfigure(1, weight=1)
        main_frame.rowconfigure(current_row, weight=1)

        # Manager list header with Select All / Deselect All
        mgr_header = ttk.Frame(manager_frame)
        mgr_header.grid(row=0, column=0, sticky=(tk.W, tk.E))

        ttk.Label(mgr_header, text="Managers:", font=("Segoe UI", 10, "bold")).grid(
            row=0, column=0, sticky=tk.W
        )

        ttk.Button(mgr_header, text="✅ Select All", command=self._select_all_managers, width=12).grid(
            row=0, column=1, padx=5
        )

        ttk.Button(mgr_header, text="❌ Deselect All", command=self._deselect_all_managers, width=12).grid(
            row=0, column=2
        )

        # Manager listbox with scrollbar
        listbox_frame = ttk.Frame(manager_frame)
        listbox_frame.grid(row=1, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), pady=5)
        listbox_frame.columnconfigure(0, weight=1)
        listbox_frame.rowconfigure(0, weight=1)

        scrollbar = ttk.Scrollbar(listbox_frame, orient=tk.VERTICAL)
        self.manager_listbox = tk.Listbox(
            listbox_frame, selectmode=tk.MULTIPLE, height=20,
            yscrollcommand=scrollbar.set
        )
        scrollbar.config(command=self.manager_listbox.yview)

        self.manager_listbox.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))

        # Selection summary
        self.selection_summary_var = tk.StringVar(value="No managers selected")
        ttk.Label(manager_frame, textvariable=self.selection_summary_var, foreground="gray").grid(
            row=2, column=0, sticky=tk.W, pady=5
        )
        current_row += 1

        # ── Step 3: Email Operations ──
        self._add_section_header(main_frame, current_row, "Step 3 — Email Operations")
        current_row += 1

        # Tab control for Draft Creation vs Sending
        self.notebook = ttk.Notebook(main_frame)
        self.notebook.grid(row=current_row, column=0, sticky=(tk.W, tk.E), pady=5)

        # Tab 1: Create Drafts
        draft_tab = ttk.Frame(self.notebook, padding="10")
        self.notebook.add(draft_tab, text="Create Drafts")

        ttk.Label(draft_tab, text="Create Outlook email drafts for selected managers:").grid(
            row=0, column=0, sticky=tk.W, pady=(0, 10)
        )

        self.create_drafts_btn = ttk.Button(
            draft_tab,
            text="📧 Create Outlook Drafts",
            command=self._create_drafts,
            state=tk.DISABLED
        )
        self.create_drafts_btn.grid(row=1, column=0, sticky=tk.W)

        # Tab 2: Send Drafts
        send_tab = ttk.Frame(self.notebook, padding="10")
        self.notebook.add(send_tab, text="Send Drafts")

        send_controls = ttk.Frame(send_tab)
        send_controls.grid(row=0, column=0, sticky=(tk.W, tk.E), pady=(0, 10))

        ttk.Button(send_controls, text="🔄 Load Drafts", command=self._load_outlook_drafts).grid(
            row=0, column=0, padx=(0, 10)
        )

        ttk.Button(send_controls, text="✅ Select All", command=self._select_all_drafts).grid(
            row=0, column=1, padx=(0, 10)
        )

        ttk.Button(send_controls, text="❌ Deselect All", command=self._deselect_all_drafts).grid(
            row=0, column=2, padx=(0, 10)
        )

        self.send_drafts_btn = ttk.Button(
            send_controls,
            text="✉️ Send Selected",
            command=self._send_drafts,
            state=tk.DISABLED
        )
        self.send_drafts_btn.grid(row=0, column=3)

        # Draft listbox
        draft_list_frame = ttk.Frame(send_tab)
        draft_list_frame.grid(row=1, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        draft_list_frame.columnconfigure(0, weight=1)
        draft_list_frame.rowconfigure(0, weight=1)

        draft_scrollbar = ttk.Scrollbar(draft_list_frame, orient=tk.VERTICAL)
        self.draft_listbox = tk.Listbox(
            draft_list_frame, selectmode=tk.MULTIPLE, height=8,
            yscrollcommand=draft_scrollbar.set
        )
        draft_scrollbar.config(command=self.draft_listbox.yview)

        self.draft_listbox.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        draft_scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))

        send_tab.rowconfigure(1, weight=1)

        current_row += 1

        # Operation status
        self.operation_status_var = tk.StringVar(value="")
        ttk.Label(main_frame, textvariable=self.operation_status_var, foreground="gray").grid(
            row=current_row, column=0, sticky=tk.W, pady=(5, 10)
        )
        current_row += 1

        # ── Progress Log ──
        self._add_section_header(main_frame, current_row, "Activity Log")
        current_row += 1

        self.log_text = scrolledtext.ScrolledText(main_frame, height=8, state=tk.DISABLED)
        self.log_text.grid(row=current_row, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), pady=5)
        main_frame.rowconfigure(current_row, weight=1)

    def _add_section_header(self, parent, row, text):
        """Add a styled section header."""
        label = ttk.Label(
            parent,
            text=text,
            font=("Segoe UI", 11, "bold"),
            foreground="#0078D4"
        )
        label.grid(row=row, column=0, sticky=tk.W, pady=(10, 5))

    def _log(self, message):
        """Add message to activity log."""
        self.log_text.configure(state=tk.NORMAL)
        timestamp = datetime.now().strftime("%H:%M:%S")
        self.log_text.insert(tk.END, f"[{timestamp}] {message}\n")
        self.log_text.see(tk.END)
        self.log_text.configure(state=tk.DISABLED)

    def _load_zip_file(self):
        """Load reports from .zip file (downloaded from web app)."""
        zip_path = filedialog.askopenfilename(
            title="Select Manager Reports .zip File",
            filetypes=[("ZIP files", "*.zip"), ("All files", "*.*")]
        )

        if not zip_path:
            return

        try:
            self._log(f"Loading .zip file: {os.path.basename(zip_path)}")

            # Extract to temp directory
            self.temp_dir = tempfile.mkdtemp()
            with zipfile.ZipFile(zip_path, 'r') as z:
                z.extractall(self.temp_dir)
                self._log(f"Extracted {len(z.namelist())} files")

            # Load metadata
            metadata_path = os.path.join(self.temp_dir, "manager_metadata.json")
            if not os.path.exists(metadata_path):
                raise FileNotFoundError("manager_metadata.json not found in .zip file")

            with open(metadata_path, 'r') as f:
                self.metadata = json.load(f)

            # Update file paths to point to temp directory
            for mgr in self.metadata["managers"]:
                mgr["filepath"] = os.path.join(self.temp_dir, mgr["filepath"])

            self.fiscal_year = self.metadata.get("fiscal_year", "FY26")

            self.load_path_var.set(zip_path)
            self.load_status_var.set(
                f"✓ Loaded {self.metadata['total_reports']} reports ({self.fiscal_year})"
            )
            self._log(f"✓ Loaded {self.metadata['total_reports']} manager reports")

            self._update_region_checkboxes()
            self._update_manager_list()
            self.create_drafts_btn.configure(state=tk.NORMAL)

        except Exception as e:
            messagebox.showerror("Error", f"Failed to load .zip file:\n{e}")
            self._log(f"✗ Error loading .zip: {e}")

    def _load_folder(self):
        """Load reports from folder (for locally generated reports)."""
        folder_path = filedialog.askdirectory(title="Select Manager Reports Folder")

        if not folder_path:
            return

        # Prompt for Sales Comp file for email mapping
        sales_comp_path = filedialog.askopenfilename(
            title="Select Sales Compensation Report (for email addresses)",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]
        )

        if not sales_comp_path:
            return

        try:
            self._log(f"Loading reports from folder: {folder_path}")
            email_map = load_email_mapping(sales_comp_path)
            self._log(f"Loaded {len(email_map)} email addresses")

            # Scan folder for report files
            managers = []
            for root, dirs, files in os.walk(folder_path):
                region = os.path.basename(root)
                for fname in files:
                    if fname.startswith("FY") and "Attainment_" in fname and fname.endswith(".xlsx"):
                        # Extract manager name from filename
                        # FY26_Attainment_John_Smith_20260213.xlsx
                        try:
                            parts = fname.replace(".xlsx", "").split("_")
                            fiscal_year = parts[0]
                            # Find "Attainment" index
                            attn_idx = parts.index("Attainment")
                            # Name is between Attainment and date (last part)
                            safe_name = "_".join(parts[attn_idx+1:-1])

                            filepath = os.path.join(root, fname)

                            managers.append({
                                "name": safe_name.replace("_", " "),
                                "safe_name": safe_name,
                                "region": region,
                                "email": None,  # TODO: match from filename to ID
                                "filepath": filepath
                            })

                            self.fiscal_year = fiscal_year
                        except Exception as e:
                            self._log(f"Skipped {fname}: {e}")

            if not managers:
                raise ValueError("No report files found in folder")

            self.metadata = {
                "fiscal_year": self.fiscal_year,
                "generated_date": datetime.today().strftime("%Y-%m-%d"),
                "total_reports": len(managers),
                "managers": managers
            }

            self.load_path_var.set(folder_path)
            self.load_status_var.set(f"✓ Loaded {len(managers)} reports ({self.fiscal_year})")
            self._log(f"✓ Loaded {len(managers)} manager reports")

            self._update_region_checkboxes()
            self._update_manager_list()
            self.create_drafts_btn.configure(state=tk.NORMAL)

        except Exception as e:
            messagebox.showerror("Error", f"Failed to load folder:\n{e}")
            self._log(f"✗ Error loading folder: {e}")

    def _update_region_checkboxes(self):
        """Update region filter checkboxes based on loaded data."""
        # Clear existing checkboxes
        for widget in self.region_filter_frame.winfo_children():
            widget.destroy()
        self.region_checkboxes.clear()

        if not self.metadata:
            return

        # Get unique regions
        regions = sorted(set(m["region"] for m in self.metadata["managers"]))

        # Create checkboxes (default: all selected)
        for i, region in enumerate(regions):
            var = tk.BooleanVar(value=True)
            cb = ttk.Checkbutton(
                self.region_filter_frame,
                text=region,
                variable=var,
                command=self._update_manager_list
            )
            cb.grid(row=i // 4, column=i % 4, sticky=tk.W, padx=10, pady=2)
            self.region_checkboxes[region] = var

    def _update_manager_list(self):
        """Update manager listbox based on region filter."""
        if not self.metadata:
            return

        # Get selected regions
        selected_regions = [
            region for region, var in self.region_checkboxes.items() if var.get()
        ]

        # Filter managers
        filtered_managers = [
            m for m in self.metadata["managers"]
            if m["region"] in selected_regions
        ]

        # Update listbox
        self.manager_listbox.delete(0, tk.END)
        for mgr in filtered_managers:
            email_status = "✓" if mgr.get("email") else "✗"
            self.manager_listbox.insert(
                tk.END,
                f"{email_status} {mgr['name']} ({mgr['region']})"
            )

        # Select all by default
        self.manager_listbox.selection_set(0, tk.END)

        self._update_selection_summary()

    def _update_selection_summary(self):
        """Update selection summary label."""
        if not self.metadata:
            return

        selected_indices = self.manager_listbox.curselection()

        # Get selected regions
        selected_regions = [
            region for region, var in self.region_checkboxes.items() if var.get()
        ]

        # Filter managers
        filtered_managers = [
            m for m in self.metadata["managers"]
            if m["region"] in selected_regions
        ]

        selected_managers = [filtered_managers[i] for i in selected_indices]

        total = len(selected_managers)
        with_email = len([m for m in selected_managers if m.get("email")])
        without_email = total - with_email

        self.selection_summary_var.set(
            f"Selected: {total} managers | ✓ Email: {with_email} | ✗ No Email: {without_email}"
        )

    def _select_all_managers(self):
        """Select all managers in the listbox."""
        self.manager_listbox.selection_set(0, tk.END)
        self._update_selection_summary()

    def _deselect_all_managers(self):
        """Deselect all managers in the listbox."""
        self.manager_listbox.selection_clear(0, tk.END)
        self._update_selection_summary()

    def _create_drafts(self):
        """Create Outlook email drafts for selected managers."""
        selected_indices = self.manager_listbox.curselection()

        if not selected_indices:
            messagebox.showwarning("Warning", "Please select at least one manager.")
            return

        # Get selected regions
        selected_regions = [
            region for region, var in self.region_checkboxes.items() if var.get()
        ]

        # Filter managers
        filtered_managers = [
            m for m in self.metadata["managers"]
            if m["region"] in selected_regions
        ]

        selected_managers = [filtered_managers[i] for i in selected_indices]

        # Filter out managers without email
        with_email = [m for m in selected_managers if m.get("email")]

        if not with_email:
            messagebox.showwarning("Warning", "No selected managers have email addresses.")
            return

        # Confirm
        response = messagebox.askyesno(
            "Confirm",
            f"Create {len(with_email)} email draft(s) in Outlook?\n\n"
            f"Drafts will be saved in: Outlook > Drafts > Manager Report"
        )

        if not response:
            return

        # Disable button during creation
        self.create_drafts_btn.configure(state=tk.DISABLED)
        self.operation_status_var.set(f"Creating {len(with_email)} drafts...")

        def run_creation():
            try:
                matched_list = [
                    (m["filepath"], m["name"], m["email"]) for m in with_email
                ]

                def progress_callback(current, total, message):
                    self.root.after(0, lambda: self._log(f"[{current}/{total}] Draft: {message}"))

                results = create_drafts_batch(
                    matched_list,
                    target_folder_name="Manager Report",
                    progress_callback=progress_callback,
                    fiscal_year=self.fiscal_year
                )

                self.root.after(0, lambda: self.operation_status_var.set(
                    f"✓ Created {results['created']} drafts, {results['failed']} failed"
                ))
                self.root.after(0, lambda: self._log(f"✓ Done! {results['created']} drafts created"))

                if results["failed"] > 0:
                    error_msg = "\n".join(
                        f"- {name} ({email}): {err}"
                        for name, email, err in results["failures_detail"]
                    )
                    self.root.after(0, lambda: self._log(f"✗ Failures:\n{error_msg}"))

            except Exception as e:
                self.root.after(0, lambda: messagebox.showerror("Error", f"Draft creation failed:\n{e}"))
                self.root.after(0, lambda: self._log(f"✗ Error: {e}"))
            finally:
                self.root.after(0, lambda: self.create_drafts_btn.configure(state=tk.NORMAL))
                self.root.after(0, lambda: self.operation_status_var.set(""))

        threading.Thread(target=run_creation, daemon=True).start()

    def _load_outlook_drafts(self):
        """Load drafts from Outlook Drafts/Manager Report folder."""
        try:
            import win32com.client
            self.outlook = win32com.client.Dispatch("Outlook.Application")
            self.draft_folder, self.draft_items = get_drafts_from_folder(self.outlook, "Manager Report")

            if self.draft_folder is None:
                messagebox.showwarning("Warning", "No 'Manager Report' folder found in Outlook Drafts.")
                self.draft_items = []
                self.send_drafts_btn.configure(state=tk.DISABLED)
                return

            # Populate draft listbox
            self.draft_listbox.delete(0, tk.END)
            for draft in self.draft_items:
                self.draft_listbox.insert(tk.END, f"{draft['subject']} → {draft['to']}")

            self._log(f"✓ Loaded {len(self.draft_items)} drafts from Outlook")

            if self.draft_items:
                self.send_drafts_btn.configure(state=tk.NORMAL)

        except ImportError:
            messagebox.showerror("Error", "pywin32 is not installed. Run: pip install pywin32")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load drafts:\n{e}")
            self._log(f"✗ Error loading drafts: {e}")

    def _select_all_drafts(self):
        """Select all drafts in the listbox."""
        self.draft_listbox.selection_set(0, tk.END)

    def _deselect_all_drafts(self):
        """Deselect all drafts in the listbox."""
        self.draft_listbox.selection_clear(0, tk.END)

    def _send_drafts(self):
        """Send selected email drafts."""
        selected_indices_ui = self.draft_listbox.curselection()

        if not selected_indices_ui:
            messagebox.showwarning("Warning", "Please select at least one draft to send.")
            return

        # Map UI indices (0-based) to Outlook indices (1-based)
        selected_indices_outlook = [self.draft_items[i]["index"] for i in selected_indices_ui]

        # Confirmation dialog
        response = messagebox.askyesno(
            "Confirm Send",
            f"⚠️ You are about to send {len(selected_indices_outlook)} email(s).\n\n"
            "This action cannot be undone. Continue?"
        )

        if not response:
            return

        self.send_drafts_btn.configure(state=tk.DISABLED)
        self.operation_status_var.set(f"Sending {len(selected_indices_outlook)} emails...")

        def run_sending():
            try:
                def progress_callback(current, total, subject):
                    self.root.after(0, lambda: self._log(f"[{current}/{total}] Sent: {subject}"))

                results = send_drafts_batch(
                    self.outlook,
                    self.draft_folder,
                    selected_indices_outlook,
                    progress_callback=progress_callback
                )

                self.root.after(0, lambda: self.operation_status_var.set(
                    f"✓ Sent {results['sent']} emails, {results['failed']} failed"
                ))
                self.root.after(0, lambda: self._log(f"✓ Done! {results['sent']} emails sent"))

                if results["failed"] > 0:
                    error_msg = "\n".join(
                        f"- {subject}: {err}"
                        for subject, err in results["failures_detail"]
                    )
                    self.root.after(0, lambda: self._log(f"✗ Failures:\n{error_msg}"))

                # Clear draft list to force reload
                self.root.after(0, lambda: self.draft_listbox.delete(0, tk.END))
                self.draft_items = []

            except Exception as e:
                self.root.after(0, lambda: messagebox.showerror("Error", f"Email sending failed:\n{e}"))
                self.root.after(0, lambda: self._log(f"✗ Error: {e}"))
            finally:
                self.root.after(0, lambda: self.send_drafts_btn.configure(state=tk.DISABLED))
                self.root.after(0, lambda: self.operation_status_var.set(""))

        threading.Thread(target=run_sending, daemon=True).start()


def main():
    root = tk.Tk()
    app = EmailManagerApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
