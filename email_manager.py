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


def plain_text_to_html(text):
    """Convert plain text to HTML email body with Calibri font styling."""
    import html as html_module
    escaped = html_module.escape(text)
    # Convert paragraphs (double newlines) to <p> tags, single newlines to <br>
    paragraphs = escaped.split('\n\n')
    html_paragraphs = []
    for p in paragraphs:
        p = p.replace('\n', '<br>\n')
        html_paragraphs.append(f'<p>{p}</p>')
    body = '\n\n'.join(html_paragraphs)
    return (
        '<html>\n<body style="font-family: Calibri, Arial, sans-serif; '
        f'font-size: 11pt; color: #333;">\n{body}\n</body>\n</html>'
    )


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
                 target_folder=None, fiscal_year="FY26",
                 subject_template=None, body_text=None):
    """Create a single Outlook draft email with the report attached.

    Args:
        subject_template: Custom subject with {manager_name}/{fiscal_year} placeholders.
                         If None, uses default template.
        body_text: Custom plain text body with {manager_name}/{fiscal_year} placeholders.
                  If None, uses default HTML template.
    """
    mail = outlook.CreateItem(0)  # olMailItem
    mail.To = to_email

    # Subject
    if subject_template is not None:
        mail.Subject = subject_template.format(
            manager_name=manager_name, fiscal_year=fiscal_year)
    else:
        subj = get_email_subject(fiscal_year)
        mail.Subject = subj.format(manager_name=manager_name)

    # Body
    if body_text is not None:
        filled = body_text.format(
            manager_name=manager_name, fiscal_year=fiscal_year)
        mail.HTMLBody = plain_text_to_html(filled)
    else:
        html_template = get_email_html(fiscal_year)
        mail.HTMLBody = html_template.format(manager_name=manager_name)

    mail.Attachments.Add(os.path.abspath(attachment_path))
    mail.Save()  # Save as Draft — does NOT send

    # Move to target subfolder if specified
    if target_folder is not None:
        mail.Move(target_folder)


def create_drafts_batch(matched_list, target_folder_name="Manager Report",
                        progress_callback=None, fiscal_year="FY26",
                        subject_template=None, body_text=None):
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
            create_draft(outlook, email, clean_name, filepath, target_folder,
                         fiscal_year, subject_template, body_text)
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
                # Store EntryID (persistent identifier) instead of Item object
                # COM objects can become invalid when stored in Python data structures
                draft_items.append({
                    "subject": item.Subject,
                    "to": item.To,
                    "entry_id": item.EntryID,  # Unique persistent identifier
                    "index": j + 1,
                })
            return folder, draft_items

    return None, []


def send_drafts_batch(draft_items, progress_callback=None):
    """Send selected draft emails from a folder.

    Args:
        draft_items: List of draft item dictionaries with 'entry_id'
        progress_callback: Optional callback function(current, total, message)

    Note:
        This function creates its own Outlook COM object to avoid threading issues.
        COM objects must be created and used in the same thread.
    """
    sent = 0
    failed = 0
    failures_detail = []
    total = len(draft_items)

    try:
        # Create new Outlook COM object in this thread
        # This is required because COM objects cannot be marshalled across threads
        import win32com.client
        outlook = win32com.client.Dispatch("Outlook.Application")
        ns = outlook.GetNamespace("MAPI")

        # Process each draft item
        for i, draft_dict in enumerate(draft_items, 1):
            subject = draft_dict.get("subject", "Unknown")
            to_addr = draft_dict.get("to", "Unknown")
            entry_id = draft_dict.get("entry_id")

            try:
                # Retrieve the actual mail item using its EntryID
                # This is the correct way to handle COM objects in Outlook
                item = ns.GetItemFromID(entry_id)

                # Verify it's still a draft (not already sent)
                if item.Sent:
                    failed += 1
                    failures_detail.append((subject, "Email already sent"))
                    if progress_callback:
                        progress_callback(i, total, f"SKIPPED: {subject} (already sent)")
                    continue

                # Send the email
                item.Send()
                sent += 1

                if progress_callback:
                    progress_callback(i, total, f"{subject} → {to_addr}")

            except Exception as e:
                error_str = str(e)
                failed += 1

                # Provide user-friendly error messages for common issues
                if "inline response" in error_str.lower():
                    friendly_error = "Cannot send inline reply draft - open draft in new window and save, then try again"
                elif "4096" in error_str and "Microsoft Outlook" in error_str:
                    friendly_error = "Outlook-specific error - draft may be in reply/forward mode"
                else:
                    friendly_error = str(e)

                failures_detail.append((subject, friendly_error))

                if progress_callback:
                    progress_callback(i, total, f"FAILED: {subject}")

    except Exception as e:
        # Failed to initialize Outlook
        failed = total
        failures_detail.append(("Outlook Connection", str(e)))

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
        self.root.geometry("900x950")
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
            row=0, column=2
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

        ttk.Button(mgr_header, text="✅ Select All", command=self._select_all_managers, width=15).grid(
            row=0, column=1, padx=5
        )

        ttk.Button(mgr_header, text="❌ Deselect All", command=self._deselect_all_managers, width=15).grid(
            row=0, column=2
        )

        # Manager listbox with scrollbar
        listbox_frame = ttk.Frame(manager_frame)
        listbox_frame.grid(row=1, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), pady=5)
        listbox_frame.columnconfigure(0, weight=1)
        listbox_frame.rowconfigure(0, weight=1)

        scrollbar = ttk.Scrollbar(listbox_frame, orient=tk.VERTICAL)
        self.manager_listbox = tk.Listbox(
            listbox_frame, selectmode=tk.MULTIPLE, height=10,
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
        draft_tab.columnconfigure(0, weight=1)

        # Email Subject
        ttk.Label(draft_tab, text="Email Subject:", font=("Segoe UI", 9, "bold")).grid(
            row=0, column=0, sticky=tk.W, pady=(0, 3)
        )

        self.subject_var = tk.StringVar(value=self._get_default_subject())
        self.subject_entry = ttk.Entry(draft_tab, textvariable=self.subject_var)
        self.subject_entry.grid(row=1, column=0, sticky=(tk.W, tk.E), pady=(0, 8))

        # Email Body
        ttk.Label(draft_tab, text="Email Body:", font=("Segoe UI", 9, "bold")).grid(
            row=2, column=0, sticky=tk.W, pady=(0, 3)
        )

        self.body_text = tk.Text(draft_tab, height=8, wrap=tk.WORD,
                                 font=("Segoe UI", 9))
        self.body_text.grid(row=3, column=0, sticky=(tk.W, tk.E), pady=(0, 5))
        self.body_text.insert("1.0", self._get_default_body())

        # Available variables hint
        ttk.Label(draft_tab, text="Available variables: {manager_name}, {fiscal_year}",
                  foreground="gray", font=("Segoe UI", 8)).grid(
            row=4, column=0, sticky=tk.W, pady=(0, 8)
        )

        # Button row
        btn_frame = ttk.Frame(draft_tab)
        btn_frame.grid(row=5, column=0, sticky=(tk.W, tk.E))

        ttk.Button(btn_frame, text="↩ Reset to Default",
                   command=self._reset_template).grid(row=0, column=0, padx=(0, 10))

        self.create_drafts_btn = ttk.Button(
            btn_frame,
            text="📧 Create Outlook Drafts",
            command=self._create_drafts,
            state=tk.DISABLED
        )
        self.create_drafts_btn.grid(row=0, column=1)

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
        label.grid(row=row, column=0, sticky=tk.W, pady=(5, 3))

    def _log(self, message):
        """Add message to activity log."""
        self.log_text.configure(state=tk.NORMAL)
        timestamp = datetime.now().strftime("%H:%M:%S")
        self.log_text.insert(tk.END, f"[{timestamp}] {message}\n")
        self.log_text.see(tk.END)
        self.log_text.configure(state=tk.DISABLED)

    def _get_default_subject(self):
        """Return default email subject template."""
        return "{fiscal_year} Attainment Report - {manager_name}"

    def _get_default_body(self):
        """Return default email body as plain text."""
        return (
            "Hi {manager_name},\n"
            "\n"
            "Please find attached your {fiscal_year} Attainment Report.\n"
            "\n"
            "This report includes attainment data for your team, organized by hierarchy "
            "with quarterly, half-year, and annual breakdowns.\n"
            "\n"
            "If you have any questions about the data, please reach out to the "
            "Sales Compensation team.\n"
            "\n"
            "Best regards,\n"
            "Sales Compensation"
        )

    def _reset_template(self):
        """Reset subject and body fields to defaults."""
        self.subject_var.set(self._get_default_subject())
        self.body_text.delete("1.0", tk.END)
        self.body_text.insert("1.0", self._get_default_body())


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

        # Read custom templates from UI
        subject_template = self.subject_var.get().strip()
        body_text = self.body_text.get("1.0", tk.END).rstrip()

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
                    fiscal_year=self.fiscal_year,
                    subject_template=subject_template,
                    body_text=body_text
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

        if not self.outlook or not self.draft_folder:
            messagebox.showerror("Error", "Outlook is not loaded. Please click 'Load Drafts' first.")
            return

        # Get the actual draft item dictionaries for selected items
        selected_draft_items = [self.draft_items[i] for i in selected_indices_ui]

        # Confirmation dialog
        response = messagebox.askyesno(
            "Confirm Send",
            f"⚠️ You are about to send {len(selected_draft_items)} email(s).\n\n"
            "This action cannot be undone. Continue?"
        )

        if not response:
            return

        self.send_drafts_btn.configure(state=tk.DISABLED)
        self.operation_status_var.set(f"Sending {len(selected_draft_items)} emails...")

        def run_sending():
            try:
                def progress_callback(current, total, subject):
                    self.root.after(0, lambda s=subject: self._log(f"[{current}/{total}] Sent: {s}"))

                # send_drafts_batch creates its own Outlook COM object in this thread
                # to avoid "marshalled for a different thread" errors
                results = send_drafts_batch(
                    selected_draft_items,
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
    # Close PyInstaller splash screen if running from bundled .exe
    try:
        import pyi_splash  # noqa: F401 — only available in PyInstaller bundle
        pyi_splash.close()
    except ImportError:
        pass

    root = tk.Tk()
    app = EmailManagerApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
