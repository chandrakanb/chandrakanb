import tkinter as tk
from tkinter import ttk, filedialog, scrolledtext, messagebox
import pandas as pd
import threading
from jira import JIRA, exceptions
import requests
from requests.auth import HTTPBasicAuth
import json
import os
from datetime import datetime

CREDENTIALS_FILE = "credentials.json"
EXCEL_CONFIG_FILE = "excel_config.json"
IMPORTANT_LOGS_FILE = "logs/important_logs.json"

class JiraUpdaterApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Jira Updater")
        self.root.state("zoomed")
        self.jira_connection = None
        self.cancel_update_flag = False
        self.log_file_path = None

        # Notebook tabs
        self.notebook = ttk.Notebook(root)
        self.notebook.pack(expand=True, fill="both")

        self.connection_tab = ttk.Frame(self.notebook)
        self.update_tab = ttk.Frame(self.notebook)
        self.logs_tab = ttk.Frame(self.notebook)

        self.notebook.add(self.connection_tab, text="Connect")
        self.notebook.add(self.update_tab, text="Update Tickets")
        self.notebook.add(self.logs_tab, text="Logs")

        # Connection status LED next to tabs
        self.status_frame = tk.Frame(root)
        self.status_frame.place(relx=1.0, y=5, anchor="ne")  # top-right corner
        self.status_canvas = tk.Canvas(self.status_frame, width=10, height=10, highlightthickness=0)
        self.status_canvas.pack(side="left")
        self.status_label = tk.Label(self.status_frame, text="Not Connected")
        self.status_label.pack(side="left", padx=2)

        # Logs Filter Flags
        self.show_important_only = tk.BooleanVar(value=False)
        self.view_mode = tk.StringVar(value="active")  

        # Create UI
        self.create_connection_ui()
        self.create_update_ui()
        self.create_logs_ui()
        self.load_credentials()
        self.load_excel_config()

    # ==================== Connection Tab ====================
    def create_connection_ui(self):
        frame = ttk.LabelFrame(self.connection_tab)
        frame.place(relx=0.5, rely=0.1, anchor="n")

        # Centered title
        title = ttk.Label(frame, text="Jira Login", anchor="center")
        frame.configure(labelwidget=title, borderwidth=2, relief="solid")

        # Make columns adjustable
        frame.columnconfigure(0, weight=0)
        frame.columnconfigure(1, weight=1)

        # Labels and Entries
        ttk.Label(frame, text="Server URL   :").grid(row=0, column=0, sticky="w")
        self.server_url_entry = ttk.Entry(frame, width=50)
        self.server_url_entry.grid(row=0, column=1, sticky="ew")

        ttk.Label(frame, text="Email ID       :").grid(row=1, column=0, sticky="w")
        self.email_entry = ttk.Entry(frame, width=50)
        self.email_entry.grid(row=1, column=1, sticky="ew")

        ttk.Label(frame, text="API Token    :").grid(row=2, column=0, sticky="w")
        self.api_token_entry = ttk.Entry(frame, width=50, show="*")
        self.api_token_entry.grid(row=2, column=1, sticky="ew")

        # Checkbox
        self.save_credentials_checkbox = tk.BooleanVar()
        ttk.Checkbutton(frame, text="Save Credentials", variable=self.save_credentials_checkbox)\
            .grid(row=3, column=0, columnspan=2, pady=5)

        # Buttons
        ttk.Button(frame, text="Connect", command=self.connect_to_jira)\
            .grid(row=4, column=0, columnspan=2, pady=10)
        ttk.Button(frame, text="Clear Credentials", command=self.clear_credentials)\
            .grid(row=5, column=0, columnspan=2, pady=5)

        # Optional: Focus first entry
        self.server_url_entry.focus()

        self.update_connection_status()

    def save_credentials(self, server_url, email, api_token):
        with open(CREDENTIALS_FILE, "w") as f:
            json.dump({"server_url": server_url, "email": email, "api_token": api_token}, f, indent=2)

    def load_credentials(self):
        if os.path.isfile(CREDENTIALS_FILE):
            try:
                with open(CREDENTIALS_FILE, "r") as f:
                    data = json.load(f)
                    self.server_url_entry.delete(0, tk.END)
                    self.server_url_entry.insert(0, data.get("server_url", ""))
                    self.email_entry.delete(0, tk.END)
                    self.email_entry.insert(0, data.get("email", ""))
                    self.api_token_entry.delete(0, tk.END)
                    self.api_token_entry.insert(0, data.get("api_token", ""))
                    self.save_credentials_checkbox.set(True)
                    self.log(f"✅ Loaded saved credentials from {CREDENTIALS_FILE}")
            except Exception as e:
                self.log(f"❌ Failed to load credentials: {e}")

    def clear_credentials(self):
        self.server_url_entry.delete(0, tk.END)
        self.email_entry.delete(0, tk.END)
        self.api_token_entry.delete(0, tk.END)
        self.save_credentials_checkbox.set(False)
        if os.path.exists(CREDENTIALS_FILE):
            os.remove(CREDENTIALS_FILE)
        self.update_connection_status()
        messagebox.showinfo("Success", "Credentials cleared!")

    def connect_to_jira(self):
        try:
            server_url = self.server_url_entry.get().strip()
            email = self.email_entry.get().strip()
            api_token = self.api_token_entry.get().strip()
            if not server_url or not email or not api_token:
                raise ValueError("All fields (Server URL, Email, API Token) are required")

            self.jira_connection = JIRA(server=server_url, basic_auth=(email, api_token))
            messagebox.showinfo("Success", "✅ Connected to Jira Successfully!")
            if self.save_credentials_checkbox.get():
                self.save_credentials(server_url, email, api_token)

        except exceptions.JIRAError as e:
            self.jira_connection = None
            messagebox.showerror("Connection Failed", f"JIRA Error: {getattr(e,'text', str(e))}")
        except requests.exceptions.RequestException:
            self.jira_connection = None
            messagebox.showerror("Network Error", "Check your internet connection or JIRA server URL")
        except Exception as e:
            self.jira_connection = None
            messagebox.showerror("Error", f"Unexpected error: {str(e)}")
        finally:
            self.update_connection_status()

    def update_connection_status(self):
        self.status_canvas.delete("all")
        if self.jira_connection:
            self.status_label.config(text="Connected")
            self.status_canvas.create_oval(0, 0, 10, 10, fill="green")
        else:
            self.status_label.config(text="Not Connected")
            self.status_canvas.create_oval(0, 0, 10, 10, fill="red")

    # ==================== Excel Config ====================
    def load_excel_config(self):
        if os.path.isfile(EXCEL_CONFIG_FILE):
            try:
                with open(EXCEL_CONFIG_FILE, "r") as f:
                    data = json.load(f)
                    excel_path = data.get("excel_path", "")
                    sheet_name = data.get("sheet_name", "")
                    if excel_path and os.path.isfile(excel_path):
                        self.excel_path.set(excel_path)
                        # Populate sheets first
                        self.populate_sheets(excel_path)
                        if sheet_name and sheet_name in self.sheet_combobox['values']:
                            self.sheet_name.set(sheet_name)
                            self.sheet_combobox.set(sheet_name)
                        self.preview_excel()
            except Exception as e:
                messagebox.showerror(f"Failed to load Excel config: {e}")

    def save_excel_config(self):
        try:
            with open(EXCEL_CONFIG_FILE, "w") as f:
                json.dump({
                    "excel_path": self.excel_path.get(),
                    "sheet_name": self.sheet_name.get()
                }, f, indent=2)
        except Exception as e:
            messagebox.showerror(f"Failed to save Excel config: {e}")

    # ==================== Update Tab ====================
    def create_update_ui(self):
        frame = ttk.LabelFrame(self.update_tab, text="Excel File and Sheet")
        frame.pack(padx=10, pady=10, fill="x")

        ttk.Label(frame, text="Excel File:").grid(row=0, column=0, padx=5, pady=5)
        self.excel_path = tk.StringVar()
        ttk.Entry(frame, textvariable=self.excel_path, width=60).grid(row=0, column=1)
        ttk.Button(frame, text="Browse", command=self.browse_excel).grid(row=0, column=2, padx=5)

        ttk.Label(frame, text="Sheet Name:").grid(row=1, column=0, padx=5, pady=5)
        self.sheet_name = tk.StringVar()
        self.sheet_combobox = ttk.Combobox(frame, textvariable=self.sheet_name, width=30)
        self.sheet_combobox.grid(row=1, column=1, sticky="w")
        self.sheet_combobox.bind("<<ComboboxSelected>>", lambda e: self.preview_excel())

        # Excel preview
        self.tree_frame = tk.Frame(self.update_tab)
        self.tree_frame.pack(fill="both", expand=True, padx=10, pady=10)
        self.tree = ttk.Treeview(self.tree_frame, show="headings")  # 'headings' hides default tree column
        self.tree.pack(side="left", fill="both", expand=True)
        self.scrollbar = ttk.Scrollbar(self.tree_frame, orient="vertical", command=self.tree.yview)
        self.scrollbar.pack(side="right", fill="y")
        self.tree.configure(yscrollcommand=self.scrollbar.set)

        ttk.Button(self.update_tab, text="Run Update", command=self.run_thread).pack(pady=10)

    def browse_excel(self):
        path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx")])
        if path:
            self.excel_path.set(path)
            self.populate_sheets(path)
            self.preview_excel()
            self.save_excel_config()

    def populate_sheets(self, path):
        try:
            sheets = pd.ExcelFile(path).sheet_names
            self.sheet_combobox['values'] = sheets
            if sheets:
                self.sheet_name.set(sheets[0])
                self.sheet_combobox.current(0)  # ensures the UI combobox shows the first sheet
        except Exception as e:
            messagebox.showerror("Error", f"Failed to read sheets: {e}")

    def preview_excel(self):
        path = self.excel_path.get()
        sheet = self.sheet_name.get()
        if not path or not os.path.isfile(path):
            return

        # Read Excel into DataFrame
        df = pd.read_excel(path, sheet_name=sheet, dtype=str, keep_default_na=False).fillna("NA")

        # Clear existing tree items
        self.tree.delete(*self.tree.get_children())

        # Set up columns only if they are different
        if list(self.tree["columns"]) != list(df.columns):
            self.tree["columns"] = list(df.columns)
            self.tree["show"] = "headings"

            # Configure columns
            for col in df.columns:
                self.tree.heading(col, text=col)
                self.tree.column(col, width=120, anchor="center")  # center-aligned, adjustable width

        # Insert rows (limit to first 100)
        for index, row in df.head(100).iterrows():
            self.tree.insert("", "end", values=list(row))

        # Make sure vertical scrollbar remains linked
        self.tree.configure(yscrollcommand=self.scrollbar.set)

    # ==================== Logs Tab ====================
    def create_logs_ui(self):
        self.log_box = scrolledtext.ScrolledText(self.logs_tab, height=30, font=("Consolas", 10))
        self.log_box.pack(fill="both", expand=True)
        ttk.Button(self.logs_tab, text="View Past Logs", command=self.view_past_logs).pack(pady=5)

    def view_past_logs(self):
        import os, json
        from datetime import datetime
        import tkinter as tk
        from tkinter import ttk

        logs_folder = "logs"
        META_FILE = os.path.join(logs_folder, "important_logs.json")

        # ---------- Single window ----------
        if hasattr(self, "logs_window") and self.logs_window.winfo_exists():
            window = self.logs_window
            for w in window.winfo_children():
                w.destroy()
        else:
            window = tk.Toplevel(self.root)
            window.state("zoomed")
            window.title("Past Logs")
            self.logs_window = window

        # ---------- Metadata ----------
        def load_metadata():
            if os.path.exists(META_FILE):
                try:
                    with open(META_FILE, "r", encoding="utf-8") as f:
                        return json.load(f)
                except:
                    pass
            return {"important_logs": [], "notes": {}}

        def save_metadata(data):
            with open(META_FILE, "w", encoding="utf-8") as f:
                json.dump(data, f, indent=4)

        metadata = load_metadata()
        important_logs = set(metadata.get("important_logs", []))
        notes = metadata.get("notes", {})

        # ---------- Actions ----------
        def refresh():
            self.view_past_logs()

        def toggle_important(file):
            if file in important_logs:
                important_logs.remove(file)
            else:
                important_logs.add(file)
            metadata["important_logs"] = sorted(important_logs)
            save_metadata(metadata)
            refresh()

        def archive_log(file):
            os.makedirs(os.path.join(logs_folder, "archive"), exist_ok=True)
            os.rename(
                os.path.join(logs_folder, file),
                os.path.join(logs_folder, "archive", file)
            )
            refresh()

        def trash_log(file):
            os.makedirs(os.path.join(logs_folder, "trash"), exist_ok=True)
            os.rename(
                os.path.join(logs_folder, file),
                os.path.join(logs_folder, "trash", file)
            )
            refresh()

        def restore_log(file, src_folder):
            os.rename(
                os.path.join(src_folder, file),
                os.path.join(logs_folder, file)
            )
            refresh()

        def edit_note(file):
            note_win = tk.Toplevel(window)
            note_win.title(f"Note – {file}")
            note_win.geometry("520x320")
            note_win.transient(window)
            note_win.grab_set()

            ttk.Label(
                note_win,
                text=f"📝 Notes for {file}",
                font=("Segoe UI", 9, "bold")
            ).pack(pady=(10, 5))

            text = tk.Text(note_win, wrap="word", height=10)
            text.pack(fill="both", expand=True, padx=12, pady=5)
            text.insert("1.0", notes.get(file, ""))

            def save_note():
                content = text.get("1.0", "end").strip()
                if content:
                    notes[file] = content
                else:
                    notes.pop(file, None)
                metadata["notes"] = notes
                save_metadata(metadata)
                note_win.destroy()
                refresh()

            btn_bar = ttk.Frame(note_win)
            btn_bar.pack(fill="x", pady=10)

            ttk.Button(btn_bar, text="❌ Cancel", width=10,
                       command=note_win.destroy).pack(side="right", padx=5)

            ttk.Button(btn_bar, text="💾 Save", width=10,
                       command=save_note).pack(side="right", padx=(5, 12))

        # ---------- Top bar ----------
        top_bar = ttk.Frame(window)
        top_bar.pack(fill="x", padx=10, pady=5)

        ttk.Checkbutton(
            top_bar,
            text="⭐ Show only important",
            variable=self.show_important_only,
            command=refresh
        ).pack(side="left", padx=5)

        for txt, val in [("📄 Active", "active"),
                         ("📦 Archive", "archive"),
                         ("🗑️ Trash", "trash")]:
            ttk.Radiobutton(
                top_bar,
                text=txt,
                value=val,
                variable=self.view_mode,
                command=refresh
            ).pack(side="left", padx=5)

        # ---------- Select folder ----------
        mode = self.view_mode.get()
        base_folder = logs_folder if mode == "active" else os.path.join(logs_folder, mode)

        # ---------- Scrollable list ----------
        container = ttk.Frame(window)
        container.pack(fill="both", expand=True)

        canvas = tk.Canvas(container)
        scrollbar = ttk.Scrollbar(container, orient="vertical", command=canvas.yview)
        list_frame = ttk.Frame(canvas)

        list_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )

        canvas.create_window((0, 0), window=list_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)

        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        canvas.bind_all(
            "<MouseWheel>",
            lambda e: canvas.yview_scroll(-1 * int(e.delta / 120), "units")
        )

        if not os.path.isdir(base_folder):
            ttk.Label(list_frame, text="No logs found.", foreground="gray").pack(pady=20)
            return

        log_files = sorted(
            [f for f in os.listdir(base_folder) if f.endswith(".log")],
            reverse=True
        )

        # ---------- Render logs ----------
        last_date = None

        for file in log_files:
            if mode == "active" and self.show_important_only.get():
                if file not in important_logs:
                    continue

            try:
                date_part = "_".join(file.split("_")[:3])
                display_date = datetime.strptime(
                    date_part, "%Y_%m_%d"
                ).strftime("%d %b %Y")
            except:
                display_date = "Unknown Date"

            if display_date != last_date:
                ttk.Label(
                    list_frame,
                    text=f"📅 {display_date}",
                    font=("Segoe UI", 10, "bold")
                ).pack(anchor="w", padx=12, pady=(12, 4))
                last_date = display_date

            row = ttk.Frame(list_frame)
            row.pack(fill="x", padx=30, pady=2)

            # ---------- LEFT (EXPANDS) ----------
            left_frame = ttk.Frame(row)
            left_frame.pack(side="left", fill="x", expand=True)

            if mode == "active":
                star = "⭐" if file in important_logs else "☆"
                ttk.Button(
                    left_frame, text=star, width=3,
                    command=lambda f=file: toggle_important(f)
                ).pack(side="left")

            ttk.Button(
                left_frame, text=file,
                command=lambda f=file: self.open_log_file(f)
            ).pack(side="left", padx=4)

            ttk.Button(
                left_frame, text="📝", width=3,
                command=lambda f=file: edit_note(f)
            ).pack(side="left", padx=4)

            # ---------- RIGHT (GRID – NO SHIFT EVER) ----------
            actions_frame = ttk.Frame(row)
            actions_frame.pack(side="right", anchor="e")

            col = 0

            if mode == "active":
                ttk.Button(
                    actions_frame, text="📦", width=3,
                    command=lambda f=file: archive_log(f)
                ).grid(row=0, column=col, padx=2)
                col += 1

                ttk.Button(
                    actions_frame, text="🗑️", width=3,
                    command=lambda f=file: trash_log(f)
                ).grid(row=0, column=col, padx=2)
                col += 1
            else:
                ttk.Button(
                    actions_frame, text="🔁 Restore", width=10,
                    command=lambda f=file, s=base_folder: restore_log(f, s)
                ).grid(row=0, column=col, padx=2)
                col += 1

            NOTE_COL = col
            actions_frame.grid_columnconfigure(NOTE_COL, minsize=320)

            note_text = ""
            if file in notes:
                preview = notes[file].replace("\n", " ").strip()
                preview = preview[:60] + "…" if len(preview) > 60 else preview
                note_text = f"* {preview}"

            ttk.Label(
                actions_frame,
                text=note_text,
                foreground="#555555",
                wraplength=300,
                justify="left"
            ).grid(row=0, column=NOTE_COL, padx=(8, 0), sticky="w")

    def open_log_file(self, file):
        top = tk.Toplevel(self.root)
        top.state('zoomed')
        top.title(file)
        txt = scrolledtext.ScrolledText(top, font=("Consolas", 10))
        txt.pack(fill="both", expand=True)
        with open(os.path.join("logs", file), "r", encoding="utf-8") as f:
            txt.insert(tk.END, f.read())
        txt.configure(state="disabled")

    def log(self, message):
        self.log_box.insert(tk.END, message + "\n")
        self.log_box.see(tk.END)
        self.root.update()

    # ==================== Threaded Update ====================
    def run_thread(self):
        self.cancel_update_flag = False
        threading.Thread(target=self.run_update_with_window, daemon=True).start()

    # ==================== JIRA Update Logic with Separate Window ====================
    def run_update_with_window(self):
        if not self.jira_connection:
            messagebox.showerror("Error", "Not connected to JIRA!")
            return
        path = self.excel_path.get()
        sheet = self.sheet_name.get()
        if not path or not sheet:
            messagebox.showerror("Error", "Select Excel file and sheet!")
            return

        df = pd.read_excel(path, sheet_name=sheet, dtype=str, keep_default_na=False).fillna("NA")
        total = len(df)

        # Create logs folder and file
        os.makedirs("logs", exist_ok=True)
        timestamp = datetime.now().strftime("%Y_%m_%d_%H_%M_%S")
        timestampstr = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        self.log_file_path = os.path.join("logs", f"{timestamp}.log")
        with open(self.log_file_path, "w", encoding="utf-8") as f:
            f.write(f"{timestampstr} | 🚀 Jira Update Started.\n\n")

        # Log window
        log_win = tk.Toplevel(self.root)
        log_win.title("Update Logs")
        log_win.geometry("900x600")

        log_text = scrolledtext.ScrolledText(log_win, font=("Consolas", 10))
        log_text.pack(fill="both", expand=True)

        progress_label = tk.Label(log_win, text="0%")
        progress_label.pack(pady=5)

        progress = ttk.Progressbar(log_win, maximum=total, length=500)
        progress.pack(pady=5)

        cancel_button = ttk.Button(log_win, text="Cancel Update")
        cancel_button.pack(pady=5)

        close_button = ttk.Button(log_win, text="Close")
        close_button.pack(pady=5)

        cancel_button.config(state="normal")
        close_button.config(state="disabled")

        def log_window_func(msg):
            timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            formatted_msg = f"{timestamp} | {msg}"

            log_text.insert(tk.END, formatted_msg + "\n")
            log_text.see(tk.END)

            self.log(formatted_msg)

            with open(self.log_file_path, "a", encoding="utf-8") as f:
                f.write(formatted_msg + "\n")

        def cancel_update_func():
            self.cancel_update_flag = True
            log_window_func("⚠ Jira Update cancelled by user.")

        cancel_button.config(command=cancel_update_func)
        close_button.config(command=log_win.destroy)

        jira_server = self.server_url_entry.get()
        username = self.email_entry.get()
        api_token = self.api_token_entry.get()

        def find_failure_sub_category_id(failure_sub_category):
            failure_sub_category_dict = {}

            if failure_category == "KITE issue":
                failure_sub_category_dict = {
                    "CAN signal issue": 16522,
                    "Image comparison": 16523,
                    "Audio comparison": 16524,
                    "Video comparison": 16525,
                    "Robot operation": 16526,
                    "Device Mapping issue": 16527,
                    "KITE Application issue": 16528
                }
            elif failure_category == "TS issue":
                failure_sub_category_dict = {
                    "TS logic incorrect" : 16529
                }
            elif failure_category == "ICB issue":
                failure_sub_category_dict = {
                    "Once seen" : 16518,
                    "Always" : 16520,
                    "GAS implementation" : 16519,
                    "Timing issue" : 16521,
                    "SRL/CR changes" : 16554
                }
            elif failure_category == "Test Env issue":
                failure_sub_category_dict = {
                    "Automation Setup malfunction" : 16514,
                    "External HW" : 16515,
                    "External Application issue" : 16516,
                    "Network Issue" : 16517,
                    "Cascading Issue" : 19025
                }            
            return failure_sub_category_dict

        for idx, row in df.iterrows():
            if self.cancel_update_flag:
                log_window_func("⚠ Jira Update stopped before completion.")
                break
            try:
                ticket_id = row['Jira ID']
                test_script_id = row['Test Script ID']
                date = row['Date']
                result = row['Result']
                failure_category = row['Failure Category']
                failure_sub_category = row['Failure Sub Category']
                root_cause_details = row['Root Cause Details']
                fixation_performed = row['Fixation Performed']
                remarks = row['Remarks']
                reviewer = row['Reviewer']
                build_number = row['Build Number']
                
                failure_category_dict = {}
                failure_category_dict = {
                    "KITE issue": 16506,
                    "TS issue": 16507,
                    "ICB issue": 16512,
                    "Test Env issue": 16508,
                    "Human error": 16509
                }

                issue = self.jira_connection.issue(ticket_id)
                log_window_func(f"⏳ Updating    		            : {ticket_id} ('{test_script_id}')")

                if result == "Fail":
                    combined_values = []
                    current_failure_category = issue.fields.customfield_10851
                    if current_failure_category:
                        current_failure_values = [option.value for option in current_failure_category]
                        previous_failure_category = issue.fields.customfield_12208
                        if previous_failure_category:
                            existing_previous_values = [option.value for option in previous_failure_category]
                            combined_values = list(set(existing_previous_values + current_failure_values))
                        else:
                            combined_values = current_failure_values
                        normalized_values = ["KITE Issue" if v=="KITE Issue" else v for v in combined_values]
                        combined_values = list(set(normalized_values))

                    def clean_value(val): return None if pd.isna(val) else val

                    failure_category = clean_value(failure_category)
                    failure_sub_category = clean_value(failure_sub_category)
                    root_cause_details = clean_value(root_cause_details)
                    fixation_performed = clean_value(fixation_performed)

                    fields = {}
                    if failure_category:
                        failure_category_id = failure_category_dict[failure_category]
                        if failure_category == "Human error":
                            fields["customfield_14461"] = {"id": str(failure_category_id)}
                        else :
                            failure_sub_category_dict = find_failure_sub_category_id(failure_sub_category)
                            failure_sub_category_id = failure_sub_category_dict[failure_sub_category]
                            fields["customfield_14461"] = {"id": str(failure_category_id), "child": {"id": str(failure_sub_category_id)}}   
                    if root_cause_details:
                        fields["customfield_10614"] = str(root_cause_details)
                    if fixation_performed:
                        fields["customfield_10618"] = str(fixation_performed)
                    if combined_values:
                        fields["customfield_12208"] = [{"value": str(v)} for v in combined_values]

                    log_window_func(f"  🔄 Updating fields          : {ticket_id} ('{test_script_id}')")
                    log_window_func(f"    Failure Category          : {failure_category}")
                    log_window_func(f"    Failure Sub Category      : {failure_sub_category}")
                    log_window_func(f"    Previous Failure Category : {combined_values}")
                    log_window_func(f"    Root Cause Details        : {root_cause_details}")
                    log_window_func(f"    Fixation Performed        : {fixation_performed}")
                    issue.update(fields=fields)
                    log_window_func(f"  ✅ Updated fields           : {ticket_id} ('{test_script_id}')")

                if result == "Fail":
                    comment_text = (
                        f" |  *Date*  |  *Result*  |  *Failure Category*  |  *Failure Sub Category*  |  *Root Cause Details*  |  *Fixation Performed*  | *Remarks* | *Reviewer* | \n"
                        f" |  {date}  *[{build_number}]*  |  {result}  |  {failure_category}  |  ~{failure_sub_category}~  | {root_cause_details}  | {fixation_performed}  | {remarks} | {reviewer} | "
                    )
                    log_window_func(f"  🔄 Commenting               : {ticket_id} ('{test_script_id}')")
                    log_window_func(f"    Date                      : {date}")
                    log_window_func(f"    Build Number              : {build_number}")
                    log_window_func(f"    Result                    : {result}")
                    log_window_func(f"    Failure Category          : {failure_category}")
                    log_window_func(f"    Failure Sub Category      : {failure_sub_category}")
                    log_window_func(f"    Root Cause Details        : {root_cause_details}")
                    log_window_func(f"    Fixation Performed        : {fixation_performed}")
                    log_window_func(f"    Remarks                   : {remarks}")
                    log_window_func(f"    Reviewer                  : {reviewer}")
                    
                else:
                    comment_text = (
                        f" |  *Date*  |  *Result*  |  *Build Number*  |\n"
                        f" |  {date}  |  {result}  |  {build_number}  |"
                    )
                    log_window_func(f"  🔄 Commenting               : {ticket_id} ('{test_script_id}')")
                    log_window_func(f"    Date                      : {date}")
                    log_window_func(f"    Result                    : {result}")
                    log_window_func(f"    Build Number              : {build_number}")


                comment_url = f"{jira_server}/rest/api/2/issue/{ticket_id}/comment"
                comment_response = requests.post(
                    comment_url,
                    json={"body": comment_text},
                    auth=HTTPBasicAuth(username, api_token),
                    headers={"Content-Type": "application/json"}
                )
                if comment_response.status_code in [200, 201]:
                    comment_id = comment_response.json().get("id", "N/A")
                    log_window_func(f"  🔄 Commented                : {ticket_id} ('{test_script_id}') | Comment ID: {comment_id}\n")
                else:
                    log_window_func(f"  ❌ Comment failed           : {ticket_id} ('{test_script_id}') | Status: {comment_response.status_code} | Response: {comment_response.text}.\n")

            except Exception as e:
                log_window_func(f"  [ERROR] Could not process       : {ticket_id} ('{test_script_id}') : {e}.\n")

            progress['value'] = idx + 1
            percent = int(((idx + 1)/total)*100)
            progress_label.config(text=f"{percent}%")
            log_win.update()

        cancel_button.config(state="disabled")
        close_button.config(state="normal")
        if not self.cancel_update_flag:
            log_window_func("✅ Jira Update Completed.")
        else:
            log_window_func("⚠ Jira Update Cancelled.")
        
        self.root.after(0, self.save_excel_config)
        self.root.after(0, self.save_excel_config)

if __name__ == "__main__":
    root = tk.Tk()
    app = JiraUpdaterApp(root)
    root.mainloop()
