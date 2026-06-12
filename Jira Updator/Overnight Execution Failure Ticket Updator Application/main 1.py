import tkinter as tk
from tkinter import ttk, messagebox, filedialog, scrolledtext
from jira import JIRA, exceptions
import pandas as pd
import threading
import calendar

class JiraUpdaterGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("JIRA Status Updater-Abhijeet")
        self.root.geometry("1200x800")

        # Configure grid weights for proper resizing
        self.root.grid_rowconfigure(7, weight=1)
        self.root.grid_columnconfigure(0, weight=1)
        self.root.grid_columnconfigure(1, weight=1)

        # Configuration
        self.jira_connection = None
        self.selected_issues = []
        self.excel_file_path = None
        self.current_mode = tk.StringVar(value="failure")

        # Create GUI components
        self.create_connection_frame()
        self.create_mode_selector()
        self.create_ticket_source_frame()
        self.create_ticket_selection_frame()
        self.create_date_picker()
        self.create_failure_fields_frame()
        self.create_pass_fields_frame()
        self.create_comment_preview_frame()
        self.create_action_frame()
        self.toggle_mode()

        # Note: You can find the KGPT Code Assist Extension for VSCode [User Guide](https://kpitc.sharepoint.com/:w:/s/CTOGenAITeam-PublicDocuments/EVaX3PLJ5ZhLiFjLj4hZcn0BGAZsYjSJD6hZ8s1PIrQ5dA) for coding-related tasks.

    def create_connection_frame(self):
        frame = ttk.LabelFrame(self.root, text="JIRA Connection")
        frame.grid(row=0, column=0, padx=10, pady=5, sticky='ew')

        ttk.Label(frame, text="Server URL:").grid(row=0, column=0)
        self.server_url = ttk.Entry(frame, width=40)
        self.server_url.insert(0, "https://kpithondajapan.atlassian.net")
        self.server_url.grid(row=0, column=1)

        ttk.Label(frame, text="Email:").grid(row=1, column=0)
        self.email = ttk.Entry(frame, width=40)
        self.email.insert(0, "abhijeet.rathore@kpit.com")
        self.email.grid(row=1, column=1)

        ttk.Label(frame, text="API Token:").grid(row=2, column=0)
        self.api_token = ttk.Entry(frame, width=40, show="*")
        self.api_token.grid(row=2, column=1)

        ttk.Button(frame, text="Connect", command=self.connect_to_jira).grid(row=3, column=1)

    def create_mode_selector(self):
        frame = ttk.Frame(self.root)
        frame.grid(row=0, column=1, padx=10, pady=5, sticky='nw')

        ttk.Label(frame, text="Update Mode:").pack(side=tk.LEFT)
        self.mode_selector = ttk.Combobox(frame, values=["Failure", "Pass"], state="readonly")
        self.mode_selector.set("Failure")
        self.mode_selector.pack(side=tk.LEFT, padx=5)
        self.mode_selector.bind("<<ComboboxSelected>>", self.on_mode_change)

    def create_ticket_source_frame(self):
        frame = ttk.LabelFrame(self.root, text="Ticket Source")
        frame.grid(row=1, column=0, columnspan=2, padx=10, pady=5, sticky='ew')

        self.source_var = tk.StringVar(value='excel')
        ttk.Radiobutton(frame, text="Excel File", variable=self.source_var,
                       value='excel', command=self.toggle_source).pack(side=tk.LEFT)
        ttk.Radiobutton(frame, text="JIRA Search", variable=self.source_var,
                       value='jira', command=self.toggle_source).pack(side=tk.LEFT)

        # Excel file components
        self.excel_btn = ttk.Button(frame, text="Browse Excel", command=self.load_excel)
        self.excel_label = ttk.Label(frame, text="No file selected")

        # JIRA search components
        self.search_entry = ttk.Entry(frame, width=30)
        self.search_btn = ttk.Button(frame, text="Search", command=self.search_tickets)

        self.toggle_source()

    def create_ticket_selection_frame(self):
        frame = ttk.LabelFrame(self.root, text="Selected Tickets")
        frame.grid(row=2, column=0, columnspan=2, padx=4, pady=2, sticky='nsew')

        self.ticket_list = tk.Listbox(frame, selectmode=tk.MULTIPLE, height=10)
        self.ticket_list.pack(fill=tk.BOTH, expand=True)

        scrollbar = ttk.Scrollbar(frame, orient="vertical", command=self.ticket_list.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.ticket_list.config(yscrollcommand=scrollbar.set)

    def create_date_picker(self):
        self.date_frame = ttk.LabelFrame(self.root, text="Date Picker")
        self.date_frame.grid(row=3, column=0, columnspan=1, padx=10, pady=5, sticky='ew')

        self.year = tk.StringVar(value="2025")
        self.month = tk.StringVar(value="")
        self.day = tk.StringVar(value="")

        ttk.Label(self.date_frame, text="Year:").pack(side=tk.LEFT)
        ttk.Entry(self.date_frame, textvariable=self.year).pack(side=tk.LEFT)

        ttk.Label(self.date_frame, text="Month:").pack(side=tk.LEFT)
        ttk.Entry(self.date_frame, textvariable=self.month).pack(side=tk.LEFT)

        ttk.Label(self.date_frame, text="Day:").pack(side=tk.LEFT)
        ttk.Entry(self.date_frame, textvariable=self.day).pack(side=tk.LEFT)

        ttk.Button(self.date_frame, text="Set Date", command=self.set_date).pack(side=tk.LEFT)

    def set_date(self):
        try:
            year = int(self.year.get())
            month = int(self.month.get())
            day = int(self.day.get())

            if not (1 <= year <= 2030 and 1 <= month <= 12 and 1 <= day <= calendar.monthrange(year, month)[1]):
                raise ValueError("Invalid date")

            self.pass_date.delete(0, tk.END)
            self.pass_date.insert(0, f"{year}-{month:02}-{day:02}")
        except ValueError as e:
            messagebox.showerror("Error", str(e))

    def create_failure_fields_frame(self):
        self.failure_frame = ttk.LabelFrame(self.root, text="Failure Details")
        self.failure_frame.grid(row=3, column=1, padx=10, pady=5, sticky='nsew')
        fields = [
            ("Failure Category", "combobox", ["ICB Issue (Always)", "ICB Issue (Once)", "KITE Issue", "Test Env Issue"]),
            ("Sub-Category", "entry", None),
            ("Root Cause", "entry", None),
            ("Fixation", "entry", None)
        ]
        for idx, (label, field_type, options) in enumerate(fields):
            ttk.Label(self.failure_frame, text=f"{label}:").grid(row=idx, column=0, sticky='w')
            if field_type == "combobox":
                cb = ttk.Combobox(self.failure_frame, values=options)
                cb.grid(row=idx, column=1, sticky='ew')
                setattr(self, label.lower().replace(" ", "_").replace("-", "_"), cb)
            else:
                entry = ttk.Entry(self.failure_frame)
                entry.grid(row=idx, column=1, sticky='ew')
                setattr(self, label.lower().replace(" ", "_").replace("-", "_"), entry)

        # Now that all the widgets are created, you can insert values into them
        # Set default values
        self.failure_category.set("Issue")
        self.sub_category.insert(0, "")  
        self.root_cause.insert(0, "")
        self.fixation.insert(0, "")

    def create_pass_fields_frame(self):
        self.pass_frame = ttk.LabelFrame(self.root, text="Pass Details")
        self.pass_frame.grid(row=3, column=1, padx=10, pady=5, sticky='nsew')

        ttk.Label(self.pass_frame, text="Date:").grid(row=0, column=0, sticky='w')
        self.pass_date = ttk.Entry(self.pass_frame)
        self.pass_date.grid(row=0, column=1, sticky='ew')

        ttk.Label(self.pass_frame, text="Status:").grid(row=1, column=0, sticky='w')
        self.pass_status = ttk.Combobox(self.pass_frame, values=["Pass", "Fail"])
        self.pass_status.grid(row=1, column=1, sticky='ew')

        # Set default values
        self.pass_date.insert(0, "")
        self.pass_status.set("")

        self.pass_frame.grid_remove()

    def create_comment_preview_frame(self):
        frame = ttk.LabelFrame(self.root, text="Comment Preview")
        frame.grid(row=5, column=0, columnspan=2, padx=10, pady=5, sticky='nsew')

        self.comment_preview = scrolledtext.ScrolledText(frame, height=5, wrap=tk.WORD)
        self.comment_preview.pack(fill=tk.BOTH, expand=True)

        ttk.Button(frame, text="Generate Comment", command=self.generate_comment).pack(side=tk.BOTTOM)

    def create_action_frame(self):
        frame = ttk.LabelFrame(self.root, text="Actions")
        frame.grid(row=0, column=2, columnspan=2, padx=10, pady=5, sticky='ew')

        # Update Ticket Button
        self.update_btn = ttk.Button(frame, text="Update Tickets", command=self.start_update)
        self.update_btn.pack(side=tk.LEFT, padx=5)

        # Status Label
        self.status_label = ttk.Label(frame, text="Ready")
        self.status_label.pack(side=tk.LEFT, padx=10)

        # Set button visibility based on connection
        self.update_connection_status()

    def generate_comment(self):
        self.comment_preview.delete(1.0, tk.END)
        self.comment_preview.insert(tk.END, self.generate_table())

    def generate_table(self):
        if self.mode_selector.get() == "Failure":
            columns = [
                {"name": "Date", "width": 20},
                {"name": "Failure Category", "width": 20},
                {"name": "Sub-Category", "width": 20},
                {"name": "Root Cause", "width": 22},
                {"name": "Fixation", "width": 20}
            ]
            data = {
                "Date": f"{self.year.get()}-{self.month.get():02}-{self.day.get():02}",
                "Failure Category": self.failure_category.get(),
                "Sub-Category": self.sub_category.get(),
                "Root Cause": self.root_cause.get(),
                "Fixation": self.fixation.get()
            }
        else:
            columns = [
                {"name": "Date", "width": 20},
                {"name": "Status", "width": 20}
            ]
            data = {
                "Date": self.pass_date.get(),
                "Status": self.pass_status.get()
            }

        # Define border characters
        borders = {
            'single': {
                'horizontal': '-',
                'vertical': '|',
                'corner': '+'
            },
            'double': {
                'horizontal': '=',
                'vertical': '||',
                'corner': '#'
            }
        }
        border_style = 'single'
        border = borders.get(border_style, borders['single'])

        # Create header
        header = border['vertical'] + border['vertical'].join(
            [col['name'].center(col.get('width', 15), ' ') for col in columns]
        ) + border['vertical'] + "\n"


        # Create row
        row = border['vertical'] + border['vertical'].join(
            [str(data.get(col['name'], 'N/A')).ljust(col.get('width', 15)) for col in columns]
        ) + border['vertical'] + "\n"

        return header + row

    # This is AI generated code, please refer KPIT AI Policy before using this in your projects

    def connect_to_jira(self):
        try:
            # Validate inputs
            if not all([self.server_url.get(), self.email.get(), self.api_token.get()]):
                raise ValueError("All connection fields are required")

            # Attempt connection
            self.jira_connection = JIRA(
                server=self.server_url.get(),
                basic_auth=(self.email.get(), self.api_token.get())
            )
            messagebox.showinfo("Success", "Connected to JIRA successfully")
            self.update_connection_status()

        except exceptions.JIRAError as e:
            self.jira_connection = None
            messagebox.showerror("Connection Failed", f"JIRA Error: {str(e)}")
            self.update_connection_status()
        except Exception as e:
            self.jira_connection = None
            messagebox.showerror("Connection Failed", f"Error: {str(e)}")
            self.update_connection_status()

    def update_connection_status(self):
        if self.jira_connection:
            self.status_label.config(text="Connected", foreground="green")
            self.update_btn.config(state="normal")
        else:
            self.status_label.config(text="Not Connected", foreground="red")
            self.update_btn.config(state="disabled")

    def start_update(self):
        if not self.jira_connection:
            messagebox.showerror("Error", "Not connected to JIRA")
            return

        selected_indices = self.ticket_list.curselection()
        if not selected_indices:
            messagebox.showwarning("No Selection", "Please select tickets to update")
            return

        ticket_keys = [self.ticket_list.get(i).split()[0] for i in selected_indices]
        threading.Thread(target=self.perform_updates, args=(ticket_keys,), daemon=True).start()

    def perform_updates(self, ticket_keys):
        try:
            self.status_label.config(text="Updating...", foreground="blue")

            for key in ticket_keys:
                issue = self.jira_connection.issue(key)

                if self.mode_selector.get() == "Failure":
                    # Update failure fields
                    issue.update(fields={
                        'customfield_10851': [{'value': self.failure_category.get()}],
                        'customfield_10617': self.sub_category.get(),
                        'customfield_10614': self.root_cause.get(),
                        'customfield_10618': self.fixation.get()
                    })

                # Add comment
                comment = self.generate_table()
                self.jira_connection.add_comment(issue, comment)

            self.status_label.config(text="Update completed successfully", foreground="green")
            messagebox.showinfo("Success", f"Updated {len(ticket_keys)} tickets")

        except exceptions.JIRAError as e:
            self.status_label.config(text="JIRA Error occurred", foreground="red")
            messagebox.showerror("Update Error", f"JIRA Error: {str(e)}")
        except Exception as e:
            self.status_label.config(text="Error occurred", foreground="red")
            messagebox.showerror("Update Error", f"Error: {str(e)}")

    def on_mode_change(self, event=None):
        self.toggle_mode()

    def toggle_mode(self):
        if self.mode_selector.get() == "Failure":
            self.failure_frame.grid()
            self.pass_frame.grid_remove()
        else:
            self.failure_frame.grid_remove()
            self.pass_frame.grid()

    def toggle_source(self):
        for widget in [self.excel_btn, self.excel_label, self.search_entry, self.search_btn]:
            widget.pack_forget()

        if self.source_var.get() == 'excel':
            self.excel_btn.pack(side=tk.LEFT, padx=5)
            self.excel_label.pack(side=tk.LEFT)
        else:
            self.search_entry.pack(side=tk.LEFT, padx=5)
            self.search_btn.pack(side=tk.LEFT)

    def load_excel(self):
        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx;*.xls")])
        if file_path:
            try:
                df = pd.read_excel(file_path)
                if 'Key' not in df.columns:
                    raise ValueError("Excel file must contain 'Key' column")

                self.ticket_list.delete(0, tk.END)
                for key in df['Key']:
                    self.ticket_list.insert(tk.END, key)
                self.excel_label.config(text=file_path.split('/')[-1])

            except Exception as e:
                messagebox.showerror("Excel Error", str(e))

    def search_tickets(self):
        if not self.jira_connection:
            messagebox.showerror("Error", "Connect to JIRA first")
            return

        try:
            issues = self.jira_connection.search_issues(self.search_entry.get())
            self.ticket_list.delete(0, tk.END)
            for issue in issues:
                self.ticket_list.insert(tk.END, f"{issue.key} - {issue.fields.summary}")
        except exceptions.JIRAError as e:
            messagebox.showerror("Search Error", str(e))

if __name__ == "__main__":
    root = tk.Tk()
    app = JiraUpdaterGUI(root)
    root.mainloop()