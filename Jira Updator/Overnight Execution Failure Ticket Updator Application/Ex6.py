import tkinter as tk
from tkinter import ttk, messagebox, filedialog, scrolledtext
import json
import requests
import pandas as pd
from jira import JIRA, JIRAError, exceptions

CREDENTIALS_FILE = "jira_credentials.json"

class JiraConnector:
    def __init__(self, root):
        self.root = root
        self.root.title("Jira Connector")
        self.root.geometry("1200x700")
        self.credentials_saved = False
        
        self.notebook = ttk.Notebook(root)
        self.notebook.pack(expand=True, fill="both")
        
        self.connection_tab = ttk.Frame(self.notebook)
        self.notebook.add(self.connection_tab, text="Connect to Jira")
        
        self.update_tab = ttk.Frame(self.notebook)
        self.notebook.add(self.update_tab, text="Update Jira Ticket")
        
        self.status_label = tk.Label(self.notebook, text="Not Connected", fg="red")
        self.status_label.pack(side=tk.TOP, anchor='ne', padx=20)
        
        self.create_connection_ui()
        self.create_update_ui()
        self.load_credentials()
        
        # Configuration
        self.jira_connection = None
        self.selected_issues = []
        self.excel_file_path = None
        self.current_mode = tk.StringVar(value="failure")
    
    def create_connection_ui(self):
        frame = ttk.LabelFrame(self.connection_tab)
        frame.pack(padx=10, pady=10)

        ttk.Label(frame, text="Server URL:").grid(row=0, column=0, padx=5, pady=5)
        self.server_url_entry = ttk.Entry(frame, width=50)
        self.server_url_entry.grid(row=0, column=1, padx=5, pady=5)

        ttk.Label(frame, text="Email:").grid(row=1, column=0, padx=5, pady=5)
        self.email_entry = ttk.Entry(frame, width=50)
        self.email_entry.grid(row=1, column=1, padx=5, pady=5)

        ttk.Label(frame, text="API Token:").grid(row=2, column=0, padx=5, pady=5)
        self.api_token_entry = ttk.Entry(frame, width=50, show="*")
        self.api_token_entry.grid(row=2, column=1, padx=5, pady=5)

        self.connect_button = ttk.Button(frame, text="Connect", command=self.connect_to_jira)
        self.connect_button.grid(row=3, column=0, columnspan=2, padx=5, pady=5)

        self.save_credentials_checkbox = tk.BooleanVar()
        self.save_credentials_checkbox.set(False)
        ttk.Checkbutton(frame, text="Save Credentials", variable=self.save_credentials_checkbox).grid(row=4, column=0, columnspan=2, padx=5, pady=5)

        self.clear_credentials_button = ttk.Button(frame, text="Clear Credentials", command=self.clear_credentials)
        self.clear_credentials_button.grid(row=5, column=0, columnspan=2, padx=5, pady=5)

    def create_update_ui(self):
        self.create_ticket_source_frame()  # Ticket Source frame at the top
        self.create_ticket_selection_frame()
        self.toggle_source()

    def create_ticket_source_frame(self):
        frame = ttk.LabelFrame(self.update_tab, text="Ticket Source")
        frame.pack(padx=10, pady=10, fill='x')

        self.source_var = tk.StringVar(value='excel')
        
        ttk.Radiobutton(frame, text="Excel File", variable=self.source_var,
                        value='excel', command=self.toggle_source).pack(side=tk.LEFT)
        ttk.Radiobutton(frame, text="JIRA Search", variable=self.source_var,
                        value='jira', command=self.toggle_source).pack(side=tk.LEFT)

        self.excel_btn = ttk.Button(frame, text="Browse Excel", command=self.load_excel)
        self.excel_btn.pack(side=tk.LEFT, padx=5)
        self.excel_label = ttk.Label(frame, text="No file selected")
        self.excel_label.pack(side=tk.LEFT, padx=5)

        self.search_entry = ttk.Entry(frame, width=30)
        self.search_entry.pack(side=tk.LEFT, padx=5)
        self.search_btn = ttk.Button(frame, text="Search", command=self.search_tickets)
        self.search_btn.pack(side=tk.LEFT, padx=5)

        self.ticket_list_status_label = tk.Label(frame, text="                    ")
        self.ticket_list_status_label.pack(side=tk.RIGHT, padx=10)
        
        self.toggle_source()

    def create_ticket_selection_frame(self):
        frame = ttk.LabelFrame(self.update_tab, text="Selected Tickets")
        frame.pack(side=tk.TOP, padx=4, pady=2, fill='both', expand=True)  # Below Ticket Source frame

        self.ticket_list = tk.Listbox(frame, selectmode=tk.MULTIPLE, height=10)
        self.ticket_list.pack(fill=tk.BOTH, expand=True)

        scrollbar = ttk.Scrollbar(frame, orient="vertical", command=self.ticket_list.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.ticket_list.config(yscrollcommand=scrollbar.set)

    def connect_to_jira(self):
        try:
            # Get input values
            server_url = self.server_url_entry.get().strip()
            email = self.email_entry.get().strip()
            api_token = self.api_token_entry.get().strip()

            # Validate input fields
            if not server_url or not email or not api_token:
                raise ValueError("All fields (Server URL, Email, API Token) are required")

            # Attempt connection
            self.jira_connection = JIRA(server=server_url, basic_auth=(email, api_token))
            messagebox.showinfo("Success", "Connected to JIRA successfully")
            if self.save_credentials_checkbox.get():
                self.save_credentials(server_url, email, api_token)
                
        except exceptions.JIRAError as e:
            self.jira_connection = None
            messagebox.showerror("Connection Failed", f"JIRA Error: {e.text if hasattr(e, 'text') else str(e)}")

        except requests.exceptions.RequestException as e:
            self.jira_connection = None
            messagebox.showerror("Network Error", "Check your internet connection or JIRA server URL")

        except Exception as e:
            self.jira_connection = None
            messagebox.showerror("Error", f"Unexpected error: {str(e)}")

        finally:
            self.update_connection_status()

    def update_connection_status(self):
        if self.jira_connection:
            self.status_label.config(text="Connected", foreground="green")
        else:
            self.status_label.config(text="Not Connected", foreground="red")

    def update_jira_ticket(self):
        if not hasattr(self, 'jira') or not self.jira:
            messagebox.showerror("Error", "Please connect to Jira first.")
            return
        
        ticket_id = self.ticket_id_entry.get()
        field = self.field_entry.get()
        new_value = self.new_value_entry.get()

        if not ticket_id or not field or not new_value:
            messagebox.showerror("Error", "Please fill in all fields")
            return
        
        try:
            issue = self.jira.issue(ticket_id)
            issue.update(fields={field: new_value})
            messagebox.showinfo("Success", "Ticket updated successfully!")
        except JIRAError as e:
            messagebox.showerror("Error", f"Failed to update ticket: {str(e)}")

    def save_credentials(self, server_url, email, api_token):
        with open(CREDENTIALS_FILE, "w") as f:
            json.dump({"server_url": server_url, "email": email, "api_token": api_token}, f)
    
    def load_credentials(self):
        try:
            with open(CREDENTIALS_FILE, "r") as f:
                data = json.load(f)
                self.server_url_entry.insert(0, data["server_url"])
                self.email_entry.insert(0, data["email"])
                self.api_token_entry.insert(0, data["api_token"])
                self.save_credentials_checkbox.set(True)
        except (FileNotFoundError, json.JSONDecodeError):
            pass
    
    def clear_credentials(self):
        self.server_url_entry.delete(0, tk.END)
        self.email_entry.delete(0, tk.END)
        self.api_token_entry.delete(0, tk.END)
        self.save_credentials_checkbox.set(False)
        try:
            open(CREDENTIALS_FILE, "w").close()
        except FileNotFoundError:
            pass
        self.status_label.config(text="Not Connected", fg="red")
        messagebox.showinfo("Success", "Credentials cleared!")

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
        file_path = filedialog.askopenfilename(filetypes=[("Excel and CSV files", "*.xlsx;*.xls;*.csv")])
        if file_path:
            try:
                # Show loading status
                self.ticket_list_status_label.config(text="     Loading...     ", foreground="blue")
                self.update_tab.update_idletasks()

                # Load file
                if file_path.endswith(('.xlsx', '.xls')):
                    df = pd.read_excel(file_path)
                elif file_path.endswith('.csv'):
                    df = pd.read_csv(file_path)
                else:
                    raise ValueError("Unsupported file format")

                 # Ensure 'Issue key' and 'Summary' columns exist
                if 'Issue key' not in df.columns or 'Summary' not in df.columns:
                    raise ValueError("File must contain 'Issue key' and 'Summary' columns")

                # Populate the ticket list with Issue key and Summary as a tuple
                self.ticket_list.delete(0, tk.END)
                for _, row in df.iterrows():
                    issue_key = row['Issue key']
                    summary = row['Summary']
                    self.ticket_list.insert(tk.END, f"{issue_key}: {summary}")

                # Display file name and status update
                self.excel_label.config(text=file_path.split('/')[-1])
                self.ticket_list_status_label.config(text="  Loading Completed ", foreground="green")

            except Exception as e:
                messagebox.showerror("File Error", str(e))
                self.ticket_list_status_label.config(text="     Load Failed    ", foreground="red")

    def search_tickets(self):
        if not self.jira_connection:
            messagebox.showerror("Error", "Connect to JIRA first")
            return

        query = self.search_entry.get().strip()
        if not query:
            messagebox.showerror("Error", "Enter a JQL query")
            return

        try:
            self.ticket_list_status_label.config(text="    Searching...    ", foreground="blue")
            self.update_tab.update_idletasks()

            self.ticket_list.delete(0, tk.END)

            start_at = 0
            batch_size = 50
            all_issues = []

            while True:
                issues = self.jira_connection.search_issues(query, startAt=start_at, maxResults=batch_size)
                if not issues:
                    break  # No more issues to fetch
                
                all_issues.extend(issues)
                start_at += batch_size

            if not all_issues:
                messagebox.showinfo("No Results", "No tickets found for the given query.")
                self.ticket_list_status_label.config(text="  No Tickets Found  ", foreground="orange")
                return

            for issue in all_issues:
                self.ticket_list.insert(tk.END, f"{issue.key}: {issue.fields.summary}")

            self.ticket_list_status_label.config(text="  Search Completed  ", foreground="green")

        except Exception as e:
            messagebox.showerror("Search Error", str(e))
            self.ticket_list_status_label.config(text="    Search Failed   ", foreground="red")

if __name__ == "__main__":
    root = tk.Tk()
    app = JiraConnector(root)
    root.mainloop()
