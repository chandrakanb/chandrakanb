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
        self.notebook.add(self.update_tab, text="Select Jira Ticket")
        
        self.action_tab = ttk.Frame(self.notebook)
        self.notebook.add(self.action_tab, text="Update Jira Ticket")
        
        self.history_tab = ttk.Frame(self.notebook)
        self.notebook.add(self.history_tab, text="History")
        
        self.status_label = tk.Label(self.notebook, text="Not Connected", fg="red")
        self.status_label.pack(side=tk.TOP, anchor='ne', padx=20)
        
        self.create_connection_ui()
        self.create_update_ui()
        self.create_action_ui()
        self.create_history_ui()
        self.load_credentials()
        
        self.jira_connection = None
    
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
        self.create_ticket_source_frame()
        self.create_ticket_selection_frame()
        self.toggle_source()
        self.create_action_selection_frame()
        self.toggle_action()

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
        frame = ttk.LabelFrame(self.update_tab, text="Select Tickets")
        frame.pack(side=tk.TOP, padx=4, pady=2, fill='both', expand=True)

        listbox_frame = tk.Frame(frame)
        listbox_frame.pack(fill=tk.BOTH, expand=True)

        self.ticket_list = tk.Listbox(listbox_frame, selectmode=tk.MULTIPLE, height=10)
        self.ticket_list.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        scrollbar = ttk.Scrollbar(listbox_frame, orient="vertical", command=self.ticket_list.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        self.ticket_list.config(yscrollcommand=scrollbar.set)

    def create_action_selection_frame(self):
        frame = ttk.LabelFrame(self.update_tab, text="Select action")
        frame.pack(padx=10, pady=10, fill='x')
        
        self.action_var = tk.StringVar(value='transition')
        
        ttk.Radiobutton(frame, text="Transition Issues", variable=self.action_var,
                        value='transition', command=self.toggle_action).pack(side=tk.LEFT)
        ttk.Radiobutton(frame, text="Edit Issues", variable=self.action_var,
                        value='edit', command=self.toggle_action).pack(side=tk.LEFT)
        ttk.Radiobutton(frame, text="Add Comment", variable=self.action_var,
                        value='comment', command=self.toggle_action).pack(side=tk.LEFT)
        ttk.Radiobutton(frame, text="Change Assignee", variable=self.action_var,
                        value='assignee', command=self.toggle_action).pack(side=tk.LEFT)
        
        
        self.action_btn = ttk.Button(frame, text="Next", command=self.toggle_action)
        self.action_btn.pack(side=tk.LEFT, padx=5)
        
        self.toggle_action()

    def toggle_action(self):
        frame = ttk.LabelFrame(self.update_tab, text="Select action")
        frame.pack(padx=10, pady=10, fill='x')
        
        for widget in [self.action_btn, self.excel_label, self.search_entry, self.search_btn]:
            widget.pack_forget()

        if self.action_var.get() == 'transition':
            self.action_btn = ttk.Button(frame, text="Next", command=self.populate_transitions)
            self.action_btn.pack(side=tk.LEFT, padx=5)
        elif self.action_var.get() == 'edit':
            self.action_btn = ttk.Button(frame, text="Next", command=self.populate_transitions1)
            self.action_btn.pack(side=tk.LEFT, padx=5)
        elif self.action_var.get() == 'comment':
            self.action_btn = ttk.Button(frame, text="Next", command=self.populate_transitions2)
            self.action_btn.pack(side=tk.LEFT, padx=5)
        elif self.action_var.get() == 'assignee':
            self.action_btn = ttk.Button(frame, text="Next", command=self.populate_transitions3)
            self.action_btn.pack(side=tk.LEFT, padx=5)
        else:
            self.action_btn = ttk.Button(frame, text="Next", command=self.populate_transitions4)
            self.action_btn.pack(side=tk.LEFT, padx=5)

    def create_action_ui(self):
        self .create_selected_ticket_frame()
        self.create_selected_action_frame()
        
    def get_selected_tickets(self):
        selected_indices = self.ticket_list.curselection()
        selected_tickets = [self.ticket_list.get(i) for i in selected_indices]
        return selected_tickets

    def create_selected_ticket_frame(self):
        self.selected_ticket_frame = ttk.LabelFrame(self.action_tab, text="Selected Tickets")
        self.selected_ticket_frame.pack(side=tk.TOP, padx=4, pady=2, fill='both', expand=True)

        listbox_frame = tk.Frame(self.selected_ticket_frame)
        listbox_frame.pack(fill=tk.BOTH, expand=True)

        self.selected_ticket_list = tk.Listbox(listbox_frame, height=10)
        self.selected_ticket_list.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        scrollbar = ttk.Scrollbar(listbox_frame, orient="vertical", command=self.selected_ticket_list.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        self.selected_ticket_list.config(yscrollcommand=scrollbar.set)
        
    def update_selected_ticket_frame(self):
        selected_tickets = self.get_selected_tickets()
        self.selected_ticket_list.delete(0, tk.END)
        for ticket in selected_tickets:
            self.selected_ticket_list.insert(tk.END, ticket)

    def create_selected_action_frame(self):
        frame = ttk.LabelFrame(self.action_tab, text="Selected action")
        frame.pack(padx=10, pady=10, fill='x')
        
        self.create_transition_frame()

    def create_transition_frame(self):
        self.transition_frame = ttk.LabelFrame(self.action_tab, text="Transition Issue")
        self.transition_frame.pack(side=tk.TOP, padx=4, pady=2, fill='both', expand=True)

        self.transition_label = ttk.Label(self.transition_frame, text="Select Transition:")
        self.transition_label.pack(side=tk.LEFT, padx=5, pady=5)

        self.transition_var = tk.StringVar()
        self.transition_dropdown = ttk.Combobox(self.transition_frame, textvariable=self.transition_var, state="readonly")
        self.transition_dropdown.pack(side=tk.LEFT, padx=5, pady=5)

        self.transition_button = ttk.Button(self.transition_frame, text="Apply Transition", command=self.apply_transition)
        self.transition_button.pack(side=tk.LEFT, padx=5, pady=5)

    def get_transitions(self, issue_key):
        try:
            transitions = self.jira_connection.transitions(issue_key)
            return {t['name']: t['id'] for t in transitions}
        except Exception as e:
            messagebox.showerror("Error", f"Failed to get transitions: {str(e)}")
            return {}

    def apply_transition(self):
        selected_transition = self.transition_var.get()
        selected_issues = self.get_selected_tickets()  # Fetch selected tickets

        if not selected_issues:
            messagebox.showwarning("Warning", "No issues selected!")
            return

        if not selected_transition:
            messagebox.showwarning("Warning", "Please select a transition!")
            return

        for issue_key in selected_issues:
            transitions = self.get_transitions(issue_key)
            if selected_transition in transitions:
                try:
                    self.jira_connection.transition_issue(issue_key, transitions[selected_transition])
                    messagebox.showinfo("Success", f"Issue {issue_key} transitioned to {selected_transition}")
                except Exception as e:
                    messagebox.showerror("Error", f"Failed to transition {issue_key}: {str(e)}")
            else:
                messagebox.showerror("Error", f"Transition '{selected_transition}' not available for {issue_key}")

    def populate_transitions(self):
        selected_issues = self.get_selected_tickets()
        if not selected_issues:
            messagebox.showwarning("Warning", "No issue selected to fetch transitions.")
            return

        first_issue = selected_issues[0]  # Fetch transitions for the first selected issue
        transitions = self.get_transitions(first_issue)

        if transitions:
            self.transition_dropdown['values'] = list(transitions.keys())
            self.transition_var.set("")  # Reset selection
        else:
            messagebox.showerror("Error", "No transitions available.")

    def connect_to_jira(self):
        try:
            server_url = self.server_url_entry.get().strip()
            email = self.email_entry.get().strip()
            api_token = self.api_token_entry.get().strip()

            if not server_url or not email or not api_token:
                raise ValueError("All fields (Server URL, Email, API Token) are required")

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
                self.ticket_list_status_label.config(text="     Loading...     ", foreground="blue")
                self.update_tab.update_idletasks()

                if file_path.endswith(('.xlsx', '.xls')):
                    df = pd.read_excel(file_path)
                elif file_path.endswith('.csv'):
                    df = pd.read_csv(file_path)
                else:
                    raise ValueError("Unsupported file format")

                if 'Issue key' not in df.columns or 'Summary' not in df.columns or 'Status' not in df.columns or 'Assignee' not in df.columns:
                    raise ValueError("File must contain 'Issue key', 'Summary', 'Status' and 'Assignee' columns")

                self.ticket_list.delete(0, tk.END)
                for _, row in df.iterrows():
                    issue_key = row['Issue key']
                    summary = row['Summary']
                    status = row['Status']
                    assignee = row['Assignee']
                    self.ticket_list.insert(tk.END, f"| {issue_key} | {summary} | {status} | {assignee} |")

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
                    break
                
                all_issues.extend(issues)
                start_at += batch_size

            if not all_issues:
                messagebox.showinfo("No Results", "No tickets found for the given query.")
                self.ticket_list_status_label.config(text="  No Tickets Found  ", foreground="orange")
                return

            for issue in all_issues:
                self.ticket_list.insert(tk.END, f"| {issue.key} | {issue.fields.summary} | {issue.fields.status} | {issue.fields.assignee} |")

            self.ticket_list_status_label.config(text="  Search Completed  ", foreground="green")

        except Exception as e:
            messagebox.showerror("Search Error", str(e))
            self.ticket_list_status_label.config(text="    Search Failed   ", foreground="red")
    
    def create_history_ui(self):
        frame = ttk.LabelFrame(self.history_tab)
        frame.pack(padx=10, pady=10, fill='both', expand=True)

        label = ttk.Label(frame, text="Feature to be implemented")
        label.pack(padx=5, pady=5)
    
if __name__ == "__main__":
    root = tk.Tk()
    app = JiraConnector(root)
    root.mainloop()
