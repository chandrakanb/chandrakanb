import tkinter as tk
from tkinter import ttk, messagebox, filedialog
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
        self.jira_connection = None

        self.setup_notebook()
        self.setup_status_label()

        self.create_connection_ui()
        self.create_update_ui()
        self.create_action_ui()
        self.create_history_ui()
        
        self.load_credentials()

    # ---------------- UI Setup ----------------

    def setup_notebook(self):
        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(expand=True, fill="both")

        self.connection_tab = ttk.Frame(self.notebook)
        self.update_tab = ttk.Frame(self.notebook)
        self.action_tab = ttk.Frame(self.notebook)
        self.history_tab = ttk.Frame(self.notebook)

        self.notebook.add(self.connection_tab, text="Connect to Jira")
        self.notebook.add(self.update_tab, text="Select Jira Ticket")
        self.notebook.add(self.action_tab, text="Update Jira Ticket")
        self.notebook.add(self.history_tab, text="History")

    def setup_status_label(self):
        self.status_label = tk.Label(self.notebook, text="Not Connected", fg="red")
        self.status_label.pack(side=tk.TOP, anchor='ne', padx=20)

    # ---------------- Connection Tab ----------------

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
        ttk.Checkbutton(frame, text="Save Credentials", variable=self.save_credentials_checkbox).grid(row=4, column=0, columnspan=2, padx=5, pady=5)

        self.clear_credentials_button = ttk.Button(frame, text="Clear Credentials", command=self.clear_credentials)
        self.clear_credentials_button.grid(row=5, column=0, columnspan=2, padx=5, pady=5)

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
        except requests.exceptions.RequestException:
            self.jira_connection = None
            messagebox.showerror("Network Error", "Check your internet connection or JIRA server URL")
        except Exception as e:
            self.jira_connection = None
            messagebox.showerror("Error", f"Unexpected error: {str(e)}")
        finally:
            self.update_connection_status()

    def update_connection_status(self):
        if self.jira_connection:
            self.status_label.config(text="Connected", fg="green")
        else:
            self.status_label.config(text="Not Connected", fg="red")

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
    # ---------------- Update Tab ----------------

    def create_update_ui(self):
        frame = ttk.LabelFrame(self.update_tab, text="Ticket Source")
        frame.pack(fill="x", padx=10, pady=5)

        self.ticket_source = tk.StringVar()
        self.ticket_source.set("excel")

        ttk.Radiobutton(frame, text="Load from Excel", variable=self.ticket_source, value="excel", command=self.update_source_view).pack(side="left", padx=5)
        ttk.Radiobutton(frame, text="Search in Jira", variable=self.ticket_source, value="search", command=self.update_source_view).pack(side="left", padx=5)

        self.source_frame = ttk.LabelFrame(self.update_tab, text="Select Ticket")
        self.source_frame.pack(fill="x", padx=10, pady=5)

        self.update_source_view()

    def update_source_view(self):
        for widget in self.source_frame.winfo_children():
            widget.destroy()

        if self.ticket_source.get() == "excel":
            self.create_excel_loader()
        else:
            self.create_jira_search_ui()

    def create_excel_loader(self):
        self.browse_button = ttk.Button(self.source_frame, text="Browse Excel", command=self.load_excel_file)
        self.browse_button.pack(pady=5)

        self.ticket_dropdown = ttk.Combobox(self.source_frame, state="readonly", width=50)
        self.ticket_dropdown.pack(pady=5)

    def create_jira_search_ui(self):
        ttk.Label(self.source_frame, text="Search Ticket:").pack(pady=5)
        self.search_entry = ttk.Entry(self.source_frame, width=50)
        self.search_entry.pack(pady=5)

        self.search_button = ttk.Button(self.source_frame, text="Search", command=self.search_jira_ticket)
        self.search_button.pack(pady=5)

        self.ticket_dropdown = ttk.Combobox(self.source_frame, state="readonly", width=50)
        self.ticket_dropdown.pack(pady=5)

    def load_excel_file(self):
        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
        if file_path:
            try:
                df = pd.read_excel(file_path)
                ticket_ids = df["Ticket ID"].dropna().tolist()
                if not ticket_ids:
                    raise ValueError("No 'Ticket ID' column found or it's empty")
                self.ticket_dropdown["values"] = ticket_ids
                self.ticket_dropdown.set(ticket_ids[0])
            except Exception as e:
                messagebox.showerror("Error", f"Failed to read Excel file: {e}")

    def search_jira_ticket(self):
        if not self.jira_connection:
            messagebox.showerror("Error", "Please connect to Jira first.")
            return
        query = self.search_entry.get().strip()
        if not query:
            messagebox.showwarning("Warning", "Please enter a search query.")
            return
        try:
            issues = self.jira_connection.search_issues(f'summary ~ "{query}"', maxResults=10)
            if not issues:
                messagebox.showinfo("Info", "No tickets found for the given query.")
                return
            ticket_ids = [issue.key for issue in issues]
            self.ticket_dropdown["values"] = ticket_ids
            self.ticket_dropdown.set(ticket_ids[0])
        except Exception as e:
            messagebox.showerror("Error", f"Search failed: {e}")

    # ---------------- Action Tab ----------------

    def create_action_ui(self):
        frame = ttk.LabelFrame(self.action_tab, text="Ticket Actions")
        frame.pack(fill="both", expand=True, padx=10, pady=10)

        # Transition
        ttk.Label(frame, text="Transition to:").grid(row=0, column=0, padx=5, pady=5)
        self.transition_entry = ttk.Entry(frame, width=30)
        self.transition_entry.grid(row=0, column=1, padx=5, pady=5)
        ttk.Button(frame, text="Transition", command=self.transition_ticket).grid(row=0, column=2, padx=5, pady=5)

        # Edit Summary
        ttk.Label(frame, text="New Summary:").grid(row=1, column=0, padx=5, pady=5)
        self.new_summary_entry = ttk.Entry(frame, width=30)
        self.new_summary_entry.grid(row=1, column=1, padx=5, pady=5)
        ttk.Button(frame, text="Edit", command=self.edit_ticket).grid(row=1, column=2, padx=5, pady=5)

        # Comment
        ttk.Label(frame, text="Comment:").grid(row=2, column=0, padx=5, pady=5)
        self.comment_entry = ttk.Entry(frame, width=30)
        self.comment_entry.grid(row=2, column=1, padx=5, pady=5)
        ttk.Button(frame, text="Add Comment", command=self.add_comment).grid(row=2, column=2, padx=5, pady=5)

        # Assignee
        ttk.Label(frame, text="Assign To:").grid(row=3, column=0, padx=5, pady=5)
        self.assignee_entry = ttk.Entry(frame, width=30)
        self.assignee_entry.grid(row=3, column=1, padx=5, pady=5)
        ttk.Button(frame, text="Assign", command=self.assign_ticket).grid(row=3, column=2, padx=5, pady=5)

    def get_selected_ticket(self):
        ticket = self.ticket_dropdown.get().strip()
        if not ticket:
            messagebox.showerror("Error", "No ticket selected.")
        return ticket

    def transition_ticket(self):
        ticket_id = self.get_selected_ticket()
        if ticket_id and self.jira_connection:
            try:
                transition = self.transition_entry.get().strip()
                self.jira_connection.transition_issue(ticket_id, transition)
                messagebox.showinfo("Success", f"Ticket {ticket_id} transitioned.")
            except Exception as e:
                messagebox.showerror("Error", str(e))

    def edit_ticket(self):
        ticket_id = self.get_selected_ticket()
        if ticket_id and self.jira_connection:
            try:
                summary = self.new_summary_entry.get().strip()
                self.jira_connection.issue(ticket_id).update(summary=summary)
                messagebox.showinfo("Success", f"Ticket {ticket_id} summary updated.")
            except Exception as e:
                messagebox.showerror("Error", str(e))

    def add_comment(self):
        ticket_id = self.get_selected_ticket()
        if ticket_id and self.jira_connection:
            try:
                comment = self.comment_entry.get().strip()
                self.jira_connection.add_comment(ticket_id, comment)
                messagebox.showinfo("Success", f"Comment added to {ticket_id}.")
            except Exception as e:
                messagebox.showerror("Error", str(e))

    def assign_ticket(self):
        ticket_id = self.get_selected_ticket()
        if ticket_id and self.jira_connection:
            try:
                assignee = self.assignee_entry.get().strip()
                self.jira_connection.assign_issue(ticket_id, assignee)
                messagebox.showinfo("Success", f"{ticket_id} assigned to {assignee}.")
            except Exception as e:
                messagebox.showerror("Error", str(e))

    # ---------------- History Tab ----------------

    def create_history_ui(self):
        frame = ttk.LabelFrame(self.history_tab, text="Change History")
        frame.pack(fill="both", expand=True, padx=10, pady=10)
        self.history_text = tk.Text(frame, height=20)
        self.history_text.pack(fill="both", expand=True, padx=5, pady=5)

# ---------------- Application Start ----------------

if __name__ == "__main__":
    root = tk.Tk()
    app = JiraConnector(root)
    root.mainloop()
