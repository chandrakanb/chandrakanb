import tkinter as tk
from tkinter import ttk, messagebox
import json
import requests
from jira import JIRA, JIRAError

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
        self.status_label.pack()
        
        self.create_connection_ui()
        self.create_update_ui()
        self.load_credentials()
    
    def create_connection_ui(self):
        frame = ttk.Frame(self.connection_tab)
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
        frame = ttk.Frame(self.update_tab)
        frame.pack(padx=10, pady=10)

        ttk.Label(frame, text="Jira Ticket ID:").grid(row=0, column=0, padx=5, pady=5)
        self.ticket_id_entry = ttk.Entry(frame, width=30)
        self.ticket_id_entry.grid(row=0, column=1, padx=5, pady=5)
        
        ttk.Label(frame, text="Field to Update:").grid(row=1, column=0, padx=5, pady=5)
        self.field_entry = ttk.Entry(frame, width=30)
        self.field_entry.grid(row=1, column=1, padx=5, pady=5)
        
        ttk.Label(frame, text="New Value:").grid(row=2, column=0, padx=5, pady=5)
        self.new_value_entry = ttk.Entry(frame, width=30)
        self.new_value_entry.grid(row=2, column=1, padx=5, pady=5)
        
        self.update_ticket_button = ttk.Button(frame, text="Update Ticket", command=self.update_jira_ticket)
        self.update_ticket_button.grid(row=3, column=0, columnspan=2, padx=5, pady=5)
    
    def connect_to_jira(self):
        server_url = self.server_url_entry.get()
        email = self.email_entry.get()
        api_token = self.api_token_entry.get()

        if not server_url or not email or not api_token:
            messagebox.showerror("Error", "Please fill in all fields")
            return

        try:
            self.jira = JIRA(server=server_url, basic_auth=(email, api_token))
            self.status_label.config(text="Connected", fg="green")
            messagebox.showinfo("Success", "Connected to Jira successfully!")
            if self.save_credentials_checkbox.get():
                self.save_credentials(server_url, email, api_token)
        except JIRAError as e:
            self.status_label.config(text="Not Connected", fg="red")
            messagebox.showerror("Error", f"Failed to connect: {str(e)}")

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

if __name__ == "__main__":
    root = tk.Tk()
    app = JiraConnector(root)
    root.mainloop()
