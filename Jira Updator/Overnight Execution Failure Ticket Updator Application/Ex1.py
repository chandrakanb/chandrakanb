import tkinter as tk
from tkinter import ttk, messagebox
import json
import requests
from requests.auth import HTTPBasicAuth

CREDENTIALS_FILE = "jira_credentials.json"

class JiraConnector:
    def __init__(self, root):
        self.root = root
        self.root.title("Jira Connector")
        self.root.geometry("1200x700")

        self.notebook = ttk.Notebook(root)
        self.notebook.pack(expand=True, fill="both")

        self.connection_tab = ttk.Frame(self.notebook)
        self.notebook.add(self.connection_tab, text="Connect to Jira")

        self.status_label = tk.Label(self.notebook, text="Not Connected", fg="red")
        self.status_label.pack()

        self.create_connection_ui()
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
        ttk.Checkbutton(frame, text="Save Credentials", variable=self.save_credentials_checkbox).grid(row=4, column=0, columnspan=2, padx=5, pady=5)

        self.clear_credentials_button = ttk.Button(frame, text="Clear Credentials", command=self.clear_credentials)
        self.clear_credentials_button.grid(row=5, column=0, columnspan=2, padx=5, pady=5)

    def connect_to_jira(self):
        server_url = self.server_url_entry.get().strip()
        email = self.email_entry.get().strip()
        api_token = self.api_token_entry.get().strip()

        if not server_url or not email or not api_token:
            messagebox.showerror("Error", "Please fill in all fields")
            return

        try:
            response = requests.get(f"{server_url}/rest/api/2/myself", auth=HTTPBasicAuth(email, api_token))

            if response.status_code == 200:
                self.status_label.config(text="Connected", fg="green")
                messagebox.showinfo("Success", "Connected to Jira successfully!")

                if self.save_credentials_checkbox.get():
                    self.save_credentials(server_url, email, api_token)
            else:
                self.status_label.config(text="Not Connected", fg="red")
                messagebox.showerror("Error", f"Failed to connect: {response.status_code}\n{response.text}")

        except requests.exceptions.RequestException as e:
            self.status_label.config(text="Not Connected", fg="red")
            messagebox.showerror("Error", f"Connection failed: {str(e)}")

    def save_credentials(self, server_url, email, api_token):
        with open(CREDENTIALS_FILE, "w") as f:
            json.dump(
				{
					"server_url": server_url, 
					"email": email, 
					"api_token": api_token
				},
				f,
                indent=2
			)

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
        
        # Clear stored credentials file
        try:
            open(CREDENTIALS_FILE, "w").close()
        except FileNotFoundError:
            pass

        # Update status to "Not Connected"
        self.status_label.config(text="Not Connected", fg="red")
        
        messagebox.showinfo("Success", "Credentials cleared!")

if __name__ == "__main__":
    root = tk.Tk()
    app = JiraConnector(root)
    root.mainloop()
