import tkinter as tk
from tkinter import messagebox

class JiraConnector:
    def __init__(self, root):
        self.root = root
        self.root.title("Jira Connector")
        self.credentials_saved = False

        # Create frames
        self.frame = tk.Frame(self.root)
        self.frame.pack(padx=10, pady=10)

        # Create labels and entries
        tk.Label(self.frame, text="Server URL:").grid(row=0, column=0, padx=5, pady=5)
        self.server_url_entry = tk.Entry(self.frame, width=50)
        self.server_url_entry.grid(row=0, column=1, padx=5, pady=5)

        tk.Label(self.frame, text="Email:").grid(row=1, column=0, padx=5, pady=5)
        self.email_entry = tk.Entry(self.frame, width=50)
        self.email_entry.grid(row=1, column=1, padx=5, pady=5)

        tk.Label(self.frame, text="API Token:").grid(row=2, column=0, padx=5, pady=5)
        self.api_token_entry = tk.Entry(self.frame, width=50, show="*")
        self.api_token_entry.grid(row=2, column=1, padx=5, pady=5)

        # Create buttons
        self.connect_button = tk.Button(self.frame, text="Connect", command=self.connect_to_jira)
        self.connect_button.grid(row=3, column=0, columnspan=2, padx=5, pady=5)

        self.save_credentials_checkbox = tk.BooleanVar()
        self.save_credentials_checkbox.set(False)
        tk.Checkbutton(self.frame, text="Save Credentials", variable=self.save_credentials_checkbox).grid(row=4, column=0, columnspan=2, padx=5, pady=5)

        self.reset_credentials_button = tk.Button(self.frame, text="Reset Credentials", command=self.reset_credentials)
        self.reset_credentials_button.grid(row=5, column=0, columnspan=2, padx=5, pady=5)

    def connect_to_jira(self):
        server_url = self.server_url_entry.get()
        email = self.email_entry.get()
        api_token = self.api_token_entry.get()

        if not server_url or not email or not api_token:
            messagebox.showerror("Error", "Please fill in all fields")
            return

        # Here you would put the actual code to connect to Jira
        # For now, it just prints the credentials
        print(f"Server URL: {server_url}")
        print(f"Email: {email}")
        print(f"API Token: {api_token}")

        if self.save_credentials_checkbox.get():
            # Here you would put the actual code to save the credentials
            # For now, it just prints a message
            print("Credentials saved")
            self.credentials_saved = True

    def reset_credentials(self):
        self.server_url_entry.delete(0, tk.END)
        self.email_entry.delete(0, tk.END)
        self.api_token_entry.delete(0, tk.END)
        self.save_credentials_checkbox.set(False)
        self.credentials_saved = False
        # Here you would put the actual code to reset the credentials
        # For now, it just prints a message
        print("Credentials reset")

if __name__ == "__main__":
    root = tk.Tk()
    app = JiraConnector(root)
    root.mainloop()