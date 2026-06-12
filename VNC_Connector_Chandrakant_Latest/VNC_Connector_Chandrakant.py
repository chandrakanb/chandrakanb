import tkinter as tk
from tkinter import ttk, messagebox
import os
import subprocess

APP_FOLDER = os.path.dirname(os.path.abspath(__file__))

root = tk.Tk()
root.title("VNC Connector - Chandrakant")
root.geometry("600x400")

notebook = ttk.Notebook(root)
notebook.pack(fill="both", expand=True)

project_tabs = {}

def refresh_dropdown_list(project):
    if project in project_tabs:
        frame = project_tabs[project]
        for widget in frame.winfo_children():
            widget.destroy()
        build_dropdown(frame, project)

def build_dropdown(frame, project):
    files = [f for f in os.listdir(os.path.join(APP_FOLDER, project)) if f.endswith(".vnc")]
    var = tk.StringVar()
    dropdown = ttk.Combobox(frame, textvariable=var, values=[f[:-4] for f in files], state="readonly")
    if files:
        var.set(files[0][:-4])
    dropdown.pack(pady=10, ipadx=5, ipady=2)

    def connect():
        selected = var.get()
        if not selected:
            messagebox.showwarning("Warning", "Please select a bench")
            return
        file_path = os.path.join(APP_FOLDER, project, f"{selected}.vnc")
        if os.path.exists(file_path):
            subprocess.Popen([file_path], shell=True)

    connect_button = ttk.Button(frame, text="Connect", command=connect)
    connect_button.pack(pady=5)

# Create a tab for each valid project folder
for project in os.listdir(APP_FOLDER):
    project_path = os.path.join(APP_FOLDER, project)
    if os.path.isdir(project_path) and any(f.endswith(".vnc") for f in os.listdir(project_path)):
        frame = ttk.Frame(notebook)
        notebook.add(frame, text=project)
        project_tabs[project] = frame
        build_dropdown(frame, project)

root.mainloop()
