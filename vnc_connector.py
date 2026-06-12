import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import os
import subprocess
import shutil

APP_FOLDER = os.path.dirname(os.path.abspath(__file__))

root = tk.Tk()
root.title("VNC Connector")
root.geometry("600x400")

notebook = ttk.Notebook(root)
notebook.pack(fill="both", expand=True)

project_tabs = {}

frame_add = tk.Frame(notebook)
notebook.add(frame_add, text="Add New File")

project_var = tk.StringVar()
bench_var = tk.StringVar()
file_path_var = tk.StringVar()

excluded_dirs = {"tcl8", "_tcl_data", "_tk_data", ".git", "__pycache__", "Private"}
project_options = [d for d in os.listdir(APP_FOLDER) if os.path.isdir(os.path.join(APP_FOLDER, d)) and d not in excluded_dirs]
project_var.set(project_options[0] if project_options else "")

project_dropdown = ttk.Combobox(frame_add, textvariable=project_var, values=project_options, state="readonly")
project_dropdown.pack(pady=5, ipadx=5, ipady=2)

new_project_frame = tk.Frame(frame_add)
new_project_label = tk.Label(new_project_frame, text="New Project Name:")
new_project_entry = tk.Entry(new_project_frame)

bench_label = tk.Label(frame_add, text="Bench Name:")
bench_entry = tk.Entry(frame_add, textvariable=bench_var)

file_button = tk.Button(frame_add, text="Select VNC File", command=lambda: file_path_var.set(filedialog.askopenfilename(filetypes=[("VNC Files", "*.vnc")]) or file_path_var.get()))
file_label = tk.Label(frame_add, textvariable=file_path_var)

add_button = tk.Button(frame_add, text="Add File", command=lambda: add_file())

bench_label.pack(pady=2)
bench_entry.pack(pady=2)
file_button.pack(pady=5)
file_label.pack(pady=2)
add_button.pack(pady=5)

def update_project_dropdown():
    global project_options
    excluded_dirs = {"tcl8", "_tcl_data", "_tk_data", ".git", "__pycache__", "Private"}
    project_options = [d for d in os.listdir(APP_FOLDER) if os.path.isdir(os.path.join(APP_FOLDER, d)) and d not in excluded_dirs]
    project_dropdown["values"] = project_options
    if project_options:
        project_var.set(project_options[0])

def add_file():
    project = project_var.get()
    if project == "Add New Project":
        project = new_project_entry.get().strip()
        if not project:
            messagebox.showerror("Error", "Please enter a new project name.")
            return
        os.makedirs(os.path.join(APP_FOLDER, project), exist_ok=True)
        build_project_tab(project)
        update_project_dropdown()

    bench_name = bench_var.get().strip()
    file_path = file_path_var.get()

    if not bench_name or not file_path:
        messagebox.showerror("Error", "Please enter bench name and select a file.")
        return

    dest_path = os.path.join(APP_FOLDER, project, f"{bench_name}.vnc")
    try:
        with open(file_path, 'rb') as src, open(dest_path, 'wb') as dst:
            dst.write(src.read())
        messagebox.showinfo("Success", f"{bench_name}.vnc added to {project}")
        refresh_dropdown_list(project)
    except Exception as e:
        messagebox.showerror("Error", str(e))

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

    def delete_file():
        selected = var.get()
        if selected:
            file_path = os.path.join(APP_FOLDER, project, f"{selected}.vnc")
            if os.path.exists(file_path):
                os.remove(file_path)
                refresh_dropdown_list(project)
    
    def delete_current_project():
        selected_tab = notebook.select()
        tab_text = notebook.tab(selected_tab, "text")
        
        if tab_text in ("Add New File", "+ Add New Project"):
            messagebox.showerror("Error", "Cannot delete this tab.")
            return

        confirm = messagebox.askyesno("Confirm Delete", f"Are you sure you want to delete the project '{tab_text}'?")
        if confirm:
            project_path = os.path.join(APP_FOLDER, tab_text)
            if os.path.exists(project_path):
                try:
                    shutil.rmtree(project_path)
                    notebook.forget(selected_tab)
                    messagebox.showinfo("Deleted", f"Project '{tab_text}' has been deleted.")
                except Exception as e:
                    messagebox.showerror("Error", f"Failed to delete project: {e}")
            else:
                messagebox.showerror("Error", f"Project folder not found: {tab_text}")
    
    tk.Button(frame, text="Connect", command=connect).pack(pady=5)
    tk.Button(frame, text="Delete .vnc file", command=delete_file).pack(pady=5)
    tk.Button(frame, text="Delete Project", command=delete_current_project).pack(pady=5)

def build_project_tab(project):
    if project in project_tabs:
        return
    frame = tk.Frame(notebook)
    notebook.insert(notebook.index("end") - 1, frame, text=project)
    project_tabs[project] = frame
    build_dropdown(frame, project)

update_project_dropdown()
project_dropdown.bind("<<ComboboxSelected>>", lambda e: show_or_hide_new_project_entry())

def show_or_hide_new_project_entry():
    if project_var.get() == "Add New Project":
        new_project_frame.pack(before=bench_label, pady=2)
        new_project_label.pack()
        new_project_entry.pack()
    else:
        new_project_frame.pack_forget()

if "Add New Project" not in project_options:
    project_options.append("Add New Project")
    project_dropdown["values"] = project_options

for project in project_options:
    if project != "Add New Project":
        build_project_tab(project)

show_or_hide_new_project_entry()

root.mainloop()
