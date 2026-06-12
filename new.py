import tkinter as tk
from tkinter import ttk, filedialog, messagebox, simpledialog
import os
import subprocess

APP_FOLDER = os.path.dirname(os.path.abspath(__file__))

root = tk.Tk()
root.title("VNC Connector")
root.geometry("700x500")

notebook = ttk.Notebook(root)
notebook.pack(fill="both", expand=True)

project_tabs = {}
excluded_dirs = {"tcl8", "_tcl_data", "_tk_data", ".git", "__pycache__", "Private"}

def get_project_list():
    return [d for d in os.listdir(APP_FOLDER) if os.path.isdir(os.path.join(APP_FOLDER, d)) and d not in excluded_dirs]

def add_file_popup():
    popup = tk.Toplevel(root)
    popup.title("Add VNC File")
    popup.geometry("350x200")
    popup.resizable(False, False)

    tk.Label(popup, text="Select Existing Project or Enter New:").pack(pady=5)

    project_var = tk.StringVar()
    project_list = get_project_list()

    project_dropdown = ttk.Combobox(popup, textvariable=project_var, values=project_list, state="normal")  # Editable combobox
    if project_list:
        project_dropdown.set(project_list[0])
    project_dropdown.pack(pady=5)

    tk.Label(popup, text="Enter Bench Name:").pack(pady=5)
    bench_entry = tk.Entry(popup)
    bench_entry.pack(pady=5)

    def browse_and_add():
        selected_project = project_var.get().strip()
        if not selected_project:
            messagebox.showerror("Error", "Enter a project name.")
            return

        bench_name = bench_entry.get().strip()
        if not bench_name:
            messagebox.showerror("Error", "Enter bench name.")
            return

        file_path = filedialog.askopenfilename(title="Select VNC File", filetypes=[("VNC Files", "*.vnc")])
        if not file_path:
            return

        dest_folder = os.path.join(APP_FOLDER, selected_project)
        os.makedirs(dest_folder, exist_ok=True)

        dest_file = os.path.join(dest_folder, f"{bench_name}.vnc")
        if os.path.exists(dest_file):
            confirm = messagebox.askyesno(
                "Overwrite File?",
                f"A bench named '{bench_name}' already exists in project '{selected_project}'.\nDo you want to overwrite it?"
            )
            if not confirm:
                return

        try:
            with open(file_path, 'rb') as src, open(dest_file, 'wb') as dst:
                dst.write(src.read())
            messagebox.showinfo("Success", f"{bench_name}.vnc added to {selected_project}")
            if selected_project not in project_tabs:
                build_project_tab(selected_project)
            else:
                refresh_project_tab(selected_project)
            popup.destroy()
        except Exception as e:
            messagebox.showerror("Error", str(e))

    tk.Button(popup, text="Select VNC File & Add", command=browse_and_add).pack(pady=10)

def build_project_tab(project):
    if project in project_tabs:
        return

    outer_frame = tk.Frame(notebook)
    canvas = tk.Canvas(outer_frame)
    scrollbar = ttk.Scrollbar(outer_frame, orient="vertical", command=canvas.yview)
    scroll_frame = tk.Frame(canvas)

    scroll_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
    canvas.create_window((0, 0), window=scroll_frame, anchor="nw")
    canvas.configure(yscrollcommand=scrollbar.set)

    canvas.pack(side="left", fill="both", expand=True)
    scrollbar.pack(side="right", fill="y")

    notebook.add(outer_frame, text=project)
    project_tabs[project] = scroll_frame

    build_bench_buttons(scroll_frame, project)

def refresh_project_tab(project):
    if project in project_tabs:
        frame = project_tabs[project]
        for widget in frame.winfo_children():
            widget.destroy()
        build_bench_buttons(frame, project)

def build_bench_buttons(frame, project):
    files = [f for f in os.listdir(os.path.join(APP_FOLDER, project)) if f.endswith(".vnc")]

    for i, file in enumerate(files):
        bench_name = file[:-4]

        def connect_bench(f=file):
            file_path = os.path.join(APP_FOLDER, project, f)
            if os.path.exists(file_path):
                subprocess.Popen([file_path], shell=True)

        btn = tk.Button(frame, text=bench_name, width=20)
        btn.grid(row=i // 3, column=i % 3, padx=5, pady=5)
        btn.bind("<Button-1>", lambda e, f=file: connect_bench(f))
        btn.bind("<Button-3>", lambda e, f=file: show_context_menu(e, project, f))

def show_context_menu(event, project, file):
    menu = tk.Menu(root, tearoff=0)
    bench_name = file[:-4]

    def delete_bench():
        path = os.path.join(APP_FOLDER, project, file)
        if messagebox.askyesno("Delete", f"Delete '{bench_name}.vnc'?"):
            os.remove(path)
            refresh_project_tab(project)

    def rename_bench():
        new_name = simpledialog.askstring("Rename", "Enter new bench name:", initialvalue=bench_name)
        if new_name and new_name != bench_name:
            confirm = messagebox.askyesno("Confirm Rename", f"Rename '{bench_name}' to '{new_name}'?")
            if confirm:
                old_path = os.path.join(APP_FOLDER, project, file)
                new_path = os.path.join(APP_FOLDER, project, f"{new_name}.vnc")
                try:
                    os.rename(old_path, new_path)
                    refresh_project_tab(project)
                except Exception as e:
                    messagebox.showerror("Rename Failed", str(e))

    menu.add_command(label="Rename", command=rename_bench)
    menu.add_command(label="Delete", command=delete_bench)
    menu.post(event.x_root, event.y_root)

def load_initial_projects():
    for project in get_project_list():
        build_project_tab(project)

load_initial_projects()

menu_bar = tk.Menu(root)
file_menu = tk.Menu(menu_bar, tearoff=0)
file_menu.add_command(label="Add VNC File", command=add_file_popup)
menu_bar.add_cascade(label="File", menu=file_menu)
root.config(menu=menu_bar)

root.mainloop()
