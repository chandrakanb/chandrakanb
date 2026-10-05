import os
import sys
import shutil
import subprocess
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, simpledialog

if getattr(sys, "frozen", False):
    APP_FOLDER = os.path.dirname(os.path.abspath(sys.executable))
else:
    APP_FOLDER = os.path.dirname(os.path.abspath(__file__))
EXCLUDED_DIRS = {"tcl8", "_tcl_data", "_tk_data", ".git", "__pycache__", "Private"}


class VNCConnector:
    def __init__(self, root):
        self.root = root
        self.root.title("VNC Bench Manager")
        self.root.geometry("800x550")
        self.setup_theme()
        self.root.minsize(650, 450)

        self.notebook = ttk.Notebook(root)
        self.notebook.pack(fill="both", expand=True, padx=5, pady=(5, 0))
        self.project_tabs = {}
        self.status_var = tk.StringVar(value="Ready")

        self.build_menu()
        status_frame = ttk.Frame(root)
        status_frame.pack(fill="x", padx=5, pady=5)

        ttk.Label(
            status_frame,
            textvariable=self.status_var,
            style="Status.TLabel",
            anchor="w"
        ).pack(side="left", fill="x", expand=True)

        ttk.Label(
            status_frame,
            text="VNC Bench Manager v1.0  •  © 2026 CB",
            style="Copyright.TLabel",
            anchor="e"
        ).pack(side="right")

        self.load_initial_projects()


    def setup_theme(self):
        """Apply a dark blue-gray + teal theme that is easy on the eyes."""
        self.colors = {
            "bg": "#18212B",
            "panel": "#202B36",
            "button": "#2C3E50",
            "hover": "#168AAD",
            "accent": "#168AAD",
            "active": "#2A9D8F",
            "text": "#E8EDF2",
            "muted": "#AAB4BE",
            "status": "#141C24",
            "border": "#344454",
        }

        self.root.configure(bg=self.colors["bg"])

        style = ttk.Style(self.root)
        try:
            style.theme_use("clam")
        except tk.TclError:
            pass

        style.configure(
            "TFrame",
            background=self.colors["bg"]
        )
        style.configure(
            "Project.TFrame",
            background=self.colors["panel"]
        )
        style.configure(
            "TLabel",
            background=self.colors["bg"],
            foreground=self.colors["text"],
            font=("Segoe UI", 10)
        )
        style.configure(
            "TNotebook",
            background=self.colors["bg"],
            borderwidth=0
        )
        style.configure(
            "TNotebook.Tab",
            background=self.colors["panel"],
            foreground=self.colors["muted"],
            padding=(16, 8),
            font=("Segoe UI", 10)
        )
        # Reverse the tab appearance:
        # inactive tabs use the teal accent, while the selected tab
        # returns to the dark panel color for a recessed/active look.
        style.map(
            "TNotebook.Tab",
            background=[
                ("selected", self.colors["panel"]),
                ("active", self.colors["hover"])
            ],
            foreground=[
                ("selected", self.colors["text"]),
                ("active", "#FFFFFF")
            ],
            padding=[
                ("selected", (20, 11)),
                ("active", (17, 9)),
                ("!selected", (15, 8))
            ]
        )
        style.configure(
            "TScrollbar",
            background=self.colors["button"],
            troughcolor=self.colors["panel"],
            bordercolor=self.colors["panel"],
            arrowcolor=self.colors["muted"]
        )
        style.map(
            "TScrollbar",
            background=[("active", self.colors["hover"])]
        )
        style.configure(
            "TButton",
            background=self.colors["button"],
            foreground=self.colors["text"],
            bordercolor=self.colors["border"],
            padding=(10, 6),
            font=("Segoe UI", 10)
        )
        style.map(
            "TButton",
            background=[
                ("pressed", self.colors["active"]),
                ("active", self.colors["hover"])
            ],
            foreground=[
                ("disabled", self.colors["muted"]),
                ("active", "#FFFFFF")
            ]
        )
        style.configure(
            "TCombobox",
            fieldbackground=self.colors["panel"],
            background=self.colors["button"],
            foreground=self.colors["text"],
            arrowcolor=self.colors["text"]
        )
        style.configure(
            "TEntry",
            fieldbackground=self.colors["panel"],
            foreground=self.colors["text"],
            insertcolor=self.colors["text"]
        )
        style.configure(
            "Status.TLabel",
            background=self.colors["status"],
            foreground=self.colors["muted"],
            padding=(5, 2)
        )
        style.configure(
            "Copyright.TLabel",
            background=self.colors["status"],
            foreground="#7F8C98",
            padding=(8, 2),
            font=("Segoe UI", 9)
        )

    def build_menu(self):
        menu = tk.Menu(self.root)
        file_menu = tk.Menu(menu, tearoff=0)
        file_menu.add_command(label="Add VNC File", command=self.add_vnc_file)
        file_menu.add_command(label="New Project", command=self.new_project)
        file_menu.add_separator()
        file_menu.add_command(label="Refresh All", command=self.refresh_all)
        file_menu.add_separator()
        file_menu.add_command(label="Exit", command=self.root.destroy)
        project_menu = tk.Menu(menu, tearoff=0)
        project_menu.add_command(label="Project Settings",
                                 command=self.project_settings)
        project_menu.add_separator()
        project_menu.add_command(label="Refresh Project",
                                 command=self.refresh_current_project)
        project_menu.add_command(label="Open Project Folder",
                                 command=self.open_current_project_folder)
        project_menu.add_separator()
        project_menu.add_command(label="Delete Project",
                                 command=self.delete_current_project)
        menu.add_cascade(label="File", menu=file_menu)
        menu.add_cascade(label="Project", menu=project_menu)
        self.root.config(menu=menu)

    def get_project_list(self):
        if not os.path.isdir(APP_FOLDER):
            return []
        return sorted(
            [d for d in os.listdir(APP_FOLDER)
             if os.path.isdir(os.path.join(APP_FOLDER, d))
             and d not in EXCLUDED_DIRS and not d.startswith(".")],
            key=str.lower
        )

    def get_vnc_files(self, project):
        folder = os.path.join(APP_FOLDER, project)
        if not os.path.isdir(folder):
            return []
        return sorted(
            [f for f in os.listdir(folder)
             if f.lower().endswith(".vnc")
             and os.path.isfile(os.path.join(folder, f))],
            key=str.lower
        )

    def _mousewheel_scroll(self, event, canvas):
        """Scroll the project canvas with the mouse wheel."""
        if not canvas.winfo_exists():
            return "break"

        # Windows / macOS commonly provide event.delta.
        if event.delta:
            direction = -1 if event.delta > 0 else 1
            canvas.yview_scroll(direction * 3, "units")

        return "break"

    def _bind_mousewheel(self, widget, canvas):
        """Make mouse-wheel scrolling work over the canvas and its children."""
        widget.bind(
            "<MouseWheel>",
            lambda event, c=canvas: self._mousewheel_scroll(event, c)
        )

        # Linux wheel events (harmless on Windows).
        widget.bind(
            "<Button-4>",
            lambda event, c=canvas: self._mousewheel_scroll(
                type("WheelEvent", (), {"delta": 120})(), c
            )
        )
        widget.bind(
            "<Button-5>",
            lambda event, c=canvas: self._mousewheel_scroll(
                type("WheelEvent", (), {"delta": -120})(), c
            )
        )

    def build_project_tab(self, project):
        if project in self.project_tabs:
            return
        outer = tk.Frame(self.notebook, bg=self.colors["panel"])
        canvas = tk.Canvas(outer, bg=self.colors["panel"], highlightthickness=0, bd=0)
        scrollbar = ttk.Scrollbar(outer, orient="vertical", command=canvas.yview)
        frame = tk.Frame(canvas, bg=self.colors["panel"])
        frame.bind("<Configure>",
                   lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        # Mouse-wheel scrolling works anywhere over the project content.
        self._bind_mousewheel(canvas, canvas)
        self._bind_mousewheel(frame, canvas)
        self.notebook.add(outer, text=project)
        self.project_tabs[project] = frame
        self.build_bench_buttons(frame, project)

    def refresh_project_tab(self, project):
        if project not in self.project_tabs:
            self.build_project_tab(project)
            return
        frame = self.project_tabs[project]
        for widget in frame.winfo_children():
            widget.destroy()
        self.build_bench_buttons(frame, project)

    def build_bench_buttons(self, frame, project):
        files = self.get_vnc_files(project)
        if not files:
            tk.Label(frame,
                      text="No .vnc files in this project.\nUse File → Add VNC File.",
                      justify="center", bg=self.colors["panel"], fg=self.colors["muted"],
                      font=("Segoe UI", 10)).grid(row=0, column=0, padx=20, pady=30)
            return

        for i, filename in enumerate(files):
            button = tk.Button(
                frame,
                text=os.path.splitext(filename)[0],
                width=24,
                height=2,
                cursor="hand2",
                bg=self.colors["button"],
                fg=self.colors["text"],
                activebackground=self.colors["hover"],
                activeforeground="#FFFFFF",
                relief="flat",
                bd=0,
                highlightthickness=1,
                highlightbackground=self.colors["border"],
                highlightcolor=self.colors["accent"],
                font=("Segoe UI", 10)
            )
            button.grid(row=i // 3, column=i % 3, padx=8, pady=8, sticky="nsew")

            # The mouse wheel must also work while the pointer is directly
            # over a Bench button.
            self._bind_mousewheel(button, frame.master)

            def on_enter(event, b=button):
                b.configure(bg=self.colors["hover"])

            def on_leave(event, b=button):
                b.configure(bg=self.colors["button"])

            button.bind("<Enter>", on_enter)
            button.bind("<Leave>", on_leave)
            button.bind("<Button-1>",
                        lambda e, p=project, f=filename: self.launch_vnc(p, f))
            button.bind("<Button-3>",
                        lambda e, p=project, f=filename:
                        self.show_context_menu(e, p, f))

    def launch_vnc(self, project, filename):
        path = os.path.join(APP_FOLDER, project, filename)
        if not os.path.isfile(path):
            messagebox.showerror("VNC File Not Found", path)
            self.refresh_project_tab(project)
            return
        try:
            # The .vnc file is opened using the OS file association.
            # The installed VNC viewer performs the actual VNC connection.
            if os.name == "nt":
                os.startfile(path)
            elif os.name == "darwin":
                subprocess.Popen(["open", path])
            else:
                subprocess.Popen(["xdg-open", path])
            self.status_var.set(f"Launched: {project}\\{filename}")
        except Exception as exc:
            messagebox.showerror("Launch Failed", f"{path}\n\n{exc}")
            self.status_var.set("Launch failed")

    def project_settings(self):
        """Show settings for the currently selected project."""
        project = self.get_current_project()
        if not project:
            messagebox.showinfo("Project Settings",
                                "Please select a project tab first.",
                                parent=self.root)
            return

        project_path = os.path.join(APP_FOLDER, project)
        if not os.path.isdir(project_path):
            messagebox.showerror("Project Settings",
                                 "The selected project folder no longer exists.",
                                 parent=self.root)
            self.refresh_all()
            return

        popup = tk.Toplevel(self.root)
        popup.title(f"Project Settings - {project}")
        popup.geometry("470x300")
        popup.resizable(False, False)
        popup.configure(bg=self.colors["bg"])
        popup.transient(self.root)
        popup.grab_set()

        tk.Label(
            popup, text="Project Settings",
            bg=self.colors["bg"], fg=self.colors["text"],
            font=("Segoe UI", 14, "bold")
        ).pack(pady=(18, 12))

        form = tk.Frame(popup, bg=self.colors["bg"])
        form.pack(fill="x", padx=25)

        tk.Label(
            form, text="Project Name:",
            bg=self.colors["bg"], fg=self.colors["muted"],
            font=("Segoe UI", 10)
        ).grid(row=0, column=0, sticky="w", pady=8)

        name_var = tk.StringVar(value=project)
        name_entry = tk.Entry(
            form, textvariable=name_var, width=36,
            bg=self.colors["panel"], fg=self.colors["text"],
            insertbackground=self.colors["text"],
            relief="flat", highlightthickness=1,
            highlightbackground=self.colors["border"],
            highlightcolor=self.colors["accent"]
        )
        name_entry.grid(row=0, column=1, sticky="ew", padx=(15, 0), pady=8)

        tk.Label(
            form, text="Folder:",
            bg=self.colors["bg"], fg=self.colors["muted"],
            font=("Segoe UI", 10)
        ).grid(row=1, column=0, sticky="w", pady=8)

        folder_label = tk.Label(
            form, text=project_path,
            bg=self.colors["bg"], fg=self.colors["text"],
            font=("Segoe UI", 9), anchor="w", justify="left",
            wraplength=300
        )
        folder_label.grid(row=1, column=1, sticky="w", padx=(15, 0), pady=8)

        vnc_count = len(self.get_vnc_files(project))
        tk.Label(
            form, text="VNC Benches:",
            bg=self.colors["bg"], fg=self.colors["muted"],
            font=("Segoe UI", 10)
        ).grid(row=2, column=0, sticky="w", pady=8)

        tk.Label(
            form, text=str(vnc_count),
            bg=self.colors["bg"], fg=self.colors["text"],
            font=("Segoe UI", 10)
        ).grid(row=2, column=1, sticky="w", padx=(15, 0), pady=8)

        form.columnconfigure(1, weight=1)

        def save_settings():
            new_name = name_var.get().strip()

            if not new_name:
                messagebox.showerror(
                    "Invalid Project Name",
                    "Project name cannot be empty.",
                    parent=popup
                )
                return

            if os.path.basename(new_name) != new_name:
                messagebox.showerror(
                    "Invalid Project Name",
                    "Project name cannot contain folder separators.",
                    parent=popup
                )
                return

            if new_name == project:
                popup.destroy()
                return

            new_path = os.path.join(APP_FOLDER, new_name)

            if os.path.exists(new_path):
                messagebox.showerror(
                    "Rename Failed",
                    f"A project named '{new_name}' already exists.",
                    parent=popup
                )
                return

            if not messagebox.askyesno(
                "Rename Project",
                f"Rename project '{project}' to '{new_name}'?",
                parent=popup
            ):
                return

            try:
                os.rename(project_path, new_path)

                # Rebuild every tab so the UI exactly matches disk.
                popup.destroy()
                self.project_tabs.clear()

                for tab_id in self.notebook.tabs():
                    try:
                        self.notebook.forget(tab_id)
                    except tk.TclError:
                        pass

                self.load_initial_projects()

                if new_name in self.project_tabs:
                    self.select_project_tab(new_name)

                self.root.update_idletasks()
                self.status_var.set(f"Project renamed to: {new_name}")
            except Exception as exc:
                messagebox.showerror(
                    "Rename Failed",
                    str(exc),
                    parent=popup
                )

        buttons = tk.Frame(popup, bg=self.colors["bg"])
        buttons.pack(side="bottom", fill="x", padx=25, pady=18)

        tk.Button(
            buttons, text="Cancel", command=popup.destroy,
            bg=self.colors["button"], fg=self.colors["text"],
            activebackground=self.colors["hover"],
            activeforeground="#FFFFFF", relief="flat", bd=0,
            padx=16, pady=7
        ).pack(side="right", padx=(8, 0))

        tk.Button(
            buttons, text="Save", command=save_settings,
            bg=self.colors["accent"], fg="#FFFFFF",
            activebackground=self.colors["hover"],
            activeforeground="#FFFFFF", relief="flat", bd=0,
            padx=16, pady=7
        ).pack(side="right")

        name_entry.focus_set()
        name_entry.selection_range(0, tk.END)

    def delete_project(self, project):
        """Delete a project folder and rebuild the Notebook tabs."""
        if not project:
            return

        if project in EXCLUDED_DIRS:
            messagebox.showerror(
                "Delete Project",
                "This project cannot be deleted.",
                parent=self.root
            )
            return

        project_path = os.path.join(APP_FOLDER, project)

        if not os.path.isdir(project_path):
            messagebox.showerror(
                "Delete Project",
                f"Project folder was not found:\n{project_path}",
                parent=self.root
            )
            self.refresh_all()
            return

        if not messagebox.askyesno(
            "Delete Project",
            f"Delete project '{project}' and ALL files inside it?\n\n"
            "This action cannot be undone.",
            parent=self.root
        ):
            return

        try:
            # Delete the actual project directory from disk.
            shutil.rmtree(project_path)

            # Remove every Notebook tab and rebuild from the folders that
            # actually remain on disk. This avoids stale tab references.
            for tab_id in list(self.notebook.tabs()):
                try:
                    self.notebook.forget(tab_id)
                except tk.TclError:
                    pass

            self.project_tabs.clear()
            self.load_initial_projects()

            # Select the first remaining project, if any.
            if self.project_tabs:
                first_project = next(iter(self.project_tabs))
                self.select_project_tab(first_project)

            self.root.update_idletasks()
            self.status_var.set(f"Deleted project: {project}")
            self.root.update_idletasks()

        except PermissionError as exc:
            messagebox.showerror(
                "Delete Project Failed",
                f"Windows denied access to the project folder.\n\n"
                f"Close any files/programs using the folder and try again.\n\n{exc}",
                parent=self.root
            )
        except OSError as exc:
            messagebox.showerror(
                "Delete Project Failed",
                f"Could not delete '{project}'.\n\n{exc}",
                parent=self.root
            )

    def delete_current_project(self):
        """Delete the currently selected project."""
        project = self.get_current_project()
        if not project:
            messagebox.showinfo(
                "Delete Project",
                "Please select a project tab first.",
                parent=self.root
            )
            return

        self.delete_project(project)

    def add_vnc_file(self):
        popup = tk.Toplevel(self.root)
        popup.title("Add VNC File")
        popup.geometry("400x240")
        popup.resizable(False, False)
        popup.configure(bg=self.colors["bg"])
        popup.transient(self.root)
        popup.grab_set()

        tk.Label(popup, text="Select Existing Project or Enter New:",
                 bg=self.colors["bg"], fg=self.colors["text"],
                 font=("Segoe UI", 10)).pack(pady=(15, 5))
        project_var = tk.StringVar()
        projects = self.get_project_list()
        combo = ttk.Combobox(popup, textvariable=project_var, values=projects,
                             state="normal", width=35)
        if projects:
            combo.set(projects[0])
        combo.pack(pady=5)

        tk.Label(popup, text="Enter Bench Name:",
                 bg=self.colors["bg"], fg=self.colors["text"],
                 font=("Segoe UI", 10)).pack(pady=(10, 5))
        bench_entry = ttk.Entry(popup, width=38)
        bench_entry.pack(pady=5)

        def add():
            project = project_var.get().strip()
            bench = bench_entry.get().strip()
            if not project or os.path.basename(project) != project:
                messagebox.showerror("Invalid Project", "Enter a valid project name.",
                                     parent=popup)
                return
            if not bench or os.path.basename(bench) != bench:
                messagebox.showerror("Invalid Bench Name",
                                     "Enter a valid bench name.", parent=popup)
                return

            source = filedialog.askopenfilename(
                parent=popup, title="Select VNC File",
                filetypes=[("VNC Files", "*.vnc"), ("All Files", "*.*")]
            )
            if not source:
                return

            folder = os.path.join(APP_FOLDER, project)
            os.makedirs(folder, exist_ok=True)
            destination = os.path.join(folder, f"{bench}.vnc")

            if os.path.exists(destination) and not messagebox.askyesno(
                "Overwrite File?",
                f"'{bench}.vnc' already exists in '{project}'. Overwrite?",
                parent=popup):
                return

            try:
                shutil.copyfile(source, destination)
                if project not in self.project_tabs:
                    self.build_project_tab(project)
                else:
                    self.refresh_project_tab(project)
                self.select_project_tab(project)
                self.status_var.set(f"Added {bench}.vnc to {project}")
                popup.destroy()
                messagebox.showinfo("Success", f"{bench}.vnc added to {project}.")
            except Exception as exc:
                messagebox.showerror("Error", str(exc), parent=popup)

        ttk.Button(popup, text="Select VNC File & Add", command=add).pack(pady=18)

    def new_project(self):
        name = simpledialog.askstring("New Project", "Enter project name:",
                                      parent=self.root)
        if not name:
            return
        name = name.strip()
        if os.path.basename(name) != name:
            messagebox.showerror("Invalid Project",
                                 "Project name cannot contain folder separators.")
            return
        path = os.path.join(APP_FOLDER, name)
        try:
            os.makedirs(path, exist_ok=True)
            if name not in self.project_tabs:
                self.build_project_tab(name)
            self.select_project_tab(name)
            self.status_var.set(f"Project ready: {name}")
        except Exception as exc:
            messagebox.showerror("Project Creation Failed", str(exc))

    def show_context_menu(self, event, project, filename):
        menu = tk.Menu(self.root, tearoff=0)
        menu.add_command(label="Open VNC",
                         command=lambda: self.launch_vnc(project, filename))
        menu.add_command(label="Rename",
                         command=lambda: self.rename_bench(project, filename))
        menu.add_command(label="Delete",
                         command=lambda: self.delete_bench(project, filename))
        menu.add_separator()
        menu.add_command(label="Open Project Folder",
                         command=lambda: self.open_project_folder(project))
        try:
            menu.tk_popup(event.x_root, event.y_root)
        finally:
            menu.grab_release()

    def rename_bench(self, project, filename):
        old = os.path.splitext(filename)[0]
        new = simpledialog.askstring("Rename", "Enter new bench name:",
                                     initialvalue=old, parent=self.root)
        if not new:
            return
        new = new.strip()
        if not new or os.path.basename(new) != new:
            messagebox.showerror("Invalid Name", "Enter a valid bench name.")
            return
        if new == old:
            return
        old_path = os.path.join(APP_FOLDER, project, filename)
        new_path = os.path.join(APP_FOLDER, project, f"{new}.vnc")
        if os.path.exists(new_path):
            messagebox.showerror("Rename Failed", f"{new}.vnc already exists.")
            return
        if not messagebox.askyesno("Confirm Rename",
                                   f"Rename '{old}' to '{new}'?",
                                   parent=self.root):
            return
        try:
            os.rename(old_path, new_path)
            self.refresh_project_tab(project)
            self.status_var.set(f"Renamed to {new}.vnc")
        except Exception as exc:
            messagebox.showerror("Rename Failed", str(exc))

    def delete_bench(self, project, filename):
        name = os.path.splitext(filename)[0]
        path = os.path.join(APP_FOLDER, project, filename)
        if not messagebox.askyesno("Delete",
                                   f"Delete '{name}.vnc'?\n\nThis cannot be undone.",
                                   parent=self.root):
            return
        try:
            os.remove(path)
            self.refresh_project_tab(project)
            self.root.update_idletasks()
            self.status_var.set(f"Deleted {project}\\{filename}")
            self.root.update_idletasks()
        except Exception as exc:
            messagebox.showerror("Delete Failed", str(exc))

    def load_initial_projects(self):
        for project in self.get_project_list():
            self.build_project_tab(project)
        if not self.project_tabs:
            self.status_var.set("No projects found. Use File → New Project.")

    def get_current_project(self):
        selected = self.notebook.select()
        for project, content_frame in self.project_tabs.items():
            # content_frame is embedded inside a Canvas, which is inside
            # the actual Notebook tab (outer frame).
            outer_frame = content_frame.master.master
            if str(outer_frame) == selected:
                return project
        return None

    def select_project_tab(self, project):
        content_frame = self.project_tabs.get(project)
        if content_frame:
            # Notebook manages the outer frame, not the embedded content frame
            # or its Canvas.
            outer_frame = content_frame.master.master
            self.notebook.select(outer_frame)

    def refresh_current_project(self):
        project = self.get_current_project()
        if project:
            self.refresh_project_tab(project)
            self.status_var.set(f"Refreshed project: {project}")
        else:
            self.status_var.set("No project selected")

    def refresh_all(self):
        current = self.get_current_project()
        discovered = set(self.get_project_list())

        for project in sorted(discovered - set(self.project_tabs), key=str.lower):
            self.build_project_tab(project)

        for project in sorted(discovered & set(self.project_tabs), key=str.lower):
            self.refresh_project_tab(project)

        for project in list(self.project_tabs):
            if project not in discovered:
                content_frame = self.project_tabs.pop(project)
                try:
                    outer_frame = content_frame.master.master
                    self.notebook.forget(outer_frame)
                except tk.TclError:
                    pass

        if current in self.project_tabs:
            self.select_project_tab(current)
        self.status_var.set("All projects refreshed")

    def open_current_project_folder(self):
        project = self.get_current_project()
        if project:
            self.open_project_folder(project)
        else:
            messagebox.showinfo("Project", "No project tab is selected.")

    def open_project_folder(self, project):
        folder = os.path.join(APP_FOLDER, project)
        if not os.path.isdir(folder):
            messagebox.showerror("Folder Not Found", folder)
            return
        try:
            if os.name == "nt":
                os.startfile(folder)
            elif os.name == "darwin":
                subprocess.Popen(["open", folder])
            else:
                subprocess.Popen(["xdg-open", folder])
            self.status_var.set(f"Opened folder: {folder}")
        except Exception as exc:
            messagebox.showerror("Open Folder Failed", str(exc))


def main():
    root = tk.Tk()
    VNCConnector(root)
    root.mainloop()


if __name__ == "__main__":
    main()
