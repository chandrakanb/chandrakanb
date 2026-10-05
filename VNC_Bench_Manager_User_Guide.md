# VNC Bench Manager — User Guide

**Version:** 1.0  
**Copyright:** © 2026 CB

---

## 1. What is VNC Bench Manager?

VNC Bench Manager is a Windows desktop utility for organizing and launching VNC connection files.

It organizes connections using:

**Project → Bench → `.vnc` file → VNC Viewer**

The application does not replace the VNC Viewer. When you click a bench, Windows opens the associated `.vnc` file using the configured VNC application.

---

## 2. Requirements

- Windows PC
- Python 3.x
- A VNC Viewer/application associated with `.vnc` files
- Valid `.vnc` connection files

For a normal Python installation, no additional Python package is required beyond the standard Tkinter installation.

---

## 3. Recommended Folder Structure

Keep the Python file and project folders together:

```text
VNC_Bench_Manager/
│
├── VNC_Bench_Manager.py
│
├── CNT PF23 Automation/
│   ├── Bench01.vnc
│   ├── Bench02.vnc
│   └── Bench03.vnc
│
├── CNT PF26 Automation/
│   ├── Bench01.vnc
│   └── Bench02.vnc
│
├── DRT Automation/
│   ├── Bench01.vnc
│   └── Bench02.vnc
│
└── OSV Automation/
    └── Bench01.vnc
```

### Important

Each folder directly inside the application folder is treated as a **Project** unless it is one of the excluded application/system folders.

The application ignores:

```text
tcl8
_tcl_data
_tk_data
.git
__pycache__
Private
```

---

## 4. Starting the Application

Open Command Prompt or PowerShell in the application folder and run:

```bat
python VNC_Bench_Manager.py
```

The main window opens as:

**VNC Bench Manager**

---

## 5. Main Interface

The application has:

- **File** menu
- **Project** menu
- Project tabs
- Bench buttons
- Vertical scrollbar
- Mouse-wheel scrolling
- Status bar
- Version/copyright footer

Example:

```text
File    Project

[CNT PF23 Automation] [CNT PF26 Automation] [DRT Automation] [OSV Automation]

┌──────────┐  ┌──────────┐  ┌──────────┐
│ Bench01  │  │ Bench02  │  │ Bench03  │
└──────────┘  └──────────┘  └──────────┘

Ready                         VNC Bench Manager v1.0 • © 2026 CB
```

---

# 6. Opening a VNC Connection

1. Select the required Project tab.
2. Find the required Bench.
3. Left-click the Bench button.
4. Windows opens the `.vnc` file.
5. Your configured VNC Viewer handles the actual connection.

Example:

```text
DRT Automation
      ↓
Bench01
      ↓
Bench01.vnc
      ↓
VNC Viewer
      ↓
Remote machine
```

The connection details such as hostname/IP, port, and other VNC settings come from the `.vnc` file.

---

# 7. Scrolling Through Many Benches

If a project contains many VNC files, the project area becomes scrollable.

You can scroll using:

### Mouse wheel

Place the mouse over:

- Bench buttons
- Empty project area
- Project content area

Then use the mouse wheel.

### Scrollbar

You can also drag the vertical scrollbar on the right side.

Both methods can be used together.

---

# 8. Adding a VNC File

Go to:

```text
File → Add VNC File
```

Then:

1. Select an existing Project or enter a new Project name.
2. Enter the Bench Name.
3. Select the `.vnc` file.
4. Click **Select VNC File & Add**.

Example:

```text
Project:    DRT Automation
Bench Name: Bench05
VNC File:   C:\VNC\Bench05.vnc
```

The application copies the selected file to:

```text
DRT Automation\Bench05.vnc
```

The project UI refreshes automatically.

---

# 9. Creating a New Project

Use:

```text
File → New Project
```

Enter a project name.

For example:

```text
Honda Automation
```

The application creates:

```text
Honda Automation/
```

and adds the project as a new tab.

You can then add VNC files using:

```text
File → Add VNC File
```

---

# 10. Adding a Bench to a New Project

You do not have to create the project separately.

From:

```text
File → Add VNC File
```

enter a project name that does not exist yet.

The application creates the project folder automatically.

---

# 11. Bench Context Menu

Right-click a Bench button.

The available actions are:

```text
Open VNC
Rename
Delete
──────────────
Open Project Folder
```

### Open VNC

Launches the selected `.vnc` file.

### Rename

Changes the Bench filename.

Example:

```text
Bench01.vnc
      ↓
Main_Bench.vnc
```

### Delete

Deletes the selected `.vnc` file after confirmation.

### Open Project Folder

Opens the project directory in Windows Explorer.

---

# 12. Renaming a Bench

1. Right-click the Bench.
2. Select **Rename**.
3. Enter the new name.
4. Confirm the rename.

The application changes:

```text
OldName.vnc
```

to:

```text
NewName.vnc
```

The UI refreshes automatically.

---

# 13. Deleting a Bench

1. Right-click the Bench.
2. Select **Delete**.
3. Confirm the deletion.

The `.vnc` file is removed from the project folder.

The project tab is refreshed immediately.

---

# 14. Project Menu

Project-level operations are available from:

```text
Project
```

The menu contains:

```text
Project Settings
────────────────────
Refresh Project
Open Project Folder
────────────────────
Delete Project
```

---

# 15. Project Settings

Select:

```text
Project → Project Settings
```

The settings window displays:

- Project Name
- Project Folder
- Number of VNC Benches

### Rename a Project

Change the Project Name and click **Save**.

Example:

```text
DRT Automation
       ↓
DRT Automation Test
```

The physical folder is renamed and the project tabs are rebuilt.

The VNC files inside the project remain intact.

---

# 16. Refresh Project

Use:

```text
Project → Refresh Project
```

This rescans the selected project and rebuilds its Bench buttons.

Useful when a `.vnc` file has been manually added, removed, or renamed through Windows Explorer.

---

# 17. Open Project Folder

Use:

```text
Project → Open Project Folder
```

Windows Explorer opens the selected project's directory.

---

# 18. Delete Project

Use:

```text
Project → Delete Project
```

The application asks for confirmation.

If confirmed:

1. The project folder is deleted.
2. All `.vnc` files inside it are deleted.
3. The project tab is removed.
4. Project tabs are rebuilt from disk.
5. The remaining project is selected automatically.

### Warning

Deleting a project deletes the actual folder and its contents.

Make sure important `.vnc` files are backed up before deleting a project.

---

# 19. File Persistence

The application uses the actual Windows folder structure as its database.

There is no separate project database to maintain.

For example:

```text
VNC_Bench_Manager/
└── Honda/
    ├── Bench01.vnc
    └── Bench02.vnc
```

If the application is closed and reopened, it scans the folders again and recreates the project/bench interface.

---

# 20. VNC File Security

A `.vnc` file may contain connection information and, depending on the VNC software/configuration, authentication-related information.

Treat `.vnc` files as potentially sensitive.

Do not share them publicly unless you have verified that they contain no sensitive credentials or connection information.

---

# 21. Troubleshooting

## Application does not start

Run:

```bat
python VNC_Bench_Manager.py
```

from Command Prompt so that any Python error is visible.

---

## Bench button does not open VNC Viewer

Check:

1. The `.vnc` file exists.
2. Double-clicking the `.vnc` file in Windows Explorer opens your VNC Viewer.
3. `.vnc` is correctly associated with your VNC application.
4. The `.vnc` file contains valid connection information.

---

## Project is not shown

Check that the project is a folder directly inside the application directory.

Correct:

```text
VNC_Bench_Manager/
└── MyProject/
```

Not:

```text
VNC_Bench_Manager/
└── Projects/
    └── MyProject/
```

unless the application is specifically changed to support nested projects.

---

## Bench is not shown

Check that the file has the extension:

```text
.vnc
```

For example:

```text
Bench01.vnc
```

---

## Mouse wheel does not scroll

Make sure the project contains enough Bench buttons to require scrolling.

Then place the mouse over the project/Bench area and use the wheel.

The scrollbar remains available as a fallback.

---

## Project deletion fails

If Windows reports an access/permission problem:

1. Close any VNC Viewer or program using files in the project.
2. Close Windows Explorer windows opened inside that project.
3. Try deleting the project again.

---

# 22. Recommended Backup

Keep a backup of the complete folder:

```text
VNC_Bench_Manager/
```

especially the `.vnc` files.

A simple backup can be:

```text
VNC_Bench_Manager_Backup/
```

or a ZIP archive.

---

# 23. Version Information

Current version:

**VNC Bench Manager v1.0**

Copyright:

**© 2026 CB**

