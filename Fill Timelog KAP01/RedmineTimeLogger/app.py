from redminelib import Redmine
from datetime import datetime, timedelta
import tkinter as tk
from tkinter import ttk, messagebox
import logging, os, sys, json

# ===================== CONFIG =====================
BASE_URL = "https://kap01.kpit.com/kap"
TASK_FILE = "tasks.json"
CONFIG_FILE = "config.json"

# ===================== LOGGING =====================
os.makedirs("logs", exist_ok=True)
logfile = f"logs/{datetime.now().strftime('%Y%m%d_%H%M%S')}.log"

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s | %(levelname)s | %(message)s",
    handlers=[
        logging.FileHandler(logfile, encoding="utf-8"),
        logging.StreamHandler(sys.stdout)
    ]
)
log = logging.getLogger()

# ===================== CONFIG LOAD/SAVE =====================
def load_config():
    if os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        except:
            return {"api_key": ""}
    return {"api_key": ""}

def save_config(cfg):
    with open(CONFIG_FILE, "w", encoding="utf-8") as f:
        json.dump(cfg, f, indent=2)

config = load_config()

# ===================== TASK LOAD/SAVE =====================
def load_tasks():
    if os.path.exists(TASK_FILE):
        try:
            with open(TASK_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        except:
            messagebox.showerror("Error", "tasks.json is invalid. Fix or delete it.")
            return []
    return []

def save_tasks():
    with open(TASK_FILE, "w", encoding="utf-8") as f:
        json.dump(tasks, f, indent=2)

tasks = load_tasks()

# ===================== DATES =====================
dates = [(datetime.today() - timedelta(days=i)).strftime("%Y-%m-%d") for i in range(10)]

# ===================== UI ROOT =====================
root = tk.Tk()
root.title("Redmine Time Logger")
root.geometry("700x560")

notebook = ttk.Notebook(root)
notebook.pack(expand=True, fill="both")

# ==================================================
# TAB 1 — TIME LOGGING
# ==================================================
tab_log = ttk.Frame(notebook)
notebook.add(tab_log, text="Time Logging")

date_vars = {}
select_all_var = tk.BooleanVar()

def toggle_all():
    for v in date_vars.values():
        v.set(select_all_var.get())

ttk.Checkbutton(
    tab_log,
    text="Select All Dates",
    variable=select_all_var,
    command=toggle_all
).pack(pady=5)

date_frame = ttk.Frame(tab_log)
date_frame.pack()

for d in dates:
    v = tk.BooleanVar()
    ttk.Checkbutton(date_frame, text=d, variable=v).pack(anchor="w")
    date_vars[d] = v

progress = ttk.Progressbar(tab_log, length=560)
progress.pack(pady=15)

def log_time():
    selected_dates = [d for d, v in date_vars.items() if v.get()]
    if not selected_dates:
        messagebox.showwarning("Missing", "Select at least one date")
        return

    api_key = config.get("api_key", "").strip()
    if not api_key:
        messagebox.showerror("API Key Missing", "Add API key in API Key tab")
        return

    if not tasks:
        messagebox.showwarning("No Tasks", "Add at least one task")
        return

    redmine = Redmine(BASE_URL, key=api_key)
    total = len(selected_dates) * len(tasks)
    progress["maximum"] = total
    progress["value"] = 0

    for d in selected_dates:
        for t in tasks:
            # Skip tasks with 0 hours
            if t["hours"] <= 0:
                log.info(f"SKIPPED | Issue {t['issue']} | {d} | 0 hours")
                progress["value"] += 1
                root.update_idletasks()
                continue

            try:
                custom_fields = [
                    {"id": 46, "value": t["cat"]},
                    {"id": 48, "value": "India"},
                    {"id": 49, "value": "Pune"}
                ]
                if t["cat"] != "TEAM MEETING" and t["type"]:
                    custom_fields.append({"id": 47, "value": t["type"]})

                redmine.time_entry.create(
                    issue_id=t["issue"],
                    spent_on=d,
                    hours=t["hours"],
                    comments=t["comment"],
                    custom_fields=custom_fields
                )
                log.info(f"OK | Issue {t['issue']} | {d} | {t['hours']}h")
            except Exception as e:
                log.error(f"FAIL | Issue {t['issue']} | {d} | {e}")

            progress["value"] += 1
            root.update_idletasks()

    messagebox.showinfo("Success", "Time logged successfully")

ttk.Button(tab_log, text="Log Time", command=log_time).pack(pady=10)

# ==================================================
# TAB 2 — TASKS
# ==================================================
tab_tasks = ttk.Frame(notebook)
notebook.add(tab_tasks, text="Tasks")

header = ttk.Frame(tab_tasks)
header.pack(fill="x", padx=10, pady=5)

for text, w in [
    ("Issue", 11),
    ("Hours", 9),
    ("Comment", 27),
    ("Category", 19),
    ("Type", 12),
]:
    ttk.Label(header, text=text, width=w).pack(side="left")

form = ttk.Frame(tab_tasks)
form.pack(fill="x", padx=10)

issue_e = ttk.Entry(form, width=10)
hours_e = ttk.Entry(form, width=8)
comment_e = ttk.Entry(form, width=26)
cat_e = ttk.Entry(form, width=18)
type_e = ttk.Entry(form, width=12)

for w in (issue_e, hours_e, comment_e, cat_e, type_e):
    w.pack(side="left", padx=2)

task_list = tk.Listbox(tab_tasks, height=9)
task_list.pack(fill="x", padx=10, pady=6)

def refresh_tasks():
    task_list.delete(0, tk.END)
    for t in tasks:
        task_list.insert(
            tk.END,
            f"{t['issue']} | {t['hours']}h | {t['comment']} | {t['cat']} | {t['type']}"
        )

def clear_task_entries():
    for e in (issue_e, hours_e, comment_e, cat_e, type_e):
        e.delete(0, tk.END)

def add_task():
    try:
        tasks.append({
            "issue": int(issue_e.get()),
            "hours": float(hours_e.get()),
            "comment": comment_e.get(),
            "cat": cat_e.get(),
            "type": type_e.get()
        })
        save_tasks()
        refresh_tasks()
        clear_task_entries()
    except:
        messagebox.showerror("Invalid Task", "Check task values")

def edit_task():
    sel = task_list.curselection()
    if not sel:
        return
    t = tasks.pop(sel[0])
    issue_e.insert(0, t["issue"])
    hours_e.insert(0, t["hours"])
    comment_e.insert(0, t["comment"])
    cat_e.insert(0, t["cat"])
    type_e.insert(0, t["type"])
    save_tasks()
    refresh_tasks()

def delete_task():
    sel = task_list.curselection()
    if sel:
        tasks.pop(sel[0])
        save_tasks()
        refresh_tasks()

btns = ttk.Frame(tab_tasks)
btns.pack(pady=5)

ttk.Button(btns, text="Add", command=add_task).pack(side="left", padx=5)
ttk.Button(btns, text="Edit", command=edit_task).pack(side="left", padx=5)
ttk.Button(btns, text="Delete", command=delete_task).pack(side="left", padx=5)

refresh_tasks()

# ==================================================
# TAB 3 — API KEY
# ==================================================
tab_api = ttk.Frame(notebook)
notebook.add(tab_api, text="API Key")

ttk.Label(tab_api, text="Redmine API Key").pack(pady=15)

api_entry = ttk.Entry(tab_api, show="*", width=50)
api_entry.pack()

if config.get("api_key"):
    api_entry.insert(0, config["api_key"])

def save_api_key():
    config["api_key"] = api_entry.get().strip()
    save_config(config)
    messagebox.showinfo("Saved", "API key saved to config.json")

ttk.Button(tab_api, text="Save API Key", command=save_api_key).pack(pady=8)

# ==================================================
# TAB 4 — LOGS
# ==================================================
tab_logs = ttk.Frame(notebook)
notebook.add(tab_logs, text="Logs")

filter_var = tk.StringVar(value="ALL")

filter_frame = ttk.Frame(tab_logs)
filter_frame.pack(pady=4)

for f in ("ALL", "INFO", "ERROR"):
    ttk.Radiobutton(filter_frame, text=f, variable=filter_var, value=f).pack(side="left")

log_text = tk.Text(tab_logs, wrap="none", height=22, state="disabled")
log_text.pack(fill="both", expand=True, padx=10)

def load_latest_log():
    files = sorted(os.listdir("logs"), reverse=True)
    if not files:
        return
    level = filter_var.get()
    with open(os.path.join("logs", files[0]), encoding="utf-8") as f:
        lines = f.readlines()

    log_text.config(state="normal")
    log_text.delete("1.0", tk.END)
    for line in lines:
        if level == "ALL" or f"| {level} |" in line:
            log_text.insert(tk.END, line)
    log_text.config(state="disabled")

ttk.Button(tab_logs, text="Refresh Logs", command=load_latest_log).pack(pady=5)
load_latest_log()

# ===================== START =====================
root.mainloop()

# This is AI generated code, please refer KPIT AI Policy before using this in your projects