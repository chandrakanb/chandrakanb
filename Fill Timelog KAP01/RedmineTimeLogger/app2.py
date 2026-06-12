from redminelib import Redmine
from datetime import datetime, timedelta
import tkinter as tk
from tkinter import ttk, messagebox
import logging, os, sys, json

BASE_URL = "https://kap01.kpit.com/kap"
TASK_FILE = "tasks.json"
CONFIG_FILE = "config.json"
FEATURE_CONFIG_FILE = "feature_config.json"

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

# ================= CONFIG =================

def load_config():
    if os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE,"r",encoding="utf-8") as f:
                return json.load(f)
        except:
            pass
    return {"api_key":""}

def save_config(cfg):
    with open(CONFIG_FILE,"w",encoding="utf-8") as f:
        json.dump(cfg,f,indent=2)

config = load_config()

# ================= TASKS =================

def load_tasks():

    if os.path.exists(TASK_FILE):
        try:
            with open(TASK_FILE,"r",encoding="utf-8") as f:
                data=json.load(f)

            if isinstance(data,list):
                return data
        except Exception as e:
            print("Task file error:",e)

    return []

def save_tasks():
    with open(TASK_FILE,"w",encoding="utf-8") as f:
        json.dump(tasks,f,indent=2)

tasks = load_tasks()

# ============ FEATURE CONFIG ============

def load_feature_config():
    if os.path.exists(FEATURE_CONFIG_FILE):
        try:
            with open(FEATURE_CONFIG_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        except:
            pass
    return {}

def save_feature_config(data):
    with open(FEATURE_CONFIG_FILE, "w", encoding="utf-8") as f:
        json.dump(data, f, indent=2)

feature_config = load_feature_config()

def save_last_selection():

    feature_config["last_feature"] = feature_entry.get()
    feature_config["last_subfeature"] = subfeature_var.get()
    feature_config["last_subtask"] = subtask_var.get()
    
    save_feature_config(feature_config)

# ================= DATES =================

dates=[(datetime.today()-timedelta(days=i)).strftime("%Y-%m-%d") for i in range(10)]

# ================= UI ROOT =================

root=tk.Tk()
root.title("Redmine Time Logger")
root.geometry("850x500")

notebook=ttk.Notebook(root)
notebook.pack(expand=True,fill="both")

# ================= TIME LOGGING TAB =================

tab_log=ttk.Frame(notebook)
notebook.add(tab_log,text="Time Logging")

date_vars={}
select_all_var=tk.BooleanVar()

def toggle_all():
    for v in date_vars.values():
        v.set(select_all_var.get())

ttk.Checkbutton(tab_log,text="Select All Dates",variable=select_all_var,command=toggle_all).pack(pady=5)

date_frame=ttk.Frame(tab_log)
date_frame.pack()

for d in dates:
    v=tk.BooleanVar()
    ttk.Checkbutton(date_frame,text=d,variable=v).pack(anchor="w")
    date_vars[d]=v

progress=ttk.Progressbar(tab_log,length=600)
progress.pack(pady=15)

def log_time():

    selected_dates=[d for d,v in date_vars.items() if v.get()]

    if not selected_dates:
        messagebox.showwarning("Select","Choose dates")
        return

    api_key=config.get("api_key","").strip()

    if not api_key:
        messagebox.showerror("API","Add API key")
        return

    redmine=Redmine(BASE_URL,key=api_key)

    total=len(selected_dates)*len(tasks)
    progress["maximum"]=total
    progress["value"]=0

    for d in selected_dates:
        for t in tasks:

            if t.get("hours",0)<=0:
                progress["value"]+=1
                root.update_idletasks()
                continue

            try:

                custom_fields=[
                    {"id":46,"value":t.get("cat","")},
                    {"id":48,"value":"India"},
                    {"id":49,"value":"Pune"}
                ]

                if t.get("cat")!="TEAM MEETING":
                    custom_fields.append({"id":47,"value":t.get("type","")})

                redmine.time_entry.create(
                    issue_id=t.get("issue"),
                    spent_on=d,
                    hours=t.get("hours"),
                    comments=t.get("comment",""),
                    custom_fields=custom_fields
                )

                log.info(f"OK | {t.get('issue')} | {d}")

            except Exception as e:

                log.error(f"FAIL | {t.get('issue')} | {d} | {e}")

            progress["value"]+=1
            root.update_idletasks()

    messagebox.showinfo("Done","Time logged")

ttk.Button(tab_log,text="Log Time",command=log_time).pack(pady=10)

# ================= TASK TAB =================

tab_tasks=ttk.Frame(notebook)
notebook.add(tab_tasks,text="Tasks")

feature_frame=ttk.Frame(tab_tasks)
feature_frame.pack(fill="x",padx=10,pady=5)

ttk.Label(feature_frame,text="Feature ID").pack(side="left")

feature_entry=ttk.Entry(feature_frame,width=12)
feature_entry.pack(side="left",padx=5)

subfeature_var=tk.StringVar()
subfeature_dropdown=ttk.Combobox(feature_frame,width=30,textvariable=subfeature_var,state="readonly")
subfeature_dropdown.pack(side="left",padx=5)
subfeature_dropdown.bind("<<ComboboxSelected>>", lambda e: save_last_selection())

subtask_var=tk.StringVar()
subtask_dropdown=ttk.Combobox(feature_frame,width=30,textvariable=subtask_var,state="readonly")
subtask_dropdown.pack(side="left",padx=5)
subtask_dropdown.bind("<<ComboboxSelected>>", lambda e: save_last_selection())

if feature_config.get("last_feature"):
    feature_entry.insert(0, feature_config["last_feature"])

if feature_config.get("last_subfeature"):
    subfeature_var.set(feature_config["last_subfeature"])

if feature_config.get("last_subtask"):
    subtask_var.set(feature_config["last_subtask"])

def fetch_subfeatures():

    api=config.get("api_key")

    if not api:
        messagebox.showerror("Error","Add API key")
        return

    fid=feature_entry.get()

    try:

        redmine=Redmine(BASE_URL,key=api)

        issues=redmine.issue.filter(parent_id=fid)

        vals=[]
        m={}

        for i in issues:
            label=f"{i.id} - {i.subject}"
            vals.append(label)
            m[label]=i.id

        subfeature_dropdown["values"]=vals
        subfeature_dropdown.map=m

        if vals:
            subfeature_dropdown.current(0)

    except Exception as e:
        messagebox.showerror("Error",str(e))

feature_config["last_feature"] = feature_entry.get()
save_feature_config(feature_config)

def fetch_subtasks():

    sel=subfeature_var.get()

    if not sel:
        messagebox.showwarning("Select","Select sub feature")
        return

    sid=subfeature_dropdown.map[sel]

    try:

        redmine=Redmine(BASE_URL,key=config["api_key"])

        issues=redmine.issue.filter(parent_id=sid)

        vals=[]
        m={}

        for i in issues:
            label=f"{i.id} - {i.subject}"
            vals.append(label)
            m[label]=i.id

        subtask_dropdown["values"]=vals
        subtask_dropdown.map=m

        if vals:
            subtask_dropdown.current(0)

    except Exception as e:
        messagebox.showerror("Error",str(e))

ttk.Button(feature_frame,text="Fetch SubFeatures",command=fetch_subfeatures).pack(side="left",padx=5)
ttk.Button(feature_frame,text="Fetch SubTasks",command=fetch_subtasks).pack(side="left",padx=5)

# ================= TASK INPUT =================

entry_frame = ttk.Frame(tab_tasks)
entry_frame.pack(pady=10, fill="x", padx=10)

category_values = [
"Please select ---",
"TECHNICAL MARKETING",
"SALES ACTIVITIES",
"PROPOSAL",
"ESTIMATION",
"PRESALES ACTIVITIES",
"REQUIREMENT",
"PLANNING",
"DESIGN",
"PROTOTYPING",
"ARCHITECTURE",
"PROJECT MANAGEMENT",
"CODING",
"OPERATIONAL ACTIVITIES",
"PRACTICE INITIATIVES",
"ORGANIZATIONAL ACTIVITIES",
"CONFIGURATION MANAGEMENT",
"INTEGRATION",
"TEST AUTOMATION",
"UNIT TESTING",
"MODULE/FUNCTION TESTING",
"INTEGRATION TESTING",
"SYSTEM TESTING",
"FIELD TESTING",
"ACCEPTANCE TESTING",
"FUNCTIONAL TESTING",
"PERFORMANCE TESTING",
"PORTING",
"TOOL EVALUATION",
"CUSTOMER MEETINGS",
"PRODUCTION SUPPORT",
"DEVELOPMENT SUPPORT",
"DEG ACTIVITIES",
"LAB SUPPORT",
"KPMS/IDP",
"TEAM MEETING",
"VENDOR MEETING",
"TRAINING",
"ROLE TRANSITION",
"INVOICING ACTIVITIES",
"OPERATIONS ACTIVITIES",
"IT PROCUREMENT",
"MISCELLANEOUS",
"TRAVEL",
"IDLE",
"SETUP HW/SW",
"PROCESS CONSULTING",
"TAG ACTIVITIES",
"CORPORATE MARKETING ACTIVITIES",
"HR ACTIVITIES",
"TIMS ACTIVITIES",
"ECODE ACTIVITIES",
"PSO ACTIVITIES",
"FINANCE AND ACCOUNTING ACTIVITIE",
"FLM ACTIVITIES"
]

type_values = [
"NA",
"WORK",
"REVIEW",
"INTERNAL REWORK",
"EXTERNAL REWORK"
]

cat_var = tk.StringVar()
type_var = tk.StringVar()

ttk.Label(entry_frame, text="Hours").pack(side="left", padx=5)

hours_e = ttk.Entry(entry_frame, width=5)
hours_e.pack(side="left")

ttk.Label(entry_frame, text="Comment").pack(side="left", padx=5)

comment_e = ttk.Entry(entry_frame, width=35)
comment_e.pack(side="left")

ttk.Label(entry_frame, text="Category").pack(side="left", padx=5)

def combobox_search(event):

    widget = event.widget
    value = widget.get().lower()

    values = widget["values"]

    for item in values:
        if item.lower().startswith(value):
            widget.set(item)
            widget.icursor(len(value))
            break

cat_dropdown = ttk.Combobox(
    entry_frame,
    textvariable=cat_var,
    values=category_values,
    width=25,
    state="normal"
)
cat_dropdown.pack(side="left", padx=5)
cat_dropdown.bind("<KeyRelease>", combobox_search)

ttk.Label(entry_frame, text="Work Type").pack(side="left", padx=5)

type_dropdown = ttk.Combobox(
    entry_frame,
    textvariable=type_var,
    values=type_values,
    width=18,
    state="normal"
)
type_dropdown.pack(side="left", padx=5)
type_dropdown.bind("<KeyRelease>", combobox_search)

# ================= TASK LIST =================

task_list = tk.Listbox(tab_tasks, height=10)
task_list.pack(fill="x", padx=10, pady=10)

def load_task_for_edit(event):

    sel = task_list.curselection()

    if not sel:
        return

    index = sel[0]

    task = tasks[index]

    hours_e.delete(0, tk.END)
    hours_e.insert(0, task.get("hours",""))

    comment_e.delete(0, tk.END)
    comment_e.insert(0, task.get("comment",""))

    cat_var.set(task.get("cat",""))

    type_var.set(task.get("type",""))

    task_list.edit_index = index

task_list.bind("<Double-Button-1>", load_task_for_edit)

def refresh_tasks():

    task_list.delete(0, tk.END)

    for t in tasks:

        issue = t.get("issue","")
        hours = t.get("hours","")
        comment = t.get("comment","")
        cat = t.get("cat","")
        typ = t.get("type","")

        display = f"{issue} | {hours}h | {comment} | {cat} | {typ}"

        task_list.insert(tk.END, display)

refresh_tasks()

# ================= ADD TASK =================

def add_task():

    if cat_var.get() == "Please select ---":
        messagebox.showwarning("Category","Select category")
        return

    try:

        sel = subtask_var.get()

        if not sel:
            messagebox.showwarning("Select","Select sub task")
            return

        issue = subtask_dropdown.map[sel]

        task = {
            "issue": int(issue),
            "hours": float(hours_e.get()),
            "comment": comment_e.get(),
            "cat": cat_var.get(),
            "type": type_var.get(),
            "subject": ""
        }

        if hasattr(task_list, "edit_index"):

            tasks[task_list.edit_index] = task
            del task_list.edit_index

        else:

            tasks.append(task)

        save_tasks()
        refresh_tasks()

        hours_e.delete(0,tk.END)
        comment_e.delete(0,tk.END)

    except Exception as e:
        messagebox.showerror("Error",str(e))

def delete_task():

    sel=task_list.curselection()

    if sel:
        tasks.pop(sel[0])
        save_tasks()
        refresh_tasks()

btn_frame=ttk.Frame(tab_tasks)
btn_frame.pack()

btn_frame = ttk.Frame(tab_tasks)
btn_frame.pack(pady=5)

ttk.Button(btn_frame, text="Add Task", command=add_task).pack(side="left", padx=5)
ttk.Button(btn_frame, text="Delete Task", command=delete_task).pack(side="left", padx=5)
# ================= API KEY TAB =================

tab_api=ttk.Frame(notebook)
notebook.add(tab_api,text="API Key")

ttk.Label(tab_api,text="Redmine API Key").pack(pady=20)

api_entry=ttk.Entry(tab_api,width=50,show="*")
api_entry.pack()

if config.get("api_key"):
    api_entry.insert(0,config["api_key"])

def save_api():

    config["api_key"]=api_entry.get().strip()
    save_config(config)
    messagebox.showinfo("Saved","API key saved")

ttk.Button(tab_api,text="Save API Key",command=save_api).pack(pady=10)

# ================= LOG TAB =================

tab_logs=ttk.Frame(notebook)
notebook.add(tab_logs,text="Logs")

log_box=tk.Text(tab_logs,height=25,state="disabled")
log_box.pack(fill="both",expand=True)

def load_log():

    files=sorted(os.listdir("logs"),reverse=True)

    if not files:
        return

    with open(os.path.join("logs",files[0])) as f:
        data=f.read()

    log_box.config(state="normal")
    log_box.delete("1.0",tk.END)
    log_box.insert(tk.END,data)
    log_box.config(state="disabled")

ttk.Button(tab_logs,text="Refresh Logs",command=load_log).pack()

load_log()

root.mainloop()