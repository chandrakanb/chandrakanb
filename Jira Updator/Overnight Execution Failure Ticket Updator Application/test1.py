import tkinter as tk
from tkinter import ttk

def transition_action_frame():
    clear_page2()
    transition_frame.pack(fill="both", expand=True)

def edit_action_frame():
    clear_page2()
    edit_frame.pack(fill="both", expand=True)

def comment_action_frame():
    clear_page2()
    comment_frame.pack(fill="both", expand=True)

def assignee_action_frame():
    clear_page2()
    assignee_frame.pack(fill="both", expand=True)

def clear_page2():
    transition_frame.pack_forget()
    edit_frame.pack_forget()
    comment_frame.pack_forget()
    assignee_frame.pack_forget()

"""
def update_action_frame():
    selection = radio_var.get()
    if selection == "Transition Issues":
        transition_action_frame()
    elif selection == "Edit Issues":
        edit_action_frame()
    elif selection == "Add Comment":
        comment_action_frame()
    elif selection == "Change Assignee":
        assignee_action_frame()
    notebook.select(1)
"""
"""
root = tk.Tk()
root.title("Notebook with Frames")

notebook = ttk.Notebook(root)
notebook.pack(expand=True, fill="both")

# Page 1
page1 = ttk.Frame(notebook)
notebook.add(page1, text="Select Option")

radio_var = tk.StringVar(value="Transition Issues")  # Default value matches option

radio_frame = ttk.Frame(page1)
radio_frame.pack(pady=20)

options = ["Transition Issues", "Edit Issues", "Add Comment", "Change Assignee"]
for option in options:
    ttk.Radiobutton(frame, text=option, variable=self.action_var, value=option), command=self.toggle_action).pack(side=tk.LEFT)

ttk.Button(page1, text="Next", command=update_frame).pack(pady=10)
"""
# Page 2
page2 = ttk.Frame(notebook)
notebook.add(page2, text="Dynamic UI")


self.transition_frame = ttk.LabelFrame(self.action_tab, text="Transition Issue")
self.transition_frame.pack(side=tk.TOP, padx=4, pady=2, fill='both', expand=True)

self.edit_frame = ttk.LabelFrame(self.action_tab, text="Edit Issues")
self.edit_frame.pack(side=tk.TOP, padx=4, pady=2, fill='both', expand=True)

self.comment_frame = ttk.LabelFrame(self.action_tab, text="Add Comment")
self.comment_frame.pack(side=tk.TOP, padx=4, pady=2, fill='both', expand=True)

self.assignee_frame = ttk.LabelFrame(self.action_tab, text="Change Assignee")
self.assignee_frame.pack(side=tk.TOP, padx=4, pady=2, fill='both', expand=True)

create_transition_frame_ui()
create_edit_frame_ui()
create_comment_frame_ui()
create_assignee_frame_ui()
root.mainloop()
