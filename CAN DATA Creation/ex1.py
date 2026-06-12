import tkinter as tk
from tkinter import messagebox

root = tk.Tk()  # Create the main window
root.title("My Tkinter App")  # Window title
root.geometry("1400x700")  # Window size (width x height)

entry = tk.Entry(root)
entry.pack()

user_input = entry.get()

def say_hello():
    print("Hello!")

button = tk.Button(root, text="Click Me", command=say_hello)
button.pack()

messagebox.showinfo("Info", "This is an info message")
messagebox.showwarning("Warning", "This is a warning")
messagebox.showerror("Error", "This is an error")

label1 = tk.Label(root, text="Name:").grid(row=0, column=0)
entry1 = tk.Entry(root).grid(row=0, column=1)

label2 = tk.Label(root, text="Age:").grid(row=1, column=0)
entry2 = tk.Entry(root).grid(row=1, column=1)


root.mainloop()  # Start GUI loop
