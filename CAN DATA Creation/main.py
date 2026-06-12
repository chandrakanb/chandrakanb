import tkinter as tk
from tkinter import messagebox

def convert():
    try:
        hex_num = entry.get().strip()
        decimal_num = int(hex_num, 16)  # Convert hex to decimal
        binary_num = bin(decimal_num)[2:]  # Convert decimal to binary

        dec_label.config(text=f"Decimal: {decimal_num}")
        bin_label.config(text=f"Binary: {binary_num}")

        copy_dec_btn.config(state=tk.NORMAL)
        copy_bin_btn.config(state=tk.NORMAL)
    except ValueError:
        messagebox.showerror("Error", "Invalid Hexadecimal Number")

def copy_to_clipboard(value):
    root.clipboard_clear()
    root.clipboard_append(value)
    root.update()
    messagebox.showinfo("Copied", f"Copied: {value}")

# Create GUI window
root = tk.Tk()
root.title("Hex Converter")
root.geometry("350x200")

# Input field
tk.Label(root, text="Enter Hex Number:").grid(row=0, column=0, padx=10, pady=5)
entry = tk.Entry(root)
entry.grid(row=0, column=1, pady=5)

# Convert Button
convert_btn = tk.Button(root, text="Convert", command=convert)
convert_btn.grid(row=1, column=0, columnspan=2, pady=5)

# Decimal Output with Copy Button
dec_label = tk.Label(root, text="Decimal: ")
dec_label.grid(row=2, column=0, sticky="w", padx=10)
copy_dec_btn = tk.Button(root, text="Copy", command=lambda: copy_to_clipboard(dec_label.cget("text").split(": ")[1]), state=tk.DISABLED)
copy_dec_btn.grid(row=2, column=1)

# Binary Output with Copy Button
bin_label = tk.Label(root, text="Binary: ")
bin_label.grid(row=3, column=0, sticky="w", padx=10)
copy_bin_btn = tk.Button(root, text="Copy", command=lambda: copy_to_clipboard(bin_label.cget("text").split(": ")[1]), state=tk.DISABLED)
copy_bin_btn.grid(row=3, column=1)

root.mainloop()
