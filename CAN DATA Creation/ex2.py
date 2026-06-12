import tkinter as tk
from tkinter import messagebox


def validate_input(event, entry):
    """Allow only '0' or '1' and replace existing value"""
    value = entry.get()
    if value not in ("0", "1"):
        entry.delete(0, tk.END)
        entry.insert(0, "0")  # Reset to default '0'
    convert()  # Automatically update the conversion when input is changed


def convert():
    decimal_values = []
    hex_values = []
    for row in bit_entries:
        binary_str = "".join(entry.get() for entry in row)
        decimal_value = int(binary_str, 2)
        hex_value = f"{decimal_value:02X}"
        decimal_values.append(decimal_value)
        hex_values.append(hex_value)

    # Update each byte's decimal and hexadecimal output
    for i in range(8):
        decimal_labels[i].config(text=f"Decimal: {decimal_values[i]}")
        hex_labels[i].config(text=f"Hex: {hex_values[i]}")

    # Update the final hex data at the bottom
    final_hex_data.set(" ".join(hex_values))


def copy_to_clipboard():
    root.clipboard_clear()
    root.clipboard_append(final_hex_data.get())
    root.update()
    messagebox.showinfo("Copied", f"Copied: {final_hex_data.get()}")


def reset_entries():
    """Reset all entry fields to 0"""
    for row in bit_entries:
        for entry in row:
            entry.delete(0, tk.END)
            entry.insert(0, "0")  # Reset to default '0'
    convert()  # Recalculate conversions after reset


# Create GUI window
root = tk.Tk()
root.title("8-Byte Binary Converter")
root.geometry("900x600")

bit_entries = []
decimal_labels = []
hex_labels = []
frame = tk.Frame(root)
frame.pack(pady=10, padx=10)

for byte in range(8):
    byte_frame = tk.Frame(frame, borderwidth=2, relief="ridge", padx=5, pady=5)
    byte_frame.grid(row=byte // 3, column=byte % 3, padx=5, pady=5)

    # "ByteX" Label (Centered at the top)
    tk.Label(byte_frame, text=f"Byte{byte}", font=("Arial", 12, "bold")).grid(row=0, column=0, columnspan=9, pady=5)

    # Bit Position Header (7 to 0)
    tk.Label(byte_frame, text="Bit Position", font=("Arial", 10, "bold")).grid(row=1, column=0, padx=10, pady=5,
                                                                               sticky="w")
    for i in range(8):
        tk.Label(byte_frame, text=str(7 - i), width=5, borderwidth=1, relief="solid").grid(row=1, column=i + 1, padx=2)

    # DATA Label
    tk.Label(byte_frame, text="DATA", font=("Arial", 10, "bold")).grid(row=2, column=0, padx=10, pady=5, sticky="w")

    # DATA Row (Entry Fields) with Default 0s
    row_entries = [tk.Entry(byte_frame, width=5, justify="center") for _ in range(8)]
    for i, entry in enumerate(row_entries):
        entry.grid(row=2, column=i + 1, padx=2)
        entry.insert(0, "0")  # Default value to 0
        entry.bind("<KeyRelease>", lambda event, e=entry: validate_input(event, e))  # Validate input and trigger conversion

    bit_entries.append(row_entries)

    # Decimal Output Label (Default)
    decimal_label = tk.Label(byte_frame, text="Decimal: 0", font=("Arial", 10))
    decimal_label.grid(row=3, column=0, columnspan=9, pady=5)
    decimal_labels.append(decimal_label)

    # Hex Output Label (Default)
    hex_label = tk.Label(byte_frame, text="Hex: 00", font=("Arial", 10))
    hex_label.grid(row=4, column=0, columnspan=9, pady=5)
    hex_labels.append(hex_label)

# Final Hex Data Display with Border
final_data_frame = tk.Frame(root, borderwidth=2, relief="ridge", padx=10, pady=10)
final_data_frame.pack(pady=10, padx=10, fill="x")

final_hex_data = tk.StringVar()
final_hex_label = tk.Label(final_data_frame, textvariable=final_hex_data, font=("Arial", 12, "bold"))
final_hex_label.pack(pady=5)

# Copy Button for Final Data with Border
copy_button = tk.Button(final_data_frame, text="Copy Final Data", command=copy_to_clipboard)
copy_button.pack(pady=5)

# Reset Button with Border
reset_button = tk.Button(final_data_frame, text="Reset", command=reset_entries)
reset_button.pack(pady=5)

root.mainloop()
