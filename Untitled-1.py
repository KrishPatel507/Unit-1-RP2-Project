import tkinter as tk
from tkinter import messagebox

# Store cases
cases = []

# ------------------------
# Functions
# ------------------------

def add_case():
    try:
        case_id = int(case_id_entry.get())
        case_name = case_name_entry.get()
        client = client_entry.get()

        if case_name == "" or client == "":
            messagebox.showerror("Error", "Please fill in all fields.")
            return

        cases.append({
            "case_id": case_id,
            "case_name": case_name,
            "client": client
        })

        messagebox.showinfo("Success", "Case Added!")

        case_id_entry.delete(0, tk.END)
        case_name_entry.delete(0, tk.END)
        client_entry.delete(0, tk.END)

        display_cases()

    except ValueError:
        messagebox.showerror("Error", "Case ID must be a number.")

def bubble_sort():
    n = len(cases)

    for i in range(n):
        for j in range(0, n - i - 1):
            if cases[j]["case_id"] > cases[j + 1]["case_id"]:
                cases[j], cases[j + 1] = cases[j + 1], cases[j]

    display_cases()
    messagebox.showinfo("Success", "Cases Sorted!")

def display_cases():
    listbox.delete(0, tk.END)

    for case in cases:
        listbox.insert(
            tk.END,
            f"ID: {case['case_id']} | {case['case_name']} | {case['client']}"
        )

def clear_cases():
    cases.clear()
    display_cases()

# ------------------------
# GUI
# ------------------------

window = tk.Tk()
window.title("GCSec Case Sorting System")
window.geometry("600x450")
window.resizable(False, False)

title = tk.Label(window,
                 text="GCSec Case Sorting System",
                 font=("Arial", 18, "bold"))
title.pack(pady=10)

# Case ID
tk.Label(window, text="Case ID").pack()
case_id_entry = tk.Entry(window, width=30)
case_id_entry.pack()

# Case Name
tk.Label(window, text="Case Name").pack()
case_name_entry = tk.Entry(window, width=30)
case_name_entry.pack()

# Client
tk.Label(window, text="Client Name").pack()
client_entry = tk.Entry(window, width=30)
client_entry.pack(pady=5)

# Buttons
button_frame = tk.Frame(window)
button_frame.pack(pady=10)

tk.Button(button_frame,
          text="Add Case",
          width=12,
          command=add_case).grid(row=0, column=0, padx=5)

tk.Button(button_frame,
          text="Bubble Sort",
          width=12,
          command=bubble_sort).grid(row=0, column=1, padx=5)

tk.Button(button_frame,
          text="Clear",
          width=12,
          command=clear_cases).grid(row=0, column=2, padx=5)

# Listbox
listbox = tk.Listbox(window, width=70, height=12)
listbox.pack(pady=10)

window.mainloop()