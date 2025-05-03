import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import os
import csv
from openpyxl import Workbook

# Theme list
themes = [
    {
        "name": "Dark",
        "bg": "#1e1e1e",
        "fg": "white",
        "button_bg": "#007acc",
        "button_fg": "white"
    },
    {
        "name": "Light",
        "bg": "#f0f0f0",
        "fg": "black",
        "button_bg": "#4caf50",
        "button_fg": "white"
    },
    {
        "name": "Alien Green",
        "bg": "#003300",
        "fg": "#00ff99",
        "button_bg": "#00cc66",
        "button_fg": "black"
    },
    {
        "name": "Sky Blue",
        "bg": "#e0f7fa",
        "fg": "#006064",
        "button_bg": "#4dd0e1",
        "button_fg": "black"
    },
    {
        "name": "Red Alert",
        "bg": "#330000",
        "fg": "#ff6666",
        "button_bg": "#cc0000",
        "button_fg": "white"
    }
]

theme_index = 0
input_file_path = ""

def apply_theme():
    t = themes[theme_index]
    root.configure(bg=t["bg"])
    for widget in [label, format_label, parts_label, status_label, file_label, output_label]:
        widget.config(bg=t["bg"], fg=t["fg"])
    for widget in [file_button, folder_button, start_button, theme_button]:
        widget.config(bg=t["button_bg"], fg=t["button_fg"])
    format_menu.config(background=t["bg"], foreground=t["fg"])
    parts_entry.config(bg="white" if t["bg"] == "#f0f0f0" else "#333", fg=t["fg"])

def toggle_theme():
    global theme_index
    theme_index = (theme_index + 1) % len(themes)
    apply_theme()

def select_file():
    global input_file_path
    input_file_path = filedialog.askopenfilename(title="Select Email List (.txt)", filetypes=[("Text Files", "*.txt")])
    if input_file_path:
        file_path_var.set(f"Selected: {input_file_path}")

def select_output_folder():
    folder = filedialog.askdirectory(title="Select Output Folder")
    if folder:
        output_dir.set(folder)

def split_email_file():
    if not input_file_path:
        messagebox.showerror("No Input File", "Please select an input email list file!")
        return

    output_path = output_dir.get()
    if not output_path:
        messagebox.showerror("No Output Folder", "Please select an output folder!")
        return

    try:
        parts = int(parts_entry.get())
        if parts < 1:
            raise ValueError("Parts must be at least 1.")

        with open(input_file_path, 'r') as f:
            emails = [line.strip() for line in f if line.strip()]

        total = len(emails)
        if parts > total:
            raise ValueError("Parts cannot exceed total emails.")

        chunk_size = total // parts
        base_name = os.path.splitext(os.path.basename(input_file_path))[0]
        selected_format = format_var.get()

        for i in range(parts):
            start = i * chunk_size
            end = (start + chunk_size) if i < parts - 1 else total
            chunk = emails[start:end]
            filename = os.path.join(output_path, f"{base_name}_part_{i+1}{selected_format}")

            if selected_format == ".txt":
                with open(filename, 'w') as f_out:
                    f_out.write('\n'.join(chunk))
            elif selected_format == ".csv":
                with open(filename, 'w', newline='') as f_out:
                    writer = csv.writer(f_out)
                    writer.writerow(["Email"])
                    for email in chunk:
                        writer.writerow([email])
            elif selected_format == ".xlsx":
                wb = Workbook()
                ws = wb.active
                ws.title = "Emails"
                ws.append(["Email"])
                for email in chunk:
                    ws.append([email])
                wb.save(filename)

        status_label.config(text=f"✅ Done! {parts} files saved in:\n{output_path}")
        messagebox.showinfo("Split Complete", f"Email list split into {parts} {selected_format} files!")

    except ValueError as ve:
        messagebox.showerror("Invalid Input", str(ve))
    except Exception as e:
        messagebox.showerror("Error", f"Something went wrong:\n{e}")

# GUI setup
root = tk.Tk()
root.title("Alien Email Splitter PRO")
root.geometry("560x480")

label = tk.Label(root, text="🚀 Split Email List into Multiple Parts", font=("Arial", 14))
label.pack(pady=10)

file_button = tk.Button(root, text="Select Email File (.txt)", command=select_file, font=("Arial", 11))
file_button.pack(pady=5)

file_path_var = tk.StringVar()
file_label = tk.Label(root, textvariable=file_path_var, font=("Arial", 9), wraplength=450, justify="center")
file_label.pack()

parts_label = tk.Label(root, text=" How many parts you want?", font=("Arial", 11))
parts_label.pack()
parts_entry = tk.Entry(root, width=10, font=("Arial", 12))
parts_entry.insert(0, "10")
parts_entry.pack(pady=5)

format_label = tk.Label(root, text=" Select Output Format:", font=("Arial", 11))
format_label.pack()
format_var = tk.StringVar()
format_var.set(".txt")
format_menu = ttk.Combobox(root, textvariable=format_var, values=[".txt", ".csv", ".xlsx"], state="readonly", width=10)
format_menu.pack(pady=5)

folder_button = tk.Button(root, text="Choose Output Folder", command=select_output_folder, font=("Arial", 11))
folder_button.pack(pady=5)

output_dir = tk.StringVar()
output_label = tk.Label(root, textvariable=output_dir, font=("Arial", 9), wraplength=400)
output_label.pack(pady=2)

start_button = tk.Button(root, text="🧠 Start Split", command=split_email_file, font=("Arial", 12), padx=10, pady=5)
start_button.pack(pady=10)

theme_button = tk.Button(root, text="🧪 Change Theme", command=toggle_theme, font=("Arial", 10))
theme_button.pack(pady=5)

status_label = tk.Label(root, text="", font=("Arial", 10), wraplength=400, justify="center")
status_label.pack(pady=10)

apply_theme()
root.mainloop()
