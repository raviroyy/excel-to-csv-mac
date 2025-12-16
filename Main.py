import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import os
import webbrowser

def convert_files():
    files = filedialog.askopenfilenames(filetypes=[("Excel files", "*.xlsx *.xls")])
    if not files:
        return

    progress["maximum"] = len(files)
    output_folder = filedialog.askdirectory(title="Select folder to save CSV files")

    if not output_folder:
        return

    for i, file in enumerate(files):
        try:
            if file.endswith(".xls"):
                df = pd.read_excel(file, engine="xlrd")
            else:
                df = pd.read_excel(file, engine="openpyxl")

            csv_name = os.path.splitext(os.path.basename(file))[0] + ".csv"
            csv_path = os.path.join(output_folder, csv_name)
            df.to_csv(csv_path, index=False)

        except Exception as e:
            messagebox.showerror("Error", f"Failed to convert:\n{file}\n\n{e}")

        progress["value"] = i + 1
        root.update_idletasks()

    messagebox.showinfo("Success", "All files converted!")
    webbrowser.open(f"file://{output_folder}")

# ---------------- GUI ----------------

root = tk.Tk()
root.title("Excel to CSV Converter")
root.geometry("420x240")
root.resizable(False, False)

# Upload button
upload_btn = tk.Button(
    root,
    text="Upload Excel File(s)",
    command=convert_files,
    font=("Arial", 12),
    width=20
)
upload_btn.pack(pady=15)

# Small download icon BELOW upload
download_icon = tk.PhotoImage(file="assets/download_icon.png")
download_icon = download_icon.subsample(4, 4)  # resize icon

icon_label = tk.Label(root, image=download_icon)
icon_label.pack(pady=5)

# Progress bar
progress = ttk.Progressbar(root, orient="horizontal", length=320, mode="determinate")
progress.pack(pady=15)

root.mainloop()
