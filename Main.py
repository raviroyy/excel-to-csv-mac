import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import os
import webbrowser

def convert_files():
    files = filedialog.askopenfilenames(filetypes=[("Excel files", "*.xlsx *.xls")])
    if not files:
        return
    progress['maximum'] = len(files)
    converted_files = []

    for i, file in enumerate(files):
        try:
            # Detect extension and read accordingly
            if file.endswith(".xls"):
                df = pd.read_excel(file, engine="xlrd")
            else:
                df = pd.read_excel(file, engine="openpyxl")
            
            # Ask for save location for CSV
            save_path = filedialog.asksaveasfilename(
                initialfile=os.path.splitext(os.path.basename(file))[0]+'.csv',
                defaultextension=".csv",
                filetypes=[("CSV files", "*.csv")]
            )
            if save_path:
                df.to_csv(save_path, index=False)
                converted_files.append(save_path)
        except Exception as e:
            messagebox.showerror("Error", f"Failed to convert {file}\n{e}")
        progress['value'] = i + 1
        root.update_idletasks()

    if converted_files:
        messagebox.showinfo("Done", f"All files converted!\nYou can open the folder to access them.")
        # Open folder of first converted file
        webbrowser.open(f'file://{os.path.dirname(converted_files[0])}')

# GUI setup
root = tk.Tk()
root.title("Excel to CSV Converter")
root.geometry("400x180")

# Optional: Add a download icon button
download_icon = tk.PhotoImage(file="assets/download_icon.png")  # Add your icon file
button = tk.Button(root, text="Upload Excel File(s)", command=convert_files, image=download_icon, compound="left")
button.pack(pady=20)

progress = ttk.Progressbar(root, orient='horizontal', length=300, mode='determinate')
progress.pack(pady=10)

root.mainloop()
