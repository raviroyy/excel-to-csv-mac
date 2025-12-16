import pandas as pd # pyright: ignore[reportMissingModuleSource]
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

def convert_files():
    files = filedialog.askopenfilenames(filetypes=[("Excel files", "*.xlsx *.xls")])
    if not files:
        return
    progress['maximum'] = len(files)
    for i, file in enumerate(files):
        try:
            df = pd.read_excel(file)
            csv_file = file.rsplit('.', 1)[0] + '.csv'
            df.to_csv(csv_file, index=False)
        except Exception as e:
            messagebox.showerror("Error", f"Failed to convert {file}\n{e}")
        progress['value'] = i + 1
        root.update_idletasks()
    messagebox.showinfo("Done", "All files converted!")

root = tk.Tk()
root.title("Excel to CSV Converter")
root.geometry("400x150")

button = tk.Button(root, text="Upload Excel File(s)", command=convert_files)
button.pack(pady=20)

progress = ttk.Progressbar(root, orient='horizontal', length=300, mode='determinate')
progress.pack(pady=10)

root.mainloop()
