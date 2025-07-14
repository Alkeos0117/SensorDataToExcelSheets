import openpyxl
from openpyxl.utils import get_column_letter
import pandas as pd
import os
import tkinter as tk
from tkinter import filedialog, messagebox

# Globals
template_file = None
data_file = None
df = None
wb = None
template_sheet = None

def select_template():
    global template_file, wb, template_sheet
    template_file = filedialog.askopenfilename(
        title="Select Excel template (.xlsx)",
        filetypes=[("Excel files", "*.xlsx")]
    )
    if template_file:
        wb = openpyxl.load_workbook(template_file)
        template_sheet = wb.active
        status_label["text"] = f"✅ Template loaded: {os.path.basename(template_file)}"
    else:
        status_label["text"] = "⚠️ No template selected."

def select_data_file():
    global data_file, df
    data_file = filedialog.askopenfilename(
        title="Select data file (.csv or .xlsx)",
        filetypes=[("Excel/CSV files", "*.xlsx *.csv")]
    )
    if data_file:
        ext = os.path.splitext(data_file)[1].lower()
        try:
            df = pd.read_csv(data_file) if ext == ".csv" else pd.read_excel(data_file)
            status_label["text"] = f"✅ Data loaded: {os.path.basename(data_file)}"
        except Exception as e:
            messagebox.showerror("Error", f"Could not load data file:\n{e}")
    else:
        status_label["text"] = "⚠️ No data file selected."

def process_and_save():
    global wb, template_sheet, df

    if wb is None or df is None:
        messagebox.showwarning("Missing files", "Please select both the template and data file first.")
        return

    base_name = template_sheet.title

    for index, row in df.iterrows():
        try:
            sn = str(row['S/N'])
            val_0 = float(row['0bar'])
            val_20 = float(row['20bar'])
            val_40 = float(row['40bar'])
            resistance = str(row['I-Widerstand'])

            sheet = wb.copy_worksheet(template_sheet)
            sheet.title = sn

            sheet["B3"] = sn
            sheet["C14"] = val_0
            sheet["C15"] = val_20
            sheet["C16"] = val_40
            sheet["B17"] = resistance + " @ 50VDC"

            # Formulas
            sheet["E14"] = "=1/$C$10*D14"
            sheet["E15"] = "=1/$C$10*D15"
            sheet["E16"] = "=1/$C$10*D16"
            sheet["A15"] = "=B8+(B10/2)"
            sheet["A16"] = "=B9"

        except Exception as e:
            print(f"❌ Error in row {index}: {e}")

    # Remove the original template sheet
    wb.remove(wb[base_name])

    # Save final output
    output_file = filedialog.asksaveasfilename(
        initialfile="final_report.xlsx",
        defaultextension=".xlsx",
        filetypes=[("Excel Workbook", "*.xlsx")],
        title="Save final report as"
    )

    if output_file:
        wb.save(output_file)
        messagebox.showinfo("Success", f"✅ Report saved as:\n{output_file}")
    else:
        messagebox.showwarning("Cancelled", "No file was saved.")

# GUI
root = tk.Tk()
root.title("Excel Report Generator")
root.geometry("520x300")

tk.Label(root, text="Step 1: Select Excel Template (.xlsx)", font=("Arial", 10)).pack(pady=5)
tk.Button(root, text="📂 Load Template", command=select_template).pack()

tk.Label(root, text="Step 2: Select Data File (.csv or .xlsx)", font=("Arial", 10)).pack(pady=5)
tk.Button(root, text="📂 Load Data", command=select_data_file).pack()

tk.Label(root, text="Step 3: Generate Final Excel File", font=("Arial", 10)).pack(pady=5)
tk.Button(root, text="📝 Generate Final Report", command=process_and_save).pack()

status_label = tk.Label(root, text="", fg="green")
status_label.pack(pady=15)

root.mainloop()
