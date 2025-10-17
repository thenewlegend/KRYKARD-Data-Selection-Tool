import customtkinter as ctk
from tkinter import filedialog
import pandas as pd
import os
import sys
from datetime import datetime

# ---------- Resource Path for PyInstaller ----------
def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

# ---------- App Configuration ----------
ctk.set_appearance_mode("System")
ctk.set_default_color_theme("blue")

root = ctk.CTk()
root.title("KSPC - KRYKARD Data Selection Tool")
root.geometry("700x650")
root.iconbitmap(default=resource_path("logo.ico"))  # ✅ Safe icon loading

# ---------- Variables ----------
file_path_var = ctk.StringVar()
checkbox_vars = []
checkbox_widgets = []
select_all_state = ctk.BooleanVar(value=False)

# ---------- Functions ----------

def browse_file():
    file_path = filedialog.askopenfilename(
        title="Select Excel File",
        filetypes=[("Excel Files", "*.xlsx *.xls")]
    )
    file_path_var.set(file_path)

    if is_valid_excel_file(file_path):
        try:
            progress_bar.set(0.2)
            progress_label.configure(text="📥 Loading file...")
            root.update_idletasks()

            df = pd.read_excel(file_path, sheet_name=0, header=1)
            show_column_selector(df.columns.tolist())

            progress_bar.set(1.0)
            progress_label.configure(text="✅ File loaded. Select columns to retain.")
        except Exception as e:
            progress_bar.set(0)
            progress_label.configure(text=f"❌ Error loading file: {str(e)}")
    else:
        progress_bar.set(0)
        progress_label.configure(text="❌ Invalid file type. Please select an Excel file.")

def is_valid_excel_file(path):
    return path.lower().endswith((".xlsx", ".xls"))

def show_column_selector(columns):
    for widget in checkbox_widgets:
        widget.destroy()
    checkbox_vars.clear()
    checkbox_widgets.clear()

    for col in columns:
        var = ctk.BooleanVar(value=False)
        checkbox = ctk.CTkCheckBox(column_frame, text=col, variable=var)
        checkbox.pack(anchor="w")
        checkbox_vars.append((col, var))
        checkbox_widgets.append(checkbox)

    select_all_state.set(False)
    toggle_btn.configure(text="✅ Select All")

def toggle_select_all():
    new_state = not select_all_state.get()
    for _, var in checkbox_vars:
        var.set(new_state)
    select_all_state.set(new_state)
    toggle_btn.configure(text="❎ Unselect All" if new_state else "✅ Select All")

def save_filtered_sheets():
    file_path = file_path_var.get()
    selected_columns = [col for col, var in checkbox_vars if var.get()]

    if not selected_columns:
        progress_label.configure(text="⚠️ No columns selected.")
        return

    try:
        progress_bar.set(0.2)
        progress_label.configure(text="🔄 Reading all sheets...")
        root.update_idletasks()

        all_sheets = pd.read_excel(file_path, sheet_name=None, header=1)
        filtered_sheets = {}

        progress_bar.set(0.4)
        progress_label.configure(text="🔍 Filtering selected columns...")
        root.update_idletasks()

        for sheet_name, df in all_sheets.items():
            valid_cols = [col for col in selected_columns if col in df.columns]
            filtered_sheets[sheet_name] = df[valid_cols]

        progress_bar.set(0.6)
        progress_label.configure(text="📁 Preparing save directory...")
        root.update_idletasks()

        base = os.path.splitext(os.path.basename(file_path))[0]
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        new_filename = f"{base}_selected_data_{timestamp}.xlsx"

        parent_dir = os.path.dirname(file_path)
        save_dir = os.path.join(parent_dir, "Selected Data")
        os.makedirs(save_dir, exist_ok=True)

        new_path = os.path.join(save_dir, new_filename)

        progress_bar.set(0.8)
        progress_label.configure(text="💾 Saving filtered sheets...")
        root.update_idletasks()

        with pd.ExcelWriter(new_path, engine="openpyxl") as writer:
            for sheet_name, filtered_df in filtered_sheets.items():
                filtered_df.to_excel(writer, sheet_name=sheet_name, index=False)

        progress_bar.set(1.0)
        progress_label.configure(text=f"✅ Saved to: {new_path}")
    except Exception as e:
        progress_bar.set(0)
        progress_label.configure(text=f"❌ Error saving file: {str(e)}")

# ---------- UI Elements ----------
file_entry = ctk.CTkEntry(root, textvariable=file_path_var, width=400)
file_entry.pack(pady=(20, 10))

browse_btn = ctk.CTkButton(root, text="📂 Browse", command=browse_file)
browse_btn.pack(pady=(0, 10))

toggle_btn = ctk.CTkButton(root, text="✅ Select All", command=toggle_select_all)
toggle_btn.pack(pady=(0, 10))

next_btn = ctk.CTkButton(root, text="➡️ Next", command=save_filtered_sheets)
next_btn.pack(pady=(0, 10))

progress_bar = ctk.CTkProgressBar(root, width=400)
progress_bar.set(0)
progress_bar.pack(pady=(10, 10))

progress_label = ctk.CTkLabel(root, text="Select a file to begin.", wraplength=600, justify="left")
progress_label.pack(pady=(10, 10))

scroll_frame = ctk.CTkScrollableFrame(root, width=600, height=300)
scroll_frame.pack(pady=(10, 10), fill="both", expand=True)

column_frame = ctk.CTkFrame(scroll_frame)
column_frame.pack(fill="both", expand=True)

# ---------- Run App ----------
root.mainloop()
