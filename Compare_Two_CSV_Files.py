import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pandas as pd

# Compare two CSV files based on the selected CSV key column
def compare_csvs(before_file, after_file, key_column):

    # Load CSVs and replace NaN with empty strings
    df_before = pd.read_csv(before_file).fillna("")
    df_after = pd.read_csv(after_file).fillna("")

    # Ensure the key column exists in both files
    if key_column not in df_before.columns or key_column not in df_after.columns:
        raise ValueError(f"Key column '{key_column}' not found in both CSV files.")
 
    # Identify added and removed keys
    before_keys = set(df_before[key_column])
    after_keys = set(df_after[key_column])
    added_rows = df_after[~df_after[key_column].isin(before_keys)]
    removed_rows = df_before[~df_before[key_column].isin(after_keys)]

    # Identify changed cell values
    changed_values = []
    common_keys = before_keys.intersection(after_keys)

    df_before_indexed = df_before.set_index(key_column)
    df_after_indexed = df_after.set_index(key_column)

    # Consider only common columns (excluding the key column)
    common_cols = df_before.columns.intersection(df_after.columns)
    common_cols = [c for c in common_cols if c != key_column]

    # Check each common key and column for changes
    for key in common_keys:
        for col in common_cols:
            before_val = str(df_before_indexed.at[key, col])
            after_val = str(df_after_indexed.at[key, col])
            if before_val != after_val:  # Only record actual changes
                changed_values.append((key, col, before_val, after_val))

    # Summary of comparison
    summary = {
        "added": len(added_rows),
        "removed": len(removed_rows),
        "changed": len(changed_values)
    }

    return summary, added_rows, removed_rows, changed_values

# Export a Word document report containing the comparison results.
def export_report(summary, added, removed, changed, key_column):

    doc = Document()
    doc.add_heading("CSV Comparison Report", level=1)

    # Add summary
    doc.add_paragraph(f"Rows added: {summary['added']}")
    doc.add_paragraph(f"Rows removed: {summary['removed']}")
    doc.add_paragraph(f"Cells changed: {summary['changed']}")

    # Add changed values
    doc.add_heading("Changed Values", level=2)
    if changed:
        for idx, col, before, after in changed[:200]:  # Limit to first 200 changes
            doc.add_paragraph(f"Key '{idx}', Column '{col}': '{before}' → '{after}'")
    else:
        doc.add_paragraph("None")

    # Add added rows (key column only)
    doc.add_heading("Added Rows (Key Column only)", level=2)
    doc.add_paragraph(added[[key_column]].to_string(index=False) if not added.empty else "None")

    # Add removed rows (key column only)
    doc.add_heading("Removed Rows (Key Column only)", level=2)
    doc.add_paragraph(removed[[key_column]].to_string(index=False) if not removed.empty else "None")

    # Prompt user to save file
    save_path = filedialog.asksaveasfilename(defaultextension=".docx",
                                             filetypes=[("Word Document", "*.docx")])
    if save_path:
        doc.save(save_path)
        messagebox.showinfo("Report Saved", f"Report saved as:\n{save_path}")

# Enable or disable the Compare button depending on CSV and key column selection.
def update_compare_button():
    before_file = before_entry.get()
    after_file = after_entry.get()
    key_selected = key_column_var.get()
    if before_file and after_file and key_selected:
        compare_btn.config(state="normal")
    else:
        compare_btn.config(state="disabled")

# Validate the key column to ensure it exists in both CSVs. Only common columns are selectable.
def validate_key_column():
    before_file = before_entry.get()
    after_file = after_entry.get()
    if not before_file or not after_file:
        return
    try:
        df_before = pd.read_csv(before_file)
        df_after = pd.read_csv(after_file)

        # Only keep columns present in both CSVs
        common_cols = list(set(df_before.columns).intersection(df_after.columns))
        if common_cols:
            key_dropdown['values'] = common_cols
            # Reset selection if current key is invalid
            if key_column_var.get() not in common_cols:
                key_dropdown.current(0)
            key_dropdown.config(state="readonly")
        else:
            key_dropdown.set("")
            key_dropdown.config(state="disabled")
            messagebox.showwarning("No Common Columns",
                                   "No common columns found between Before and After CSVs.")
    except Exception as e:
        key_dropdown.set("")
        key_dropdown.config(state="disabled")
        messagebox.showerror("Error", f"Error validating key column:\n{e}")
    finally:
        update_compare_button()

# Populate the Key Column dropdown with columns from the Before CSV.
def populate_key_dropdown(file_path):
    try:
        df = pd.read_csv(file_path)
        if len(df.columns) > 0:
            key_dropdown['values'] = list(df.columns)
            key_dropdown.current(0)
            key_dropdown.config(state="readonly")
        else:
            key_dropdown.set("")
            key_dropdown.config(state="disabled")
            messagebox.showwarning("No Columns", "CSV has no columns.")
    except Exception as e:
        key_dropdown.set("")
        key_dropdown.config(state="disabled")
        messagebox.showerror("Error", f"Cannot read CSV:\n{e}")
    finally:
        update_compare_button()

#  Browse and select the Before CSV file. Resets and validates the key column dropdown.
def browse_before():
    file_path = filedialog.askopenfilename(filetypes=[("CSV files", "*.csv")])
    before_entry.delete(0, tk.END)
    before_entry.insert(0, file_path)
    # Reset dropdown
    key_dropdown.set("")
    key_dropdown.config(state="disabled")

    if file_path:
        populate_key_dropdown(file_path)
        validate_key_column()
    update_compare_button()

#  Browse and select the After CSV file. Validates the key column dropdown after selection.
def browse_after():
    file_path = filedialog.askopenfilename(filetypes=[("CSV files", "*.csv")])
    after_entry.delete(0, tk.END)
    after_entry.insert(0, file_path)
    validate_key_column()
    update_compare_button()

# Run the CSV comparison and display results in the GUI.
def run_comparison():
    global summary, added, removed, changed
    before_file = before_entry.get()
    after_file = after_entry.get()
    key_column = key_column_var.get()

    if not before_file or not after_file or not key_column:
        messagebox.showwarning("Missing Input",
                               "Please select both CSV files and choose the Key Column.")
        return
    try:
        summary, added, removed, changed = compare_csvs(before_file, after_file, key_column)

        # Display results in the text area
        result_text.configure(state="normal")
        result_text.delete(1.0, tk.END)
        result_text.insert(tk.END, f"Comparison Complete!\n\n")
        result_text.insert(tk.END, f"Rows Added: {summary['added']}\n")
        result_text.insert(tk.END, f"Rows Removed: {summary['removed']}\n")
        result_text.insert(tk.END, f"Cells Changed: {summary['changed']}\n\n")
        result_text.insert(tk.END, "Preview of Changed Values (first 10):\n")

        if changed:
            for i, (idx, col, before_val, after_val) in enumerate(changed[:10]):
                result_text.insert(tk.END, f"- Key '{idx}', {col}: '{before_val}' → '{after_val}'\n")
        else:
            result_text.insert(tk.END, "None\n")

        if not added.empty:
            result_text.insert(tk.END, "\nAdded Rows (Key Column only):\n")
            for key in added[key_column]:
                result_text.insert(tk.END, f"- {key}\n")

        if not removed.empty:
            result_text.insert(tk.END, "\nRemoved Rows (Key Column only):\n")
            for key in removed[key_column]:
                result_text.insert(tk.END, f"- {key}\n")

        result_text.configure(state="disabled")
        export_btn.config(state="normal")

    except Exception as e:
        messagebox.showerror("Error", f"Something went wrong:\n{e}")

# GUI Layout
root = tk.Tk()
root.title("CSV Compare Tool")
root.geometry("700x700")
root.configure(bg="#f5f5f5")

# Main frame
main_frame = ttk.Frame(root, padding=15)
main_frame.pack(fill="both", expand=True)

# CSV upload frame
csv_frame = ttk.LabelFrame(main_frame, text="Upload CSV Files", padding=15)
csv_frame.grid(row=0, column=0, columnspan=3, sticky="ew", pady=10)

# Before CSV
ttk.Label(csv_frame, text="Before CSV:").grid(row=0, column=0, sticky="w", pady=5)
before_entry = ttk.Entry(csv_frame, width=70)
before_entry.grid(row=0, column=1, padx=5)
ttk.Button(csv_frame, text="Browse", command=browse_before).grid(row=0, column=2, padx=5)

# After CSV
ttk.Label(csv_frame, text="After CSV:").grid(row=1, column=0, sticky="w", pady=5)
after_entry = ttk.Entry(csv_frame, width=70)
after_entry.grid(row=1, column=1, padx=5)
ttk.Button(csv_frame, text="Browse", command=browse_after).grid(row=1, column=2, padx=5)

# Key Column dropdown
ttk.Label(csv_frame, text="Key Column: ").grid(row=2, column=0, sticky="w", pady=5)
key_column_var = tk.StringVar()
key_dropdown = ttk.Combobox(csv_frame, textvariable=key_column_var, state="disabled", width=30)
key_dropdown.grid(row=2, column=1, sticky="w", pady=10, padx=5)
key_dropdown.bind("<<ComboboxSelected>>", lambda e: update_compare_button())

# Compare button
compare_btn = tk.Button(main_frame, text="Compare CSVs", font=("Helvetica", 12, "bold"),
                        bg="#4CAF50", command=run_comparison, state="disabled")
compare_btn.grid(row=1, column=0, columnspan=3, pady=15, sticky="ew")

# Results frame
results_frame = ttk.LabelFrame(main_frame, text="Results:", padding=10)
results_frame.grid(row=2, column=0, columnspan=3, sticky="nsew")
main_frame.rowconfigure(2, weight=1)
main_frame.columnconfigure(0, weight=1)

# Text area for results
result_text = tk.Text(results_frame, height=25, width=115, wrap="none", bg="#fdfdfd", relief="sunken")
result_text.pack(side="left", fill="both", expand=True)
result_text.configure(state="disabled")

# Scrollbars
v_scroll = ttk.Scrollbar(results_frame, orient="vertical", command=result_text.yview)
v_scroll.pack(side="right", fill="y")
result_text.configure(yscrollcommand=v_scroll.set)

h_scroll = ttk.Scrollbar(results_frame, orient="horizontal", command=result_text.xview)
h_scroll.pack(side="bottom", fill="x")
result_text.configure(xscrollcommand=h_scroll.set)

# Export button
export_btn = tk.Button(main_frame, text="Export Report (.docx)", font=("Helvetica", 12, "bold"),
                       bg="#2196F3",
                       command=lambda: export_report(summary, added, removed, changed, key_column_var.get()))
export_btn.grid(row=4, column=0, columnspan=3, pady=10, sticky="ew")
export_btn.config(state="disabled")

# Start GUI
root.mainloop()
