import os
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import threading
import math
from datetime import datetime


# =====================================================
# READ EXCEL (.xls / .xlsx)
# =====================================================
def read_excel_auto(path):
    ext = os.path.splitext(path)[1].lower()

    if ext == ".xls":
        import xlwings as xw
        app = xw.App(visible=False)
        wb = app.books.open(path)
        data = wb.sheets[0].used_range.value
        wb.close()
        app.quit()
        return pd.DataFrame(data[1:], columns=data[0])

    elif ext == ".xlsx":
        return pd.read_excel(path, engine="openpyxl")

    else:
        raise ValueError("Unsupported file format")


# =====================================================
# MAIN PROCESS
# =====================================================
def process_file(path, project_qty, project_code, designer, revision, progress, root):
    try:
        print("\n========== PROCESS STARTED ==========\n")
        progress["value"] = 5

        df = read_excel_auto(path)
        
        # Print initial data for debugging
        print(f"Initial rows (including empty): {len(df)}")
        print(f"Initial columns: {df.columns.tolist()}")

        # =================================================
        # STEP 1: Remove completely empty rows
        # =================================================
        initial_count = len(df)
        
        # Remove rows where Hierarchy is empty/NaN
        df = df.dropna(subset=['Hierarchy']).copy()
        df = df[df['Hierarchy'].astype(str).str.strip() != ''].copy()
        df = df[df['Hierarchy'].astype(str).str.strip() != 'nan'].copy()
        
        empty_removed = initial_count - len(df)
        df = df.reset_index(drop=True)
        
        print(f"Empty rows removed: {empty_removed}")
        print(f"Rows remaining: {len(df)}")
        
        if len(df) == 0:
            messagebox.showerror("Error", "No valid data rows found in the file!")
            return

        # =================================================
        # REQUIRED COLUMNS - Add if missing
        # =================================================
        required = [
            "Hierarchy", "Description", "Part Code",
            "Unit Qty", "Unit Weight (kg)", "Thickness (mm)",
            "Process 1", "Process 2"
        ]

        for col in required:
            if col not in df.columns:
                df[col] = ""

        # Map "Part Name" to "Description" if Description is empty
        if 'Part Name' in df.columns:
            df['Description'] = df['Description'].fillna('')
            df['Description'] = df.apply(
                lambda row: row['Part Name'] if (row['Description'] == '' or pd.isna(row['Description'])) 
                else row['Description'], 
                axis=1
            )

        progress["value"] = 10

        # =================================================
        # NORMALIZATION
        # =================================================
        df["Hierarchy"] = df["Hierarchy"].astype(str).str.strip()
        df["Hierarchy"] = df["Hierarchy"].str.replace(r"\.0$", "", regex=True)
        
        df["Description"] = df["Description"].fillna("").astype(str).str.strip()
        df["Part Code"] = df["Part Code"].fillna("").astype(str).str.strip()
        
        # Handle Unit Qty
        df["Unit Qty"] = pd.to_numeric(df["Unit Qty"], errors="coerce").fillna(0)
        
        # Handle Unit Weight
        df["Unit Weight (kg)"] = pd.to_numeric(df["Unit Weight (kg)"], errors="coerce")
        
        # Handle Thickness
        df["Thickness (mm)"] = pd.to_numeric(df["Thickness (mm)"], errors='coerce')
        
        # Add Level column for hierarchy depth
        df["Level"] = df["Hierarchy"].apply(lambda x: x.count(".") + 1 if x and x != 'nan' else 0)

        print(f"\nFirst 5 rows after normalization:")
        print(df[['Hierarchy', 'Description', 'Part Code', 'Unit Qty']].head())
        
        print(f"\nSample descriptions (first 20):")
        print(df['Description'].head(20).tolist())

        progress["value"] = 20

        # =================================================
        # SHEET → PARENT THICKNESS LOGIC
        # =================================================
        hierarchy_index = {}
        for idx, row in df.iterrows():
            hierarchy_index[row["Hierarchy"]] = idx

        sheet_rows_to_delete = []
        thickness_copied = 0

        # Find all "Sheet" items (case-insensitive)
        sheet_mask = df["Description"].str.lower().str.strip() == "sheet"
        sheet_items = df[sheet_mask]
        
        print(f"\nFound {len(sheet_items)} items with Description='Sheet'")
        
        if len(sheet_items) > 0:
            print("Sheet items found:")
            print(sheet_items[['Hierarchy', 'Description', 'Thickness (mm)', 'Unit Weight (kg)']])

        for idx, row in sheet_items.iterrows():
            # Check if sheet has thickness but no weight
            if pd.isna(row["Unit Weight (kg)"]) and not pd.isna(row["Thickness (mm)"]):
                parts = row["Hierarchy"].split(".")
                # Look for parent items (going up the hierarchy)
                parent_found = False
                for i in range(len(parts) - 1, 0, -1):
                    parent = ".".join(parts[:i])
                    if parent in hierarchy_index:
                        p_idx = hierarchy_index[parent]
                        # Copy thickness to parent
                        df.at[p_idx, "Thickness (mm)"] = row["Thickness (mm)"]
                        sheet_rows_to_delete.append(idx)
                        thickness_copied += 1
                        parent_found = True
                        print(f"✓ Copied thickness {row['Thickness (mm)']} from {row['Hierarchy']} to parent {parent}")
                        break
                
                if not parent_found:
                    print(f"✗ No parent found for sheet at {row['Hierarchy']}")

        # Remove sheet rows that were processed
        if sheet_rows_to_delete:
            df = df.drop(index=sheet_rows_to_delete)
            df = df.reset_index(drop=True)
            print(f"Deleted {len(sheet_rows_to_delete)} sheet rows")

        print(f"Total thickness copied: {thickness_copied}")

        progress["value"] = 40

        # =================================================
        # QTY EXPLOSION (Multiply quantities through hierarchy)
        # =================================================
        # Sort by hierarchy depth to process parents first
        df["__depth"] = df["Hierarchy"].str.count(r"\.")
        df = df.sort_values(["__depth", "Hierarchy"]).reset_index(drop=True)
        
        qty_map = {}
        
        for idx, row in df.iterrows():
            h = row["Hierarchy"]
            q = row["Unit Qty"]
            
            if "." in h:  # Has parent
                parent_h = h.rsplit(".", 1)[0]
                parent_qty = qty_map.get(parent_h, 1)
                q = q * parent_qty
            
            qty_map[h] = q
            df.at[idx, "Unit Qty"] = q

        # Calculate final quantities
        df["Total Qty"] = df["Unit Qty"] * project_qty
        df["Total Weight (kg)"] = df["Total Qty"] * df["Unit Weight (kg)"].fillna(0)
        
        # Round for cleaner output
        df["Unit Qty"] = df["Unit Qty"].round(3)
        df["Total Qty"] = df["Total Qty"].round(3)
        df["Total Weight (kg)"] = df["Total Weight (kg)"].round(3)

        print(f"\nAfter QTY explosion (first 5 rows):")
        print(df[['Hierarchy', 'Description', 'Unit Qty', 'Total Qty']].head())

        progress["value"] = 60

        # =================================================
        # PROCESS LOGIC (Based on Part Code prefix)
        # =================================================
        df["Process 1"] = "-"
        df["Process 2"] = "-"
        
        for i, r in df.iterrows():
            code = str(r["Part Code"]).strip().upper()
            if code.startswith("FL"):
                df.at[i, "Process 1"] = "Laser Cutting"
                df.at[i, "Process 2"] = "-"
            elif code.startswith("BE"):
                df.at[i, "Process 1"] = "Laser Cutting"
                df.at[i, "Process 2"] = "Bending"
            elif code.startswith("PL"):
                df.at[i, "Process 1"] = "Plasma Cutting"
                df.at[i, "Process 2"] = "-"
            elif code.startswith("SAW"):
                df.at[i, "Process 1"] = "Saw Cutting"
                df.at[i, "Process 2"] = "-"

        # Count processes assigned
        process_count = len(df[df["Process 1"] != "-"])
        print(f"\nProcess assigned to {process_count} rows")

        progress["value"] = 70

        # =================================================
        # FINAL CLEANUP
        # =================================================
        # Remove temporary columns
        if "__depth" in df.columns:
            df = df.drop(columns=["__depth"])
        
        # Keep only useful columns for output
        output_columns = [
            "Hierarchy", "Level", "Part Code", "Description", 
            "Unit Qty", "Unit Weight (kg)", "Thickness (mm)",
            "Total Qty", "Total Weight (kg)", 
            "Process 1", "Process 2"
        ]
        
        # Add any extra columns from original file
        extra_cols = [c for c in df.columns if c not in output_columns 
                     and c not in ["__depth"]]
        final_columns = output_columns + extra_cols
        
        # Keep only columns that exist
        final_columns = [c for c in final_columns if c in df.columns]
        df = df[final_columns]

        progress["value"] = 85

        # =================================================
        # SAVE OUTPUT FILE
        # =================================================
        out = os.path.splitext(path)[0] + "_updated.xlsx"
        
        print(f"\nSaving to: {out}")
        print(f"Output columns: {final_columns}")
        print(f"Total data rows: {len(df)}")
        
        # Save with openpyxl directly
        with pd.ExcelWriter(out, engine='openpyxl') as writer:
            # Write data starting from row 7 (0-indexed row 6)
            df.to_excel(writer, index=False, startrow=6, sheet_name='BOM')
            
            # Get the workbook and worksheet
            workbook = writer.book
            worksheet = writer.sheets['BOM']
            
            # Write header information in the first 6 rows
            from openpyxl.styles import Font, PatternFill
            
            header_font = Font(bold=True, size=11)
            
            worksheet['A1'] = "Project Code:"
            worksheet['A1'].font = header_font
            worksheet['B1'] = project_code
            
            worksheet['A2'] = "Designer:"
            worksheet['A2'].font = header_font
            worksheet['B2'] = designer
            
            worksheet['A3'] = "Generated Date:"
            worksheet['A3'].font = header_font
            worksheet['B3'] = datetime.today().strftime("%d-%m-%Y")
            
            worksheet['A4'] = "Revision:"
            worksheet['A4'].font = header_font
            worksheet['B4'] = revision
            
            worksheet['A5'] = "Project Quantity:"
            worksheet['A5'].font = header_font
            worksheet['B5'] = project_qty
            
            # Row 6 is intentionally blank
            
            # Auto-adjust column widths
            for column in worksheet.columns:
                max_length = 0
                column_letter = column[0].column_letter
                for cell in column:
                    try:
                        if len(str(cell.value)) > max_length:
                            max_length = len(str(cell.value))
                    except:
                        pass
                adjusted_width = min(max_length + 2, 50)
                worksheet.column_dimensions[column_letter].width = adjusted_width

        progress["value"] = 100

        print(f"\n========== PROCESS COMPLETED ==========")
        print(f"Total rows in output: {len(df)}")
        print(f"Output saved to: {out}")

        # Show detailed success message
        messagebox.showinfo(
            "✅ PROCESS COMPLETED",
            f"Empty rows removed: {empty_removed}\n"
            f"Valid data rows: {len(df)}\n"
            f"Sheet rows deleted: {len(sheet_rows_to_delete)}\n"
            f"Thickness copied: {thickness_copied}\n"
            f"Process assigned: {process_count}\n\n"
            f"📄 BOM data starts from Row 7\n"
            f"📋 Headers in Rows 1-5\n\n"
            f"💾 Saved as:\n{os.path.basename(out)}"
        )

    except Exception as e:
        messagebox.showerror("❌ ERROR", f"Error: {str(e)}")
        import traceback
        print("\n" + "="*50)
        print("ERROR DETAILS:")
        print("="*50)
        traceback.print_exc()


# =====================================================
# GUI
# =====================================================
def start_gui():
    root = tk.Tk()
    root.title("BOM Qty Processor v2.0")
    root.geometry("600x480")
    root.resizable(False, False)
    
    # Title
    title_frame = tk.Frame(root, bg="#2196F3", height=60)
    title_frame.pack(fill=tk.X)
    tk.Label(
        title_frame, 
        text="BOM Quantity Explosion & Process Assignment", 
        font=("Arial", 14, "bold"),
        fg="white",
        bg="#2196F3"
    ).pack(pady=15)

    # File selection
    file_frame = tk.LabelFrame(root, text="Input File", padx=10, pady=10)
    file_frame.pack(fill=tk.X, padx=20, pady=(15, 5))
    
    file_entry = tk.Entry(file_frame, width=55, font=("Arial", 9))
    file_entry.pack(side=tk.LEFT, padx=5)
    
    tk.Button(
        file_frame,
        text="📂 Browse",
        command=lambda: file_entry.delete(0, tk.END) or 
                       file_entry.insert(0, filedialog.askopenfilename(
                           filetypes=[("Excel Files", "*.xls *.xlsx")]
                       )),
        bg="#FF9800",
        fg="white",
        font=("Arial", 9, "bold")
    ).pack(side=tk.LEFT)

    # Parameters
    param_frame = tk.LabelFrame(root, text="Parameters", padx=10, pady=10)
    param_frame.pack(fill=tk.X, padx=20, pady=5)
    
    def param_field(parent, label, default, row):
        tk.Label(parent, text=label, width=15, anchor='w', font=("Arial", 9)).grid(
            row=row, column=0, padx=5, pady=5, sticky='w'
        )
        entry = tk.Entry(parent, width=30, font=("Arial", 9))
        entry.insert(0, str(default))
        entry.grid(row=row, column=1, padx=5, pady=5)
        return entry

    proj = param_field(param_frame, "Project Code:", "NA910", 0)
    des = param_field(param_frame, "Designer:", "Akshay", 1)
    rev = param_field(param_frame, "Revision:", "R00", 2)
    qty = param_field(param_frame, "Project Qty:", "8", 3)

    # Progress bar
    progress_frame = tk.Frame(root)
    progress_frame.pack(pady=15)
    
    bar = ttk.Progressbar(progress_frame, length=500, mode='determinate')
    bar.pack()

    # Start button
    def start():
        if not file_entry.get():
            messagebox.showerror("Error", "Please select an Excel file first!")
            return

        try:
            project_qty_val = float(qty.get())
            if project_qty_val <= 0:
                messagebox.showerror("Error", "Project Qty must be greater than 0!")
                return
        except ValueError:
            messagebox.showerror("Error", "Project Qty must be a valid number!")
            return

        bar["value"] = 0
        
        threading.Thread(
            target=process_file,
            args=(
                file_entry.get(),
                project_qty_val,
                proj.get(),
                des.get(),
                rev.get(),
                bar,
                root
            ),
            daemon=True
        ).start()

    tk.Button(
        root, 
        text="▶ START PROCESSING", 
        command=start,
        bg="#4CAF50",
        fg="white",
        font=("Arial", 11, "bold"),
        padx=30,
        pady=8,
        cursor="hand2"
    ).pack(pady=10)

    # Footer
    tk.Label(root, text="© BOM Processor v2.0 | Developed for Manufacturing", 
             fg="gray", font=("Arial", 8)).pack(side=tk.BOTTOM, pady=10)

    root.mainloop()


# =====================================================
# RUN
# =====================================================
if __name__ == "__main__":
    start_gui()