import os
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox
from datetime import datetime

print("Starting Part Code Analyzer...")

def read_excel_file(path):
    ext = os.path.splitext(path)[1].lower()
    if ext == ".xlsx":
        return pd.read_excel(path, engine="openpyxl")
    elif ext == ".xls":
        import xlwings as xw
        app = xw.App(visible=False)
        wb = app.books.open(path)
        data = wb.sheets[0].used_range.value
        wb.close()
        app.quit()
        return pd.DataFrame(data[1:], columns=data[0])

def analyze_part_codes(file_path):
    df = read_excel_file(file_path)
    
    print("\n" + "="*70)
    print("PART CODE INDENTATION ANALYSIS")
    print("="*70)
    
    # Show columns
    print(f"\nColumns in file: {df.columns.tolist()}")
    
    # Clean Part Code column
    df['Part Code'] = df['Part Code'].fillna('').astype(str)
    
    # Analyze Part Code indentation
    print("\n" + "-"*70)
    print("ALL ROWS - PART CODE ANALYSIS")
    print("-"*70)
    print(f"{'Row#':<6} {'Part Code':<30} {'Leading Spaces':<15} {'Length':<8} {'First Char'}")
    print("-"*70)
    
    for idx, row in df.iterrows():
        part_code = str(row['Part Code'])
        leading_spaces = len(part_code) - len(part_code.lstrip())
        stripped_code = part_code.strip()
        first_char = stripped_code[0] if stripped_code else 'N/A'
        
        # Show first 30 rows
        if idx < 30:
            print(f"{idx+2:<6} '{part_code[:28]:<28}' {leading_spaces:<15} {len(part_code):<8} '{first_char}'")
        elif idx == 30:
            print(f"... showing only first 30 rows ...")
            print(f"(Total rows: {len(df)})")
    
    # Count unique indentation levels
    print("\n" + "-"*70)
    print("INDENTATION LEVEL SUMMARY")
    print("-"*70)
    
    indent_counts = {}
    for idx, row in df.iterrows():
        part_code = str(row['Part Code'])
        spaces = len(part_code) - len(part_code.lstrip())
        indent_counts[spaces] = indent_counts.get(spaces, 0) + 1
    
    print(f"{'Spaces':<10} {'Count':<10}")
    print("-"*20)
    for spaces in sorted(indent_counts.keys()):
        print(f"{spaces:<10} {indent_counts[spaces]:<10}")
    
    # Show Hierarchy patterns
    print("\n" + "-"*70)
    print("SAMPLE HIERARCHY PATTERN (First 30 rows)")
    print("-"*70)
    print(f"{'Row#':<6} {'Part Code':<35} {'Spaces':<8} {'Level'}")
    print("-"*70)
    
    prev_spaces = 0
    level = 0
    
    for idx, row in df.iterrows():
        part_code = str(row['Part Code'])
        spaces = len(part_code) - len(part_code.lstrip())
        stripped = part_code.strip()
        
        # Simple level calculation based on indentation
        if spaces > prev_spaces:
            level += 1
        elif spaces < prev_spaces:
            level -= (prev_spaces - spaces) // 2  # Assuming 2 spaces per level
        
        if idx < 30:
            print(f"{idx+2:<6} '{stripped[:33]:<33}' {spaces:<8} {level}")
        
        prev_spaces = spaces

    print("\n" + "="*70)
    print("ANALYSIS COMPLETE")
    print("="*70)
    
    return df

# Simple GUI
def create_gui():
    root = tk.Tk()
    root.title("Part Code Analyzer")
    root.geometry("500x200")
    
    root.lift()
    root.attributes('-topmost', True)
    
    tk.Label(root, text="Part Code Indentation Analyzer", font=("Arial", 14, "bold")).pack(pady=20)
    tk.Label(root, text="Select your Excel file to analyze Part Code patterns", font=("Arial", 10)).pack()
    
    file_entry = tk.Entry(root, width=50)
    file_entry.pack(pady=10)
    
    def browse():
        filename = filedialog.askopenfilename(
            filetypes=[("Excel Files", "*.xlsx *.xls")]
        )
        if filename:
            file_entry.delete(0, tk.END)
            file_entry.insert(0, filename)
    
    btn_frame = tk.Frame(root)
    btn_frame.pack()
    
    tk.Button(btn_frame, text="Browse File", command=browse, bg="#FF9800", fg="white", padx=10).pack(side=tk.LEFT, padx=5)
    
    def analyze():
        path = file_entry.get()
        if path:
            analyze_part_codes(path)
            messagebox.showinfo("Done", "Check terminal for detailed analysis")
    
    tk.Button(btn_frame, text="Analyze", command=analyze, bg="#4CAF50", fg="white", padx=10).pack(side=tk.LEFT, padx=5)
    
    root.mainloop()

if __name__ == "__main__":
    create_gui()