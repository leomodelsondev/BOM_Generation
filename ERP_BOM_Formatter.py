import os
import re
import shutil
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import threading
from datetime import datetime
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment

# =====================================================
# TEMPLATE PATH  — update this to your local path
# =====================================================
TEMPLATE_PATH = r'D:\OneDrive - Nido Machineries Pvt Ltd\Desktop\MAC\BOM Testing\Template-Standard_BOM.xlsx'

# Fills used in template (exact hex from template inspection)
FILL_GREEN  = PatternFill(start_color="92D050", end_color="92D050", fill_type="solid")
FILL_YELLOW = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
FILL_NONE   = PatternFill(fill_type=None)

FONT_BOLD   = Font(name="Calibri", size=11, bold=True)
FONT_NORMAL = Font(name="Calibri", size=11, bold=False)


# =====================================================
# READ EXCEL  — force Hierarchy as string
# =====================================================
def read_excel_file(path):
    ext = os.path.splitext(path)[1].lower()
    if ext == ".xlsx":
        df_peek = pd.read_excel(path, engine="openpyxl", nrows=0)
        dtypes  = {"Hierarchy": str} if "Hierarchy" in df_peek.columns else {}
        return pd.read_excel(path, engine="openpyxl", dtype=dtypes)
    elif ext == ".xls":
        df_peek = pd.read_excel(path, engine="xlrd", nrows=0)
        dtypes  = {"Hierarchy": str} if "Hierarchy" in df_peek.columns else {}
        return pd.read_excel(path, engine="xlrd", dtype=dtypes)
    else:
        raise ValueError("Unsupported format. Use .xls or .xlsx")


# =====================================================
# EXTRACT THICKNESS FROM PART CODE
# BE-MS-T1.6-30971 -> 1.6  |  FL-GI-T3-6109 -> 3.0
# =====================================================
def extract_thickness(part_code):
    code = str(part_code).strip().upper()
    m = re.search(r'-T(\d+\.?\d*)(?:-|$)', code)
    if not m:
        m = re.search(r'T(\d+\.?\d*)', code)
    if m:
        try:
            return float(m.group(1))
        except ValueError:
            pass
    return None


# =====================================================
# LEAF / PARENT TAGGING
# =====================================================
def tag_leaf_parent(df):
    hierarchies = df['Hierarchy'].astype(str).tolist()
    def is_leaf(h):
        prefix = str(h) + '.'
        return not any(str(o).startswith(prefix) for o in hierarchies if str(o) != str(h))
    df = df.copy()
    df['__is_leaf'] = df['Hierarchy'].apply(is_leaf)
    return df


# =====================================================
# DIRECT CHILDREN
# =====================================================
def get_direct_children_rows(parent_h, df):
    parent_h = str(parent_h)
    prefix   = parent_h + '.'
    result   = []
    for _, row in df.iterrows():
        h = str(row['Hierarchy'])
        if h.startswith(prefix) and '.' not in h[len(prefix):]:
            result.append(row)
    return result


# =====================================================
# BUILD RM LOOKUP FROM TEMPLATE
# =====================================================
def build_rm_lookup(template_path):
    rm_df = pd.read_excel(template_path, sheet_name='Raw Material Part Code')
    # columns: Item Code, Item Name, Default Unit of Measure, Item Group, Item Sub Category, HSN/SAC

    fl_lookup  = {}   # thickness(float) -> (item_code, item_name)
    be_options = []   # (item_code, item_name) where name contains 'sheet'

    for _, row in rm_df.iterrows():
        ic    = str(row['Item Code']).strip()
        iname = str(row['Item Name']).strip()
        # Extract trailing thickness number from RM item code
        # e.g. RM-SH-HRPO-2500X1250-1.6  ->  1.6
        m = re.search(r'-(\d+\.?\d*)$', ic)
        if m:
            try:
                fl_lookup[float(m.group(1))] = (ic, iname)
            except ValueError:
                pass
        if 'sheet' in iname.lower():
            be_options.append((ic, iname))

    print(f"  FL lookup entries : {len(fl_lookup)}")
    print(f"  BE sheet options  : {len(be_options)}")
    return fl_lookup, be_options


# =====================================================
# WRITE CHILD PART BOM SHEET
# Exact format: Row 1 = headers (already in template)
# Data from Row 2 onwards, no extra rows
# Columns: A=Item Code, B=Item Name, C=Description,
#          D=Default UOM, E=Routing, F=Qty, G=Unit Weight(Kgs),
#          H=Item(BOM Details), I=Item Name(BOM Details),
#          J=Quantity(BOM Details), K=UOM(BOM Details),
#          L=Do Not Explode(Items)
# =====================================================
def write_child_sheet(ws, leaf_rows):
    # Clear existing data (keep row 1 headers)
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
        for cell in row:
            cell.value = None
            cell.fill  = FILL_NONE

    for i, r in enumerate(leaf_rows, start=2):
        # A-C: left-aligned, no fill
        ws.cell(row=i, column=1,  value=r['item_code']).font        = FONT_NORMAL
        ws.cell(row=i, column=2,  value=r['item_name']).font        = FONT_NORMAL
        ws.cell(row=i, column=3,  value=r['description']).font      = FONT_NORMAL
        # D-G: centered
        for col, val in [(4, r['uom']), (5, r['routing']),
                         (6, r['qty']), (7, r['weight'])]:
            c = ws.cell(row=i, column=col, value=val)
            c.font      = FONT_NORMAL
            c.alignment = Alignment(horizontal='center')
        # H: yellow fill (Item BOM Details)
        ch = ws.cell(row=i, column=8,  value=r['item_bom'])
        ch.font = FONT_NORMAL
        ch.fill = FILL_YELLOW if r['item_bom'] and not str(r['item_bom']).startswith('[') else FILL_NONE
        # I-L: centered
        for col, val in [(9,  r['item_name_bom']),
                         (10, r['qty_bom']),
                         (11, r['uom_bom']),
                         (12, r['do_not_explode'])]:
            c = ws.cell(row=i, column=col, value=val)
            c.font      = FONT_NORMAL
            c.alignment = Alignment(horizontal='center')

    print(f"  Written {len(leaf_rows)} rows to Child Part BOM sheet")


# =====================================================
# WRITE PARENT BOM SHEET (Bom Template-1)
# Parent row: A=Item Code (bold+green), B-G filled, H-L blank
# Child rows: A-G blank, H=Item Code (yellow), I-L filled
# =====================================================
def write_parent_sheet(ws, parent_sections):
    # Clear existing data (keep row 1 headers)
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
        for cell in row:
            cell.value = None
            cell.fill  = FILL_NONE
            cell.font  = FONT_NORMAL

    current_row = 2
    for section in parent_sections:
        p = section['parent']

        # ── PARENT ROW ──────────────────────────────────────
        # A: bold + green fill
        ca = ws.cell(row=current_row, column=1, value=p['item_code'])
        ca.font = FONT_BOLD
        ca.fill = FILL_GREEN

        # B: Item Name
        ws.cell(row=current_row, column=2,  value=p['item_name']).alignment  = Alignment(horizontal='center')
        ws.cell(row=current_row, column=3,  value=p['description']).alignment = Alignment(horizontal='center')
        ws.cell(row=current_row, column=4,  value=p['uom']).alignment         = Alignment(horizontal='center')
        ws.cell(row=current_row, column=5,  value=p['routing']).alignment     = Alignment(horizontal='center')

        # F: Standard BOM Qty = 1 (default)
        cf = ws.cell(row=current_row, column=6, value=1)
        cf.alignment = Alignment(horizontal='center')

        # G: Unit Weight
        ws.cell(row=current_row, column=7, value=p['weight']).alignment = Alignment(horizontal='center')

        # H-L blank for parent row (children fill these)
        current_row += 1

        # ── CHILD ROWS ───────────────────────────────────────
        for child in section['children']:
            # A-G: blank
            # H: Item Code — yellow fill
            ch = ws.cell(row=current_row, column=8, value=child['item_code'])
            ch.fill = FILL_YELLOW

            # I: Item Name, J: Qty, K: UOM, L: Do Not Explode
            ws.cell(row=current_row, column=9,  value=child['item_name']).alignment  = Alignment(horizontal='center')
            ws.cell(row=current_row, column=10, value=child['qty']).alignment        = Alignment(horizontal='center')
            ws.cell(row=current_row, column=11, value=child['uom']).alignment        = Alignment(horizontal='center')
            ws.cell(row=current_row, column=12, value=child['do_not_explode']).alignment = Alignment(horizontal='center')

            current_row += 1

    print(f"  Written {current_row - 2} rows to Parent BOM sheet")
    return current_row - 2


# =====================================================
# MAIN PROCESS
# =====================================================
def process_file(file_path, template_path, progress_bar, root_window, status_label):
    try:
        def upd(msg, pct):
            status_label.config(text=msg)
            progress_bar["value"] = pct
            root_window.update()

        print("\n" + "="*60)
        print("  ERP BOM FORMATTER  v3.0")
        print("="*60)
        print(f"  Input    : {file_path}")
        print(f"  Template : {template_path}")

        # ── READ INPUT ───────────────────────────────────────
        upd("Reading input file ...", 5)
        df_raw = read_excel_file(file_path)
        original_count = len(df_raw)
        print(f"\n  Rows loaded : {original_count}")
        print(f"  Columns     : {list(df_raw.columns)}")

        # Normalise key columns
        upd("Normalising ...", 10)
        df_raw['Hierarchy']   = df_raw['Hierarchy'].astype(str).str.strip()
        df_raw['Part Code']   = df_raw['Part Code'].fillna('').astype(str).str.strip()
        df_raw['Description'] = df_raw['Description'].fillna('').astype(str).str.strip()
        if 'Part Name' in df_raw.columns:
            df_raw['Part Name'] = df_raw['Part Name'].fillna('').astype(str).str.strip()

        # ── LOAD RM LOOKUP ───────────────────────────────────
        upd("Loading RM reference ...", 14)
        fl_lookup, be_options = build_rm_lookup(template_path)

        # ── STEP 1: DELETE UNWANTED ROWS ─────────────────────
        upd("Step 1: Deleting unwanted rows ...", 20)
        print("\nSTEP 1 - DELETE ROWS WHERE Part Code IS BLANK")
        is_blank  = df_raw['Part Code'] == ''
        rule_a    = is_blank & (df_raw['Description'].str.lower() == 'sheet')
        rule_b    = is_blank & (df_raw['Description'].str.lower() != 'sheet')
        to_del    = rule_a | rule_b
        del_count = int(to_del.sum())
        print(f"  Rule A (Sheet + blank Part Code) : {rule_a.sum()}")
        print(f"  Rule B (Other + blank Part Code) : {rule_b.sum()}")
        print(f"  Total deleted                    : {del_count}")
        df = df_raw[~to_del].reset_index(drop=True)
        print(f"  Rows remaining                   : {len(df)}")

        # ── STEP 2: TAG LEAF / PARENT ────────────────────────
        upd("Step 2: Tagging hierarchy ...", 30)
        df = tag_leaf_parent(df)

        # ── STEP 3: BUILD CHILD PART BOM DATA ────────────────
        upd("Step 3: Building Child Part BOM ...", 45)
        print("\nSTEP 3 - CHILD PART BOM (leaf parts)")
        leaf_rows        = []
        be_manual_needed = []

        for _, row in df[df['__is_leaf']].iterrows():
            pc    = str(row['Part Code']).strip()
            pname = str(row.get('Part Name', '')).strip() or str(row.get('Description', '')).strip()
            desc  = str(row.get('Description', '')).strip()
            try:    qty = float(row.get('Unit Qty', 1))
            except: qty = 1.0
            try:    wt  = float(row.get('Unit Weight (kg)', row.get('Unit Weight (Kgs)', 0)))
            except: wt  = ''

            item_bom, item_name_bom = '', ''
            pc_up = pc.upper()

            if pc_up.startswith('FL'):
                t = extract_thickness(pc)
                if t is not None and t in fl_lookup:
                    item_bom, item_name_bom = fl_lookup[t]
                    print(f"  FL: {pc} -> T{t} -> {item_bom}")
                else:
                    item_bom      = f'[LOOKUP FAILED T={t}]'
                    item_name_bom = '[MANUAL REQUIRED]'
                    print(f"  FL FAILED: {pc} T={t}")
            elif pc_up.startswith('BE'):
                item_bom      = '[SELECT FROM DROPDOWN]'
                item_name_bom = '[SELECT FROM DROPDOWN]'
                be_manual_needed.append(pc)

            leaf_rows.append({
                'item_code'      : pc,
                'item_name'      : pname,
                'description'    : desc,
                'uom'            : 'Nos',
                'routing'        : '',
                'qty'            : qty,
                'weight'         : wt,
                'item_bom'       : item_bom,
                'item_name_bom'  : item_name_bom,
                'qty_bom'        : wt,   # same as unit weight per spec
                'uom_bom'        : 'Kg',
                'do_not_explode' : 1,
            })

        print(f"  Leaf rows : {len(leaf_rows)}")
        if be_manual_needed:
            print(f"  BE parts needing manual RM : {len(be_manual_needed)}")

        # ── STEP 4: BUILD PARENT BOM DATA ────────────────────
        upd("Step 4: Building Parent BOM ...", 62)
        print("\nSTEP 4 - PARENT BOM")
        parent_sections = []

        for _, prow in df[~df['__is_leaf']].iterrows():
            pc    = str(prow['Part Code']).strip()
            pname = str(prow.get('Part Name', '')).strip() or str(prow.get('Description', '')).strip()
            desc  = str(prow.get('Description', '')).strip()
            try:    wt = float(prow.get('Unit Weight (kg)', prow.get('Unit Weight (Kgs)', 0)))
            except: wt = ''

            children_out = []
            for child in get_direct_children_rows(prow['Hierarchy'], df):
                cpc   = str(child['Part Code']).strip()
                cname = str(child.get('Part Name', '')).strip() or str(child.get('Description', '')).strip()
                try:    cqty = float(child.get('Unit Qty', 1))
                except: cqty = 1.0
                children_out.append({
                    'item_code'     : cpc,
                    'item_name'     : cname,
                    'qty'           : cqty,
                    'uom'           : 'Nos',
                    'do_not_explode': 1,
                })

            parent_sections.append({
                'parent': {
                    'item_code'  : pc,
                    'item_name'  : pname,
                    'description': desc,
                    'uom'        : 'Nos',
                    'routing'    : '',
                    'weight'     : wt,
                },
                'children': children_out,
            })

        print(f"  Parent sections : {len(parent_sections)}")

        # ── SAVE INTO TEMPLATE COPY ───────────────────────────
        upd("Saving into template ...", 78)
        base        = os.path.splitext(file_path)[0]
        output_path = f"{base}_ERP_BOM.xlsx"

        # Copy template so original is untouched
        shutil.copy2(template_path, output_path)
        wb = load_workbook(output_path)

        # Write Child Part BOM
        ws_child = wb['Child Part Bom With RM']
        write_child_sheet(ws_child, leaf_rows)

        # Write Parent BOM
        ws_parent = wb['Bom Template-1']
        write_parent_sheet(ws_parent, parent_sections)

        # Add BE reference sheet if needed
        if be_options:
            if 'BE_RM_Reference' in wb.sheetnames:
                del wb['BE_RM_Reference']
            ws_ref = wb.create_sheet('BE_RM_Reference')
            ws_ref['A1'] = 'BE Part RM Options -- Item Name contains Sheet'
            ws_ref['A1'].font = Font(bold=True, size=11, name="Calibri")
            ws_ref['A2'] = 'Item Code'
            ws_ref['B2'] = 'Item Name'
            ws_ref['A2'].font = Font(bold=True, name="Calibri")
            ws_ref['B2'].font = Font(bold=True, name="Calibri")
            for ri, (ic, iname) in enumerate(be_options, start=3):
                ws_ref.cell(row=ri, column=1, value=ic)
                ws_ref.cell(row=ri, column=2, value=iname)
            ws_ref.column_dimensions['A'].width = 32
            ws_ref.column_dimensions['B'].width = 52

        wb.save(output_path)
        upd("Done!", 100)

        print(f"\n  Saved : {output_path}")
        print(f"  Child Part BOM rows : {len(leaf_rows)}")
        print(f"  Parent BOM sections : {len(parent_sections)}")
        print("="*60)

        be_note = ""
        if be_manual_needed:
            be_note = (f"\n\nNOTE: {len(be_manual_needed)} BE parts need manual RM selection."
                       f"\nSee 'BE_RM_Reference' sheet for options.")

        messagebox.showinfo(
            "Success",
            f"ERP BOM generated!\n\n"
            f"Summary\n"
            f"{'-'*36}\n"
            f"  Input rows          : {original_count}\n"
            f"  Deleted rows        : {del_count}\n"
            f"  Child BOM rows      : {len(leaf_rows)}\n"
            f"  Parent BOM sections : {len(parent_sections)}\n\n"
            f"Saved as:\n  {os.path.basename(output_path)}"
            + be_note
        )

    except Exception as e:
        progress_bar["value"] = 0
        status_label.config(text="Error -- check terminal")
        print(f"\nERROR: {e}")
        import traceback; traceback.print_exc()
        messagebox.showerror("Error", f"An error occurred:\n\n{e}")


# =====================================================
# GUI
# =====================================================
def create_gui():
    NAVY   = "#1B2A4A"
    TEAL   = "#1E6E6E"
    ACCENT = "#E8A838"
    BG     = "#F0F4F8"
    WHITE  = "#FFFFFF"
    GREY   = "#7A8A99"
    DARK   = "#2C3E50"

    root = tk.Tk()
    root.title("ERP BOM Formatter  v3.0")
    root.configure(bg=BG)
    root.geometry("740x560")
    root.resizable(False, False)
    root.lift()
    root.focus_force()

    # Header
    hdr = tk.Frame(root, bg=NAVY, height=68)
    hdr.pack(fill=tk.X)
    hdr.pack_propagate(False)
    tk.Label(hdr, text="ERP BOM Formatter",
             font=("Montserrat", 18, "bold"), bg=NAVY, fg=WHITE
             ).pack(side=tk.LEFT, padx=22, pady=14)
    tk.Label(hdr, text="v3.0",
             font=("Montserrat", 10), bg=NAVY, fg=ACCENT
             ).pack(side=tk.LEFT, pady=20)

    # Footer (packed before body)
    ftr = tk.Frame(root, bg=NAVY, height=28)
    ftr.pack(fill=tk.X, side=tk.BOTTOM)
    ftr.pack_propagate(False)
    tk.Label(ftr,
             text="Concept & Developed by Mahesh Arvind Chavan  |  ERP BOM Formatter v3.0  |  Nido Automation",
             bg=NAVY, fg=GREY, font=("Montserrat", 8)
             ).pack(pady=6)

    # Body
    body = tk.Frame(root, bg=BG)
    body.pack(fill=tk.BOTH, expand=True, padx=22, pady=14)

    # Steps panel
    sf = tk.LabelFrame(body, text="  Processing Steps  ",
                       bg=BG, fg=TEAL, font=("Montserrat", 9, "bold"),
                       bd=1, relief=tk.GROOVE, padx=8, pady=6)
    sf.pack(fill=tk.X, pady=(0, 10))
    for num, text in [
        ("1", "Delete rows where Part Code is blank"),
        ("2", "Tag each row as Leaf (no children) or Parent (has children)"),
        ("3", "Child Part BOM sheet: FL auto-filled, BE marked for manual selection"),
        ("4", "Parent BOM sheet: parent rows (green) + direct children (yellow)"),
        (" ", "Output written directly into a copy of the template -- exact same format"),
    ]:
        rf = tk.Frame(sf, bg=BG)
        rf.pack(fill=tk.X, pady=1)
        lb, lf = (TEAL, WHITE) if num.strip().isdigit() else (BG, BG)
        tk.Label(rf, text=f" {num} ", bg=lb, fg=lf,
                 font=("Montserrat", 8, "bold")).pack(side=tk.LEFT, padx=(0, 7))
        tk.Label(rf, text=text, bg=BG, fg=DARK,
                 font=("Montserrat", 9), anchor='w').pack(side=tk.LEFT)

    # Template file
    tf = tk.LabelFrame(body, text="  Template File  ",
                       bg=BG, fg=TEAL, font=("Montserrat", 9, "bold"),
                       bd=1, relief=tk.GROOVE, padx=10, pady=8)
    tf.pack(fill=tk.X, pady=(0, 8))
    tfi = tk.Frame(tf, bg=BG)
    tfi.pack(fill=tk.X)
    tmpl_entry = tk.Entry(tfi, width=66, font=("Montserrat", 9),
                          bg=WHITE, relief=tk.FLAT,
                          highlightthickness=1, highlightbackground="#CCCCCC")
    tmpl_entry.insert(0, TEMPLATE_PATH)
    tmpl_entry.pack(side=tk.LEFT, padx=(0, 8), ipady=5)
    def browse_tmpl():
        fn = filedialog.askopenfilename(title="Select Template File",
                                        filetypes=[("Excel Files", "*.xlsx"), ("All Files", "*.*")])
        if fn:
            tmpl_entry.delete(0, tk.END)
            tmpl_entry.insert(0, fn)
    tb = tk.Button(tfi, text="Browse...", command=browse_tmpl,
                   bg="#7A8A99", fg=WHITE, font=("Montserrat", 9, "bold"),
                   relief=tk.FLAT, cursor="hand2", padx=10, pady=5)
    tb.pack(side=tk.LEFT)

    # Input BOM file
    ff = tk.LabelFrame(body, text="  Input BOM File  ",
                       bg=BG, fg=TEAL, font=("Montserrat", 9, "bold"),
                       bd=1, relief=tk.GROOVE, padx=10, pady=8)
    ff.pack(fill=tk.X, pady=(0, 10))
    fi = tk.Frame(ff, bg=BG)
    fi.pack(fill=tk.X)
    file_entry = tk.Entry(fi, width=66, font=("Montserrat", 9),
                          bg=WHITE, relief=tk.FLAT,
                          highlightthickness=1, highlightbackground="#CCCCCC")
    file_entry.pack(side=tk.LEFT, padx=(0, 8), ipady=5)
    def browse():
        fn = filedialog.askopenfilename(title="Select BOM Input File",
                                        filetypes=[("Excel Files", "*.xlsx *.xls"), ("All Files", "*.*")])
        if fn:
            file_entry.delete(0, tk.END)
            file_entry.insert(0, fn)
    brbtn = tk.Button(fi, text="Browse...", command=browse,
                      bg=ACCENT, fg=NAVY, font=("Montserrat", 9, "bold"),
                      relief=tk.FLAT, cursor="hand2", padx=10, pady=5)
    brbtn.pack(side=tk.LEFT)
    brbtn.bind("<Enter>", lambda e: brbtn.config(bg="#D4932A"))
    brbtn.bind("<Leave>", lambda e: brbtn.config(bg=ACCENT))

    # Progress
    pgf = tk.Frame(body, bg=BG)
    pgf.pack(fill=tk.X, pady=(6, 0))
    style = ttk.Style()
    style.theme_use('clam')
    style.configure("ERP.Horizontal.TProgressbar",
                    troughcolor="#D0D8E4", background=TEAL, thickness=14)
    progress_bar = ttk.Progressbar(pgf, length=696, mode='determinate',
                                   style="ERP.Horizontal.TProgressbar")
    progress_bar.pack(pady=(6, 3))

    status_lbl = tk.Label(body, text="Ready -- select files and press START",
                          bg=BG, fg=GREY, font=("Montserrat", 9, "italic"))
    status_lbl.pack(pady=(3, 8))

    def start():
        fp = file_entry.get().strip()
        tp = tmpl_entry.get().strip()
        if not fp:
            messagebox.showerror("Missing File", "Please select the Input BOM file.")
            return
        if not tp:
            messagebox.showerror("Missing Template", "Please select the Template file.")
            return
        if not os.path.exists(fp):
            messagebox.showerror("Not Found", f"Input file not found:\n{fp}")
            return
        if not os.path.exists(tp):
            messagebox.showerror("Not Found", f"Template file not found:\n{tp}")
            return
        progress_bar["value"] = 0
        status_lbl.config(text="Processing ...", fg=TEAL)
        root.update()
        threading.Thread(
            target=process_file,
            args=(fp, tp, progress_bar, root, status_lbl),
            daemon=True
        ).start()

    btn_row = tk.Frame(body, bg=BG)
    btn_row.pack(pady=2)
    start_btn = tk.Button(btn_row, text="START PROCESSING",
                          command=start,
                          bg=TEAL, fg=WHITE,
                          font=("Montserrat", 12, "bold"),
                          relief=tk.FLAT, cursor="hand2",
                          padx=44, pady=10)
    start_btn.pack()
    start_btn.bind("<Enter>", lambda e: start_btn.config(bg=NAVY))
    start_btn.bind("<Leave>", lambda e: start_btn.config(bg=TEAL))

    root.mainloop()


# =====================================================
if __name__ == "__main__":
    print("\n" + "="*60)
    print("  INITIALIZING ERP BOM FORMATTER  v3.0")
    print("="*60)
    create_gui()