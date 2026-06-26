import os, re, shutil
import pandas as pd
import xlrd
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import threading
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment

TEMPLATE_PATH = r'D:\MAC Excel Code\VS Code\BOM_Generation\Template-Standard_BOM.xlsx'

FILL_GREEN  = PatternFill(start_color="92D050", end_color="92D050", fill_type="solid")
FILL_YELLOW = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
FILL_NONE   = PatternFill(fill_type=None)
FONT_BOLD   = Font(name="Calibri", size=11, bold=True)
FONT_NORMAL = Font(name="Calibri", size=11, bold=False)

# Material code in part code -> RM sheet material keyword
MAT_MAP = {
    'MS'  : ['HRPO', 'MS'],
    'GI'  : ['GI'],
    'SS'  : ['SS'],
    'ACRY': ['ACRY'],
}

# Developer identity -- do not modify
_DEV_NAME = ''.join(['M','a','h','e','s','h',' ','A','r','v','i','n','d',' ','C','h','a','v','a','n'])


# =====================================================
# READ FILE + REBUILD HIERARCHY FROM ROW ORDER
# =====================================================
def read_and_fix_hierarchy(path):
    ext = os.path.splitext(path)[1].lower()
    if ext == ".xls":
        wb  = xlrd.open_workbook(path)
        ws  = wb.sheets()[0]
        headers = [str(ws.cell(0, c).value).strip() for c in range(ws.ncols)]
        rows = []
        for ri in range(1, ws.nrows):
            row = {}
            for ci, h in enumerate(headers):
                row[h] = ws.cell(ri, ci).value
            rows.append(row)
        df = pd.DataFrame(rows)
    else:
        df = pd.read_excel(path, engine="openpyxl", dtype=str)

    for col in df.columns:
        if df[col].dtype == object:
            df[col] = df[col].fillna('').astype(str).str.strip()

    # Rebuild hierarchy by walking rows in order using depth from raw string
    raw_hierarchies = df['Hierarchy'].astype(str).tolist()
    counters  = {}
    new_hier  = []
    prev_depth = 0

    for raw in raw_hierarchies:
        raw_clean = str(raw).strip().rstrip('0').rstrip('.')
        parts = [p for p in raw_clean.split('.') if p != '']
        depth = max(len(parts), 1)

        if depth > prev_depth:
            for d in range(prev_depth + 1, depth + 1):
                counters[d] = 0
            counters[depth] += 1
        elif depth == prev_depth:
            counters[depth] += 1
        else:
            for d in list(counters.keys()):
                if d > depth:
                    del counters[d]
            counters[depth] = counters.get(depth, 0) + 1

        hier = '.'.join(str(counters.get(d, 1)) for d in range(1, depth + 1))
        new_hier.append(hier)
        prev_depth = depth

    df['Hierarchy'] = new_hier
    print("\n  Hierarchy rebuilt:")
    for i, (r, n) in enumerate(zip(raw_hierarchies, new_hier)):
        print(f"    row {i+1:2d}: raw='{r}'  ->  '{n}'")
    return df


# =====================================================
# PARSE PREFIX, MATERIAL, THICKNESS FROM PART CODE
# FL-MS-T2-30965  -> ('FL', 'MS', 2.0)
# BE-MS-480-83-T3-8772 -> ('BE', 'MS', 3.0)
# =====================================================
def parse_part_code(pc):
    code   = str(pc).strip().upper()
    segs   = code.split('-')
    prefix = segs[0] if segs else ''
    material = segs[1] if len(segs) > 1 else ''
    thickness = None
    for seg in segs:
        m = re.match(r'^T(\d+\.?\d*)$', seg)
        if m:
            thickness = float(m.group(1))
            break
    return prefix, material, thickness


# =====================================================
# VALID CHILD PART -- must end with -<2 to 7 digits>
# =====================================================
def is_valid_child_part(pc):
    return bool(re.search(r'-\d{2,7}$', str(pc).strip()))


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
# DIRECT CHILDREN OF A PARENT
# =====================================================
def get_direct_children(parent_h, df):
    prefix = str(parent_h) + '.'
    result = []
    for _, row in df.iterrows():
        h = str(row['Hierarchy'])
        if h.startswith(prefix) and '.' not in h[len(prefix):]:
            result.append(row)
    return result


# =====================================================
# BUILD RM LOOKUP FROM TEMPLATE
# Returns: dict {(mat_keyword, thickness): (item_code, item_name)}
#          list [(item_code, item_name)] for BE dropdown (name has 'sheet')
# =====================================================
def build_rm_lookup(template_path):
    rm = pd.read_excel(template_path, sheet_name='Raw Material Part Code')
    fl_lookup  = {}
    be_options = []

    for _, row in rm.iterrows():
        ic    = str(row['Item Code']).strip()
        iname = str(row['Item Name']).strip()

        # Only sheet items for FL auto-fill
        if '-SH-' in ic:
            m = re.search(r'-(\d+\.?\d*)$', ic)
            if m:
                t = float(m.group(1))
                parts = ic.split('-')
                mat_kw = parts[2] if len(parts) > 2 else ''
                fl_lookup[(mat_kw, t)] = (ic, iname)

        # BE dropdown: item name contains 'sheet'
        if 'sheet' in iname.lower():
            be_options.append((ic, iname))

    print(f"  FL lookup keys : {list(fl_lookup.keys())}")
    print(f"  BE options     : {len(be_options)}")
    return fl_lookup, be_options


# =====================================================
# LOOKUP RM FOR FL PART CODE
# FL-MS-T2-30965 -> material=MS -> try HRPO, MS -> T=2 -> match
# =====================================================
def lookup_fl_rm(pc, fl_lookup):
    _, material, thickness = parse_part_code(pc)
    if thickness is None:
        return '', ''
    keywords = MAT_MAP.get(material, [material])
    for kw in keywords:
        result = fl_lookup.get((kw, thickness))
        if result:
            return result
    return f'[NOT FOUND: {material} T{thickness}]', '[MANUAL REQUIRED]'


# =====================================================
# WRITE CHILD PART BOM SHEET
# =====================================================
def write_child_sheet(ws, leaf_rows):
    for row in ws.iter_rows(min_row=2, max_row=max(ws.max_row, 2)):
        for cell in row:
            cell.value = None
            cell.fill  = FILL_NONE

    for i, r in enumerate(leaf_rows, start=2):
        ws.cell(row=i, column=1, value=r['item_code']).font   = FONT_NORMAL
        ws.cell(row=i, column=2, value=r['item_name']).font   = FONT_NORMAL
        ws.cell(row=i, column=3, value=r['description']).font = FONT_NORMAL
        for col, val in [(4, r['uom']), (5, r['routing']),
                         (6, r['qty']), (7, r['weight'])]:
            c = ws.cell(row=i, column=col, value=val)
            c.font = FONT_NORMAL
            c.alignment = Alignment(horizontal='center')
        # H: yellow fill only when value is properly auto-filled
        ch = ws.cell(row=i, column=8, value=r['item_bom'])
        ch.font = FONT_NORMAL
        is_filled = r['item_bom'] and not str(r['item_bom']).startswith('[')
        ch.fill = FILL_YELLOW if is_filled else FILL_NONE
        for col, val in [(9,  r['item_name_bom']),
                         (10, r['qty_bom']),
                         (11, r['uom_bom']),
                         (12, r['do_not_explode'])]:
            c = ws.cell(row=i, column=col, value=val)
            c.font = FONT_NORMAL
            c.alignment = Alignment(horizontal='center')

    print(f"  Written {len(leaf_rows)} rows to Child Part BOM")


# =====================================================
# WRITE PARENT BOM SHEET
# =====================================================
def write_parent_sheet(ws, parent_sections):
    for row in ws.iter_rows(min_row=2, max_row=max(ws.max_row, 2)):
        for cell in row:
            cell.value = None
            cell.fill  = FILL_NONE
            cell.font  = FONT_NORMAL

    cur = 2
    for section in parent_sections:
        p        = section['parent']
        children = section['children']

        # Write parent data (cols A-G) on current row
        ca = ws.cell(row=cur, column=1, value=p['item_code'])
        ca.font = FONT_BOLD
        ca.fill = FILL_GREEN
        for col, val in [(2, p['item_name']), (3, p['description']),
                         (4, p['uom']),        (5, p['routing']),
                         (6, 1),               (7, p['weight'])]:
            c = ws.cell(row=cur, column=col, value=val)
            c.alignment = Alignment(horizontal='center')

        # Write ALL children (cols H-L)
        # First child shares the SAME row as parent (template pattern)
        # Subsequent children go on the next rows (cols A-G blank)
        for idx, child in enumerate(children):
            # First child: same row as parent (cur stays)
            # Subsequent children: next row
            if idx > 0:
                cur += 1
            ch = ws.cell(row=cur, column=8, value=child['item_code'])
            ch.fill = FILL_YELLOW
            for col, val in [(9,  child['item_name']),
                             (10, child['qty']),
                             (11, child['uom']),
                             (12, child['do_not_explode'])]:
                c = ws.cell(row=cur, column=col, value=val)
                c.alignment = Alignment(horizontal='center')

        cur += 1  # Move to next row after this parent+children block

    print(f"  Written {cur - 2} rows to Parent BOM")


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
        print("  ERP BOM FORMATTER  v4.0")
        print("="*60)

        upd("Reading and rebuilding hierarchy ...", 5)
        df_raw = read_and_fix_hierarchy(file_path)
        original_count = len(df_raw)

        pc_col = 'PART NUMBER' if 'PART NUMBER' in df_raw.columns else 'Part Code'
        print(f"\n  Rows: {original_count}  |  Part Code column: '{pc_col}'")

        df_raw[pc_col]        = df_raw[pc_col].fillna('').astype(str).str.strip()
        df_raw['Description'] = df_raw['Description'].fillna('').astype(str).str.strip()
        if 'Part Name' in df_raw.columns:
            df_raw['Part Name'] = df_raw['Part Name'].fillna('').astype(str).str.strip()

        upd("Loading RM reference ...", 12)
        fl_lookup, be_options = build_rm_lookup(template_path)

        # STEP 1: DELETE BLANK PART CODE ROWS
        upd("Step 1: Cleaning rows ...", 20)
        print("\nSTEP 1 - DELETE BLANK Part Code ROWS")
        is_blank  = df_raw[pc_col] == ''
        rule_a    = is_blank &  (df_raw['Description'].str.lower() == 'sheet')
        rule_b    = is_blank & ~(df_raw['Description'].str.lower() == 'sheet')
        del_count = int((rule_a | rule_b).sum())
        df = df_raw[~(rule_a | rule_b)].reset_index(drop=True)
        print(f"  Deleted {del_count} rows | Remaining: {len(df)}")

        # STEP 2: TAG
        upd("Step 2: Tagging hierarchy nodes ...", 30)
        df = tag_leaf_parent(df)

        # STEP 3: CHILD PART BOM
        upd("Step 3: Building Child Part BOM ...", 45)
        print("\nSTEP 3 - CHILD PART BOM")

        leaf_rows         = []
        seen_child        = set()
        skipped_hw        = []
        be_manual_needed  = []

        for _, row in df[df['__is_leaf']].iterrows():
            pc    = str(row[pc_col]).strip()
            pname = str(row.get('Part Name', '')).strip() or str(row.get('Description', '')).strip()
            desc  = str(row.get('Description', '')).strip()

            if not is_valid_child_part(pc):
                skipped_hw.append(pc)
                print(f"  SKIP HW/BO: '{pc}'")
                continue

            if pc in seen_child:
                print(f"  SKIP DUP : '{pc}'")
                continue
            seen_child.add(pc)

            try:    wt = float(row.get('Unit Weight (kg)', row.get('Unit Weight (Kgs)', 0)) or 0)
            except: wt = ''

            prefix, _, _ = parse_part_code(pc)
            item_bom, item_name_bom = '', ''

            if prefix == 'FL':
                item_bom, item_name_bom = lookup_fl_rm(pc, fl_lookup)
                print(f"  FL: {pc} -> '{item_bom}'")
            elif prefix == 'BE':
                item_bom, item_name_bom = lookup_fl_rm(pc, fl_lookup)
                if item_bom and not item_bom.startswith('['):
                    print(f"  BE auto-fill: {pc} -> '{item_bom}'")
                else:
                    item_bom      = '[SELECT FROM DROPDOWN]'
                    item_name_bom = '[SELECT FROM DROPDOWN]'
                    be_manual_needed.append(pc)
                    print(f"  BE manual needed: {pc}")

            leaf_rows.append({
                'item_code'     : pc,
                'item_name'     : pname,
                'description'   : desc,
                'uom'           : 'Nos',
                'routing'       : '',
                'qty'           : 1,
                'weight'        : wt,
                'item_bom'      : item_bom,
                'item_name_bom' : item_name_bom,
                'qty_bom'       : wt,
                'uom_bom'       : 'Kg',
                'do_not_explode': 1,
            })

        print(f"\n  Leaf rows added   : {len(leaf_rows)}")
        print(f"  HW/BO skipped     : {len(skipped_hw)}")
        print(f"  BE manual needed  : {len(be_manual_needed)}")

        # STEP 4: PARENT BOM
        upd("Step 4: Building Parent BOM ...", 62)
        print("\nSTEP 4 - PARENT BOM")

        parent_sections  = []
        seen_parents     = set()

        for _, prow in df[~df['__is_leaf']].iterrows():
            pc    = str(prow[pc_col]).strip()
            pname = str(prow.get('Part Name', '')).strip() or str(prow.get('Description', '')).strip()
            desc  = str(prow.get('Description', '')).strip()

            if pc in seen_parents:
                print(f"  SKIP DUP parent: '{pc}'")
                continue
            seen_parents.add(pc)

            try:    wt = float(prow.get('Unit Weight (kg)', prow.get('Unit Weight (Kgs)', 0)) or 0)
            except: wt = ''

            children_out = []
            for child in get_direct_children(prow['Hierarchy'], df):
                cpc   = str(child[pc_col]).strip()
                cname = str(child.get('Part Name', '')).strip() or str(child.get('Description', '')).strip()
                try:    cqty = float(child.get('Unit Qty', 1) or 1)
                except: cqty = 1.0
                children_out.append({
                    'item_code'     : cpc,
                    'item_name'     : cname,
                    'qty'           : cqty,
                    'uom'           : 'Nos',
                    'do_not_explode': 1,
                })

            parent_sections.append({
                'parent'  : {'item_code': pc, 'item_name': pname,
                             'description': desc, 'uom': 'Nos',
                             'routing': '', 'weight': wt},
                'children': children_out,
            })

        print(f"  Parent sections : {len(parent_sections)}")

        # SAVE
        upd("Saving into template ...", 80)
        base        = os.path.splitext(file_path)[0]
        output_path = f"{base}_ERP_BOM.xlsx"

        shutil.copy2(template_path, output_path)
        wb = load_workbook(output_path)

        write_child_sheet(wb['Child Part Bom With RM'], leaf_rows)
        write_parent_sheet(wb['Bom Template-1'], parent_sections)

        if be_options:
            if 'BE_RM_Reference' in wb.sheetnames:
                del wb['BE_RM_Reference']
            wsr = wb.create_sheet('BE_RM_Reference')
            wsr['A1'] = 'BE Part RM Options - Item Name contains Sheet'
            wsr['A1'].font = Font(bold=True, size=11, name="Calibri")
            wsr['A2'].font = wsr['B2'].font = Font(bold=True, name="Calibri")
            wsr['A2'] = 'Item Code'
            wsr['B2'] = 'Item Name'
            for ri, (ic, iname) in enumerate(be_options, start=3):
                wsr.cell(row=ri, column=1, value=ic)
                wsr.cell(row=ri, column=2, value=iname)
            wsr.column_dimensions['A'].width = 34
            wsr.column_dimensions['B'].width = 54

        wb.save(output_path)
        upd("Done!", 100)

        print(f"\n  Saved: {output_path}")
        print("="*60)

        be_note = ""
        if be_manual_needed:
            be_note = (f"\n\nNOTE: {len(be_manual_needed)} BE parts need manual RM selection."
                       f"\nSee BE_RM_Reference sheet for options.")

        messagebox.showinfo("Success",
            f"ERP BOM generated!\n\n"
            f"Summary\n"
            f"{'-'*36}\n"
            f"  Input rows          : {original_count}\n"
            f"  Blank PC deleted    : {del_count}\n"
            f"  HW/BO skipped       : {len(skipped_hw)}\n"
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
# GUI  -- all widget text uses ASCII only
# =====================================================
def create_gui():
    NAVY = "#1B2A4A"; TEAL = "#1E6E6E"; ACCENT = "#E8A838"
    BG   = "#F0F4F8"; WHITE = "#FFFFFF"; GREY = "#7A8A99"; DARK = "#2C3E50"

    root = tk.Tk()
    root.title("ERP BOM Formatter  v4.0")
    root.configure(bg=BG)
    root.geometry("740x590")
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
    tk.Label(hdr, text="v4.0",
             font=("Montserrat", 10), bg=NAVY, fg=ACCENT
             ).pack(side=tk.LEFT, pady=20)

    # Footer -- packed before body
    ftr = tk.Frame(root, bg=NAVY, height=28)
    ftr.pack(fill=tk.X, side=tk.BOTTOM)
    ftr.pack_propagate(False)
    tk.Label(ftr,
             text="Concept & Developed by " + _DEV_NAME + "  |  ERP BOM Formatter v4.0  |  NIDO Automation",
             bg=NAVY, fg=GREY, font=("Montserrat", 8)
             ).pack(pady=6)

    # Body
    body = tk.Frame(root, bg=BG)
    body.pack(fill=tk.BOTH, expand=True, padx=22, pady=14)

    # Steps
    sf = tk.LabelFrame(body, text="  Processing Steps  ", bg=BG, fg=TEAL,
                       font=("Montserrat", 9, "bold"), bd=1, relief=tk.GROOVE,
                       padx=8, pady=6)
    sf.pack(fill=tk.X, pady=(0, 10))

    steps = [
        ("1", "Rebuild Hierarchy from row order -- fixes 1.10 vs 1.1 Excel float issue"),
        ("2", "Delete rows where Part Code is blank"),
        ("3", "Child BOM: valid parts only (end -digits); no HW/BO; no duplicates; Qty=1"),
        ("4", "FL: Item(BOM Details) auto-filled by material+thickness from RM sheet"),
        ("5", "BE: marked for manual selection; reference sheet provided"),
        ("6", "Parent BOM: parent row (green) + direct children (yellow); no duplicates"),
    ]
    for num, text in steps:
        rf = tk.Frame(sf, bg=BG)
        rf.pack(fill=tk.X, pady=1)
        bg_num = TEAL if num.strip().isdigit() else BG
        fg_num = WHITE if num.strip().isdigit() else BG
        tk.Label(rf, text=" " + num + " ", bg=bg_num, fg=fg_num,
                 font=("Montserrat", 8, "bold")).pack(side=tk.LEFT, padx=(0, 7))
        tk.Label(rf, text=text, bg=BG, fg=DARK,
                 font=("Montserrat", 9), anchor='w').pack(side=tk.LEFT)

    # File row helper
    def make_file_row(label, default='', filetypes=None):
        lf = tk.LabelFrame(body, text="  " + label + "  ", bg=BG, fg=TEAL,
                            font=("Montserrat", 9, "bold"), bd=1, relief=tk.GROOVE,
                            padx=10, pady=8)
        lf.pack(fill=tk.X, pady=(0, 8))
        fi = tk.Frame(lf, bg=BG)
        fi.pack(fill=tk.X)
        entry = tk.Entry(fi, width=66, font=("Montserrat", 9), bg=WHITE,
                         relief=tk.FLAT, highlightthickness=1,
                         highlightbackground="#CCCCCC")
        if default:
            entry.insert(0, default)
        entry.pack(side=tk.LEFT, padx=(0, 8), ipady=5)
        ft = filetypes or [("Excel Files", "*.xlsx *.xls"), ("All Files", "*.*")]
        def browse(e=entry, ft=ft):
            fn = filedialog.askopenfilename(title="Select " + label, filetypes=ft)
            if fn:
                e.delete(0, tk.END)
                e.insert(0, fn)
        btn = tk.Button(fi, text="Browse...", command=browse,
                        bg=ACCENT, fg=NAVY, font=("Montserrat", 9, "bold"),
                        relief=tk.FLAT, cursor="hand2", padx=10, pady=5)
        btn.pack(side=tk.LEFT)
        btn.bind("<Enter>", lambda e, b=btn: b.config(bg="#D4932A"))
        btn.bind("<Leave>", lambda e, b=btn: b.config(bg=ACCENT))
        return entry

    tmpl_entry = make_file_row("Template File (.xlsx)", default=TEMPLATE_PATH,
                               filetypes=[("Excel Files", "*.xlsx"), ("All Files", "*.*")])
    file_entry = make_file_row("Input BOM File (.xls / .xlsx)")

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
        for p, lbl in [(fp, "Input BOM"), (tp, "Template")]:
            if not os.path.exists(p):
                messagebox.showerror("Not Found", lbl + " file not found:\n" + p)
                return
        progress_bar["value"] = 0
        status_lbl.config(text="Processing ...", fg=TEAL)
        root.update()
        threading.Thread(target=process_file,
                         args=(fp, tp, progress_bar, root, status_lbl),
                         daemon=True).start()

    btn_row = tk.Frame(body, bg=BG)
    btn_row.pack(pady=2)
    start_btn = tk.Button(btn_row, text="START PROCESSING", command=start,
                          bg=TEAL, fg=WHITE, font=("Montserrat", 12, "bold"),
                          relief=tk.FLAT, cursor="hand2", padx=44, pady=10)
    start_btn.pack()
    start_btn.bind("<Enter>", lambda e: start_btn.config(bg=NAVY))
    start_btn.bind("<Leave>", lambda e: start_btn.config(bg=TEAL))

    root.mainloop()


if __name__ == "__main__":
    print("\n" + "="*60)
    print("  INITIALIZING ERP BOM FORMATTER  v4.0")
    print("="*60)
    create_gui()