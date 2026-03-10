import os
import re
import smartsheet
import pdfrw
from pdfrw import PdfDict, PdfObject, PdfString
from datetime import datetime


# =========================================================
# CONFIG  (ONLY EDIT THESE)
# =========================================================

SMARTSHEET_TOKEN = os.environ["SMARTSHEET_TOKEN"] 
SHEET_ID = 6053038855769988

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
TEMPLATE_FILE = os.path.join(BASE_DIR, "Transfer of Assets Form testing 13.pdf")

MAX_LINES_PER_PAGE = 30   # increase to reduce number of pages


# =========================================================
# COLUMN NAMES (must match exactly in sheet)
# =========================================================

COL_ROW_ID = "Row ID"
COL_DATE_REQUESTED = "Date requested"
COL_FULL_REQUEST = "Full request"
COL_CELL_NAME = "Cell name"
COL_REQUESTOR_NAME = "Requesters name"
COL_NEED_BY = "Need by date"
COL_ARI_COMPLETE = "ARI complete"


# =========================================================
# SAFETY CHECK
# =========================================================

print("Using template:", TEMPLATE_FILE)

if not os.path.exists(TEMPLATE_FILE):
    raise FileNotFoundError("❌ Template not found — fix TEMPLATE_FILE path")


# =========================================================
# PDF HELPERS
# =========================================================

def set_need_appearances(pdf):
    if not getattr(pdf.Root, "AcroForm", None):
        pdf.Root.AcroForm = PdfDict()
    pdf.Root.AcroForm.update(PdfDict(NeedAppearances=PdfObject("true")))

def iter_fields(pdf):
    def walk(f):
        if isinstance(f, list):
            for x in f:
                yield from walk(x)
            return
        kids = getattr(f, "Kids", None)
        if kids:
            for k in kids:
                yield from walk(k)
        else:
            if getattr(f, "T", None):
                yield (str(f.T).strip("()"), f)
    yield from walk(pdf.Root.AcroForm.Fields)

def fill(field, value):
    field.V = PdfString.encode(str(value or ""))
    if hasattr(field, "AP"):
        field.AP = None

def chunk_by_lines(text, max_lines):
    lines = [ln.strip() for ln in str(text or "").splitlines() if ln.strip()]
    chunks = []
    for i in range(0, len(lines), max_lines):
        chunks.append("\n".join(lines[i:i+max_lines]))
    return chunks

def clean_filename(text):
    return re.sub(r'[\\/*?:"<>|]', "", str(text))


# =========================================================
# SMARTSHEET HELPERS
# =========================================================

def build_col_map(sheet):
    return {c.title.strip(): c.id for c in sheet.columns}

def get_cell(row, col_id):
    for c in row.cells:
        if c.column_id == col_id:
            return c
    return None

def get_cell_value(row, col_id):
    c = get_cell(row, col_id)
    if not c:
        return None
    return c.value if c.value is not None else c.display_value

def set_checkbox_true(ss, sheet_id, row_id, col_id):
    new_row = smartsheet.models.Row()
    new_row.id = row_id

    cell = smartsheet.models.Cell()
    cell.column_id = col_id
    cell.value = True

    new_row.cells.append(cell)
    ss.Sheets.update_rows(sheet_id, [new_row])


# =========================================================
# FIND PAGE SUFFIXES (Row1, Row1_2, Row1_3...)
# =========================================================

def get_suffixes(field_map):
    suffixes = []
    for name in field_map:
        m = re.match(r"Item part number and descriptionRow1(_\d+)?$", name)
        if m:
            suffixes.append(m.group(1) or "")
    def key(x):
        return 0 if x == "" else int(x.replace("_",""))
    return sorted(set(suffixes), key=key)


# =========================================================
# RUN
# =========================================================

print("\n--- Script Started ---")

ss = smartsheet.Smartsheet(SMARTSHEET_TOKEN)
ss.errors_as_exceptions(True)

sheet = ss.Sheets.get_sheet(SHEET_ID)
col_map = build_col_map(sheet)

processed = 0
skipped = 0

for row in sheet.rows:

    # Skip if ARI already complete
    ari_complete = get_cell_value(row, col_map[COL_ARI_COMPLETE])
    if ari_complete is True:
        skipped += 1
        continue

    date_val = str(get_cell_value(row, col_map[COL_DATE_REQUESTED]) or "").strip()
    req_val = get_cell_value(row, col_map[COL_FULL_REQUEST])

    if not date_val or not req_val:
        continue

    print("Processing row:", row.row_number)

    # Pull values
    row_id_val = str(get_cell_value(row, col_map[COL_ROW_ID]) or "").strip()
    cell_name = str(get_cell_value(row, col_map[COL_CELL_NAME]) or "")
    requester = str(get_cell_value(row, col_map[COL_REQUESTOR_NAME]) or "")
    need_by = str(get_cell_value(row, col_map[COL_NEED_BY]) or "")

    if not row_id_val:
        row_id_val = str(row.id)

    req_val = "\n".join([ln.strip() for ln in str(req_val).splitlines() if ln.strip()])

    # Load PDF
    pdf = pdfrw.PdfReader(TEMPLATE_FILE)
    set_need_appearances(pdf)
    field_map = {n:f for n,f in iter_fields(pdf)}

    # Header
    if "Date of request" in field_map:
        fill(field_map["Date of request"], date_val)

    if "Requestors Name" in field_map:
        fill(field_map["Requestors Name"], requester)

    if "Need by date" in field_map:
        fill(field_map["Need by date"], need_by)

    # Rows
    suffixes = get_suffixes(field_map)
    chunks = chunk_by_lines(req_val, MAX_LINES_PER_PAGE)

    for i, suf in enumerate(suffixes):
        desc = f"Item part number and descriptionRow1{suf}"
        cell = f"Cell nameRow1{suf}"
        req  = f"Requestors nameRow1{suf}"
        line = f"Line numberRow1{suf}"

        text = chunks[i] if i < len(chunks) else ""

        if desc in field_map: fill(field_map[desc], text)
        if cell in field_map: fill(field_map[cell], cell_name)
        if req in field_map:  fill(field_map[req], requester)
        if line in field_map: fill(field_map[line], i+1)

    # Footer
    if "DATE" in field_map:
        fill(field_map["DATE"], date_val)
    if "DATE_2" in field_map:
        fill(field_map["DATE_2"], date_val)

    # =============================
    # FILENAME (RowID + Cell + Date + Pale)
    # =============================

    safe_row = clean_filename(row_id_val)
    safe_cell = clean_filename(cell_name)
    safe_date = clean_filename(date_val.replace("/", "-"))

    out_name = f"{safe_row} - {safe_cell} - {safe_date}.pdf"

    pdfrw.PdfWriter().write(out_name, pdf)

    # Upload
    with open(out_name, "rb") as f:
        ss.Attachments.attach_file_to_row(SHEET_ID, row.id, (out_name, f, "application/pdf"))

    # Mark complete
    set_checkbox_true(ss, SHEET_ID, row.id, col_map[COL_ARI_COMPLETE])

    print("✅ Uploaded + marked complete:", out_name)
    processed += 1


print("\n--- Summary ---")
print("Processed:", processed)
print("Skipped (already complete):", skipped)
print("--- Done ---\n")



