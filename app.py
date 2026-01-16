import streamlit as st
import openpyxl
from openpyxl.styles import PatternFill, Border, Side, Font
import copy
from datetime import datetime
from dateutil.relativedelta import relativedelta
import re
import random
import os
from io import BytesIO

# -----------------------------
# Page config
# -----------------------------
st.set_page_config(page_title="Man-Month Allocation", layout="wide")

# -----------------------------
# Header with logo (top right)
# -----------------------------
col1, col2 = st.columns([5,1])
with col1:
    st.title("📊 Man-Month Allocation Tool")
with col2:
    st.image(
        "https://raw.githubusercontent.com/dimitrisaronis1-dev/MANMONTHS-6/main/SPACE%20LOGO_colored%20horizontal.png",
        use_container_width=True
    )

st.markdown("---")

# -----------------------------
# Constants
# -----------------------------
TEMPLATE_FILE = "AM TEST 1.xlsx"
MAX_YEARLY_CAPACITY = 11

yellow = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
red_fill = PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid")
green_fill = PatternFill(start_color="00FF00", end_color="00FF00", fill_type="solid")

thin_border = Border(
    left=Side(style='thin'),
    right=Side(style='thin'),
    top=Side(style='thin'),
    bottom=Side(style='thin')
)

# -----------------------------
# Date functions
# -----------------------------
def parse_date(text, is_start=True):
    text = str(text).strip()

    if "σήμερα" in text.lower() or "simera" in text.lower():
        return datetime.today()

    if re.match(r"^\d{4}$", text):
        return datetime.strptime(("01/01/" if is_start else "31/12/") + text, "%d/%m/%Y")

    elif re.match(r"^\d{1,2}/\d{4}$", text):
        d = datetime.strptime("01/" + text, "%d/%m/%Y")
        return d if is_start else d + relativedelta(months=1) - relativedelta(days=1)

    elif re.match(r"^\d{1,2}/\d{1,2}/\d{4}$", text):
        return datetime.strptime(text, "%d/%m/%Y")

    else:
        raise ValueError(f"Unsupported date format: {text}")

def parse_period(p):
    p_cleaned = str(p).strip().replace("—", "-").replace("–", "-")

    if re.match(r"^\d{4}$", p_cleaned):
        return parse_date(p_cleaned, True), parse_date(p_cleaned, False)

    parts = p_cleaned.split("-")
    if len(parts) != 2:
        raise ValueError(f"Invalid period format: {p}")

    return parse_date(parts[0], True), parse_date(parts[1], False)

def month_range(start, end):
    current = datetime(start.year, start.month, 1)
    end = datetime(end.year, end.month, 1)
    out = []
    while current <= end:
        out.append((current.year, current.month))
        current += relativedelta(months=1)
    return out

def is_light_color(hex_color):
    hex_color = hex_color.lstrip('#')
    rgb = tuple(int(hex_color[i:i+2], 16) for i in (0, 2, 4))
    luminance = (0.299*rgb[0] + 0.587*rgb[1] + 0.114*rgb[2]) / 255
    return luminance > 0.5

# -----------------------------
# UI
# -----------------------------
st.subheader("📤 Ανέβασε το INPUT Excel (2 στήλες)")
uploaded_file = st.file_uploader("Excel file", type=["xlsx"])

if uploaded_file:

    if not os.path.exists(TEMPLATE_FILE):
        st.error("❌ Δεν βρέθηκε το AM TEST 1.xlsx στο repo.")
        st.stop()

    with st.spinner("Επεξεργασία αρχείου..."):

        wb_in = openpyxl.load_workbook(uploaded_file)
        ws_in = wb_in.active

        headers = {}
        for c in range(1, ws_in.max_column + 1):
            headers[str(ws_in.cell(1,c).value).strip()] = c

        if "ΧΡΟΝΙΚΟ ΔΙΑΣΤΗΜΑ" not in headers or "ΑΝΘΡΩΠΟΜΗΝΕΣ" not in headers:
            st.error("Το input πρέπει να έχει στήλες: ΧΡΟΝΙΚΟ ΔΙΑΣΤΗΜΑ και ΑΝΘΡΩΠΟΜΗΝΕΣ")
            st.stop()

        PERIOD_COL = headers["ΧΡΟΝΙΚΟ ΔΙΑΣΤΗΜΑ"]
        AM_COL = headers["ΑΝΘΡΩΠΟΜΗΝΕΣ"]

        data = []
        all_months = set()
        project_counter = 0

        for r in range(2, ws_in.max_row + 1):
            period = ws_in.cell(r, PERIOD_COL).value
            am_raw = ws_in.cell(r, AM_COL).value
            try:
                am = int(am_raw) if am_raw else 0
            except:
                am = 0

            if not period or am == 0:
                continue

            try:
                start, end = parse_period(period)
            except:
                continue

            months = month_range(start, end)

            data.append({
                "project_id": project_counter,
                "period_str": period,
                "original_am": am,
                "months_in_period": months
            })
            project_counter += 1

            for m in months:
                all_months.add(m)

        all_months = sorted(all_months)
        years = sorted(set(y for y,m in all_months))
        data.sort(key=lambda x: len(x["months_in_period"]))

        wb = openpyxl.load_workbook(TEMPLATE_FILE)
        ws = wb.active

        ws.freeze_panes = 'D1'

        START_ROW_DATA = 4
        YEAR_ROW = 2
        MONTH_ROW = 3
        YEARLY_TOTAL_ROW = START_ROW_DATA + 1
        START_COL = 5

        yearly_am_totals = {year: 0 for year in years}
        month_allocation_status = {(y, m): None for y in years for m in range(1,13)}

        col = START_COL
        month_col_map = {}

        for y in years:
            year_start_col = col
            year_header_cell = ws.cell(YEAR_ROW, col)

            rand_hex = '%02X%02X%02X' % (random.randint(0,255),random.randint(0,255),random.randint(0,255))
            year_header_cell.fill = PatternFill(start_color=rand_hex, end_color=rand_hex, fill_type="solid")
            year_header_cell.font = Font(color="000000" if is_light_color(rand_hex) else "FFFFFF", bold=True)

            for m in range(1,13):
                ws.cell(MONTH_ROW, col).value = m
                month_col_map[(y,m)] = col
                col += 1

            ws.merge_cells(start_row=YEAR_ROW, start_column=year_start_col, end_row=YEAR_ROW, end_column=col-1)
            year_header_cell.value = y

        row = START_ROW_DATA + 2

        for project in data:
            ws.cell(row,2).value = project["period_str"]
            ws.cell(row,3).value = project["original_am"]

            allocated = 0

            for (y,m) in project["months_in_period"]:
                if allocated >= project["original_am"]:
                    break
                if yearly_am_totals[y] >= MAX_YEARLY_CAPACITY:
                    continue
                if month_allocation_status[(y,m)] is not None:
                    continue

                cell = ws.cell(row, month_col_map[(y,m)])
                cell.value = "X"
                cell.fill = yellow

                yearly_am_totals[y] += 1
                month_allocation_status[(y,m)] = project["project_id"]
                allocated += 1

            row += 1

        ws.title = "ΑΝΑΛΥΣΗ"
        cv_sheet = wb.create_sheet(title='CV', index=0)

        for r_idx, row_cells in enumerate(ws_in.iter_rows()):
            for c_idx, cell in enumerate(row_cells):
                new_cell = cv_sheet.cell(row=r_idx+1, column=c_idx+1, value=cell.value)
                if cell.has_style:
                    new_cell.font = copy.copy(cell.font)
                    new_cell.border = copy.copy(cell.border)
                    new_cell.fill = copy.copy(cell.fill)
                    new_cell.number_format = cell.number_format

        ws['A6'] = '=MATCH(B6,CV!$B$2:$B$100,0)'

        output = BytesIO()
        wb.save(output)
        output.seek(0)

    st.success("✅ Ολοκληρώθηκε!")
    st.download_button(
        "⬇️ Κατέβασε το αποτέλεσμα",
        data=output,
        file_name="KATANOMI_AM.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
