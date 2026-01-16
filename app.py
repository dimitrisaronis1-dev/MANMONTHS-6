import streamlit as st
import openpyxl
from openpyxl.styles import PatternFill, Border, Side, Font
import copy
from datetime import datetime
from dateutil.relativedelta import relativedelta
import re
import random
import os
import pandas as pd
import requests
import io

# --- Streamlit Page Configuration ---
st.set_page_config(
    page_title='Person-Month Allocation Tool',
    page_icon='📊',
    layout='wide'
)

# --- Main Title ---
st.title('📊 Person-Month Allocation Tool')

# --- Logo and Custom CSS for positioning ---
# Corrected LOGO_URL for your repository
LOGO_URL = "https://raw.githubusercontent.com/dimitrisaronis1-dev/MANMONTHS-6/main/SPACE%20LOGO_colored%20horizontal.png"

st.markdown(
    """
    <style>
        .logo-container {
            position: absolute;
            top: 20px;
            right: 20px;
            width: 150px;
            z-index: 1000;
        }
        .logo-container img {
            max-width: 100%;
            height: auto;
        }
    </style>
    """,
    unsafe_allow_html=True
)

response_logo = requests.get(LOGO_URL)
if response_logo.status_code == 200:
    st.markdown(
        f"<div class='logo-container'>",
        unsafe_allow_html=True
    )
    st.image(response_logo.content, use_column_width=False)
    st.markdown(
        f"</div>",
        unsafe_allow_html=True
    )
else:
    st.error(f"Could not load logo image from: {LOGO_URL}. Please check the URL and that the file exists in your GitHub repository.")

# --- Template File Download ---
# Corrected TEMPLATE_FILE_URL for your repository
TEMPLATE_FILE_URL = "https://raw.githubusercontent.com/dimitrisaronis1-dev/MANMONTHS-6/main/AM%20TEST%201.xlsx"

st.write("Fetching template file from GitHub...")
try:
    response_template = requests.get(TEMPLATE_FILE_URL)
    response_template.raise_for_status() # Raise an exception for HTTP errors (4xx or 5xx)
    template_file_content = io.BytesIO(response_template.content)
    wb_template = openpyxl.load_workbook(template_file_content) # Use wb_template for the template workbook
    st.success("Template file downloaded and loaded successfully.")
except requests.exceptions.RequestException as e:
    st.error(f"Error downloading template file from GitHub: {e}. Please check the URL or your internet connection. Expected URL: {TEMPLATE_FILE_URL}")
    st.stop() # Stop the app if template cannot be loaded
except Exception as e:
    st.error(f"Error loading template file: {e}. Ensure it's a valid Excel file.")
    st.stop() # Stop the app if template cannot be loaded

# ------------------------------------------------
# Συναρτήσεις ημερομηνιών (Functions for dates)
# ------------------------------------------------
def parse_date(text, is_start=True):
    text = str(text).strip()

    if "σήμερα" in text.lower() or "simera" in text.lower():
        if not is_start:
            return datetime.today()
        else:
            pass

    if re.match(r"^\d{4}$", text):
        if is_start:
            return datetime.strptime("01/01/" + text, "%d/%m/%Y")
        else:
            return datetime.strptime("31/12/" + text, "%d/%m/%Y")
    elif re.match(r"^\d{1,2}/\d{4}$", text):
        if is_start:
            return datetime.strptime("01/" + text, "%d/%m/%Y")
        else:
            d = datetime.strptime("01/" + text, "%d/%m/%Y")
            return d + relativedelta(months=1) - relativedelta(days=1)
    elif re.match(r"^\d{1,2}/\d{1,2}/\d{4}$", text):
        return datetime.strptime(text, "%d/%m/%Y")
    else:
        raise ValueError(f"Unsupported date format: '{text}'. Expected 'YYYY', 'M/YYYY' or 'MM/YYYY', 'D/M/YYYY' or 'DD/MM/YYYY', or 'Σήμερα' (for end date).")

def parse_period(p):
    p_cleaned = str(p).strip().replace("—", "-").replace("–", "-")

    if re.match(r"^\d{4}$", p_cleaned):
        return parse_date(p_cleaned, True), parse_date(p_cleaned, False)

    parts = p_cleaned.split("-")
    if len(parts) != 2:
        raise ValueError(f"Invalid period format: '{p}'. Expected 'YYYY' or 'START_DATE-END_DATE'.")
    a, b = parts
    return parse_date(a, True), parse_date(b, False)

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
    luminance = (0.299 * rgb[0] + 0.587 * rgb[1] + 0.114 * rgb[2]) / 255
    return luminance > 0.5

# ------------------------------------------------
# Core Processing Logic Function
# ------------------------------------------------
def process_allocation(uploaded_input_file, template_workbook):
    # Define colors and border style
    yellow = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
    red_fill = PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid")
    green_fill = PatternFill(start_color="00FF00", end_color="00FF00", fill_type="solid")
    thin_border = Border(left=Side(style='thin'),
                         right=Side(style='thin'),
                         top=Side(style='thin'),
                         bottom=Side(style='thin'))

    # ------------------------------------------------
    # Διαβάζουμε INPUT και βρίσκουμε στήλες (Read INPUT and find columns)
    # ------------------------------------------------
    try:
        # Load input workbook from uploaded file
        wb_in = openpyxl.load_workbook(uploaded_input_file)
        ws_in = wb_in.active
    except Exception as e:
        st.error(f"Error loading input file: {e}. Please ensure it's a valid Excel file.")
        return None, None, None, None, None # Return None for all if error

    headers = {}
    for c in range(1, ws_in.max_column + 1):
        val = str(ws_in.cell(1,c).value).strip() if ws_in.cell(1,c).value is not None else ""
        headers[val] = c

    if "ΧΡΟΝΙΚΟ ΔΙΑΣΤΗΜΑ" not in headers or "ΑΝΘΡΩΠΟΜΗΝΕΣ" not in headers:
        st.error("Το input πρέπει να έχει στήλες: 'ΧΡΟΝΙΚΟ ΔΙΑΣΤΗΜΑ' και 'ΑΝΘΡΩΠΟΜΗΝΕΣ' στην πρώτη γραμμή.")
        return None, None, None, None, None

    PERIOD_COL = headers["ΧΡΟΝΙΚΟ ΔΙΑΣΤΗΜΑ"]
    AM_COL = headers["ΑΝΘΡΩΠΟΜΗΝΕΣ"]

    data = []
    all_months = set()
    project_counter = 0

    for r in range(2, ws_in.max_row + 1):
        period = ws_in.cell(r, PERIOD_COL).value
        am_raw = ws_in.cell(r, AM_COL).value
        try:
            am = int(am_raw) if am_raw is not None else 0
        except (ValueError, TypeError):
            am = 0

        if not period or am == 0:
            continue

        try:
            start, end = parse_period(str(period))
        except ValueError as e:
            st.warning(f"Skipping row {r} due to period parsing error: {e}")
            continue

        months = month_range(start, end)
        months_in_period_count = len(months)

        if months_in_period_count > 0:
            am_per_month_ratio = am / months_in_period_count
        else:
            am_per_month_ratio = 0

        data.append({
            "project_id": project_counter,
            "period_str": period,
            "original_am": am,
            "months_in_period": months,
            "months_in_period_count": months_in_period_count,
            "am_per_month_ratio": am_per_month_ratio,
            "allocated_am": 0,
            "unallocated_am": am
        })
        project_counter += 1

        for m in months:
            all_months.add(m)

    if not data:
        st.error("No valid project data found in the input file. Please check the 'ΧΡΟΝΙΚΟ ΔΙΑΣΤΗΜΑ' και 'ΑΝΘΡΩΠΟΜΗΝΕΣ' columns.")
        return None, None, None, None, None

    all_months = sorted(all_months)
    years = sorted(set(y for y,m in all_months))

    data.sort(key=lambda x: x["months_in_period_count"])

    # ------------------------------------------------
    # Ανοίγουμε TEMPLATE (Opening TEMPLATE) - Now use the loaded wb_template
    # ------------------------------------------------
    wb = copy.deepcopy(template_workbook) # Make a deep copy of the loaded template workbook
    ws = wb.active

    # Freeze the first 3 columns (A, B, C) - this means the freeze point is at D1
    ws.freeze_panes = 'D1'

    # Adjusted START_ROW, YEAR_ROW, MONTH_ROW
    START_ROW_DATA = 4
    YEAR_ROW = 2
    MONTH_ROW = 3
    YEARLY_TOTAL_ROW = START_ROW_DATA + 1
    START_COL = 5
    MAX_YEARLY_CAPACITY = 11 # Define MAX_YEARLY_CAPACITY here

    # ------------------------------------------------
    # Καθαρισμός παλιάς περιοχής (Clearing old area)
    # ------------------------------------------------
    merged_cells_to_unmerge = []
    for cell_range_str in list(ws.merged_cells.ranges):
        min_col_mc, min_row_mc, max_col_mc, max_row_mc = openpyxl.utils.cell.range_boundaries(str(cell_range_str))
        if (
           (min_row_mc <= YEAR_ROW <= max_row_mc) or \
           (min_row_mc <= MONTH_ROW <= max_row_mc) or \
           (min_row_mc <= START_ROW_DATA <= max_row_mc) or \
           (min_row_mc <= YEARLY_TOTAL_ROW <= max_row_mc)
        ):
            merged_cells_to_unmerge.append(cell_range_str)

    for cell_range_str in merged_cells_to_unmerge:
        ws.unmerge_cells(str(cell_range_str))

    max_col_to_clear = max(START_COL + len(years) * 12, ws.max_column + 1)

    rows_to_clear_completely = [YEAR_ROW, MONTH_ROW, START_ROW_DATA, YEARLY_TOTAL_ROW]

    for r_clear in rows_to_clear_completely:
        for c_clear in range(1, max_col_to_clear):
            ws.cell(r_clear, c_clear).value = None
            ws.cell(r_clear, c_clear).fill = PatternFill() # Also clear fill in these rows

    for r_clear in range(START_ROW_DATA + 2, ws.max_row + 1):
        for c_clear in range(1, max_col_to_clear):
            ws.cell(r_clear,c_clear).value = None
            ws.cell(r_clear,c_clear).fill = PatternFill() # Clear fill as well

    # ------------------------------------------------
    # Initialize Allocation Data Structures
    # ------------------------------------------------
    yearly_am_totals = {year: 0 for year in years}
    month_allocation_status = {(y, m): None for y in years for m in range(1, 13)}

    # ------------------------------------------------
    # Χτίσιμο ετών & μηνών (Building years & months)
    # ------------------------------------------------
    col = START_COL
    month_col_map = {}

    for y in years:
        year_start_col = col
        year_header_cell = ws.cell(YEAR_ROW, year_start_col) # Reference specific cell for this year
        
        # Apply random fill and font to the first cell of the merged year block
        r_color_func = lambda: random.randint(0,255)
        random_color_hex = '%02X%02X%02X' % (r_color_func(), r_color_func(), r_color_func())
        year_header_cell.fill = PatternFill(start_color=random_color_hex, end_color=random_color_hex, fill_type="solid")

        if not is_light_color(random_color_hex):
            year_header_cell.font = Font(color="FFFFFF")
        else:
            year_header_cell.font = Font(color="000000")

        # Populate month cells for the current year and apply their borders
        for m in range(1,13):
            ws.cell(MONTH_ROW, col).value = m
            ws.cell(MONTH_ROW, col).border = thin_border # Apply border to month cells as they are created
            month_col_map[(y,m)] = col
            col += 1
        year_end_col = col - 1

        # Now merge the cells for the current year
        ws.merge_cells(start_row=YEAR_ROW, start_column=year_start_col, end_row=YEAR_ROW, end_column=year_end_col)
        ws.cell(YEAR_ROW, year_start_col).value = y # Set value to the top-left cell of the merged block
        ws.cell(YEAR_ROW, year_start_col).border = thin_border # Apply border *after* merge to the top-left cell of the merged region.

    # This loop is removed as borders are applied during cell creation or after merge for year header
    # for c_border in range(START_COL, col):
    #     ws.cell(YEAR_ROW, c_border).border = thin_border
    #     ws.cell(MONTH_ROW, c_border).border = thin_border

    ws.cell(YEARLY_TOTAL_ROW, 2).value = "ΕΤΗΣΙΑ ΣΥΝΟΛΑ"
    ws.cell(YEARLY_TOTAL_ROW, 2).font = Font(bold=True)
    ws.cell(YEARLY_TOTAL_ROW, 2).border = thin_border

    # ------------------------------------------------
    # Γραμμές & μπάρες (Greedy Allocation)
    # ------------------------------------------------
    row = START_ROW_DATA + 2

    unallocated_projects = []
    yearly_overages = {}

    for project_idx, project_data in enumerate(data):
        period_str = project_data["period_str"]
        original_am = project_data["original_am"]
        months_in_period = project_data["months_in_period"]
        project_id = project_data["project_id"]
        allocated_count = 0
        unallocated_count = original_am
        project_unallocated_reason = []

        ws.cell(row,2).value = period_str
        ws.cell(row,2).border = thin_border
        ws.cell(row,3).value = original_am
        ws.cell(row,3).border = thin_border

        for (y, m) in sorted(months_in_period):
            if allocated_count >= original_am:
                break

            if (y,m) in month_col_map:
                if yearly_am_totals[y] >= MAX_YEARLY_CAPACITY:
                    if f"Year {y} capacity reached" not in project_unallocated_reason:
                        project_unallocated_reason.append(f"Year {y} capacity reached")
                    continue

                if month_allocation_status[(y,m)] is not None:
                    occupying_project_id = month_allocation_status[(y,m)]
                    if f"Month {m}/{y} already allocated by Project {occupying_project_id}" not in project_unallocated_reason:
                        project_unallocated_reason.append(f"Month {m}/{y} already allocated by Project {occupying_project_id}")
                    continue

                cell_to_fill = ws.cell(row, month_col_map[(y,m)])
                cell_to_fill.value = 'X'
                cell_to_fill.fill = yellow
                cell_to_fill.border = thin_border

                yearly_am_totals[y] += 1
                month_allocation_status[(y,m)] = project_id
                allocated_count += 1
                unallocated_count -= 1

        project_data["allocated_am"] = allocated_count
        project_data["unallocated_am"] = unallocated_count

        if unallocated_count > 0:
            unallocated_projects.append({
                "period": period_str,
                "original_am": original_am,
                "allocated_am": allocated_count,
                "unallocated_am": unallocated_count,
                "reasons": "; ".join(list(set(project_unallocated_reason)))
            })
            ws.cell(row, 3).font = Font(color="FF0000", bold=True)
        else:
            ws.cell(row, 3).font = Font(color="000000")

        for c_border in range(START_COL, col):
            ws.cell(row, c_border).border = thin_border

        row += 1

    for y in years:
        if y in yearly_am_totals:
            col_for_year_total = month_col_map[(y, 1)]
            year_month_cols = [month_col_map[(y, m)] for m in range(1, 13) if (y, m) in month_col_map]
            if year_month_cols:
                year_start_col = min(year_month_cols)
                year_end_col = max(year_month_cols)
                if year_start_col != year_end_col:
                    ws.merge_cells(start_row=YEARLY_TOTAL_ROW, start_column=year_start_col, end_row=YEARLY_TOTAL_ROW, end_column=year_end_col)
                col_for_year_total = year_start_col
            else:
                continue

            total_cell = ws.cell(YEARLY_TOTAL_ROW, col_for_year_total)
            total_cell.value = yearly_am_totals[y]
            total_cell.font = Font(bold=True)
            total_cell.border = thin_border
            total_cell.alignment = openpyxl.styles.Alignment(horizontal='center', vertical='center')

            if yearly_am_totals[y] >= MAX_YEARLY_CAPACITY:
                year_header_cell = ws.cell(YEAR_ROW, col_for_year_total)
                year_header_cell.fill = red_fill
                year_header_cell.font = Font(color="FFFFFF", bold=True)

                total_cell.fill = red_fill
                total_cell.font = Font(color="FFFFFF", bold=True)
                if yearly_am_totals[y] > MAX_YEARLY_CAPACITY:
                    yearly_overages[y] = yearly_am_totals[y] - MAX_YEARLY_CAPACITY
            elif yearly_am_totals[y] > 0:
                total_cell.fill = green_fill
                total_cell.font = Font(color="000000", bold=True)

    for c_width in range(START_COL, col):
        ws.column_dimensions[openpyxl.utils.get_column_letter(c_width)].width = 2.5

    # ------------------------------------------------
    # Save & download
    # ------------------------------------------------
    ws.title = 'ΑΝΑΛΥΣΗ'

    cv_sheet = wb.create_sheet(title='CV', index=0)
    for row_idx, row_data in enumerate(ws_in.iter_rows()):
        for col_idx, cell in enumerate(row_data):
            new_cell = cv_sheet.cell(row=row_idx + 1, column=col_idx + 1, value=cell.value)
            if cell.has_style:
                new_cell.font = copy.copy(cell.font)
                new_cell.border = copy.copy(cell.border)
                new_cell.fill = copy.copy(cell.fill)
                new_cell.number_format = cell.number_format

    for col_idx in range(1, ws_in.max_column + 1):
        col_letter = openpyxl.utils.get_column_letter(col_idx)
        if col_letter in ws_in.column_dimensions:
            cv_sheet.column_dimensions[col_letter].width = ws_in.column_dimensions[col_letter].width

    # Add the formula to cell A6 of the 'ΑΝΑΛΥΣΗ' sheet
    ws['A6'] = '=MATCH(B6,CV!$B$2:$B$100,0)'

    # Save the workbook to a BytesIO object
    excel_file_buffer = io.BytesIO()
    wb.save(excel_file_buffer)
    excel_file_buffer.seek(0) # Rewind the buffer to the beginning

    return excel_file_buffer, yearly_am_totals, unallocated_projects, MAX_YEARLY_CAPACITY, yearly_overages


# --- Streamlit UI for file upload and processing ---
st.markdown("---")
st.header("1. Upload Input Excel File")
uploaded_input_file = st.file_uploader("Παρακαλώ ανεβάστε το INPUT excel (μόνο 2 στήλες: ΧΡΟΝΙΚΟ ΔΙΑΣΤΗΜΑ και ΑΝΘΡΩΠΟΜΗΝΕΣ)", type=["xlsx", "xls"])

if uploaded_input_file is None:
    st.info("Παρακαλώ ανεβάστε ένα αρχείο Excel για να ξεκινήσετε.")
else:
    with st.spinner("Processing allocation..."):
        excel_buffer, yearly_totals, unallocated_projs, max_cap, yearly_over = process_allocation(uploaded_input_file, wb_template)

    if excel_buffer is not None:
        # Extract base name without extension from the uploaded file's name
        if uploaded_input_file.name:
            base_name = os.path.splitext(uploaded_input_file.name)[0]
        else:
            base_name = "output" # Default if file name somehow missing

        output_filename = f"{base_name}_ΚΑΤΑΝΟΜΗ ΑΜ.xlsx"

        st.download_button(
            label="Download Resulting Excel File",
            data=excel_buffer,
            file_name=output_filename,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

        # ------------------------------------------------
        # Output Summary for Streamlit
        # ------------------------------------------------
        st.markdown("--- ")
        st.header("Allocation Summary")

        st.write(f"Max yearly capacity per year: {max_cap} person-months")

        st.subheader("Yearly Person-Month Totals:")
        for year, total_am in sorted(yearly_totals.items()):
            status = "(Capacity Reached)" if total_am >= max_cap else ""
            if year in yearly_over:
                status = f"(OVER CAPACITY by {yearly_over[year]})"
            st.write(f"  Year {year}: {total_am} {status}")

        if unallocated_projs:
            st.subheader("Projects with Unallocated Person-Months:")
            for proj in unallocated_projs:
                st.write(f"  Period: {proj['period']}, Original AM: {proj['original_am']}, Allocated AM: {proj['allocated_am']}, Unallocated AM: {proj['unallocated_am']}")
                if proj['reasons']:
                    st.write(f"    Reasons for unallocation: {proj['reasons']}")
        else:
            st.success("All person-months were allocated successfully.")