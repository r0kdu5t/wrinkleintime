import os
from datetime import date, timedelta
import openpyxl
from openpyxl.styles import Alignment, Font


def find_sunday(input_date: date) -> str:
    """Returns the date of the first Sunday following the input_date as a string (e.g., '25Oct2026')."""
    vba_weekday = (
        input_date.isoweekday() + 1 if input_date.isoweekday() < 7 else 1
    )
    days_to_add = 8 - vba_weekday
    sunday_date = input_date + timedelta(days=days_to_add)
    return sunday_date.strftime("%d%b%Y")


def prepend_time_sheet(file_name: str = "Master_Timesheets.xlsx"):
    """Creates a weekly timesheet and prepends it to an existing or brand-new workbook."""
    d_week_end = find_sunday(date.today())
    sheet_title = f"WE{d_week_end}"

    # 1. Load or Create the Workbook
    if os.path.exists(file_name):
        wb = openpyxl.load_workbook(file_name)
        print(f"Found existing workbook '{file_name}'.")

        # Skip if this week's sheet already exists to avoid duplicates
        if sheet_title in wb.sheetnames:
            print(f"Sheet '{sheet_title}' already exists. Skipping creation.")
            return
        # Create a new sheet
        ws = wb.create_sheet(title=sheet_title)
    else:
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = sheet_title
        print(f"Workbook not found. Creating new file: '{file_name}'.")

    # 2. Set Column Widths (Columns A to I)
    column_letters = ["A", "B", "C", "D", "E", "F", "G", "H", "I"]
    for col in column_letters:
        ws.column_dimensions[col].width = 10

    # 3. Populate Headers (Row 1)
    headers = [
        "Date",
        "Job / FX#",
        "Start",
        "Time",
        "T1.5",
        "CSR#",
        "Comment_1",
        "Comment_2",
        d_week_end,
    ]
    for col_idx, header_value in enumerate(headers, start=1):
        ws.cell(row=1, column=col_idx, value=header_value)

    # 4. Apply Header Fonts (Bold)
    bold_font = Font(bold=True)
    for col_idx in range(1, 10):
        ws.cell(row=1, column=col_idx).font = bold_font

    # 5. Apply Number Formatting and Alignment
    center_alignment = Alignment(horizontal="center")
    for row in range(2, 101):
        ws.cell(row=row, column=3).number_format = "hh:mm"

        # Columns D & E: Number format and Center Alignment
        for col_idx in range(4, 5):
            cell = ws.cell(row=row, column=col_idx)
            cell.number_format = "0.00"
            cell.alignment = center_alignment

    # 6. Add Formulas and Basic Hours
    ws["A2"] = "=I1-6"
    for row in range(2, 10):
        ws[f"D{row}"] = 1
        ws[f"E{row}"] = 0

    ws["D10"] = "=SUM(D2:D9)"
    ws["E10"] = "=SUM(E2:E9)"

    # 7. Set Window Zoom to 80%
    ws.sheet_view.zoomScale = 80

    # 8. PREPEND: Move the newly created sheet to the very first position (index 0)
    # openpyxl keeps an internal list of sheets layout called _sheets
    sheet_index = wb._sheets.index(ws)
    wb._sheets.insert(0, wb._sheets.pop(sheet_index))

    # 9. Clean up default sheet if this was a first run creation
    if "Sheet" in wb.sheetnames and len(wb.sheetnames) > 1:
        wb.remove(wb["Sheet"])

    # Save the updated workbook
    wb.save(file_name)
    print(f"Successfully prepended sheet '{sheet_title}' to '{file_name}'.")


# Execute the script
if __name__ == "__main__":
    prepend_time_sheet()
