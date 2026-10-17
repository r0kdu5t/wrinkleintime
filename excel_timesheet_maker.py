from datetime import date, timedelta
import openpyxl
from openpyxl.styles import Alignment, Font


def find_sunday(input_date: date) -> str:
    """Returns the date of the first Sunday following the input_date as a string (e.g., '18Oct2026')."""
    # Excel's Weekday(date) defaults to 1 for Sunday, 2 for Monday, etc.
    # Python's isoweekday() returns 1 for Monday through 7 for Sunday.
    # Map Python's weekday to replicate the VBA logic: InputDate + (8 - Weekday(InputDate))
    vba_weekday = (
        input_date.isoweekday() + 1 if input_date.isoweekday() < 7 else 1
    )
    days_to_add = 8 - vba_weekday
    sunday_date = input_date + timedelta(days=days_to_add)

    # Format matching VBA's "dmmmyyyy" (e.g., 25Oct2026)
    return sunday_date.strftime("%d%b%Y")


def do_time_sheet():
    """Creates an Excel Worksheet For My Weekly Timesheet."""
    # Create a new workbook and select active sheet (equivalent to GetExcelWS)
    wb = openpyxl.Workbook()
    ws = wb.active

    # 1. Get the upcoming Sunday's date string
    d_week_end = find_sunday(date.today())

    # 2. Set Column Widths (Columns A to I)
    column_letters = ["A", "B", "C", "D", "E", "F", "G", "H", "I"]
    for col in column_letters:
        ws.column_dimensions[col].width = 10

    # 3. Name the Sheet
    ws.title = f"WE{d_week_end}"

    # 4. Populate Headers (Row 1)
    headers = [
        "Date",
        "Job / FX#",
        "Start",
        "Time",
        "T1.5",
        "CSR#",
        "Comment_1",
        "Comment_2",
        d_week_end,  # Cell I1
    ]

    for col_idx, header_value in enumerate(headers, start=1):
        ws.cell(row=1, column=col_idx, value=header_value)

    # 5. Apply Header Fonts (Bold)
    bold_font = Font(bold=True)
    for col_idx in range(1, 10):
        ws.cell(row=1, column=col_idx).font = bold_font

    # 6. Apply Number Formatting and Alignment to entire columns dynamically
    center_alignment = Alignment(horizontal="center")

    # While openpyxl doesn't format an "infinite" column easily, we can format
    # the maximum potential rows used (e.g., rows 2 to 100)
    for row in range(2, 101):
        # Column C: Time format
        ws.cell(row=row, column=3).number_format = "hh:mm"

        # Columns D & E: Number format and Center Alignment
        for col_idx in:
            cell = ws.cell(row=row, column=col_idx)
            cell.number_format = "0.00"
            cell.alignment = center_alignment

    # 7. Add Formulas and Hardcoded Values
    ws["A2"] = "=I1-6"  # Formulates the start of the week date

    # Fill rows 2 to 9 for Ordinary (D) and Overtime (E) hours
    for row in range(2, 10):
        ws[f"D{row}"] = 1
        ws[f"E{row}"] = 0

    # Total Sum Formulas in Row 10
    ws["D10"] = "=SUM(D2:D9)"
    ws["E10"] = "=SUM(E2:E9)"

    # 8. Set Window Zoom to 80%
    ws.sheet_view.zoomScale = 80

    # Save the file
    file_name = f"Timesheet_WE_{d_week_end}.xlsx"
    wb.save(file_name)
    print(f"Timesheet successfully created and saved as '{file_name}'")


# Execute the script
if __name__ == "__main__":
    do_time_sheet()
