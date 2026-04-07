from datetime import datetime, date, timedelta
from openpyxl import load_workbook
from io import BytesIO


# -----------------------------
# Helpers
# -----------------------------
def normalize_holidays(holidays):
    return {
        datetime.strptime(h, "%Y-%m-%d").date()
        if isinstance(h, str) else h
        for h in holidays or []
    }


def is_working_day(d, holidays):
    return d.weekday() < 5 and d not in holidays


def get_next_working_day(d, holidays):
    current = d
    while not is_working_day(current, holidays):
        current += timedelta(days=1)
    return current


def add_business_days(start_date, duration, holidays):
    """
    Correct logic:
    - Start date is Day 1
    """

    duration = int(duration)

    if duration <= 1:
        return start_date

    current = start_date
    days_counted = 1

    while days_counted < duration:
        current += timedelta(days=1)
        if is_working_day(current, holidays):
            days_counted += 1

    return current


# -----------------------------
# CORE FUNCTION (File Path)
# -----------------------------
def update_project_schedule(
    input_xlsx_path,
    output_xlsx_path,
    project_start_date,
    sheet_name="Performance-Plan",
    holidays=None,
):
    wb = load_workbook(input_xlsx_path, data_only=True)
    ws = wb[sheet_name]

    headers = {
        str(cell.value).strip(): cell.column
        for cell in ws[1]
        if cell.value
    }

    start_col = headers["Planned Start Date"]
    end_col = headers["Planned End Date"]
    duration_col = headers["Duration"]

    holiday_set = normalize_holidays(holidays)

    current_start = datetime.strptime(project_start_date, "%Y-%m-%d").date()
    current_start = get_next_working_day(current_start, holiday_set)

    for row in range(2, ws.max_row + 1):
        duration_val = ws.cell(row=row, column=duration_col).value

        if duration_val in (None, "", 0, "0"):
            ws.cell(row=row, column=start_col).value = None
            ws.cell(row=row, column=end_col).value = None
            continue

        start_date = current_start
        end_date = add_business_days(start_date, duration_val, holiday_set)

        ws.cell(row=row, column=start_col).value = start_date
        ws.cell(row=row, column=end_col).value = end_date

        current_start = get_next_working_day(end_date + timedelta(days=1), holiday_set)

    wb.save(output_xlsx_path)
    return output_xlsx_path


# -----------------------------
# CORE FUNCTION (API STREAM)
# -----------------------------
def update_project_schedule_stream(file_bytes, project_start_date, holidays):
    wb = load_workbook(BytesIO(file_bytes), data_only=True)
    ws = wb["Performance-Plan"]

    headers = {
        str(cell.value).strip(): cell.column
        for cell in ws[1]
        if cell.value
    }

    start_col = headers["Planned Start Date"]
    end_col = headers["Planned End Date"]
    duration_col = headers["Duration"]

    holiday_set = normalize_holidays(holidays)

    current_start = datetime.strptime(project_start_date, "%Y-%m-%d").date()
    current_start = get_next_working_day(current_start, holiday_set)

    for row in range(2, ws.max_row + 1):
        duration_val = ws.cell(row=row, column=duration_col).value

        if duration_val in (None, "", 0, "0"):
            ws.cell(row=row, column=start_col).value = None
            ws.cell(row=row, column=end_col).value = None
            continue

        start_date = current_start
        end_date = add_business_days(start_date, duration_val, holiday_set)

        ws.cell(row=row, column=start_col).value = start_date
        ws.cell(row=row, column=end_col).value = end_date

        current_start = get_next_working_day(end_date + timedelta(days=1), holiday_set)

    output = BytesIO()
    wb.save(output)
    output.seek(0)

    return output