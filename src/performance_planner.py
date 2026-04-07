from __future__ import annotations

import re
from datetime import date, datetime, timedelta
from io import BytesIO
from pathlib import Path
from typing import Dict, List, Optional, Set, Tuple

from openpyxl import load_workbook

PERFORMANCE_SHEET = "Performance-Plan"


# -----------------------------
# Helpers
# -----------------------------
def normalize_text(value: object) -> str:
    return re.sub(r"\s+", " ", str(value).strip()).lower() if value is not None else ""


def parse_date_value(value: object) -> Optional[date]:
    if value is None or value == "":
        return None
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value

    text = str(value).strip()
    if not text:
        return None

    for fmt in ("%Y-%m-%d", "%d-%m-%Y", "%m/%d/%Y", "%d/%m/%Y"):
        try:
            return datetime.strptime(text, fmt).date()
        except ValueError:
            pass

    try:
        return datetime.fromisoformat(text).date()
    except ValueError:
        return None


def parse_holidays(values: List[object], base_year: int) -> Set[date]:
    holidays: Set[date] = set()
    for value in values or []:
        if value is None or value == "":
            continue
        if isinstance(value, datetime):
            holidays.add(value.date())
        elif isinstance(value, date):
            holidays.add(value)
        else:
            holidays.add(datetime.strptime(str(value).strip(), "%Y-%m-%d").date())
    return holidays


def is_working_day(d: date, holidays: Set[date]) -> bool:
    return d.weekday() < 5 and d not in holidays


def next_working_day(d: date, holidays: Set[date]) -> date:
    while not is_working_day(d, holidays):
        d += timedelta(days=1)
    return d


def excel_workday(start_date: date, days: int, holidays: Set[date]) -> date:
    """
    Excel-style WORKDAY behavior:
    - days = 0 => return start_date if it is a working day, else next working day
    - days > 0 => move forward by that many working days, excluding start_date
    - days < 0 => move backward by that many working days, excluding start_date
    """
    if days == 0:
        return next_working_day(start_date, holidays)

    current = start_date
    remaining = abs(days)
    step = 1 if days > 0 else -1

    while remaining > 0:
        current += timedelta(days=step)
        if is_working_day(current, holidays):
            remaining -= 1

    return current


# -----------------------------
# Formula-based schedule builder
# -----------------------------
def build_schedule_dates(
    project_start_date: date,
    holidays: Set[date],
) -> Tuple[Dict[int, date], Dict[int, date], Dict[int, int]]:
    """
    Recreates the workbook's planned start/end formula chain for rows 2..33.
    Returns:
      planned_start[row], planned_end[row], duration_override[row]
    """
    ps: Dict[int, date] = {}
    pe: Dict[int, date] = {}
    dur_override: Dict[int, int] = {}

    # Row 2
    ps[2] = next_working_day(project_start_date, holidays)

    # Row 3
    ps[3] = ps[2]
    pe[3] = ps[3]

    # Row 4
    ps[4] = excel_workday(pe[3], 1, holidays)
    pe[4] = ps[4]

    # Row 5
    ps[5] = excel_workday(pe[3], 1, holidays)
    pe[5] = excel_workday(ps[5], 15, holidays)

    # Row 6
    ps[6] = excel_workday(pe[3], 5, holidays)
    pe[6] = excel_workday(ps[6], 0, holidays)

    # Row 7
    ps[7] = excel_workday(pe[6], 2, holidays)
    pe[7] = ps[7]

    # Row 8
    ps[8] = excel_workday(pe[6], 3, holidays)
    pe[8] = ps[8]

    # Row 9
    ps[9] = excel_workday(pe[3], 5, holidays)
    pe[9] = ps[9]

    # Row 10
    ps[10] = excel_workday(pe[7], 2, holidays)
    pe[10] = excel_workday(ps[10], 0, holidays)

    # Row 11
    ps[11] = excel_workday(pe[7], 2, holidays)
    pe[11] = excel_workday(ps[11], 5, holidays)

    # Row 12
    ps[12] = pe[7]
    pe[12] = excel_workday(ps[12], 2, holidays)

    # Row 13
    ps[13] = pe[12]
    pe[13] = excel_workday(ps[13], 2, holidays)

    # Row 14
    ps[14] = pe[13]
    pe[14] = ps[14]

    # Row 15
    ps[15] = ps[3] + timedelta(days=8)

    # Row 16
    ps[16] = excel_workday(ps[15], 5, holidays)
    pe[16] = excel_workday(ps[16], 2, holidays)

    # Row 17
    ps[17] = pe[16]
    pe[17] = excel_workday(ps[17], 2, holidays)

    # Row 18
    ps[18] = pe[17]
    pe[18] = excel_workday(ps[18], 2, holidays)

    # Row 19
    ps[19] = excel_workday(pe[14], 3, holidays)
    pe[19] = ps[19]

    # Row 20
    ps[20] = excel_workday(pe[19], 2, holidays)
    pe[20] = excel_workday(ps[20], 0, holidays)

    # Row 21
    ps[21] = excel_workday(pe[20], 3, holidays)
    pe[21] = ps[21]

    # Row 22
    ps[22] = excel_workday(pe[21], 2, holidays)
    pe[22] = excel_workday(ps[22], 0, holidays)

    # Row 23
    ps[23] = excel_workday(pe[22], 3, holidays)
    pe[23] = ps[23]

    # Row 24
    ps[24] = excel_workday(pe[23], 0, holidays)
    pe[24] = ps[24]

    # Row 25
    ps[25] = excel_workday(pe[24], 3, holidays)
    pe[25] = excel_workday(ps[25], 0, holidays)

    # Row 26
    ps[26] = excel_workday(pe[25], 4, holidays)
    pe[26] = excel_workday(ps[26], 0, holidays)

    # Row 27
    ps[27] = excel_workday(pe[26], 2, holidays)
    pe[27] = ps[27]

    # Row 28
    ps[28] = pe[25]
    pe[28] = excel_workday(ps[28], 2, holidays)

    # Row 29
    ps[29] = pe[28]
    pe[29] = pe[27]

    # Row 30
    ps[30] = pe[29]
    pe[30] = excel_workday(ps[30], 2, holidays)

    # Row 31
    ps[31] = pe[27]
    pe[31] = excel_workday(ps[31], 3, holidays)

    # Row 32
    ps[32] = excel_workday(pe[31], 1, holidays)
    pe[32] = excel_workday(ps[32], 2, holidays)

    # Row 33
    ps[33] = excel_workday(pe[32], 0, holidays)
    pe[33] = excel_workday(ps[33], 2, holidays)

    # Summary row overrides
    pe[15] = pe[22] - timedelta(days=1)
    dur_override[15] = (pe[15] - ps[15]).days
    pe[2] = pe[33]

    return ps, pe, dur_override


# -----------------------------
# Role -> Owner mapping from same workbook
# -----------------------------
def build_role_owner_map(ws, role_col: int, owner_col: int) -> Dict[str, str]:
    role_owner_map: Dict[str, str] = {}
    for r in range(2, ws.max_row + 1):
        role_val = ws.cell(r, role_col).value
        owner_val = ws.cell(r, owner_col).value

        if role_val and owner_val:
            key = normalize_text(role_val)
            if key not in role_owner_map:
                role_owner_map[key] = str(owner_val).strip()
    return role_owner_map


def apply_owner_mapping(
    ws,
    role_owner_map: Dict[str, str],
    role_col: int,
    owner_col: int,
    selected_role: str = "",
):
    selected_role_key = normalize_text(selected_role)

    for r in range(2, ws.max_row + 1):
        row_role = ws.cell(r, role_col).value
        current_owner = ws.cell(r, owner_col).value

        if row_role is None:
            continue

        row_role_key = normalize_text(row_role)

        # If role input is provided, only update rows matching that role.
        if selected_role_key and row_role_key != selected_role_key:
            continue

        # Only fill if owner is empty
        if current_owner is None or str(current_owner).strip() == "":
            if row_role_key in role_owner_map:
                ws.cell(r, owner_col).value = role_owner_map[row_role_key]


# -----------------------------
# Local file update
# -----------------------------
def update_project_schedule(
    project_start_date: str,
    holidays: Optional[List[object]] = None,
    role: str = "",
    template_path: Optional[str] = None,
    output_path: Optional[str] = None,
) -> str:
    """
    Reads the template workbook from project_root/data/Yes Prep Performance_Project Plan (Formulated).xlsx by default,
    updates it, and writes to output_path.
    """
    if template_path is None:
        template_path = str(Path(__file__).resolve().parent.parent / "data" / "Yes Prep Performance_Project Plan (Formulated).xlsx")

    if output_path is None:
        output_path = str(Path(__file__).resolve().parent.parent / "data" / "output.xlsx")

    wb = load_workbook(template_path)
    if PERFORMANCE_SHEET not in wb.sheetnames:
        raise ValueError(f"Sheet '{PERFORMANCE_SHEET}' not found.")

    ws = wb[PERFORMANCE_SHEET]

    headers = {normalize_text(c.value): c.column for c in ws[1] if c.value}

    required = {
        "planned start date": "Planned Start Date",
        "planned end date": "Planned End Date",
        "owner": "Owner",
        "role": "Role",
    }
    missing = [label for key, label in required.items() if key not in headers]
    if missing:
        raise ValueError(f"Missing required columns in '{PERFORMANCE_SHEET}': {', '.join(missing)}")

    start_col = headers["planned start date"]
    end_col = headers["planned end date"]
    duration_col = headers.get("duration")
    owner_col = headers["owner"]
    role_col = headers["role"]

    project_start = parse_date_value(project_start_date)
    if project_start is None:
        raise ValueError("Invalid project_start_date. Use YYYY-MM-DD, for example 2026-01-01.")

    holiday_set = parse_holidays(holidays or [], project_start.year)

    planned_start, planned_end, duration_overrides = build_schedule_dates(project_start, holiday_set)
    role_owner_map = build_role_owner_map(ws, role_col, owner_col)

    for row in range(2, ws.max_row + 1):
        if row not in planned_start or row not in planned_end:
            continue

        ws.cell(row=row, column=start_col).value = planned_start[row]
        ws.cell(row=row, column=end_col).value = planned_end[row]

        if duration_col and row in duration_overrides:
            ws.cell(row=row, column=duration_col).value = duration_overrides[row]

    apply_owner_mapping(ws, role_owner_map, role_col, owner_col, selected_role=role)

    wb.save(output_path)
    return output_path


# -----------------------------
# API/stream version
# -----------------------------
def update_project_schedule_stream(
    project_start_date: str,
    holidays: Optional[List[object]] = None,
    role: str = "",
    template_path: Optional[str] = None,
) -> BytesIO:
    if template_path is None:
        template_path = str(Path(__file__).resolve().parent.parent / "data" / "Yes Prep Performance_Project Plan (Formulated).xlsx")

    wb = load_workbook(template_path)
    if PERFORMANCE_SHEET not in wb.sheetnames:
        raise ValueError(f"Sheet '{PERFORMANCE_SHEET}' not found.")

    ws = wb[PERFORMANCE_SHEET]

    headers = {normalize_text(c.value): c.column for c in ws[1] if c.value}

    required = {
        "planned start date": "Planned Start Date",
        "planned end date": "Planned End Date",
        "owner": "Owner",
        "role": "Role",
    }
    missing = [label for key, label in required.items() if key not in headers]
    if missing:
        raise ValueError(f"Missing required columns in '{PERFORMANCE_SHEET}': {', '.join(missing)}")

    start_col = headers["planned start date"]
    end_col = headers["planned end date"]
    duration_col = headers.get("duration")
    owner_col = headers["owner"]
    role_col = headers["role"]

    project_start = parse_date_value(project_start_date)
    if project_start is None:
        raise ValueError("Invalid project_start_date. Use YYYY-MM-DD, for example 2026-01-01.")

    holiday_set = parse_holidays(holidays or [], project_start.year)

    planned_start, planned_end, duration_overrides = build_schedule_dates(project_start, holiday_set)
    role_owner_map = build_role_owner_map(ws, role_col, owner_col)

    for row in range(2, ws.max_row + 1):
        if row not in planned_start or row not in planned_end:
            continue

        ws.cell(row=row, column=start_col).value = planned_start[row]
        ws.cell(row=row, column=end_col).value = planned_end[row]

        if duration_col and row in duration_overrides:
            ws.cell(row=row, column=duration_col).value = duration_overrides[row]

    apply_owner_mapping(ws, role_owner_map, role_col, owner_col, selected_role=role)

    output = BytesIO()
    wb.save(output)
    output.seek(0)
    return output