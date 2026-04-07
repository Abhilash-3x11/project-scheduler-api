from __future__ import annotations

import json
import re
from collections import defaultdict
from datetime import date, datetime, timedelta
from io import BytesIO
from typing import Dict, List, Optional, Set, Tuple

from openpyxl import load_workbook
from openpyxl.workbook.workbook import Workbook


PERFORMANCE_SHEET = "Performance-Plan"
AVAILABILITY_SHEET = "Availability"


# -----------------------------
# Basic helpers
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
    # Accept common date string formats
    for fmt in ("%Y-%m-%d", "%d-%m-%Y", "%m/%d/%Y", "%d/%m/%Y"):
        try:
            return datetime.strptime(text, fmt).date()
        except ValueError:
            pass
    try:
        return datetime.fromisoformat(text).date()
    except ValueError:
        return None


def is_working_day(d: date, holidays: Set[date]) -> bool:
    return d.weekday() < 5 and d not in holidays


def next_working_day(d: date, holidays: Set[date]) -> date:
    current = d
    while not is_working_day(current, holidays):
        current += timedelta(days=1)
    return current


def excel_workday(start_date: date, days: int, holidays: Set[date]) -> date:
    """
    Excel WORKDAY-style logic:
    - days = 0 -> return the start date if it's a working day, otherwise the next working day
    - days > 0 -> count forward, excluding the start date
    - days < 0 -> count backward, excluding the start date
    """
    if days == 0:
        return next_working_day(start_date, holidays)

    step = 1 if days > 0 else -1
    remaining = abs(days)
    current = start_date

    while remaining > 0:
        current += timedelta(days=step)
        if is_working_day(current, holidays):
            remaining -= 1

    return current


def parse_holiday_entry(value: object, base_year: int) -> Set[date]:
    """
    Supports:
    - 2026-01-26
    - Nov 17-21
    - Nov 17-21, 2026  (also accepted)
    """
    result: Set[date] = set()

    if value is None or value == "":
        return result

    if isinstance(value, (datetime, date)):
        result.add(parse_date_value(value))
        return result

    text = str(value).strip()
    if not text:
        return result

    # Full date string
    parsed = parse_date_value(text)
    if parsed:
        result.add(parsed)
        return result

    # Range like "Nov 17-21"
    m = re.fullmatch(
        r"([A-Za-z]{3,9})\s+(\d{1,2})-(\d{1,2})(?:,\s*(\d{4}))?",
        text,
    )
    if m:
        month_text = m.group(1)
        start_day = int(m.group(2))
        end_day = int(m.group(3))
        year = int(m.group(4)) if m.group(4) else base_year

        month_lookup = {
            "jan": 1, "january": 1,
            "feb": 2, "february": 2,
            "mar": 3, "march": 3,
            "apr": 4, "april": 4,
            "may": 5,
            "jun": 6, "june": 6,
            "jul": 7, "july": 7,
            "aug": 8, "august": 8,
            "sep": 9, "sept": 9, "september": 9,
            "oct": 10, "october": 10,
            "nov": 11, "november": 11,
            "dec": 12, "december": 12,
        }
        month = month_lookup[month_text.lower()]

        for day in range(start_day, end_day + 1):
            result.add(date(year, month, day))
        return result

    raise ValueError(f"Unsupported holiday format: {value!r}")


def parse_holidays(values: List[object], base_year: int) -> Set[date]:
    holidays: Set[date] = set()
    for item in values or []:
        holidays |= parse_holiday_entry(item, base_year)
    return holidays


# -----------------------------
# Availability sheet parsing
# -----------------------------
def parse_availability_sheet(ws) -> List[dict]:
    """
    Returns blocks like:
    [
      {
        "title": "Cornerstone",
        "roles": {"Project Manager": ["Hari Rs"], ...},
        "holidays": ["Nov 17-21", ...]
      },
      ...
    ]
    """
    header_rows: List[int] = []
    for r in range(1, ws.max_row + 1):
        if normalize_text(ws.cell(r, 1).value) == "role":
            header_rows.append(r)

    blocks = []

    for idx, header_row in enumerate(header_rows):
        title = str(ws.cell(header_row, 2).value).strip() if ws.cell(header_row, 2).value else ""
        start_row = header_row + 1
        end_row = (header_rows[idx + 1] - 1) if idx + 1 < len(header_rows) else ws.max_row

        role_to_owners: Dict[str, List[str]] = defaultdict(list)
        holidays: List[object] = []

        for r in range(start_row, end_row + 1):
            role_val = ws.cell(r, 1).value
            owner_val = ws.cell(r, 2).value
            holiday_val = ws.cell(r, 3).value

            if role_val is None and owner_val is None and holiday_val is None:
                continue

            if role_val is not None and owner_val is not None:
                role_key = str(role_val).strip()
                owner_text = str(owner_val).strip()
                if role_key and owner_text:
                    if owner_text not in role_to_owners[role_key]:
                        role_to_owners[role_key].append(owner_text)

            if holiday_val not in (None, ""):
                holidays.append(holiday_val)

        blocks.append(
            {
                "title": title,
                "roles": dict(role_to_owners),
                "holidays": holidays,
            }
        )

    return blocks


def build_role_owner_lookup(
    availability_blocks: List[dict],
    preferred_block_title: Optional[str] = None,
) -> Tuple[Dict[str, List[str]], Set[date]]:
    """
    Returns:
      - role lookup dict (normalized role -> list of owners)
      - holiday set from availability sheet
    If preferred_block_title is given and matches a block title, that block's
    holidays are used first; otherwise all blocks' holidays are used.
    """
    lookup: Dict[str, List[str]] = defaultdict(list)
    collected_holidays: Set[date] = set()

    preferred_norm = normalize_text(preferred_block_title) if preferred_block_title else ""

    # Prefer the selected block first, then the rest
    sorted_blocks = sorted(
        availability_blocks,
        key=lambda b: 0 if normalize_text(b["title"]) == preferred_norm else 1,
    )

    for block in sorted_blocks:
        for role, owners in block["roles"].items():
            key = normalize_text(role)
            for owner in owners:
                if owner not in lookup[key]:
                    lookup[key].append(owner)

        # Holidays from preferred block first, but all blocks are still allowed
        if preferred_norm:
            if normalize_text(block["title"]) == preferred_norm:
                collected_holidays |= parse_holidays(block["holidays"], base_year=date.today().year)
        else:
            collected_holidays |= parse_holidays(block["holidays"], base_year=date.today().year)

    # If preferred block was requested but not found, fall back to all blocks
    if preferred_norm and not collected_holidays:
        for block in availability_blocks:
            collected_holidays |= parse_holidays(block["holidays"], base_year=date.today().year)

    return lookup, collected_holidays


def resolve_owner_for_role(
    role_text: object,
    role_owner_lookup: Dict[str, List[str]],
    existing_owner: object = None,
) -> str:
    """
    - Exact role match first
    - If role text contains multiple roles (e.g. "Customer + Business Consultant"),
      try each token.
    - If no mapping is found, keep the existing owner.
    """
    if role_text is None or str(role_text).strip() == "":
        return str(existing_owner).strip() if existing_owner not in (None, "") else ""

    raw = str(role_text).strip()
    exact_key = normalize_text(raw)

    if exact_key in role_owner_lookup:
        owners = role_owner_lookup[exact_key]
        if len(owners) == 1:
            return owners[0]
        return " / ".join(owners)

    # Split combined roles
    parts = [p.strip() for p in re.split(r"\s*(?:\+|/|,|&| and )\s*", raw) if p.strip()]
    resolved_parts: List[str] = []

    for part in parts:
        key = normalize_text(part)
        if key in role_owner_lookup:
            owners = role_owner_lookup[key]
            if owners:
                resolved_parts.append(owners[0])

    # If everything could be resolved, use the mapped names
    if resolved_parts and len(resolved_parts) == len(parts):
        deduped = []
        for owner in resolved_parts:
            if owner not in deduped:
                deduped.append(owner)
        return " / ".join(deduped)

    # Otherwise, preserve the existing owner if present
    if existing_owner not in (None, ""):
        return str(existing_owner).strip()

    # Partial mapping fallback
    if resolved_parts:
        deduped = []
        for owner in resolved_parts:
            if owner not in deduped:
                deduped.append(owner)
        return " / ".join(deduped)

    return ""


# -----------------------------
# Formula-based schedule builder
# -----------------------------
def build_schedule_dates(project_start_date: date, holidays: Set[date]) -> Tuple[Dict[int, date], Dict[int, date], Dict[int, int]]:
    """
    Recreates the workbook formula chain for rows 2..33.
    Returns:
      planned_start[row], planned_end[row], duration_override[row]
    """
    ps: Dict[int, date] = {}
    pe: Dict[int, date] = {}
    dur_override: Dict[int, int] = {}

    # Row 2 summary start, end is linked later
    ps[2] = next_working_day(project_start_date, holidays)

    # Row 3
    ps[3] = ps[2]
    pe[3] = ps[3]  # =D3+0

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

    # Row 15 summary block
    ps[15] = ps[3] + timedelta(days=8)     # =D3+8
    # end is =E22-1, calculated after row 22

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

    # Summary rows
    pe[15] = pe[22] - timedelta(days=1)   # =E22-1
    dur_override[15] = (pe[15] - ps[15]).days

    pe[2] = pe[33]                        # =E33

    return ps, pe, dur_override


# -----------------------------
# Workbook update
# -----------------------------
def update_project_schedule(
    input_xlsx_path: str,
    output_xlsx_path: str,
    project_start_date: str,
    holidays: Optional[List[object]] = None,
    role: str = "",
    sheet_name: str = PERFORMANCE_SHEET,
) -> str:
    wb = load_workbook(input_xlsx_path, data_only=False)

    if sheet_name not in wb.sheetnames:
        raise ValueError(f"Sheet '{sheet_name}' not found.")

    if AVAILABILITY_SHEET not in wb.sheetnames:
        raise ValueError(f"Sheet '{AVAILABILITY_SHEET}' not found.")

    ws = wb[sheet_name]
    ws_avail = wb[AVAILABILITY_SHEET]

    # Read headers
    headers = {}
    for cell in ws[1]:
        if cell.value is not None:
            headers[normalize_text(cell.value)] = cell.column

    required = {
        "task": "Task",
        "duration": "Duration",
        "planned start date": "Planned Start Date",
        "planned end date": "Planned End Date",
        "owner": "Owner",
        "role": "Role",
    }

    missing = [label for key, label in required.items() if key not in headers]
    if missing:
        raise ValueError(f"Missing required columns in '{sheet_name}': {', '.join(missing)}")

    task_col = headers["task"]
    duration_col = headers["duration"]
    start_col = headers["planned start date"]
    end_col = headers["planned end date"]
    owner_col = headers["owner"]
    role_col = headers["role"]

    # Availability parsing
    availability_blocks = parse_availability_sheet(ws_avail)
    preferred_block_title = None
    if role:
        # If the input role matches a block title (Cornerstone / Munchkin), prefer that block.
        role_norm = normalize_text(role)
        for block in availability_blocks:
            if normalize_text(block["title"]) == role_norm:
                preferred_block_title = block["title"]
                break

    role_owner_lookup, availability_holidays = build_role_owner_lookup(
        availability_blocks,
        preferred_block_title=preferred_block_title,
    )

    project_start = parse_date_value(project_start_date)
    if project_start is None:
        raise ValueError("Invalid project_start_date. Use YYYY-MM-DD (for example, 2026-01-01).")

    # Combine holidays: user holidays + availability holidays
    user_holiday_set = parse_holidays(holidays or [], base_year=project_start.year)
    all_holidays = user_holiday_set | availability_holidays

    # Compute dates according to workbook formulas
    planned_start, planned_end, duration_overrides = build_schedule_dates(project_start, all_holidays)

    # Write computed dates and owners
    for row in range(2, ws.max_row + 1):
        if row not in planned_start or row not in planned_end:
            continue

        # Planned dates
        ws.cell(row=row, column=start_col).value = planned_start[row]
        ws.cell(row=row, column=end_col).value = planned_end[row]

        # Duration override only for summary row 15
        if row in duration_overrides:
            ws.cell(row=row, column=duration_col).value = duration_overrides[row]

        # Role -> Owner mapping
        row_role = ws.cell(row=row, column=role_col).value
        current_owner = ws.cell(row=row, column=owner_col).value
        mapped_owner = resolve_owner_for_role(row_role, role_owner_lookup, existing_owner=current_owner)

        if mapped_owner:
            ws.cell(row=row, column=owner_col).value = mapped_owner

    wb.save(output_xlsx_path)
    return output_xlsx_path


def update_project_schedule_stream(
    file_bytes: bytes,
    project_start_date: str,
    holidays: Optional[List[object]] = None,
    role: str = "",
    sheet_name: str = PERFORMANCE_SHEET,
) -> BytesIO:
    wb = load_workbook(BytesIO(file_bytes), data_only=False)

    if sheet_name not in wb.sheetnames:
        raise ValueError(f"Sheet '{sheet_name}' not found.")

    if AVAILABILITY_SHEET not in wb.sheetnames:
        raise ValueError(f"Sheet '{AVAILABILITY_SHEET}' not found.")

    ws = wb[sheet_name]
    ws_avail = wb[AVAILABILITY_SHEET]

    headers = {}
    for cell in ws[1]:
        if cell.value is not None:
            headers[normalize_text(cell.value)] = cell.column

    required = {
        "task": "Task",
        "duration": "Duration",
        "planned start date": "Planned Start Date",
        "planned end date": "Planned End Date",
        "owner": "Owner",
        "role": "Role",
    }

    missing = [label for key, label in required.items() if key not in headers]
    if missing:
        raise ValueError(f"Missing required columns in '{sheet_name}': {', '.join(missing)}")

    task_col = headers["task"]
    duration_col = headers["duration"]
    start_col = headers["planned start date"]
    end_col = headers["planned end date"]
    owner_col = headers["owner"]
    role_col = headers["role"]

    availability_blocks = parse_availability_sheet(ws_avail)
    preferred_block_title = None
    if role:
        role_norm = normalize_text(role)
        for block in availability_blocks:
            if normalize_text(block["title"]) == role_norm:
                preferred_block_title = block["title"]
                break

    role_owner_lookup, availability_holidays = build_role_owner_lookup(
        availability_blocks,
        preferred_block_title=preferred_block_title,
    )

    project_start = parse_date_value(project_start_date)
    if project_start is None:
        raise ValueError("Invalid project_start_date. Use YYYY-MM-DD (for example, 2026-01-01).")

    user_holiday_set = parse_holidays(holidays or [], base_year=project_start.year)
    all_holidays = user_holiday_set | availability_holidays

    planned_start, planned_end, duration_overrides = build_schedule_dates(project_start, all_holidays)

    for row in range(2, ws.max_row + 1):
        if row not in planned_start or row not in planned_end:
            continue

        ws.cell(row=row, column=start_col).value = planned_start[row]
        ws.cell(row=row, column=end_col).value = planned_end[row]

        if row in duration_overrides:
            ws.cell(row=row, column=duration_col).value = duration_overrides[row]

        row_role = ws.cell(row=row, column=role_col).value
        current_owner = ws.cell(row=row, column=owner_col).value
        mapped_owner = resolve_owner_for_role(row_role, role_owner_lookup, existing_owner=current_owner)

        if mapped_owner:
            ws.cell(row=row, column=owner_col).value = mapped_owner

    output = BytesIO()
    wb.save(output)
    output.seek(0)
    return output