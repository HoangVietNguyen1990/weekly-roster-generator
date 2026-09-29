import pandas as pd
from datetime import datetime, date

# Victoria, Australia Official Public Holidays (2026 & 2027)
VIC_PUBLIC_HOLIDAYS = {
    # 2026
    date(2026, 1, 1): "New Year's Day",
    date(2026, 1, 26): "Australia Day",
    date(2026, 3, 9): "Labour Day",
    date(2026, 4, 3): "Good Friday",
    date(2026, 4, 4): "Easter Saturday",
    date(2026, 4, 5): "Easter Sunday",
    date(2026, 4, 6): "Easter Monday",
    date(2026, 4, 25): "ANZAC Day",
    date(2026, 6, 8): "King's Birthday",
    date(2026, 9, 25): "Friday before AFL Grand Final",
    date(2026, 11, 3): "Melbourne Cup Day",
    date(2026, 12, 25): "Christmas Day",
    date(2026, 12, 28): "Boxing Day (Observed)",

    # 2027
    date(2027, 1, 1): "New Year's Day",
    date(2027, 1, 26): "Australia Day",
    date(2027, 3, 8): "Labour Day",
    date(2027, 3, 26): "Good Friday",
    date(2027, 3, 27): "Easter Saturday",
    date(2027, 3, 28): "Easter Sunday",
    date(2027, 3, 29): "Easter Monday",
    date(2027, 4, 25): "ANZAC Day",
    date(2027, 6, 14): "King's Birthday",
    date(2027, 9, 24): "Friday before AFL Grand Final (TBA)",
    date(2027, 11, 2): "Melbourne Cup Day",
    date(2027, 12, 27): "Christmas Day (Observed)",
    date(2027, 12, 28): "Boxing Day (Observed)",
}

# Victoria Government School Holiday Periods (2026 & 2027)
VIC_SCHOOL_HOLIDAYS = [
    # 2026
    {"name": "2026 Autumn School Holidays", "start": date(2026, 4, 3), "end": date(2026, 4, 19)},
    {"name": "2026 Winter School Holidays", "start": date(2026, 6, 27), "end": date(2026, 7, 12)},
    {"name": "2026 Spring School Holidays", "start": date(2026, 9, 19), "end": date(2026, 10, 4)},
    {"name": "2026-2027 Summer School Holidays", "start": date(2026, 12, 19), "end": date(2027, 1, 26)},

    # 2027
    {"name": "2027 Autumn School Holidays", "start": date(2027, 3, 26), "end": date(2027, 4, 11)},
    {"name": "2027 Winter School Holidays", "start": date(2027, 6, 26), "end": date(2027, 7, 11)},
    {"name": "2027 Spring School Holidays", "start": date(2027, 9, 18), "end": date(2027, 10, 3)},
    {"name": "2027-2028 Summer School Holidays", "start": date(2027, 12, 18), "end": date(2028, 1, 26)},
]

def to_date_obj(dt_val):
    if dt_val is None:
        return None
    if isinstance(dt_val, datetime):
        return dt_val.date()
    if isinstance(dt_val, date):
        return dt_val
    try:
        dt_p = pd.to_datetime(dt_val, dayfirst=True)
        if pd.notna(dt_p):
            return dt_p.date()
    except Exception:
        pass
    return None

def is_vic_public_holiday(dt_val):
    """
    Returns (True, holiday_name) if date is a Victorian public holiday, else (False, None).
    """
    d_obj = to_date_obj(dt_val)
    if not d_obj:
        return False, None
    if d_obj in VIC_PUBLIC_HOLIDAYS:
        return True, VIC_PUBLIC_HOLIDAYS[d_obj]
    return False, None

def is_vic_school_holiday(dt_val):
    """
    Returns (True, holiday_name) if date falls within a Victorian school holiday period, else (False, None).
    """
    d_obj = to_date_obj(dt_val)
    if not d_obj:
        return False, None
    for hol in VIC_SCHOOL_HOLIDAYS:
        if hol["start"] <= d_obj <= hol["end"]:
            return True, hol["name"]
    return False, None
