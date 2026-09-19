import sys
import os
import datetime
import pandas as pd
import pytest

# Ensure workspace root is in sys.path
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import app

# 1. TEST DATE & TIME PARSERS
def test_parse_date_robust():
    assert app.parse_date_robust("2026-08-10") == datetime.date(2026, 8, 10)
    assert app.parse_date_robust("10-08-2026") == datetime.date(2026, 8, 10)
    assert app.parse_date_robust("10/08/2026") == datetime.date(2026, 8, 10)
    assert app.parse_date_robust(datetime.datetime(2026, 8, 10)) == datetime.date(2026, 8, 10)
    assert app.parse_date_robust("") is None
    assert app.parse_date_robust("nan") is None

def test_melbourne_time_utilities():
    now = app.get_melbourne_now()
    today = app.get_melbourne_today()
    assert isinstance(now, (datetime.datetime, datetime.date))
    assert isinstance(today, datetime.date)

def test_time_and_shift_parsers():
    assert app.parse_time_to_decimal("06:30") == 6.5
    assert app.parse_time_to_decimal("14:15") == 14.25
    assert app.parse_time_to_decimal("invalid") == 0.0

    shift_tuple = app.parse_shift_range("06:00-14:30")
    assert shift_tuple is not None
    start_dec, end_dec, duration = shift_tuple
    assert start_dec == 6.0
    assert end_dec == 14.5
    assert duration == 8.5

    assert app.parse_shift_range("OFF") is None
    assert app.parse_shift_range("Unavailable") is None

# 2. TEST AWARD BREAK & WAGE CALCULATION ENGINE
def test_calculate_award_break():
    # < 4 hours -> 0 min break
    mins0, msg0 = app.calculate_award_break(3.5)
    assert mins0 == 0
    # < 5 hours -> 10 min break
    mins10, msg10 = app.calculate_award_break(4.5)
    assert mins10 == 10
    # < 7 hours -> 30 min break
    mins30, msg30 = app.calculate_award_break(6.0)
    assert mins30 == 30

def test_calculate_roster_wages():
    sample_roster = pd.DataFrame([
        {"NAME": "Test Employee", "AGE": 25, "ROLE": "Baker", "Monday": "06:00-14:00", "Tuesday": "06:00-14:00", "Wednesday": "OFF", "Thursday": "OFF", "Friday": "OFF", "Saturday": "OFF", "Sunday": "OFF"}
    ])
    summary = app.calculate_roster_wages(sample_roster)
    assert isinstance(summary, dict)
    assert "total_gross" in summary
    assert "total_tax" in summary
    assert "total_net" in summary
    assert summary["total_gross"] >= 0

def test_calculate_weekly_hour_rate_breakdown():
    sample_roster = pd.DataFrame([
        {"NAME": "John Baker", "AGE": 25, "ROLE": "Baker", "Monday": "06:00-14:00", "Tuesday": "OFF", "Wednesday": "OFF", "Thursday": "OFF", "Friday": "OFF", "Saturday": "OFF", "Sunday": "OFF"}
    ])
    breakdown = app.calculate_weekly_hour_rate_breakdown(sample_roster)
    assert isinstance(breakdown, pd.DataFrame)

# 3. TEST DATA CLEANSING & STANDARDIZATION
def test_sanitize_dataframe():
    df = pd.DataFrame({"A": [1, 2], "A ": [3, 4]})
    clean_df = app.sanitize_dataframe(df)
    assert clean_df.shape[1] == 1

def test_reorder_roster_dataframe():
    df = pd.DataFrame({"Tuesday": ["OFF"], "NAME": ["Alice"], "Monday": ["06:00-12:00"]})
    reordered = app.reorder_roster_dataframe(df)
    cols = list(reordered.columns)
    assert cols[0] == "NAME"
    assert cols[1] == "Monday"
    assert cols[2] == "Tuesday"

def test_standardize_unavailability_df():
    df = pd.DataFrame([
        {"Staff": "Bob", "Date": "Monday", "Notes": "Unavailable morning"}
    ])
    std_df = app.standardize_unavailability_df(df)
    assert not std_df.empty

# 4. TEST FILE PERSISTENCE ENGINE
def test_persistence_engine():
    test_df = pd.DataFrame({"Name": ["Test User"], "Value": [100]})
    filename = "test_persistence_sample.csv"
    app.save_persisted_df(test_df, filename)
    loaded_df = app.load_persisted_df(filename)
    assert loaded_df is not None
    assert not loaded_df.empty
    assert loaded_df.iloc[0]["Name"] == "Test User"
    
    # Cleanup test file
    filepath = os.path.join(app.DATA_DIR, filename)
    if os.path.exists(filepath):
        os.remove(filepath)

# 5. TEST EXCEL & EXPORT GENERATORS
def test_build_roster_excel_bytes():
    sample_df = pd.DataFrame([
        {"NAME": "Alice", "ROLE": "Sales", "Monday": "07:00-15:00", "Tuesday": "OFF"}
    ])
    excel_bytes = app.build_roster_excel_bytes(sample_df, datetime.date(2026, 8, 10))
    assert isinstance(excel_bytes, bytes)
    assert len(excel_bytes) > 0

def test_generate_xero_timesheet_csv():
    sample_breakdown = pd.DataFrame([
        {"Employee": "John Baker", "Ordinary Hours": 38.0, "Saturday Hours": 0.0, "Sunday Hours": 0.0, "Public Holiday Hours": 0.0}
    ])
    csv_str = app.generate_xero_timesheet_csv(sample_breakdown)
    assert isinstance(csv_str, str)
    assert "John Baker" in csv_str or len(csv_str) > 0

# 6. TEST USER PROFILES & AUTH ENGINE
def test_user_profiles_management():
    profiles = app.load_user_profiles()
    assert isinstance(profiles, dict)
    active_profiles = app.get_active_user_profiles()
    assert isinstance(active_profiles, dict)

def test_build_welcome_email_content():
    subj, body = app.build_welcome_email_content("John Doe", "johndoe", "Temp1234!")
    assert "John Doe" in body or "johndoe" in body
    assert "Brumby's Bakery" in subj or "Brumby's" in body
