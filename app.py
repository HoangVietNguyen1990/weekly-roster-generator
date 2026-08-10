import streamlit as st
import pandas as pd
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
import io
import os
from datetime import datetime, timedelta

# Set page configuration with a modern design
st.set_page_config(
    page_title="Brumby's Bakery Roster Creator",
    page_icon="🥐",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom high-contrast Emerald & Golden Wheat Bakery styling
st.markdown("""
<style>
    /* Main App Background - Deep Emerald & Forest Teal Gradient */
    .stApp {
        background: linear-gradient(135deg, #0e2b26 0%, #16443c 50%, #1f574d 100%);
        color: #ffffff;
    }

    /* GLOBAL TEXT OVERRIDE FOR ALL LABELS & WIDGET TITLES */
    label, 
    label p, 
    label span, 
    div[data-testid="stWidgetLabel"], 
    div[data-testid="stWidgetLabel"] *, 
    .stMarkdown, 
    .stMarkdown p, 
    .stMarkdown span {
        color: #ffffff !important;
        font-weight: 600 !important;
        opacity: 1 !important;
    }
    
    /* Header Banner Styling - Ultra Standout Title */
    .header-style {
        background: linear-gradient(135deg, #ffffff 0%, #f7d594 40%, #e5a93c 100%);
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
        font-weight: 900;
        font-size: 3.2rem;
        margin-bottom: 6px;
        letter-spacing: -0.8px;
        filter: drop-shadow(0 2px 8px rgba(0,0,0,0.5));
    }
    .sub-header-style {
        color: #c8e6e0 !important;
        font-size: 1.2rem;
        margin-bottom: 25px;
        font-weight: 400;
    }

    /* Subheadings */
    h1, h2, h3, h4, h5, h6 {
        color: #ffffff !important;
        font-weight: 700 !important;
    }

    /* TABLE HEADERS: GREEN BACKGROUND (#081D19) WITH BOLD WHITE TEXT (#FFFFFF) ACROSS ALL TABS */
    div[data-testid="stDataEditor"], div[data-testid="stDataFrame"] {
        border-radius: 0 0 14px 14px !important;
        overflow: hidden !important;
        box-shadow: 0 8px 28px rgba(0, 0, 0, 0.45) !important;
        border: 2px solid #e5a93c !important;
        background-color: #ffffff !important;
        
        /* Glide Data Grid Header Canvas CSS Variables */
        --gdg-bg-header: #081d19 !important;
        --gdg-bg-header-has-focus: #133b34 !important;
        --gdg-text-header: #ffffff !important;
        --gdg-text-header-selected: #ffffff !important;
        --gdg-font-family: 'Inter', sans-serif !important;
    }
    
    /* Fallback DOM Header Element Styling - Green Background & White Text */
    div[data-testid="stDataEditor"] th, 
    div[data-testid="stDataFrame"] th, 
    div[class*="header"], 
    div[class*="Header"],
    div[role="columnheader"],
    div[class*="gdg-header"] {
        background-color: #081d19 !important;
        background: linear-gradient(135deg, #081d19 0%, #16443c 100%) !important;
        color: #ffffff !important;
        font-weight: 900 !important;
        font-size: 1.05rem !important;
        border-bottom: 2px solid #e5a93c !important;
    }
    div[role="columnheader"] *, div[role="columnheader"] p, div[role="columnheader"] span {
        color: #ffffff !important;
        font-weight: 900 !important;
    }

    /* Tab Bar Customization - High Contrast Inactive Tabs */
    .stTabs [data-baseweb="tab-list"] {
        gap: 8px;
        background-color: rgba(9, 32, 28, 0.85);
        padding: 8px;
        border-radius: 14px;
        border: 1px solid rgba(229, 169, 60, 0.35);
    }
    .stTabs button[role="tab"],
    .stTabs [data-baseweb="tab"],
    .stTabs [data-baseweb="tab"] *,
    .stTabs [aria-selected="false"],
    .stTabs [aria-selected="false"] * {
        color: #ffffff !important;
        opacity: 1 !important;
        font-weight: 600 !important;
    }
    .stTabs [aria-selected="false"] {
        background-color: rgba(255, 255, 255, 0.12) !important;
        border-radius: 10px !important;
    }
    .stTabs [data-baseweb="tab"]:hover {
        background-color: rgba(229, 169, 60, 0.3) !important;
    }
    .stTabs [aria-selected="true"] {
        background: linear-gradient(135deg, #e5a93c 0%, #d48827 100%) !important;
        border: none !important;
        box-shadow: 0 4px 14px rgba(229, 169, 60, 0.4) !important;
    }
    .stTabs [aria-selected="true"],
    .stTabs [aria-selected="true"] * {
        color: #0e2b26 !important;
        font-weight: 800 !important;
        opacity: 1 !important;
    }

    /* Radio Buttons High Contrast */
    .stRadio label, .stRadio div, .stRadio span, .stRadio p, .stRadio * {
        color: #ffffff !important;
        font-size: 1rem !important;
        font-weight: 600 !important;
        opacity: 1 !important;
    }

    /* File Uploader Dropzone & Instruction Text High Contrast */
    [data-testid="stFileUploaderDropzone"] {
        background-color: rgba(9, 32, 28, 0.85) !important;
        border: 2px dashed rgba(229, 169, 60, 0.6) !important;
        border-radius: 12px !important;
        padding: 15px !important;
    }
    [data-testid="stFileUploaderDropzone"]:hover {
        border-color: #e5a93c !important;
        background-color: rgba(9, 32, 28, 0.95) !important;
    }
    [data-testid="stFileUploaderDropzone"] *,
    [data-testid="stFileUploaderInstructions"] *,
    section[data-testid="stFileUploader"] * {
        color: #ffffff !important;
        opacity: 1 !important;
    }

    /* Hero Generate Button Styling - Tight Under Date Picker */
    .hero-generate-btn {
        margin-top: -10px !important;
    }
    .hero-generate-btn button {
        width: 100% !important;
        background: linear-gradient(135deg, #e5a93c 0%, #d48827 100%) !important;
        color: #0e2b26 !important;
        border: none !important;
        padding: 14px 28px !important;
        border-radius: 12px !important;
        font-weight: 900 !important;
        font-size: 1.2rem !important;
        letter-spacing: 0.5px !important;
        transition: all 0.3s ease !important;
        box-shadow: 0 6px 20px rgba(229, 169, 60, 0.45) !important;
        margin-top: 0px !important;
        margin-bottom: 5px !important;
    }
    .hero-generate-btn button p, .hero-generate-btn button span {
        color: #0e2b26 !important;
        font-weight: 900 !important;
        font-size: 1.2rem !important;
    }
    .hero-generate-btn button:hover {
        transform: translateY(-2px) scale(1.01) !important;
        box-shadow: 0 10px 28px rgba(229, 169, 60, 0.65) !important;
        background: linear-gradient(135deg, #f7d594 0%, #e5a93c 100%) !important;
    }

    /* General Action Buttons */
    .stButton>button {
        background: linear-gradient(135deg, #e5a93c 0%, #d48827 100%);
        color: #0e2b26;
        border: none;
        padding: 12px 28px;
        border-radius: 10px;
        font-weight: 800;
        font-size: 1.05rem;
        transition: all 0.3s ease;
        box-shadow: 0 4px 14px rgba(229, 169, 60, 0.35);
    }

    /* Download Button Specific Accent */
    .stDownloadButton>button {
        background: linear-gradient(135deg, #2e7d6e 0%, #1b5349 100%);
        color: #ffffff !important;
        border: 1px solid #e5a93c;
        padding: 14px 32px;
        border-radius: 12px;
        font-weight: 800;
        font-size: 1.1rem;
        transition: all 0.3s ease;
    }
    .stDownloadButton>button:hover {
        background: linear-gradient(135deg, #e5a93c 0%, #d48827 100%);
        color: #0e2b26 !important;
        transform: translateY(-2px);
        box-shadow: 0 6px 18px rgba(229, 169, 60, 0.4);
    }

    /* Sidebar Styling High Contrast */
    section[data-testid="stSidebar"] {
        background-color: #09201c;
        border-right: 1px solid rgba(229, 169, 60, 0.2);
    }
    section[data-testid="section-sidebar"] * {
        color: #ffffff !important;
    }
</style>
""", unsafe_allow_html=True)

st.markdown('<h1 class="header-style">🥐 Brumby\'s Pakenham — Weekly Staff Roster</h1>', unsafe_allow_html=True)
st.markdown('<p class="sub-header-style">Law-compliant, automated shift scheduling tailored for artisan bakeries.</p>', unsafe_allow_html=True)

# --- SIDEBAR CONFIGURATION ---
st.sidebar.image("https://img.icons8.com/fluency/96/bakery.png", width=80)
st.sidebar.title("🍞 Bakery Roster Controls")
st.sidebar.info("This application runs locally and uses an offline deterministic constraint solver to optimize staff rostering based on General Retail Award rules and bakery template requirements.")

# --- DISK PERSISTENCE ENGINE ---
DATA_DIR = "data"
os.makedirs(DATA_DIR, exist_ok=True)

def load_persisted_df(filename, default_df):
    path = os.path.join(DATA_DIR, filename)
    if os.path.exists(path):
        try:
            # Read all columns as string to avoid formatting locks
            df = pd.read_csv(path, dtype=str)
            return df
        except:
            return default_df
    return default_df

def save_persisted_df(df, filename):
    path = os.path.join(DATA_DIR, filename)
    try:
        df.astype(str).to_csv(path, index=False)
    except:
        pass

# Helper to read excel sheets robustly, converting everything to strings for easy editing
def read_excel_robust(uploaded_file):
    if uploaded_file is None:
        return None
    try:
        # Load without headers first to scan rows
        df_raw = pd.read_excel(uploaded_file, header=None)
        
        # Look for the header row containing key columns
        keywords = ["name", "employee", "staff", "role", "age", "dob", "commence", "start", "shift", "day", "monday", "tuesday", "unavailability", "fixed", "status", "type"]
        header_row_idx = 0
        
        for idx, row in df_raw.iterrows():
            row_vals = [str(x).strip().lower() for x in row if pd.notna(x)]
            if len(row_vals) >= 2:
                matches = sum(1 for val in row_vals for kw in keywords if kw in val)
                if matches >= 2:
                    header_row_idx = idx
                    break
                
        # Re-read from that header row index
        df = pd.read_excel(uploaded_file, header=header_row_idx)
        df.columns = [str(c).strip() for c in df.columns]
        
        # Convert all columns to strings to make them fully editable in st.data_editor
        for col in df.columns:
            if pd.api.types.is_datetime64_any_dtype(df[col]):
                df[col] = df[col].dt.strftime('%Y-%m-%d')
            else:
                df[col] = df[col].astype(str).replace("NaT", "").replace("nan", "")
        return df
    except Exception as e:
        st.error(f"Error parsing Excel file structure: {e}")
        return None

# Initialize Session States with Disk Cache
if 'manual_employees' not in st.session_state:
    default_emp = pd.DataFrame([
        {"Name": "Elizabeth", "Role": "Senior Team Member", "Age": "28", "Employment Type": "Part-Time", "Start Date": "2024-01-01"},
        {"Name": "Stella", "Role": "Junior Team Member", "Age": "17", "Employment Type": "Casual", "Start Date": "2024-03-15"},
        {"Name": "Ainsley Mactier", "Role": "Junior Team Member", "Age": "19", "Employment Type": "Casual", "Start Date": "2024-05-10"},
        {"Name": "Aimi", "Role": "Junior Team Member", "Age": "20", "Employment Type": "Casual", "Start Date": "2024-06-01"},
        {"Name": "Jude", "Role": "Senior Team Member", "Age": "25", "Employment Type": "Full-Time", "Start Date": "2024-01-01"},
        {"Name": "Aroha", "Role": "Senior Team Member", "Age": "32", "Employment Type": "Full-Time", "Start Date": "2023-11-01"},
        {"Name": "Robert", "Role": "Senior Team Member", "Age": "45", "Employment Type": "Full-Time", "Start Date": "2023-10-01"},
    ])
    st.session_state.manual_employees = load_persisted_df("employees.csv", default_emp)

if 'manual_unavailability' not in st.session_state:
    default_unavail = pd.DataFrame([
        {"Employee": "Elizabeth", "Day": "Saturday", "Time Window": "All Day"},
        {"Employee": "Elizabeth", "Day": "Sunday", "Time Window": "All Day"},
        {"Employee": "Stella", "Day": "Monday", "Time Window": "Before 3:30pm"},
        {"Employee": "Stella", "Day": "Tuesday", "Time Window": "Before 3:30pm"},
        {"Employee": "Stella", "Day": "Thursday", "Time Window": "Before 3:30pm"},
        {"Employee": "Stella", "Day": "Friday", "Time Window": "Before 3:30pm"},
    ])
    st.session_state.manual_unavailability = load_persisted_df("unavailability.csv", default_unavail)

if 'manual_requirements' not in st.session_state:
    default_req = pd.DataFrame([
        {"Shift": "7:30am-12:30pm", "Monday": "2", "Tuesday": "2", "Wednesday": "2", "Thursday": "2", "Friday": "2", "Saturday": "0", "Sunday": "0"},
        {"Shift": "12:30pm-5:30pm", "Monday": "1", "Tuesday": "1", "Wednesday": "1", "Thursday": "1", "Friday": "1", "Saturday": "2", "Sunday": "2"},
    ])
    st.session_state.manual_requirements = load_persisted_df("requirements.csv", default_req)

if 'manual_fixed' not in st.session_state:
    default_fixed = pd.DataFrame([
        {"Employee": "Elizabeth", "Monday": "7:30am-12:30pm", "Tuesday": "off", "Wednesday": "off", "Thursday": "off", "Friday": "off", "Saturday": "off", "Sunday": "off"},
        {"Employee": "Aroha", "Monday": "6:00am-1:00pm", "Tuesday": "6:00am-1:00pm", "Wednesday": "6:00am-1:00pm", "Thursday": "off", "Friday": "off", "Saturday": "6:00am-2:00pm", "Sunday": "6:00am-11:00am"},
    ])
    st.session_state.manual_fixed = load_persisted_df("fixed.csv", default_fixed)

# --- NEW TAB ORDER: Home Page as Tab 1 ---
tab_home, tab_emp, tab_unavail, tab_req, tab_fixed = st.tabs([
    "🏠 Home / Roster Generator",
    "👥 Staff Members", 
    "🚫 Unavailability", 
    "📋 Daily Requirements", 
    "📌 Fixed Shifts"
])

# Helpers for parsing times
def parse_time_to_decimal(time_str):
    try:
        time_str = str(time_str).strip().lower().replace(" ", "")
        is_pm = "pm" in time_str
        time_str = time_str.replace("am", "").replace("pm", "")
        if ":" in time_str:
            parts = time_str.split(":")
            hours = int(parts[0])
            minutes = int(parts[1])
        else:
            hours = int(time_str)
            minutes = 0
        if is_pm and hours < 12:
            hours += 12
        elif not is_pm and hours == 12:
            hours = 0
        return hours + minutes / 60.0
    except:
        return 0.0

def parse_shift_range(shift_str):
    if not shift_str or str(shift_str).strip().lower() in ["off", "unavailable", "nan", ""]:
        return None
    try:
        parts = str(shift_str).split("-")
        start = parse_time_to_decimal(parts[0])
        end = parse_time_to_decimal(parts[1])
        duration = end - start if end > start else (24 - start) + end
        return start, end, duration
    except:
        return None

def is_overlapping_unavailability(unavail_str, shift_start, shift_end):
    unavail_str = str(unavail_str).strip().lower()
    if "all day" in unavail_str or "anytime" in unavail_str:
        return True
    if "before" in unavail_str:
        time_part = unavail_str.replace("before", "").strip()
        t = parse_time_to_decimal(time_part)
        return shift_start < t
    if "after" in unavail_str:
        time_part = unavail_str.replace("after", "").strip()
        t = parse_time_to_decimal(time_part)
        return shift_end > t
    r = parse_shift_range(unavail_str)
    if r:
        u_start, u_end, _ = r
        return max(shift_start, u_start) < min(shift_end, u_end)
    return False

def find_column(df, candidates, default=""):
    for c in df.columns:
        if str(c).strip().lower() in candidates:
            return c
    return default

# Robust Name Matcher
def find_matching_employee(raw_name, name_map):
    raw_name_clean = str(raw_name).strip().lower()
    if not raw_name_clean or raw_name_clean == "nan":
        return None
    # 1. Exact match
    if raw_name_clean in name_map:
        return name_map[raw_name_clean]
    # 2. Prefix match (e.g. "Ainsley" matches "Ainsley Mactier", "Ana" matches "anastasia")
    for norm_name, display_name in name_map.items():
        if norm_name.startswith(raw_name_clean) or raw_name_clean.startswith(norm_name):
            return display_name
    # 3. First name matches first name
    for norm_name, display_name in name_map.items():
        first_raw = raw_name_clean.split()[0]
        first_norm = norm_name.split()[0]
        if first_raw == first_norm:
            return display_name
    return None

# Deterministic solver
def solve_roster(employees_raw, unavailability_raw, requirements_raw, fixed_raw, start_dt, debug_logs=None):
    # Standardize column headers to lowercase and stripped
    employees = employees_raw.copy()
    employees.columns = [str(c).strip().lower() for c in employees.columns]
    
    unavailability = unavailability_raw.copy()
    unavailability.columns = [str(c).strip().lower() for c in unavailability.columns]
    
    requirements = requirements_raw.copy()
    requirements.columns = [str(c).strip().lower() for c in requirements.columns]
    
    fixed = fixed_raw.copy()
    fixed.columns = [str(c).strip().lower() for c in fixed.columns]

    # Map employees columns
    name_col = find_column(employees, ["name", "employee", "employee name", "staff name", "staff"])
    start_col = find_column(employees, ["start date", "commencement date", "started", "startdate", "commence date", "commence"])
    age_col = find_column(employees, ["age", "years"])
    dob_col = find_column(employees, ["dob", "date of birth", "birth date"])
    type_col = find_column(employees, ["employment type", "type", "status", "ft/pt/casual", "employmenttype"])

    # Build active employees list
    active_employees = []
    for _, row in employees.iterrows():
        raw_name = str(row.get(name_col, "")).strip() if name_col else ""
        if not raw_name or raw_name.lower() in ["nan", ""]:
            continue
            
        is_active = True
        if start_col:
            try:
                start_val = row.get(start_col)
                if pd.notna(start_val):
                    start_date_parsed = pd.to_datetime(start_val)
                    if start_date_parsed.date() > start_dt:
                        is_active = False
            except:
                pass
                
        if not is_active:
            continue
            
        age = 25
        if age_col and pd.notna(row.get(age_col)):
            try:
                age = int(row.get(age_col))
            except:
                pass
        elif dob_col and pd.notna(row.get(dob_col)):
            try:
                dob_parsed = pd.to_datetime(row.get(dob_col))
                today = datetime.now()
                age = today.year - dob_parsed.year - ((today.month, today.day) < (dob_parsed.month, dob_parsed.day))
            except:
                pass
                
        emp_type = str(row.get(type_col, "Casual")).strip() if type_col else "Casual"
        
        active_employees.append({
            "Name": raw_name,
            "NormalizedName": raw_name.lower(),
            "Age": age,
            "Type": emp_type
        })

    if not active_employees:
        return pd.DataFrame()

    # Create mapping from normalized name to original name
    name_map = {emp["NormalizedName"]: emp["Name"] for emp in active_employees}
    
    # Initialize Roster
    days_of_week = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"]
    roster_output = {emp["Name"]: {day: "off" for day in days_of_week} for emp in active_employees}
    weekly_shifts_count = {emp["Name"]: 0 for emp in active_employees}
    elizabeth_weekday_shifts = 0

    # 1. Apply fixed shifts first
    fixed_name_col = find_column(fixed, ["employee", "name", "employee name", "staff name", "staff"])
    if not fixed_name_col and not fixed.empty:
        fixed_name_col = fixed.columns[0]
        
    if fixed_name_col:
        for _, fix_row in fixed.iterrows():
            raw_fixed_name = str(fix_row.get(fixed_name_col, "")).strip().lower()
            name = find_matching_employee(raw_fixed_name, name_map)
            if name:
                for day in days_of_week:
                    day_col = find_column(fixed, [day.lower(), day.lower()[:3]])
                    if day_col:
                        val = str(fix_row.get(day_col, "off")).strip()
                        if val.lower() not in ["off", "nan", ""]:
                            roster_output[name][day] = val
                            weekly_shifts_count[name] += 1
                            if name == "Elizabeth" and day not in ["Saturday", "Sunday"]:
                                elizabeth_weekday_shifts += 1

    # 2. Check unavailability
    unavail_name_col = find_column(unavailability, ["employee", "name", "employee name", "staff name", "staff"])
    unavail_day_col = find_column(unavailability, ["day", "date", "weekday"])
    unavail_window_col = find_column(unavailability, ["time window", "window", "time", "unavailability", "reason", "time constraint", "constraint"])

    unavail_map = {}
    if unavail_name_col and unavail_day_col and unavail_window_col:
        for _, un_row in unavailability.iterrows():
            raw_unavail_name = str(un_row.get(unavail_name_col, "")).strip().lower()
            unavail_day = str(un_row.get(unavail_day_col, "")).strip().lower()
            window = str(un_row.get(unavail_window_col, "All Day")).strip()
            matched_name = find_matching_employee(raw_unavail_name, name_map)
            if matched_name and unavail_day:
                for day in days_of_week:
                    if unavail_day == day.lower() or unavail_day == day.lower()[:3]:
                        key = (matched_name.lower(), day.lower())
                        if key not in unavail_map:
                            unavail_map[key] = []
                        unavail_map[key].append(window)

    # Mark day-off unavailability
    for (norm_name, day_lower), windows in unavail_map.items():
        name = name_map.get(norm_name)
        if name:
            for window in windows:
                if "all day" in window.lower() or "anytime" in window.lower():
                    for day in days_of_week:
                        if day.lower() == day_lower:
                            if roster_output[name][day] == "off":
                                roster_output[name][day] = " unavailable"

    # 3. Schedule required shifts day by day
    req_day_col = find_column(requirements, ["day", "date", "weekday"])
    req_shift_col = find_column(requirements, ["shift", "time", "hours", "shift time"])
    has_day_cols = any(find_column(requirements, [day.lower(), day.lower()[:3]]) for day in days_of_week)

    for day in days_of_week:
        shifts_to_fill = []
        
        if has_day_cols and req_shift_col:
            day_col = find_column(requirements, [day.lower(), day.lower()[:3]])
            if day_col:
                for _, req in requirements.iterrows():
                    shift = str(req.get(req_shift_col, "")).strip()
                    count_val = req.get(day_col)
                    if pd.notna(count_val):
                        try:
                            count = int(float(count_val))
                        except:
                            count = 0
                        if count > 0:
                            parsed = parse_shift_range(shift)
                            if parsed:
                                start, end, duration = parsed
                                for _ in range(count):
                                    shifts_to_fill.append({"shift": shift, "start": start, "end": end, "duration": duration})
        elif req_day_col and req_shift_col:
            req_count_col = find_column(requirements, ["count required", "count", "required", "staff needed", "personnel", "quantity", "countrequired"])
            clean_day_series = requirements[req_day_col].astype(str).str.strip().str.lower()
            day_reqs = requirements[clean_day_series.str.startswith(day.lower()[:3])]
            
            for _, req in day_reqs.iterrows():
                shift = str(req.get(req_shift_col, "")).strip()
                count_val = req.get(req_count_col, 1)
                try:
                    count = int(count_val) if pd.notna(count_val) else 1
                except:
                    count = 1
                parsed = parse_shift_range(shift)
                if parsed:
                    start, end, duration = parsed
                    for _ in range(count):
                        shifts_to_fill.append({"shift": shift, "start": start, "end": end, "duration": duration})
                        
        fixed_shifts_today = []
        for name in roster_output:
            val = roster_output[name][day]
            if val != "off" and val != " unavailable":
                fixed_shifts_today.append(val)
                
        remaining_shifts_to_fill = []
        for shift_req in shifts_to_fill:
            filled_idx = -1
            req_clean = shift_req["shift"].strip().lower().replace(" ", "")
            for idx, fixed_shift in enumerate(fixed_shifts_today):
                fixed_clean = str(fixed_shift).strip().lower().replace(" ", "")
                if req_clean == fixed_clean:
                    filled_idx = idx
                    break
            if filled_idx >= 0:
                fixed_shifts_today.pop(filled_idx)
            else:
                remaining_shifts_to_fill.append(shift_req)
                
        shifts_to_fill = remaining_shifts_to_fill
        shifts_to_fill = sorted(shifts_to_fill, key=lambda x: x["duration"], reverse=True)

        for shift_info in shifts_to_fill:
            best_candidate = None
            best_score = -999999

            for emp in active_employees:
                name = emp["Name"]
                norm_name = emp["NormalizedName"]
                age = emp["Age"]
                
                if roster_output[name][day] != "off" and roster_output[name][day] != " unavailable":
                    continue
                
                if weekly_shifts_count[name] >= 5:
                    continue

                is_emp_unavailable = False
                key = (norm_name, day.lower())
                if key in unavail_map:
                    for window in unavail_map[key]:
                        if is_overlapping_unavailability(window, shift_info["start"], shift_info["end"]):
                            is_emp_unavailable = True
                            break
                if is_emp_unavailable:
                    continue

                if name == "Elizabeth":
                    if day in ["Saturday", "Sunday"]:
                        continue
                    if elizabeth_weekday_shifts >= 2:
                        continue

                if age < 18 and day not in ["Saturday", "Sunday"]:
                    if max(shift_info["start"], 9.0) < min(shift_info["end"], 15.5):
                        continue

                score = 0
                score -= weekly_shifts_count[name] * 20
                
                if age < 21:
                    score += (21 - age) * shift_info["duration"] * 2
                
                if name == "Elizabeth" and shift_info["start"] in [7.0, 7.5]:
                    score += 100

                if score > best_score:
                    best_score = score
                    best_candidate = name

            if best_candidate:
                roster_output[best_candidate][day] = shift_info["shift"]
                weekly_shifts_count[best_candidate] += 1
                if best_candidate == "Elizabeth" and day not in ["Saturday", "Sunday"]:
                    elizabeth_weekday_shifts += 1

    roster_rows = []
    for name, sched in roster_output.items():
        row = {"Employee": name}
        row.update(sched)
        roster_rows.append(row)
    return pd.DataFrame(roster_rows)

# --- TAB 1: HOME PAGE & ROSTER GENERATOR ---
with tab_home:
    st.markdown("""
    <div style="background: rgba(9, 32, 28, 0.7); border: 2px solid #e5a93c; border-radius: 16px; padding: 25px; margin-bottom: 25px; box-shadow: 0 8px 30px rgba(0,0,0,0.4);">
        <h2 style="color: #f7d594 !important; margin-top: 0; font-size: 1.8rem; font-weight: 800;">⚡ Weekly Roster Generator</h2>
        <p style="color: #ffffff !important; font-size: 1.05rem; margin-bottom: 0;">Configure your target week period below and hit the <b>Generate Weekly Roster</b> button to instantly build an award-compliant bakery schedule.</p>
    </div>
    """, unsafe_allow_html=True)
    
    col1, col2 = st.columns([1, 1])
    with col1:
        st.markdown('<div style="background: rgba(9, 32, 28, 0.6); padding: 16px; border-radius: 14px; border: 1px solid rgba(229,169,60,0.35);">', unsafe_allow_html=True)
        start_date = st.date_input("🗓️ Roster Start Date (Monday)", datetime.now() + timedelta(days=(0 - datetime.now().weekday())))
        
        # Position Hero Generate Button TIGHT right under the Date Picker
        st.markdown('<div class="hero-generate-btn">', unsafe_allow_html=True)
        if st.button("🚀 GENERATE WEEKLY ROSTER", key="btn_hero_generate"):
            with st.spinner("Calculating optimal bakery roster locally..."):
                try:
                    emp_data = st.session_state.manual_employees
                    unavail_data = st.session_state.manual_unavailability
                    req_data = st.session_state.manual_requirements
                    fixed_data = st.session_state.manual_fixed
                    
                    roster_out_df = solve_roster(emp_data, unavail_data, req_data, fixed_data, start_date)
                    st.session_state.final_roster_df = roster_out_df
                    st.success("🎉 Weekly Roster successfully generated!")
                except Exception as e:
                    st.error(f"Failed to generate roster: {e}")
        st.markdown('</div></div>', unsafe_allow_html=True)

    with col2:
        upload_template = st.file_uploader("📋 Custom Layout Template (.xlsx)", type=["xlsx"], key="home_template_upload")

    # Display generated roster & export controls
    if "final_roster_df" in st.session_state:
        final_df = st.session_state.final_roster_df
        
        if final_df.empty:
            st.warning("⚠️ Roster generated is empty. Please check your Staff Members tab to ensure active employees are listed.")
        else:
            st.markdown("""
            <div style="background: linear-gradient(135deg, #081d19 0%, #16443c 100%); padding: 12px 20px; border-radius: 14px 14px 0 0; color: #ffffff !important; font-weight: 800; font-size: 1.15rem; letter-spacing: 0.3px; border: 2px solid #e5a93c; border-bottom: none; margin-top: 15px;">
                📝 Generated Roster Preview & Manual Override Editor
            </div>
            """, unsafe_allow_html=True)
            edited_final_df = st.data_editor(final_df, num_rows="dynamic", key="home_final_roster_editor")
            st.session_state.final_roster_df = edited_final_df
            
            # Create Excel download
            if upload_template is not None:
                wb = openpyxl.load_workbook(upload_template)
            else:
                wb = openpyxl.Workbook()
                
            ws = wb.active
            ws.title = f"Roster {start_date.strftime('%d.%m.%Y')}"
            
            headers = ["Employee", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"]
            
            # Title banner
            ws.merge_cells('A1:H1')
            cell_title = ws['A1']
            cell_title.value = "Brumby's Pakenham Weekly Roster"
            cell_title.font = Font(name="Calibri", size=16, bold=True, color="1F4E78")
            cell_title.alignment = Alignment(horizontal="left", vertical="center")
            
            # Week period banner
            end_dt = start_date + timedelta(days=6)
            period_str = f"Week Period: {start_date.strftime('%d.%m.%Y')} to {end_dt.strftime('%d.%m.%Y')}"
            ws.merge_cells('A2:H2')
            cell_period = ws['A2']
            cell_period.value = period_str
            cell_period.font = Font(name="Calibri", size=11, italic=True, color="595959")
            cell_period.alignment = Alignment(horizontal="left", vertical="center")
            
            # Header Row
            for col_num, header in enumerate(headers, 1):
                cell = ws.cell(row=4, column=col_num)
                cell.value = header
                header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
                header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
                cell.fill = header_fill
                cell.font = header_font
                cell.alignment = Alignment(horizontal="center", vertical="center")
                
            thin_border = Border(
                left=Side(style='thin', color='BFBFBF'),
                right=Side(style='thin', color='BFBFBF'),
                top=Side(style='thin', color='BFBFBF'),
                bottom=Side(style='thin', color='BFBFBF')
            )
            
            off_fill = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")
            white_fill = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")
            text_font = Font(name="Calibri", size=11)
            
            from openpyxl.cell.cell import MergedCell
            
            for row_idx, row_data in enumerate(edited_final_df.itertuples(index=False), 5):
                for col_idx, value in enumerate(row_data, 1):
                    cell = ws.cell(row=row_idx, column=col_idx)
                    
                    target_cell = cell
                    if isinstance(cell, MergedCell):
                        for merged_range in ws.merged_cells.ranges:
                            if merged_range.min_row <= row_idx <= merged_range.max_row and merged_range.min_col <= col_idx <= merged_range.max_col:
                                target_cell = ws.cell(row=merged_range.min_row, column=merged_range.min_col)
                                break
                    
                    val_str = str(value).strip().lower()
                    if "unavailable" in val_str or val_str == "off" or val_str == "nan" or val_str == "":
                        target_cell.fill = off_fill
                        target_cell.value = ""
                    else:
                        target_cell.fill = white_fill
                        target_cell.value = value
                        
                    target_cell.font = text_font
                    target_cell.alignment = Alignment(horizontal="center", vertical="center")
                    target_cell.border = thin_border

            ws.column_dimensions['A'].width = 30
            for col_letter in ['B', 'C', 'D', 'E', 'F', 'G', 'H']:
                ws.column_dimensions[col_letter].width = 25
                
            ws.row_dimensions[1].height = 32
            ws.row_dimensions[2].height = 20
            ws.row_dimensions[3].height = 12
            ws.row_dimensions[4].height = 28
            
            last_data_row = len(edited_final_df) + 4
            for r in range(5, last_data_row + 1):
                ws.row_dimensions[r].height = 22
                
            # Award Break Card Note
            note_start_row = last_data_row + 3
            ws.row_dimensions[note_start_row].height = 26
            ws.merge_cells(start_row=note_start_row, start_column=1, end_row=note_start_row, end_column=8)
            note_hdr_cell = ws.cell(row=note_start_row, column=1)
            note_hdr_cell.value = "General Retail Industry Award - Required Breaks Reference Card:"
            note_hdr_fill = PatternFill(start_color="DDEBF7", end_color="DDEBF7", fill_type="solid")
            note_hdr_cell.fill = note_hdr_fill
            note_hdr_cell.font = Font(name="Calibri", size=11, bold=True, color="1F4E78")
            note_hdr_cell.alignment = Alignment(horizontal="left", vertical="center")
            
            sub_hdr_row = note_start_row + 1
            ws.row_dimensions[sub_hdr_row].height = 22
            
            ws.merge_cells(start_row=sub_hdr_row, start_column=1, end_row=sub_hdr_row, end_column=3)
            c1 = ws.cell(row=sub_hdr_row, column=1)
            c1.value = "Shift Duration"
            
            ws.merge_cells(start_row=sub_hdr_row, start_column=4, end_row=sub_hdr_row, end_column=6)
            c2 = ws.cell(row=sub_hdr_row, column=4)
            c2.value = "Paid Rest Break(s)"
            
            ws.merge_cells(start_row=sub_hdr_row, start_column=7, end_row=sub_hdr_row, end_column=8)
            c3 = ws.cell(row=sub_hdr_row, column=7)
            c3.value = "Unpaid Meal Break(s)"
            
            sub_hdr_fill = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")
            sub_font = Font(name="Calibri", size=10, bold=True)
            for c in [c1, c2, c3]:
                c.fill = sub_hdr_fill
                c.font = sub_font
                c.alignment = Alignment(horizontal="center", vertical="center")
                c.border = thin_border

            breaks_data = [
                ("Less than 4 hours", "None", "None"),
                ("4 hours up to 5 hours", "1 x 10 minutes", "None"),
                ("5 hours up to 7 hours", "1 x 10 minutes", "1 x 30 to 60 minutes"),
                ("7 hours up to 10 hours", "2 x 10 minutes", "1 x 30 to 60 minutes"),
                ("More than 10 hours", "2 x 10 minutes", "2 x 30 to 60 minutes")
            ]
            
            curr_row = sub_hdr_row + 1
            for duration_text, paid_breaks, unpaid_breaks in breaks_data:
                ws.row_dimensions[curr_row].height = 20
                
                ws.merge_cells(start_row=curr_row, start_column=1, end_row=curr_row, end_column=3)
                cell_dur = ws.cell(row=curr_row, column=1)
                cell_dur.value = duration_text
                
                ws.merge_cells(start_row=curr_row, start_column=4, end_row=curr_row, end_column=6)
                cell_paid = ws.cell(row=curr_row, column=4)
                cell_paid.value = paid_breaks
                
                ws.merge_cells(start_row=curr_row, start_column=7, end_row=curr_row, end_column=8)
                cell_unpaid = ws.cell(row=curr_row, column=7)
                cell_unpaid.value = unpaid_breaks
                
                for c in [cell_dur, cell_paid, cell_unpaid]:
                    c.font = Font(name="Calibri", size=10)
                    c.alignment = Alignment(horizontal="center", vertical="center")
                    c.border = thin_border
                
                curr_row += 1
            
            excel_data = io.BytesIO()
            wb.save(excel_data)
            excel_data.seek(0)
            
            st.markdown("<br>", unsafe_allow_html=True)
            st.download_button(
                label="📥 DOWNLOAD ROSTER EXCEL FILE (.XLSX)",
                data=excel_data,
                file_name=f"Team Roster {start_date.strftime('%d.%m.%Y')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="btn_download_excel_home"
            )

# --- TAB 2: STAFF MEMBERS ---
with tab_emp:
    st.subheader("Manage Bakery Employees")
    emp_mode = st.radio("Upload Mode:", ["Replace current data", "Append to current data"], key="emp_upload_mode", horizontal=True)
    upload_emp = st.file_uploader("Upload EMPLOYEE LIST.xlsx (Optional)", type=["xlsx"], key="emp_upload")
    
    if upload_emp is not None:
        file_key = f"processed_{upload_emp.name}_{upload_emp.size}_{emp_mode}"
        if st.session_state.get("last_emp_file") != file_key:
            loaded = read_excel_robust(upload_emp)
            if loaded is not None:
                if emp_mode == "Replace current data":
                    st.session_state.manual_employees = loaded
                else:
                    combined = pd.concat([st.session_state.manual_employees, loaded], ignore_index=True).drop_duplicates()
                    st.session_state.manual_employees = combined
                st.session_state.last_emp_file = file_key
                save_persisted_df(st.session_state.manual_employees, "employees.csv")
                
    st.markdown("""
    <div style="background: linear-gradient(135deg, #081d19 0%, #16443c 100%); padding: 10px 18px; border-radius: 12px 12px 0 0; color: #ffffff !important; font-weight: 800; font-size: 1.1rem; letter-spacing: 0.3px; border: 2px solid #e5a93c; border-bottom: none; margin-top: 15px;">
        👥 Bakery Staff Members List
    </div>
    """, unsafe_allow_html=True)
    employees_df = st.data_editor(st.session_state.manual_employees, num_rows="dynamic", key="edit_employees")
    st.session_state.manual_employees = employees_df
    save_persisted_df(employees_df, "employees.csv")

# --- TAB 3: UNAVAILABILITY ---
with tab_unavail:
    st.subheader("Log Staff Unavailability")
    unavail_mode = st.radio("Upload Mode:", ["Replace current data", "Append to current data"], key="unavail_upload_mode", horizontal=True)
    upload_unavail = st.file_uploader("Upload unavailability list.xlsx (Optional)", type=["xlsx"], key="unavail_upload")
    
    if upload_unavail is not None:
        file_key = f"processed_{upload_unavail.name}_{upload_unavail.size}_{unavail_mode}"
        if st.session_state.get("last_unavail_file") != file_key:
            loaded = read_excel_robust(upload_unavail)
            if loaded is not None:
                if unavail_mode == "Replace current data":
                    st.session_state.manual_unavailability = loaded
                else:
                    combined = pd.concat([st.session_state.manual_unavailability, loaded], ignore_index=True).drop_duplicates()
                    st.session_state.manual_unavailability = combined
                st.session_state.last_unavail_file = file_key
                save_persisted_df(st.session_state.manual_unavailability, "unavailability.csv")
                
    st.markdown("""
    <div style="background: linear-gradient(135deg, #081d19 0%, #16443c 100%); padding: 10px 18px; border-radius: 12px 12px 0 0; color: #ffffff !important; font-weight: 800; font-size: 1.1rem; letter-spacing: 0.3px; border: 2px solid #e5a93c; border-bottom: none; margin-top: 15px;">
        🚫 Staff Weekly Unavailability Constraints
    </div>
    """, unsafe_allow_html=True)
    unavailability_df = st.data_editor(st.session_state.manual_unavailability, num_rows="dynamic", key="edit_unavailability")
    st.session_state.manual_unavailability = unavailability_df
    save_persisted_df(unavailability_df, "unavailability.csv")

# --- TAB 4: DAILY REQUIREMENTS ---
with tab_req:
    st.subheader("Daily Bakery Shift Requirements")
    req_mode = st.radio("Upload Mode:", ["Replace current data", "Append to current data"], key="req_upload_mode", horizontal=True)
    upload_req = st.file_uploader("Upload Daily Shift personel requirement.xlsx (Optional)", type=["xlsx"], key="req_upload")
    
    if upload_req is not None:
        file_key = f"processed_{upload_req.name}_{upload_req.size}_{req_mode}"
        if st.session_state.get("last_req_file") != file_key:
            loaded = read_excel_robust(upload_req)
            if loaded is not None:
                if req_mode == "Replace current data":
                    st.session_state.manual_requirements = loaded
                else:
                    combined = pd.concat([st.session_state.manual_requirements, loaded], ignore_index=True).drop_duplicates()
                    st.session_state.manual_requirements = combined
                st.session_state.last_req_file = file_key
                save_persisted_df(st.session_state.manual_requirements, "requirements.csv")
                
    st.markdown("""
    <div style="background: linear-gradient(135deg, #081d19 0%, #16443c 100%); padding: 10px 18px; border-radius: 12px 12px 0 0; color: #ffffff !important; font-weight: 800; font-size: 1.1rem; letter-spacing: 0.3px; border: 2px solid #e5a93c; border-bottom: none; margin-top: 15px;">
        📋 Daily Shift Coverage Requirements (Mon-Sun)
    </div>
    """, unsafe_allow_html=True)
    requirements_df = st.data_editor(st.session_state.manual_requirements, num_rows="dynamic", key="edit_requirements")
    st.session_state.manual_requirements = requirements_df
    save_persisted_df(requirements_df, "requirements.csv")

# --- TAB 5: FIXED SHIFTS ---
with tab_fixed:
    st.subheader("Fixed Baseline Shifts")
    fixed_mode = st.radio("Upload Mode:", ["Replace current data", "Append to current data"], key="fixed_upload_mode", horizontal=True)
    upload_fixed = st.file_uploader("Upload Roster fixed - dont change.xlsx (Optional)", type=["xlsx"], key="fixed_upload")
    
    if upload_fixed is not None:
        file_key = f"processed_{upload_fixed.name}_{upload_fixed.size}_{fixed_mode}"
        if st.session_state.get("last_fixed_file") != file_key:
            loaded = read_excel_robust(upload_fixed)
            if loaded is not None:
                if fixed_mode == "Replace current data":
                    st.session_state.manual_fixed = loaded
                else:
                    combined = pd.concat([st.session_state.manual_fixed, loaded], ignore_index=True).drop_duplicates()
                    st.session_state.manual_fixed = combined
                st.session_state.last_fixed_file = file_key
                save_persisted_df(st.session_state.manual_fixed, "fixed.csv")
                
    st.markdown("""
    <div style="background: linear-gradient(135deg, #081d19 0%, #16443c 100%); padding: 10px 18px; border-radius: 12px 12px 0 0; color: #ffffff !important; font-weight: 800; font-size: 1.1rem; letter-spacing: 0.3px; border: 2px solid #e5a93c; border-bottom: none; margin-top: 15px;">
        📌 Fixed Baseline Staff Shifts
    </div>
    """, unsafe_allow_html=True)
    fixed_df = st.data_editor(st.session_state.manual_fixed, num_rows="dynamic", key="edit_fixed")
    st.session_state.manual_fixed = fixed_df
    save_persisted_df(fixed_df, "fixed.csv")
