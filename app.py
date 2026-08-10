import streamlit as st
import pandas as pd
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
import io
import os
from datetime import datetime, timedelta

# Set page configuration with a modern design
st.set_page_config(
    page_title="Weekly Roster Creator",
    page_icon="📅",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom premium styling
st.markdown("""
<style>
    .main {
        background-color: #0f111a;
        color: #ffffff;
    }
    .stTabs [data-baseweb="tab-list"] {
        gap: 8px;
    }
    .stTabs [data-baseweb="tab"] {
        height: 50px;
        white-space: pre-wrap;
        background-color: #1e2235;
        border-radius: 8px 8px 0px 0px;
        color: #a0aec0;
        padding-left: 20px;
        padding-right: 20px;
        font-weight: 600;
    }
    .stTabs [aria-selected="true"] {
        background-color: #3b82f6 !important;
        color: white !important;
    }
    .stButton>button {
        background: linear-gradient(135deg, #3b82f6 0%, #1d4ed8 100%);
        color: white;
        border: none;
        padding: 10px 24px;
        border-radius: 8px;
        font-weight: bold;
        transition: all 0.3s ease;
    }
    .stButton>button:hover {
        transform: translateY(-2px);
        box-shadow: 0 4px 12px rgba(59, 130, 246, 0.4);
    }
    .header-style {
        background: linear-gradient(135deg, #3b82f6 0%, #8b5cf6 100%);
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
        font-weight: 800;
        font-size: 3rem;
        margin-bottom: 20px;
    }
</style>
""", unsafe_allow_html=True)

st.markdown('<h1 class="header-style">Weekly Roster Creator</h1>', unsafe_allow_html=True)
st.write("Generate optimized, law-compliant rosters locally and instantly (no AI API Keys required).")

# --- SIDEBAR CONFIGURATION ---
st.sidebar.image("https://img.icons8.com/fluency/96/calendar.png", width=80)
st.sidebar.title("App Controls")
st.sidebar.info("This app runs locally on your browser/server. It uses a deterministic constraint solver to optimize staff rostering based on your rules and templates.")

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

# --- TABS FOR INPUT ---
tab1, tab2, tab3, tab4, tab5 = st.tabs([
    "👥 Employees", 
    "🚫 Unavailability", 
    "📋 Daily Requirements", 
    "📌 Fixed Shifts", 
    "⚙️ Generate & Export"
])

# Tab 1: Employees
with tab1:
    st.subheader("Manage Employees")
    upload_emp = st.file_uploader("Upload EMPLOYEE LIST.xlsx (Optional)", type=["xlsx"], key="emp_upload")
    emp_mode = st.radio("Upload Mode:", ["Replace current data", "Append to current data"], key="emp_upload_mode", horizontal=True)
    
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
                
    employees_df = st.data_editor(st.session_state.manual_employees, num_rows="dynamic", key="edit_employees")
    st.session_state.manual_employees = employees_df
    save_persisted_df(employees_df, "employees.csv")

# Tab 2: Unavailability
with tab2:
    st.subheader("Log Unavailability")
    upload_unavail = st.file_uploader("Upload unavailability list.xlsx (Optional)", type=["xlsx"], key="unavail_upload")
    unavail_mode = st.radio("Upload Mode:", ["Replace current data", "Append to current data"], key="unavail_upload_mode", horizontal=True)
    
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
                
    unavailability_df = st.data_editor(st.session_state.manual_unavailability, num_rows="dynamic", key="edit_unavailability")
    st.session_state.manual_unavailability = unavailability_df
    save_persisted_df(unavailability_df, "unavailability.csv")

# Tab 3: Daily Requirements
with tab3:
    st.subheader("Daily Shift Requirements")
    upload_req = st.file_uploader("Upload Daily Shift personel requirement.xlsx (Optional)", type=["xlsx"], key="req_upload")
    req_mode = st.radio("Upload Mode:", ["Replace current data", "Append to current data"], key="req_upload_mode", horizontal=True)
    
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
                
    requirements_df = st.data_editor(st.session_state.manual_requirements, num_rows="dynamic", key="edit_requirements")
    st.session_state.manual_requirements = requirements_df
    save_persisted_df(requirements_df, "requirements.csv")

# Tab 4: Fixed Shifts
with tab4:
    st.subheader("Fixed Baseline Shifts")
    upload_fixed = st.file_uploader("Upload Roster fixed - dont change.xlsx (Optional)", type=["xlsx"], key="fixed_upload")
    fixed_mode = st.radio("Upload Mode:", ["Replace current data", "Append to current data"], key="fixed_upload_mode", horizontal=True)
    
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
                
    fixed_df = st.data_editor(st.session_state.manual_fixed, num_rows="dynamic", key="edit_fixed")
    st.session_state.manual_fixed = fixed_df
    save_persisted_df(fixed_df, "fixed.csv")

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

    if debug_logs is not None:
        debug_logs["employees_cols"] = list(employees.columns)
        debug_logs["unavail_cols"] = list(unavailability.columns)
        debug_logs["requirements_cols"] = list(requirements.columns)
        debug_logs["fixed_cols"] = list(fixed.columns)

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

    if debug_logs is not None:
        debug_logs["active_employees_count"] = len(active_employees)
        debug_logs["active_employees_list"] = [e["Name"] for e in active_employees]

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
    # Fallback to column 0 if no matching header name
    if not fixed_name_col and not fixed.empty:
        fixed_name_col = fixed.columns[0]
        
    fixed_applied = 0
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
                            fixed_applied += 1
                            if name == "Elizabeth" and day not in ["Saturday", "Sunday"]:
                                elizabeth_weekday_shifts += 1

    if debug_logs is not None:
        debug_logs["fixed_applied_count"] = fixed_applied

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
    
    # Check if requirements contains columns matching days of the week (Grid Format)
    has_day_cols = any(find_column(requirements, [day.lower(), day.lower()[:3]]) for day in days_of_week)

    if debug_logs is not None:
        debug_logs["shifts_per_day"] = {}

    for day in days_of_week:
        shifts_to_fill = []
        
        if has_day_cols and req_shift_col:
            # Grid Format: Days are columns, rows are shifts, cell values are counts
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
            # Flat Format: Rows are (Day, Shift, Count)
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
                        
        # Get fixed shifts already applied for this day to subtract from requirements
        fixed_shifts_today = []
        for name in roster_output:
            val = roster_output[name][day]
            if val != "off" and val != " unavailable":
                fixed_shifts_today.append(val)
                
        # Subtract fixed shifts from required shifts to fill
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
                fixed_shifts_today.pop(filled_idx)  # Consume this fixed shift
            else:
                remaining_shifts_to_fill.append(shift_req)
                
        shifts_to_fill = remaining_shifts_to_fill
        shifts_to_fill = sorted(shifts_to_fill, key=lambda x: x["duration"], reverse=True)

        if debug_logs is not None:
            debug_logs["shifts_per_day"][day] = len(shifts_to_fill)

        for shift_info in shifts_to_fill:
            best_candidate = None
            best_score = -999999

            for emp in active_employees:
                name = emp["Name"]
                norm_name = emp["NormalizedName"]
                age = emp["Age"]
                
                # Check if already working a shift today
                if roster_output[name][day] != "off" and roster_output[name][day] != " unavailable":
                    continue
                
                # Check weekly shift limit
                if weekly_shifts_count[name] >= 5:
                    continue

                # Check unavailability overlap
                is_emp_unavailable = False
                key = (norm_name, day.lower())
                if key in unavail_map:
                    for window in unavail_map[key]:
                        if is_overlapping_unavailability(window, shift_info["start"], shift_info["end"]):
                            is_emp_unavailable = True
                            break
                if is_emp_unavailable:
                    continue

                # Rule: Elizabeth cap and weekends
                if name == "Elizabeth":
                    if day in ["Saturday", "Sunday"]:
                        continue
                    if elizabeth_weekday_shifts >= 2:
                        continue

                # Rule: Under-18 school hours (Mon-Fri 9:00 AM - 3:30 PM)
                if age < 18 and day not in ["Saturday", "Sunday"]:
                    if max(shift_info["start"], 9.0) < min(shift_info["end"], 15.5):
                        continue

                # Calculate heuristic score
                score = 0
                score -= weekly_shifts_count[name] * 20
                
                # Junior rate preference
                if age < 21:
                    score += (21 - age) * shift_info["duration"] * 2
                
                # Elizabeth early morning shift preference
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

    # Form dataframe output
    roster_rows = []
    for name, sched in roster_output.items():
        row = {"Employee": name}
        row.update(sched)
        roster_rows.append(row)
    return pd.DataFrame(roster_rows)


# Tab 5: Generate & Export
with tab5:
    st.subheader("Roster Configurations & Generation")
    
    col1, col2 = st.columns(2)
    with col1:
        start_date = st.date_input("Roster Start Date (Monday)", datetime.now() + timedelta(days=(0 - datetime.now().weekday())))
        upload_template = st.file_uploader("Upload Roster Layout Template (Optional)", type=["xlsx"], key="template_upload")
    
    if st.button("Generate Weekly Roster"):
        with st.spinner("Calculating optimal roster locally..."):
            try:
                # Debug logging structure
                debug_logs = {}
                roster_out_df = solve_roster(employees_df, unavailability_df, requirements_df, fixed_df, start_date, debug_logs)
                
                # Store in session state
                st.session_state.final_roster_df = roster_out_df
                st.success("Roster successfully generated!")
            except Exception as e:
                st.error(f"Failed to generate roster: {e}")

    if "final_roster_df" in st.session_state:
        final_df = st.session_state.final_roster_df
        
        if final_df.empty:
            st.warning("⚠️ Roster generated is empty. Please verify that your Employees tab lists active employee names.")
        else:
            st.markdown("### 📝 Final Generated Roster (Double-click any cell to manually override/edit)")
            edited_final_df = st.data_editor(final_df, num_rows="dynamic", key="final_roster_editor")
            st.session_state.final_roster_df = edited_final_df
            
            # Create Excel download
            if upload_template is not None:
                wb = openpyxl.load_workbook(upload_template)
            else:
                wb = openpyxl.Workbook()
                
            ws = wb.active
            ws.title = f"Roster {start_date.strftime('%d.%m.%Y')}"
            
            headers = ["Employee", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"]
            
            # If using template, clean up rows or write in place
            if upload_template is None:
                ws.append(headers)
                
            header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
            header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
            thin_border = Border(
                left=Side(style='thin', color='BFBFBF'),
                right=Side(style='thin', color='BFBFBF'),
                top=Side(style='thin', color='BFBFBF'),
                bottom=Side(style='thin', color='BFBFBF')
            )
            
            if upload_template is None:
                for col_num, header in enumerate(headers, 1):
                    cell = ws.cell(row=1, column=col_num)
                    cell.fill = header_fill
                    cell.font = header_font
                    cell.alignment = Alignment(horizontal="center", vertical="center")
                    cell.border = thin_border
            
            unavail_fill = PatternFill(start_color="FCE4D6", end_color="FCE4D6", fill_type="solid")
            off_fill = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")
            text_font = Font(name="Calibri", size=11)
            
            from openpyxl.cell.cell import MergedCell
            
            for row_idx, row_data in enumerate(edited_final_df.itertuples(index=False), 2):
                for col_idx, value in enumerate(row_data, 1):
                    cell = ws.cell(row=row_idx, column=col_idx)
                    
                    target_cell = cell
                    if isinstance(cell, MergedCell):
                            for merged_range in ws.merged_cells.ranges:
                                if merged_range.min_row <= row_idx <= merged_range.max_row and merged_range.min_col <= col_idx <= merged_range.max_col:
                                    target_cell = ws.cell(row=merged_range.min_row, column=merged_range.min_col)
                                    break
                                
                    target_cell.value = value
                    target_cell.font = text_font
                    target_cell.alignment = Alignment(horizontal="center", vertical="center")
                    target_cell.border = thin_border
                    
                    val_str = str(value).strip().lower()
                    if "unavailable" in val_str:
                        target_cell.fill = unavail_fill
                    elif val_str == "off":
                        target_cell.fill = off_fill
            
            excel_data = io.BytesIO()
            wb.save(excel_data)
            excel_data.seek(0)
            
            st.download_button(
                label="📥 Download Roster Excel File",
                data=excel_data,
                file_name=f"Team Roster {start_date.strftime('%d.%m.%Y')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
