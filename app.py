import streamlit as st
import pandas as pd
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
import io
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

# Helper to read excel sheets robustly, skipping empty headers and titles
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
            # Ensure it is a multi-column row (skips titles/merged header banners) and has keywords
            if len(row_vals) >= 2:
                matches = sum(1 for val in row_vals for kw in keywords if kw in val)
                if matches >= 2:  # Must match at least 2 headers (e.g. NAME and Commence)
                    header_row_idx = idx
                    break
                
        # Re-read from that header row index
        df = pd.read_excel(uploaded_file, header=header_row_idx)
        df.columns = [str(c).strip() for c in df.columns]
        return df
    except Exception as e:
        st.error(f"Error parsing Excel file structure: {e}")
        return None

# Initialize Session States for manual inputs if they don't exist
if 'manual_employees' not in st.session_state:
    st.session_state.manual_employees = pd.DataFrame([
        {"Name": "Elizabeth", "Role": "Senior Team Member", "Age": 28, "Employment Type": "Part-Time", "Start Date": "2024-01-01"},
        {"Name": "Stella", "Role": "Junior Team Member", "Age": 17, "Employment Type": "Casual", "Start Date": "2024-03-15"},
        {"Name": "Ainsley", "Role": "Junior Team Member", "Age": 19, "Employment Type": "Casual", "Start Date": "2024-05-10"},
        {"Name": "Aimi", "Role": "Junior Team Member", "Age": 20, "Employment Type": "Casual", "Start Date": "2024-06-01"},
        {"Name": "Jude", "Role": "Senior Team Member", "Age": 25, "Employment Type": "Full-Time", "Start Date": "2024-01-01"},
        {"Name": "Aroha", "Role": "Senior Team Member", "Age": 32, "Employment Type": "Full-Time", "Start Date": "2023-11-01"},
        {"Name": "Robert", "Role": "Senior Team Member", "Age": 45, "Employment Type": "Full-Time", "Start Date": "2023-10-01"},
    ])

if 'manual_unavailability' not in st.session_state:
    st.session_state.manual_unavailability = pd.DataFrame([
        {"Employee": "Elizabeth", "Day": "Saturday", "Time Window": "All Day"},
        {"Employee": "Elizabeth", "Day": "Sunday", "Time Window": "All Day"},
        {"Employee": "Stella", "Day": "Monday", "Time Window": "Before 3:30pm"},
        {"Employee": "Stella", "Day": "Tuesday", "Time Window": "Before 3:30pm"},
        {"Employee": "Stella", "Day": "Thursday", "Time Window": "Before 3:30pm"},
        {"Employee": "Stella", "Day": "Friday", "Time Window": "Before 3:30pm"},
    ])

if 'manual_requirements' not in st.session_state:
    st.session_state.manual_requirements = pd.DataFrame([
        {"Day": "Monday", "Shift": "7:30am-12:30pm", "Count Required": 2},
        {"Day": "Monday", "Shift": "12:30pm-5:30pm", "Count Required": 1},
        {"Day": "Tuesday", "Shift": "7:30am-12:30pm", "Count Required": 2},
        {"Day": "Tuesday", "Shift": "12:30pm-5:30pm", "Count Required": 1},
        {"Day": "Wednesday", "Shift": "7:30am-12:30pm", "Count Required": 2},
        {"Day": "Wednesday", "Shift": "12:30pm-5:30pm", "Count Required": 1},
        {"Day": "Thursday", "Shift": "7:30am-12:30pm", "Count Required": 2},
        {"Day": "Thursday", "Shift": "12:30pm-5:30pm", "Count Required": 1},
        {"Day": "Friday", "Shift": "7:30am-12:30pm", "Count Required": 2},
        {"Day": "Friday", "Shift": "12:30pm-5:30pm", "Count Required": 1},
        {"Day": "Saturday", "Shift": "8:30am-1:30pm", "Count Required": 2},
        {"Day": "Sunday", "Shift": "9:00am-2:00pm", "Count Required": 2},
    ])

if 'manual_fixed' not in st.session_state:
    st.session_state.manual_fixed = pd.DataFrame([
        {"Employee": "Elizabeth", "Monday": "7:30am-12:30pm", "Tuesday": "off", "Wednesday": "off", "Thursday": "off", "Friday": "off", "Saturday": "off", "Sunday": "off"},
        {"Employee": "Aroha", "Monday": "6:00am-1:00pm", "Tuesday": "6:00am-1:00pm", "Wednesday": "6:00am-1:00pm", "Thursday": "off", "Friday": "off", "Saturday": "6:00am-2:00pm", "Sunday": "6:00am-11:00am"},
    ])

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
    if upload_emp is not None:
        loaded = read_excel_robust(upload_emp)
        if loaded is not None:
            st.session_state.manual_employees = loaded
    employees_df = st.data_editor(st.session_state.manual_employees, num_rows="dynamic", key="edit_employees")

# Tab 2: Unavailability
with tab2:
    st.subheader("Log Unavailability")
    upload_unavail = st.file_uploader("Upload unavailability list.xlsx (Optional)", type=["xlsx"], key="unavail_upload")
    if upload_unavail is not None:
        loaded = read_excel_robust(upload_unavail)
        if loaded is not None:
            st.session_state.manual_unavailability = loaded
    unavailability_df = st.data_editor(st.session_state.manual_unavailability, num_rows="dynamic", key="edit_unavailability")

# Tab 3: Daily Requirements
with tab3:
    st.subheader("Daily Shift Requirements")
    upload_req = st.file_uploader("Upload Daily Shift personel requirement.xlsx (Optional)", type=["xlsx"], key="req_upload")
    if upload_req is not None:
        loaded = read_excel_robust(upload_req)
        if loaded is not None:
            st.session_state.manual_requirements = loaded
    requirements_df = st.data_editor(st.session_state.manual_requirements, num_rows="dynamic", key="edit_requirements")

# Tab 4: Fixed Shifts
with tab4:
    st.subheader("Fixed Baseline Shifts")
    upload_fixed = st.file_uploader("Upload Roster fixed - dont change.xlsx (Optional)", type=["xlsx"], key="fixed_upload")
    if upload_fixed is not None:
        loaded = read_excel_robust(upload_fixed)
        if loaded is not None:
            st.session_state.manual_fixed = loaded
    fixed_df = st.data_editor(st.session_state.manual_fixed, num_rows="dynamic", key="edit_fixed")

# Helpers for parsing times
def parse_time_to_decimal(time_str):
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

# Deterministic solver
def solve_roster(employees_raw, unavailability_raw, requirements_raw, fixed_raw, start_dt):
    employees = employees_raw.copy()
    unavailability = unavailability_raw.copy()
    requirements = requirements_raw.copy()
    fixed = fixed_raw.copy()

    # Standardize headers in case of leading/trailing spaces
    employees.columns = [str(c).strip() for c in employees.columns]
    unavailability.columns = [str(c).strip() for c in unavailability.columns]
    requirements.columns = [str(c).strip() for c in requirements.columns]
    fixed.columns = [str(c).strip() for c in fixed.columns]

    # Employees mapping
    emp_rename = {}
    c_name = find_column(employees, ["name", "employee", "employee name", "staff name", "staff"])
    if c_name: emp_rename[c_name] = "Name"
    c_start = find_column(employees, ["start date", "commencement date", "started", "startdate", "commence date", "commence"])
    if c_start: emp_rename[c_start] = "Start Date"
    c_age = find_column(employees, ["age", "years"])
    if c_age: emp_rename[c_age] = "Age"
    c_dob = find_column(employees, ["dob", "date of birth", "birth date"])
    if c_dob: emp_rename[c_dob] = "DOB"
    c_type = find_column(employees, ["employment type", "type", "status", "ft/pt/casual", "employmenttype"])
    if c_type: emp_rename[c_type] = "Employment Type"
    employees = employees.rename(columns=emp_rename)

    # Calculate Age dynamically if DOB is present and Age is missing
    if "DOB" in employees.columns and ("Age" not in employees.columns or employees["Age"].isna().all()):
        try:
            dob_parsed = pd.to_datetime(employees["DOB"], errors='coerce')
            today = datetime.now()
            employees["Age"] = dob_parsed.apply(lambda x: today.year - x.year - ((today.month, today.day) < (x.month, x.day)) if pd.notna(x) else 25)
        except:
            employees["Age"] = 25

    # Unavailability mapping
    unavail_rename = {}
    c_u_name = find_column(unavailability, ["employee", "name", "employee name", "staff name", "staff"])
    if c_u_name: unavail_rename[c_u_name] = "Employee"
    c_u_day = find_column(unavailability, ["day", "date", "weekday"])
    if c_u_day: unavail_rename[c_u_day] = "Day"
    c_u_window = find_column(unavailability, ["time window", "window", "time", "unavailability", "reason"])
    if c_u_window: unavail_rename[c_u_window] = "Time Window"
    unavailability = unavailability.rename(columns=unavail_rename)

    # Requirements mapping
    req_rename = {}
    c_r_day = find_column(requirements, ["day", "date", "weekday"])
    if c_r_day: req_rename[c_r_day] = "Day"
    c_r_shift = find_column(requirements, ["shift", "time", "hours", "shift time"])
    if c_r_shift: req_rename[c_r_shift] = "Shift"
    c_r_count = find_column(requirements, ["count required", "count", "required", "staff needed", "personnel", "quantity", "countrequired"])
    if c_r_count: req_rename[c_r_count] = "Count Required"
    requirements = requirements.rename(columns=req_rename)

    # Fixed mapping
    fixed_rename = {}
    c_f_name = find_column(fixed, ["employee", "name", "employee name", "staff name", "staff"])
    if c_f_name: fixed_rename[c_f_name] = "Employee"
    fixed = fixed.rename(columns=fixed_rename)

    # Filter active employees
    active_employees = []
    for _, emp in employees.iterrows():
        name_val = emp.get("Name")
        if pd.isna(name_val) or str(name_val).strip() == "":
            continue
        try:
            start_date_parsed = pd.to_datetime(emp.get("Start Date"))
            if start_date_parsed.date() <= start_dt:
                active_employees.append(emp.to_dict())
        except:
            active_employees.append(emp.to_dict())

    if not active_employees:
        return pd.DataFrame()

    days_of_week = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"]
    roster_output = {emp["Name"]: {day: "off" for day in days_of_week} for emp in active_employees}
    weekly_shifts_count = {emp["Name"]: 0 for emp in active_employees}
    elizabeth_weekday_shifts = 0

    # 1. Apply fixed shifts first
    for _, fix_row in fixed.iterrows():
        name = fix_row.get("Employee")
        if name in roster_output:
            for day in days_of_week:
                val = str(fix_row.get(day, "off")).strip()
                if val.lower() not in ["off", "nan", ""]:
                    roster_output[name][day] = val
                    weekly_shifts_count[name] += 1
                    if name == "Elizabeth" and day not in ["Saturday", "Sunday"]:
                        elizabeth_weekday_shifts += 1

    # 2. Check general day-off unavailability
    for _, un_row in unavailability.iterrows():
        name = un_row.get("Employee")
        day = un_row.get("Day")
        window = str(un_row.get("Time Window", "All Day")).strip().lower()
        if name in roster_output and day in days_of_week:
            if "all day" in window or "anytime" in window:
                if roster_output[name][day] == "off":
                    roster_output[name][day] = " unavailable"

    # 3. Schedule required shifts day by day
    for day in days_of_week:
        day_reqs = requirements[requirements["Day"].astype(str).str.lower() == day.lower()]
        shifts_to_fill = []
        for _, req in day_reqs.iterrows():
            shift = req.get("Shift")
            count = int(req.get("Count Required", 1))
            parsed = parse_shift_range(shift)
            if parsed:
                start, end, duration = parsed
                for _ in range(count):
                    shifts_to_fill.append({"shift": shift, "start": start, "end": end, "duration": duration})
                    
        # Sort shifts by duration descending (longest shift first)
        shifts_to_fill = sorted(shifts_to_fill, key=lambda x: x["duration"], reverse=True)

        for shift_info in shifts_to_fill:
            best_candidate = None
            best_score = -999999

            for emp in active_employees:
                name = emp["Name"]
                try:
                    age = int(emp.get("Age", 25))
                except:
                    age = 25
                
                # Check if already working a shift today
                if roster_output[name][day] != "off" and roster_output[name][day] != " unavailable":
                    continue
                
                # Check weekly shift limit
                if weekly_shifts_count[name] >= 5:
                    continue

                # Check day-specific unavailability
                is_emp_unavailable = False
                emp_unavail = unavailability[(unavailability["Employee"] == name) & (unavailability["Day"].astype(str).str.lower() == day.lower())]
                for _, un_row in emp_unavail.iterrows():
                    window = str(un_row.get("Time Window", "All Day"))
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
                roster_out_df = solve_roster(employees_df, unavailability_df, requirements_df, fixed_df, start_date)
                
                if roster_out_df.empty:
                    st.warning("⚠️ Roster generated is empty. Please verify that your Employees tab lists active employee names, and the Start Dates are not in the future.")
                    st.markdown("### Debug Information:")
                    st.write("Active Employees Table Loaded in App:", employees_df)
                else:
                    st.success("Roster successfully generated!")
                    st.dataframe(roster_out_df)
                    
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
                    
                    for row_idx, row_data in enumerate(roster_out_df.itertuples(index=False), 2):
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
                
            except Exception as e:
                st.error(f"Failed to generate roster: {e}")
