import streamlit as st
import pandas as pd
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
import google.generativeai as genai
import json
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
st.write("Generate optimized, law-compliant rosters using Google AI Studio's Gemini models.")

# --- SIDEBAR CONFIGURATION ---
st.sidebar.image("https://img.icons8.com/color/96/google-logo.png", width=80)
st.sidebar.title("Google AI Studio Setup")

api_key = st.sidebar.text_input("Enter Gemini API Key", type="password", help="Get your API key from Google AI Studio")
model_option = st.sidebar.selectbox(
    "Select Gemini Model",
    ["gemini-3.6-flash", "gemini-1.5-flash", "gemini-1.5-pro"],
    help="Pro/newer models are recommended for complex logic, while Flash models are faster and have higher free-tier quotas."
)

st.sidebar.markdown("---")
st.sidebar.markdown("### How to Get an API Key:")
st.sidebar.markdown("1. Go to [Google AI Studio](https://aistudio.google.com/)\n2. Click **Get API Key**\n3. Create key in a new or existing project\n4. Paste it here!")

# Initialize Session States for manual inputs if they don't exist
if 'manual_employees' not in st.session_state:
    st.session_state.manual_employees = pd.DataFrame([
        {"Name": "Elizabeth", "Role": "Senior Team Member", "Age": 28, "Employment Type": "Part-Time", "Start Date": "2024-01-01"},
        {"Name": "Stella", "Role": "Junior Team Member", "Age": 18, "Employment Type": "Casual", "Start Date": "2024-03-15"},
        {"Name": "Ainsley", "Role": "Junior Team Member", "Age": 19, "Employment Type": "Casual", "Start Date": "2024-05-10"},
        {"Name": "Aimi", "Role": "Junior Team Member", "Age": 20, "Employment Type": "Casual", "Start Date": "2024-06-01"},
    ])

if 'manual_unavailability' not in st.session_state:
    st.session_state.manual_unavailability = pd.DataFrame([
        {"Employee": "Elizabeth", "Day": "Saturday", "Time Window": "All Day"},
        {"Employee": "Elizabeth", "Day": "Sunday", "Time Window": "All Day"},
        {"Employee": "Stella", "Day": "Monday", "Time Window": "Before 3:30pm"},
    ])

if 'manual_requirements' not in st.session_state:
    st.session_state.manual_requirements = pd.DataFrame([
        {"Day": "Monday", "Shift": "7:30am-12:30pm", "Count Required": 2},
        {"Day": "Monday", "Shift": "12:30pm-5:30pm", "Count Required": 1},
        {"Day": "Tuesday", "Shift": "7:30am-12:30pm", "Count Required": 2},
    ])

if 'manual_fixed' not in st.session_state:
    st.session_state.manual_fixed = pd.DataFrame([
        {"Employee": "Elizabeth", "Monday": "7:30am-12:30pm", "Tuesday": "off", "Wednesday": "off", "Thursday": "off", "Friday": "off", "Saturday": "off", "Sunday": "off"},
    ])

# --- TABS FOR INPUT ---
tab1, tab2, tab3, tab4, tab5 = st.tabs([
    "👥 Employees", 
    "🚫 Unavailability", 
    "📋 Daily Requirements", 
    "📌 Fixed Shifts", 
    "⚙️ Generate & Export"
])

# Helper function to check/load uploaded files or fallback to manual data
def get_data(uploaded_file, manual_df):
    if uploaded_file is not None:
        try:
            return pd.read_excel(uploaded_file)
        except Exception as e:
            st.error(f"Error reading file: {e}")
            return manual_df
    return manual_df

# Tab 1: Employees
with tab1:
    st.subheader("Manage Employees")
    upload_emp = st.file_uploader("Upload EMPLOYEE LIST.xlsx (Optional)", type=["xlsx"], key="emp_upload")
    
    st.markdown("### Edit Employee Data Manually")
    employees_df = st.data_editor(
        st.session_state.manual_employees,
        num_rows="dynamic",
        key="edit_employees"
    )
    if upload_emp:
        emp_loaded = pd.read_excel(upload_emp)
        st.dataframe(emp_loaded)

# Tab 2: Unavailability
with tab2:
    st.subheader("Log Unavailability")
    upload_unavail = st.file_uploader("Upload unavailability list.xlsx (Optional)", type=["xlsx"], key="unavail_upload")
    
    st.markdown("### Edit Unavailability Data Manually")
    unavailability_df = st.data_editor(
        st.session_state.manual_unavailability,
        num_rows="dynamic",
        key="edit_unavailability"
    )
    if upload_unavail:
        unavail_loaded = pd.read_excel(upload_unavail)
        st.dataframe(unavail_loaded)

# Tab 3: Daily Requirements
with tab3:
    st.subheader("Daily Shift Requirements")
    upload_req = st.file_uploader("Upload Daily Shift personel requirement.xlsx (Optional)", type=["xlsx"], key="req_upload")
    
    st.markdown("### Edit Shift Requirements Manually")
    requirements_df = st.data_editor(
        st.session_state.manual_requirements,
        num_rows="dynamic",
        key="edit_requirements"
    )
    if upload_req:
        req_loaded = pd.read_excel(upload_req)
        st.dataframe(req_loaded)

# Tab 4: Fixed Shifts
with tab4:
    st.subheader("Fixed Baseline Shifts")
    upload_fixed = st.file_uploader("Upload Roster fixed - dont change.xlsx (Optional)", type=["xlsx"], key="fixed_upload")
    
    st.markdown("### Edit Fixed Shifts Manually")
    fixed_df = st.data_editor(
        st.session_state.manual_fixed,
        num_rows="dynamic",
        key="edit_fixed"
    )
    if upload_fixed:
        fixed_loaded = pd.read_excel(upload_fixed)
        st.dataframe(fixed_loaded)

# Tab 5: Generate & Export
with tab5:
    st.subheader("Roster Configurations & Generation")
    
    col1, col2 = st.columns(2)
    with col1:
        start_date = st.date_input("Roster Start Date (Monday)", datetime.now() + timedelta(days=(0 - datetime.now().weekday())))
        upload_template = st.file_uploader("Upload Roster Layout Template (Optional)", type=["xlsx"], key="template_upload")
    
    with col2:
        st.info("Ensure you have input your Google AI Studio API Key in the sidebar before proceeding.")
    
    if st.button("Generate Weekly Roster"):
        if not api_key:
            st.error("Please enter a valid Gemini API Key in the sidebar.")
        else:
            with st.spinner("Analyzing files and generating roster using Google AI Studio Gemini API..."):
                try:
                    # Configure API
                    genai.configure(api_key=api_key)
                    
                    # Prepare DataFrames to send
                    final_employees = get_data(upload_emp, employees_df)
                    final_unavail = get_data(upload_unavail, unavailability_df)
                    final_req = get_data(upload_req, requirements_df)
                    final_fixed = get_data(upload_fixed, fixed_df)
                    
                    # Construct Prompt for Gemini
                    prompt = f"""
                    You are an expert workforce scheduler. You must generate a weekly roster starting on {start_date.strftime('%Y-%m-%d')} (Monday).
                    
                    ### INPUT DATA:
                    1. Employees List:
                    {final_employees.to_markdown(index=False)}
                    
                    2. Unavailability List:
                    {final_unavail.to_markdown(index=False)}
                    
                    3. Daily Shift Requirements:
                    {final_req.to_markdown(index=False)}
                    
                    4. Fixed baseline shifts:
                    {final_fixed.to_markdown(index=False)}
                    
                    ### ROSTER CONSTRAINTS & RULES:
                    - Only roster active employees whose Start Date is before or equal to {start_date.strftime('%Y-%m-%d')}.
                    - **Elizabeth**: Max 2 weekday shifts, no weekend shifts (Saturday and Sunday must be 'off' or ' unavailable'). Prefer early morning shifts starting at 7:00 AM or 7:30 AM.
                    - **School Students (Under 18)**: Cannot work weekday (Mon-Fri) shifts during school hours (9:00 AM - 3:30 PM).
                    - **Cost-Optimization**: Sort daily shifts by duration descending. Prioritize junior staff (under 21) for longer shifts first.
                    - **Workweek Limits**: Max 5 workdays per employee per week. Distribute shifts evenly.
                    - **Fixed Shifts**: Keep baseline fixed shifts as specified. Treat any cells with "off" as empty/schedulable.
                    - **Unavailability**: Mark cells as " unavailable" if an employee is unavailable for the day/shift.
                    
                    ### OUTPUT FORMAT:
                    You must return a JSON object with a single key "roster" mapping to an array of employee schedules.
                    Do not return any markdown wraps (like ```json), just the raw JSON text.
                    Each employee schedule must contain:
                    - "Employee": Name of the employee.
                    - "Monday": Shift description or "off" or " unavailable"
                    - "Tuesday": Shift description or "off" or " unavailable"
                    - "Wednesday": Shift description or "off" or " unavailable"
                    - "Thursday": Shift description or "off" or " unavailable"
                    - "Friday": Shift description or "off" or " unavailable"
                    - "Saturday": Shift description or "off" or " unavailable"
                    - "Sunday": Shift description or "off" or " unavailable"
                    
                    Example Output Structure:
                    {{
                        "roster": [
                            {{
                                "Employee": "Elizabeth",
                                "Monday": "7:30am-12:30pm",
                                "Tuesday": "off",
                                "Wednesday": "7:30am-12:30pm",
                                "Thursday": "off",
                                "Friday": "off",
                                "Saturday": " unavailable",
                                "Sunday": " unavailable"
                            }}
                        ]
                    }}
                    """
                    
                    # Call Gemini
                    model = genai.GenerativeModel(model_option)
                    response = model.generate_content(prompt)
                    
                    # Parse JSON Output
                    raw_text = response.text.strip()
                    if raw_text.startswith("```json"):
                        raw_text = raw_text[7:]
                    if raw_text.endswith("```"):
                        raw_text = raw_text[:-3]
                    raw_text = raw_text.strip()
                    
                    roster_data = json.loads(raw_text)
                    roster_out_df = pd.DataFrame(roster_data["roster"])
                    
                    st.success("Roster successfully generated!")
                    st.dataframe(roster_out_df)
                    
                    # Create downloadable Excel workbook
                    # Use uploaded template or create new
                    wb = None
                    if upload_template is not None:
                        wb = openpyxl.load_workbook(upload_template)
                    else:
                        wb = openpyxl.Workbook()
                        
                    ws = wb.active
                    ws.title = f"Roster {start_date.strftime('%d.%m.%Y')}"
                    
                    # Write headers
                    headers = ["Employee", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"]
                    ws.append(headers)
                    
                    # Style headers
                    header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
                    header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
                    thin_border = Border(
                        left=Side(style='thin', color='BFBFBF'),
                        right=Side(style='thin', color='BFBFBF'),
                        top=Side(style='thin', color='BFBFBF'),
                        bottom=Side(style='thin', color='BFBFBF')
                    )
                    
                    for col_num, header in enumerate(headers, 1):
                        cell = ws.cell(row=1, column=col_num)
                        cell.fill = header_fill
                        cell.font = header_font
                        cell.alignment = Alignment(horizontal="center", vertical="center")
                        cell.border = thin_border
                    
                    # Write rows and apply styles
                    unavail_fill = PatternFill(start_color="FCE4D6", end_color="FCE4D6", fill_type="solid")
                    off_fill = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")
                    text_font = Font(name="Calibri", size=11)
                    
                    from openpyxl.cell.cell import MergedCell
                    
                    for row_idx, row_data in enumerate(roster_out_df.itertuples(index=False), 2):
                        for col_idx, value in enumerate(row_data, 1):
                            cell = ws.cell(row=row_idx, column=col_idx)
                            
                            # If it is a merged cell, write to the top-left cell of the merged range
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
                            
                            # Color coding for readability
                            val_str = str(value).strip().lower()
                            if "unavailable" in val_str:
                                target_cell.fill = unavail_fill
                            elif val_str == "off":
                                target_cell.fill = off_fill
                    
                    # Save to byte stream for Streamlit download
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
                    st.error(f"Failed to generate roster or parse Gemini response: {e}")
                    if 'response' in locals() and response:
                        st.text_area("Raw API Response for Debugging", response.text)
