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

    /* TAB BAR CUSTOMIZATION - ULTRA-WIDE ROUNDED PILL BUTTONS */
    .stTabs [data-baseweb="tab-list"] {
        gap: 16px !important;
        background-color: rgba(9, 32, 28, 0.85) !important;
        padding: 14px 20px !important;
        border-radius: 40px !important;
        border: 1px solid rgba(229, 169, 60, 0.35) !important;
    }
    .stTabs button[role="tab"],
    .stTabs [data-baseweb="tab"] {
        height: 54px !important;
        padding-left: 55px !important;
        padding-right: 55px !important;
        border-radius: 30px !important;
        font-size: 1.05rem !important;
        transition: all 0.25s ease !important;
        margin: 0 4px !important;
    }
    .stTabs button[role="tab"] p,
    .stTabs [data-baseweb="tab"] p,
    .stTabs button[role="tab"] span,
    .stTabs [data-baseweb="tab"] span {
        padding-left: 12px !important;
        padding-right: 12px !important;
        white-space: nowrap !important;
    }
    .stTabs button[role="tab"] *,
    .stTabs [data-baseweb="tab"] *,
    .stTabs [aria-selected="false"] * {
        color: #ffffff !important;
        opacity: 1 !important;
        font-weight: 700 !important;
    }
    .stTabs [aria-selected="false"] {
        background-color: rgba(255, 255, 255, 0.12) !important;
        border: 1px solid transparent !important;
        border-radius: 30px !important;
    }
    .stTabs [data-baseweb="tab"]:hover {
        background-color: rgba(229, 169, 60, 0.3) !important;
    }
    .stTabs [aria-selected="true"] {
        background: linear-gradient(135deg, #e5a93c 0%, #d48827 100%) !important;
        border: none !important;
        border-radius: 30px !important;
        box-shadow: 0 6px 22px rgba(229, 169, 60, 0.55) !important;
    }
    .stTabs [aria-selected="true"] * {
        color: #0e2b26 !important;
        font-weight: 900 !important;
        opacity: 1 !important;
    }
    .stTabs [data-baseweb="tab-highlight-container"] {
        display: none !important;
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
        padding: 12px !important;
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

    /* FILE UPLOADER BUTTON SPECIFIC - CLEAN SINGLE GOLDEN BUTTON NO NESTING */
    [data-testid="stFileUploaderDropzone"] button {
        background-color: #e5a93c !important;
        background: linear-gradient(135deg, #e5a93c 0%, #d48827 100%) !important;
        color: #0e2b26 !important;
        border: none !important;
        border-radius: 8px !important;
        box-shadow: none !important;
        padding: 8px 18px !important;
    }
    [data-testid="stFileUploaderDropzone"] button * {
        background: transparent !important;
        color: #0e2b26 !important;
        border: none !important;
        box-shadow: none !important;
        font-weight: 800 !important;
    }

    /* Hero Generate Button Styling - Tight Under Date Picker */
    .hero-generate-btn {
        margin-top: 5px !important;
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

import json

# --- DISK PERSISTENCE ENGINE ---
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DATA_DIR = os.path.join(BASE_DIR, "data")
os.makedirs(DATA_DIR, exist_ok=True)

USER_PROFILES_FILE = os.path.join(DATA_DIR, "user_profiles.json")

DEFAULT_PROFILES = {
  "admin": {
    "username": "admin",
    "password": "admin123",
    "role": "Manager",
    "employee_name": "Bakery Manager",
    "profile": {
      "full_name": "Bakery Manager", "address": "", "home_phone": "", "mobile": "",
      "email": "manager@brumbys.com.au", "dob": "", "gender": "", "tfn": "",
      "store": "Brumby's Pakenham", "classification": "Full-Time", "commencement_date": "",
      "employment_level": "Store Manager", "super_fund": "", "super_policy": "",
      "super_address": "", "super_contact": "", "super_abn": "", "bank_name": "",
      "bank_branch": "", "bank_bsb": "", "bank_account": "", "account_name": ""
    }
  },
  "ainsley.mactier": {
    "username": "ainsley.mactier",
    "password": "TempPass123!",
    "role": "Employee",
    "employee_name": "Ainsley Mactier",
    "profile": {
      "full_name": "Ainsley Brenda Mactier",
      "address": "8 Knapton Ave, Beaconsfield Upper, Vic 3808",
      "home_phone": "0359192106",
      "mobile": "0479122444",
      "email": "ainsley.mac@outlook.com",
      "dob": "14th August 2006",
      "gender": "Female",
      "tfn": "520700",
      "store": "Brumby's Pakenham",
      "classification": "Casual",
      "commencement_date": "2024-05-10",
      "employment_level": "Junior Team Member",
      "super_fund": "Australian Super",
      "super_policy": "9124950",
      "super_address": "",
      "super_contact": "",
      "super_abn": "",
      "bank_name": "Commonwealth",
      "bank_branch": "",
      "bank_bsb": "062 948",
      "bank_account": "2847 7286",
      "account_name": "Ainsley Mactier"
    }
  },
  "elizabeth": {
    "username": "elizabeth", "password": "TempPass123!", "role": "Employee", "employee_name": "Elizabeth",
    "profile": { "full_name": "Elizabeth", "address": "", "home_phone": "", "mobile": "", "email": "", "dob": "", "gender": "Female", "tfn": "", "store": "Brumby's Pakenham", "classification": "Part-Time", "commencement_date": "2024-01-01", "employment_level": "Senior Team Member", "super_fund": "", "super_policy": "", "super_address": "", "super_contact": "", "super_abn": "", "bank_name": "", "bank_branch": "", "bank_bsb": "", "bank_account": "", "account_name": "" }
  },
  "stella": {
    "username": "stella", "password": "TempPass123!", "role": "Employee", "employee_name": "Stella",
    "profile": { "full_name": "Stella", "address": "", "home_phone": "", "mobile": "", "email": "", "dob": "", "gender": "Female", "tfn": "", "store": "Brumby's Pakenham", "classification": "Casual", "commencement_date": "2024-03-15", "employment_level": "Junior Team Member", "super_fund": "", "super_policy": "", "super_address": "", "super_contact": "", "super_abn": "", "bank_name": "", "bank_branch": "", "bank_bsb": "", "bank_account": "", "account_name": "" }
  },
  "aimi": {
    "username": "aimi", "password": "TempPass123!", "role": "Employee", "employee_name": "Aimi",
    "profile": { "full_name": "Aimi", "address": "", "home_phone": "", "mobile": "", "email": "", "dob": "", "gender": "Female", "tfn": "", "store": "Brumby's Pakenham", "classification": "Casual", "commencement_date": "2024-06-01", "employment_level": "Junior Team Member", "super_fund": "", "super_policy": "", "super_address": "", "super_contact": "", "super_abn": "", "bank_name": "", "bank_branch": "", "bank_bsb": "", "bank_account": "", "account_name": "" }
  },
  "jude": {
    "username": "jude", "password": "TempPass123!", "role": "Employee", "employee_name": "Jude",
    "profile": { "full_name": "Jude", "address": "", "home_phone": "", "mobile": "", "email": "", "dob": "", "gender": "Male", "tfn": "", "store": "Brumby's Pakenham", "classification": "Full-Time", "commencement_date": "2024-01-01", "employment_level": "Senior Team Member", "super_fund": "", "super_policy": "", "super_address": "", "super_contact": "", "super_abn": "", "bank_name": "", "bank_branch": "", "bank_bsb": "", "bank_account": "", "account_name": "" }
  },
  "aroha": {
    "username": "aroha", "password": "TempPass123!", "role": "Employee", "employee_name": "Aroha",
    "profile": { "full_name": "Aroha", "address": "", "home_phone": "", "mobile": "", "email": "", "dob": "", "gender": "Female", "tfn": "", "store": "Brumby's Pakenham", "classification": "Full-Time", "commencement_date": "2023-11-01", "employment_level": "Senior Team Member", "super_fund": "", "super_policy": "", "super_address": "", "super_contact": "", "super_abn": "", "bank_name": "", "bank_branch": "", "bank_bsb": "", "bank_account": "", "account_name": "" }
  },
  "robert": {
    "username": "robert", "password": "TempPass123!", "role": "Employee", "employee_name": "Robert",
    "profile": { "full_name": "Robert", "address": "", "home_phone": "", "mobile": "", "email": "", "dob": "", "gender": "Male", "tfn": "", "store": "Brumby's Pakenham", "classification": "Full-Time", "commencement_date": "2023-10-01", "employment_level": "Senior Team Member", "super_fund": "", "super_policy": "", "super_address": "", "super_contact": "", "super_abn": "", "bank_name": "", "bank_branch": "", "bank_bsb": "", "bank_account": "", "account_name": "" }
  }
}

def load_user_profiles():
    profiles = {}
    if os.path.exists(USER_PROFILES_FILE):
        try:
            with open(USER_PROFILES_FILE, "r", encoding="utf-8") as f:
                profiles = json.load(f)
        except Exception:
            profiles = {}
            
    if not profiles or "admin" not in profiles:
        profiles = DEFAULT_PROFILES.copy()
        save_user_profiles(profiles)
        
    return profiles

def save_user_profiles(profiles):
    try:
        with open(USER_PROFILES_FILE, "w", encoding="utf-8") as f:
            json.dump(profiles, f, indent=2)
    except Exception as e:
        st.error(f"Error saving user profiles: {e}")

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

# Initialize Session States with Disk Cache
if 'authenticated' not in st.session_state:
    st.session_state.authenticated = False
if 'logged_in_user' not in st.session_state:
    st.session_state.logged_in_user = None
if 'user_role' not in st.session_state:
    st.session_state.user_role = None

user_profiles = load_user_profiles()

# LOGIN PAGE IF NOT AUTHENTICATED
if not st.session_state.authenticated:
    st.markdown('<h1 class="header-style">🥐 Brumby\'s Pakenham — Portal</h1>', unsafe_allow_html=True)
    st.markdown('<p class="sub-header-style">Please log in with your username and password to access your bakery account.</p>', unsafe_allow_html=True)
    
    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        st.markdown("""
        <div style="background: linear-gradient(135deg, #081d19 0%, #16443c 100%); padding: 30px; border-radius: 20px; border: 2px solid #e5a93c; box-shadow: 0 10px 30px rgba(0,0,0,0.5);">
            <h2 style="color: #e5a93c !important; margin-top: 0; text-align: center;">🔐 Staff Portal Login</h2>
        """, unsafe_allow_html=True)
        
        login_user = st.text_input("Username", key="input_user")
        login_pass = st.text_input("Password", type="password", key="input_pass")
        
        if st.button("🚀 Login to Portal", key="btn_login"):
            login_user_clean = login_user.strip().lower()
            current_profiles = load_user_profiles()
            if login_user_clean in current_profiles:
                account = current_profiles[login_user_clean]
                if account.get("password") == login_pass:
                    st.session_state.authenticated = True
                    st.session_state.logged_in_user = login_user_clean
                    st.session_state.user_role = account.get("role", "Employee")
                    st.success(f"Welcome back, {account.get('employee_name', login_user_clean)}!")
                    st.rerun()
                else:
                    st.error("Incorrect password. Please check your credentials.")
            else:
                st.error(f"Username '{login_user_clean}' not found. Available accounts: {', '.join(current_profiles.keys())}")
                
        st.markdown("""
            <hr style="border-color: rgba(229, 169, 60, 0.3); margin: 20px 0;">
            <p style="font-size: 0.88rem; color: #c8e6e0 !important; text-align: center;">
                <b>Default Credentials:</b><br>
                👑 Manager: <code>admin</code> / <code>admin123</code><br>
                👤 Employee (e.g. Ainsley Mactier): <code>ainsley.mactier</code> / <code>TempPass123!</code>
            </p>
        </div>
        """, unsafe_allow_html=True)
    st.stop()

# --- SIDEBAR CONFIGURATION (AUTHENTICATED) ---
st.sidebar.image("https://img.icons8.com/fluency/96/bakery.png", width=80)
st.sidebar.title("🍞 Bakery Portal Controls")

curr_user_key = st.session_state.logged_in_user
curr_user_info = user_profiles.get(curr_user_key, {})
display_name = curr_user_info.get("employee_name", curr_user_key)
role_title = curr_user_info.get("role", "Employee")

st.sidebar.markdown(f"""
<div style="background-color: rgba(229, 169, 60, 0.15); padding: 12px 16px; border-radius: 12px; border: 1px solid #e5a93c; margin-bottom: 15px;">
    <div style="color: #e5a93c; font-size: 0.85rem; font-weight: 700;">LOGGED IN ACCOUNT</div>
    <div style="color: #ffffff; font-size: 1.1rem; font-weight: 900;">👤 {display_name}</div>
    <div style="color: #c8e6e0; font-size: 0.85rem;">Role: <b>{role_title}</b></div>
</div>
""", unsafe_allow_html=True)

if st.sidebar.button("🚪 Logout", key="btn_logout"):
    st.session_state.authenticated = False
    st.session_state.logged_in_user = None
    st.session_state.user_role = None
    st.rerun()

st.sidebar.info("This application runs locally and optimizes staff rostering and confidential profile management.")

col_head1, col_head2 = st.columns([3.5, 1])
with col_head1:
    st.markdown('<h1 class="header-style">🥐 Brumby\'s Pakenham Portal</h1>', unsafe_allow_html=True)
    st.markdown(f'<p class="sub-header-style">Welcome back, <b>{display_name}</b> ({role_title})</p>', unsafe_allow_html=True)
with col_head2:
    st.markdown("<div style='margin-top: 15px;'></div>", unsafe_allow_html=True)
    if st.button("🚪 Logout", key="btn_logout_header", use_container_width=True):
        st.session_state.authenticated = False
        st.session_state.logged_in_user = None
        st.session_state.user_role = None
        st.rerun()

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

# --- ROLE-BASED TAB NAVIGATION ---
is_manager = (st.session_state.user_role == "Manager")

if is_manager:
    tab_home, tab_emp, tab_unavail, tab_req, tab_fixed = st.tabs([
        "🏠 Home / Roster Generator",
        "👥 Staff Members", 
        "🚫 Unavailability", 
        "📋 Daily Requirements", 
        "📌 Fixed Shifts"
    ])
else:
    # Employee sees only their personal information form tab
    tab_my_info, = st.tabs(["📋 My Personal Information Form"])

# Helper function to render Confidential Profile Form
def render_confidential_profile_form(user_key, is_admin=False):
    user_data = user_profiles.get(user_key, {})
    prof = user_data.get("profile", {})
    emp_name = user_data.get("employee_name", user_key)
    
    st.markdown(f"""
    <div style="background: linear-gradient(135deg, #081d19 0%, #16443c 100%); padding: 16px 24px; border-radius: 14px; border: 2px solid #e5a93c; margin-bottom: 20px;">
        <h2 style="color: #e5a93c !important; margin: 0;">📜 Employee Confidential Information Form</h2>
        <p style="color: #c8e6e0 !important; margin-top: 4px; margin-bottom: 0;">Employee: <b>{emp_name}</b> &nbsp;|&nbsp; Store: <b>{prof.get('store', "Brumby's Pakenham")}</b></p>
    </div>
    """, unsafe_allow_html=True)
    
    with st.form(key=f"form_profile_{user_key}"):
        st.markdown("### 1. 👤 Employee Personal Details")
        c1, c2 = st.columns(2)
        with c1:
            full_name = st.text_input("Full Name", value=str(prof.get("full_name", emp_name)))
            address = st.text_area("Home Address", value=str(prof.get("address", "")), height=100)
            home_phone = st.text_input("Home Phone Number", value=str(prof.get("home_phone", "")))
            mobile = st.text_input("Mobile Phone Number", value=str(prof.get("mobile", "")))
            email = st.text_input("Email Address", value=str(prof.get("email", "")))
            dob = st.text_input("Date of Birth", value=str(prof.get("dob", "")))
        with c2:
            gender_options = ["Female", "Male", "Other", "Prefer not to say"]
            curr_gender = str(prof.get("gender", "Female"))
            g_idx = gender_options.index(curr_gender) if curr_gender in gender_options else 0
            gender = st.selectbox("Gender", gender_options, index=g_idx)
            
            tfn = st.text_input("Tax File Number (TFN)", value=str(prof.get("tfn", "")))
            store = st.text_input("Store Location", value=str(prof.get("store", "Brumby's Pakenham")))
            
            class_options = ["Casual", "Part-Time", "Full-Time"]
            curr_class = str(prof.get("classification", "Casual"))
            c_idx = class_options.index(curr_class) if curr_class in class_options else 0
            classification = st.selectbox("Employment Classification", class_options, index=c_idx)
            
            commencement_date = st.text_input("Commencement Date", value=str(prof.get("commencement_date", "")))
            employment_level = st.text_input("Employment Level / Role", value=str(prof.get("employment_level", "")))

        st.markdown("---")
        st.markdown("### 2. 🏦 Bank Account Details")
        b1, b2 = st.columns(2)
        with b1:
            bank_name = st.text_input("Bank Name", value=str(prof.get("bank_name", "")))
            bank_branch = st.text_input("Branch Location", value=str(prof.get("bank_branch", "")))
            account_name = st.text_input("Name on Account", value=str(prof.get("account_name", "")))
        with b2:
            bank_bsb = st.text_input("BSB Number (e.g. 062 948)", value=str(prof.get("bank_bsb", "")))
            bank_account = st.text_input("Account Number", value=str(prof.get("bank_account", "")))

        st.markdown("---")
        st.markdown("### 3. 📈 Superannuation Details")
        s1, s2 = st.columns(2)
        with s1:
            super_fund = st.text_input("Name of Super Fund", value=str(prof.get("super_fund", "")))
            super_policy = st.text_input("Policy / Membership Number", value=str(prof.get("super_policy", "")))
            super_address = st.text_input("Super Fund Address", value=str(prof.get("super_address", "")))
        with s2:
            super_contact = st.text_input("Fund Contact Number", value=str(prof.get("super_contact", "")))
            super_abn = st.text_input("Fund ABN / SPIN / USI", value=str(prof.get("super_abn", "")))

        submit_btn = st.form_submit_button("💾 Save Confidential Profile Information")
        
        if submit_btn:
            user_profiles[user_key]["profile"] = {
                "full_name": full_name,
                "address": address,
                "home_phone": home_phone,
                "mobile": mobile,
                "email": email,
                "dob": dob,
                "gender": gender,
                "tfn": tfn,
                "store": store,
                "classification": classification,
                "commencement_date": commencement_date,
                "employment_level": employment_level,
                "bank_name": bank_name,
                "bank_branch": bank_branch,
                "account_name": account_name,
                "bank_bsb": bank_bsb,
                "bank_account": bank_account,
                "super_fund": super_fund,
                "super_policy": super_policy,
                "super_address": super_address,
                "super_contact": super_contact,
                "super_abn": super_abn
            }
            save_user_profiles(user_profiles)
            st.success("✅ Profile information updated and saved successfully!")

# Helper function for Employee Availability Management & Auto-Sync
def render_employee_availability_manager(user_key):
    user_data = user_profiles.get(user_key, {})
    emp_name = user_data.get("employee_name", user_key)
    
    st.markdown(f"""
    <div style="background: linear-gradient(135deg, #081d19 0%, #16443c 100%); padding: 16px 24px; border-radius: 14px; border: 2px solid #e5a93c; margin-bottom: 20px;">
        <h2 style="color: #e5a93c !important; margin: 0;">📅 My Unavailability Calendar & Shift Constraints</h2>
        <p style="color: #c8e6e0 !important; margin-top: 4px; margin-bottom: 0;">Logged in Employee: <b>{emp_name}</b>. Add or update your weekly availability constraints below. Entries auto-sync directly to the master bakery roster.</p>
    </div>
    """, unsafe_allow_html=True)
    
    # 1. Add Constraint Controls
    with st.form(key=f"form_add_unavail_{user_key}"):
        st.markdown("### ➕ Log New Unavailability Constraint")
        col1, col2, col3 = st.columns([1, 1.2, 1.5])
        with col1:
            sel_day = st.selectbox("Select Day of Week", ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"], key=f"sel_day_{user_key}")
        with col2:
            preset_window = st.selectbox("Constraint / Window", ["All Day", "Before 3:30pm", "After 5:00pm", "Custom Shift Window"], key=f"sel_window_{user_key}")
        with col3:
            custom_window = st.text_input("Custom Window Details (if selected)", value="7:30am-12:30pm", key=f"input_custom_win_{user_key}")
            
        submit_unavail = st.form_submit_button("📌 Add Unavailability Constraint")
        
        if submit_unavail:
            target_window = custom_window.strip() if preset_window == "Custom Shift Window" else preset_window
            if target_window:
                df_curr = st.session_state.manual_unavailability.copy()
                emp_col = find_column(df_curr, ["employee", "name", "staff"], "Employee")
                day_col = find_column(df_curr, ["day", "date", "weekday"], "Day")
                win_col = find_column(df_curr, ["time window", "window", "time", "unavailability"], "Time Window")
                
                new_entry = {emp_col: emp_name, day_col: sel_day, win_col: target_window}
                df_updated = pd.concat([df_curr, pd.DataFrame([new_entry])], ignore_index=True)
                st.session_state.manual_unavailability = df_updated
                save_persisted_df(df_updated, "unavailability.csv")
                st.success(f"✅ Added unavailability constraint for {sel_day}: {target_window}")

    st.markdown("---")
    st.markdown("### 📋 My Active Logged Constraints")
    df_curr = st.session_state.manual_unavailability.copy()
    emp_col = find_column(df_curr, ["employee", "name", "staff"], "Employee")
    day_col = find_column(df_curr, ["day", "date", "weekday"], "Day")
    win_col = find_column(df_curr, ["time window", "window", "time", "unavailability"], "Time Window")
    
    # Match employee name robustly
    mask = df_curr[emp_col].astype(str).str.strip().str.lower() == emp_name.lower()
    user_rows = df_curr[mask].copy()
    
    if user_rows.empty:
        st.info("🎉 You have no logged unavailability constraints. You are listed as Available all week!")
    else:
        for idx, row in user_rows.iterrows():
            day_val = row.get(day_col, "")
            win_val = row.get(win_col, "")
            c1, c2, c3 = st.columns([2, 3, 1])
            with c1:
                st.markdown(f"🗓️ **{day_val}**")
            with c2:
                badge_bg = "#9b2c2c" if "all day" in str(win_val).lower() else "#b7791f"
                st.markdown(f'<span style="background-color: {badge_bg}; color: white; padding: 4px 14px; border-radius: 12px; font-weight: 700;">{win_val}</span>', unsafe_allow_html=True)
            with c3:
                if st.button("🗑️ Delete", key=f"btn_delete_unavail_{idx}"):
                    st.session_state.manual_unavailability = df_curr.drop(idx).reset_index(drop=True)
                    save_persisted_df(st.session_state.manual_unavailability, "unavailability.csv")
                    st.success("Constraint deleted.")
                    st.rerun()

# Helper function to render Whole Team Unavailability Calendar Matrix for Manager
def render_team_unavailability_matrix():
    st.markdown("""
    <div style="background: linear-gradient(135deg, #081d19 0%, #16443c 100%); padding: 12px 20px; border-radius: 12px 12px 0 0; color: #ffffff !important; font-weight: 800; font-size: 1.15rem; letter-spacing: 0.3px; border: 2px solid #e5a93c; border-bottom: none; margin-top: 15px;">
        📅 Whole Team Weekly Unavailability Matrix (Real-Time Live Sync Calendar)
    </div>
    """, unsafe_allow_html=True)
    
    unavail_df = st.session_state.manual_unavailability
    emp_col = find_column(unavail_df, ["employee", "name", "staff"], "Employee")
    day_col = find_column(unavail_df, ["day", "date", "weekday"], "Day")
    win_col = find_column(unavail_df, ["time window", "window", "time", "unavailability"], "Time Window")
    
    days = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"]
    
    emp_list = []
    if "manual_employees" in st.session_state and "Name" in st.session_state.manual_employees.columns:
        emp_list = [str(x).strip() for x in st.session_state.manual_employees["Name"].dropna().unique() if str(x).strip()]
    if not emp_list:
        emp_list = ["Elizabeth", "Stella", "Ainsley Mactier", "Aimi", "Jude", "Aroha", "Robert"]
        
    matrix_rows = []
    for emp in emp_list:
        r = {"Employee": emp}
        for d in days:
            matches = unavail_df[(unavail_df[emp_col].astype(str).str.strip().str.lower() == emp.lower()) & 
                                 (unavail_df[day_col].astype(str).str.strip().str.lower().str.startswith(d.lower()[:3]))]
            if not matches.empty:
                r[d] = " | ".join(matches[win_col].astype(str).tolist())
            else:
                r[d] = "Available"
        matrix_rows.append(r)
        
    df_matrix = pd.DataFrame(matrix_rows)
    st.dataframe(df_matrix, use_container_width=True, hide_index=True)

# IF EMPLOYEE, RENDER PERSONAL INFORMATION FORM
if not is_manager:
    with tab_my_info:
        render_confidential_profile_form(st.session_state.logged_in_user)


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

if is_manager:
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
            st.markdown('</div>', unsafe_allow_html=True)

        with col2:
            st.markdown("""
            <div style="background: rgba(9, 32, 28, 0.5); border: 1px solid rgba(229, 169, 60, 0.4); border-radius: 14px; padding: 15px; height: 100%;">
                <h4 style="color: #e5a93c !important; margin-top: 0;">📋 Generator Rules Summary</h4>
                <ul style="margin-bottom: 0; padding-left: 20px; font-size: 0.95rem; color: #ffffff !important;">
                    <li>Respects staff unavailability constraints</li>
                    <li>Fulfills daily shift requirements</li>
                    <li>Ensures mandatory award break times</li>
                    <li>Enforces minimum rest periods between shifts</li>
                </ul>
            </div>
            """, unsafe_allow_html=True)

        if 'final_roster_df' in st.session_state and not st.session_state.final_roster_df.empty:
            st.markdown("<br>", unsafe_allow_html=True)
            st.markdown("""
            <div style="background: linear-gradient(135deg, #081d19 0%, #16443c 100%); padding: 12px 20px; border-radius: 12px 12px 0 0; color: #ffffff !important; font-weight: 800; font-size: 1.2rem; letter-spacing: 0.5px; border: 2px solid #e5a93c; border-bottom: none;">
                📅 Generated Weekly Roster Schedule (Editable)
            </div>
            """, unsafe_allow_html=True)
            
            edited_final_df = st.data_editor(st.session_state.final_roster_df, num_rows="dynamic", key="edit_generated_roster")
            st.session_state.final_roster_df = edited_final_df

            # Prepare Excel workbook in memory using openpyxl for exact styling match
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "Weekly Roster"

            # Header styling
            ws.merge_cells("A1:H1")
            title_cell = ws.cell(row=1, column=1)
            title_cell.value = "BRUMBY'S PAKENHAM - WEEKLY STAFF ROSTER"
            title_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
            title_cell.fill = title_fill
            title_cell.font = Font(name="Calibri", size=14, bold=True, color="FFFFFF")
            title_cell.alignment = Alignment(horizontal="center", vertical="center")

            ws.merge_cells("A2:H2")
            sub_cell = ws.cell(row=2, column=1)
            end_date = start_date + timedelta(days=6)
            sub_cell.value = f"Week Period: Monday {start_date.strftime('%d/%m/%Y')} to Sunday {end_date.strftime('%d/%m/%Y')}"
            sub_cell.fill = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
            sub_cell.font = Font(name="Calibri", size=11, bold=True, color="1F4E78")
            sub_cell.alignment = Alignment(horizontal="center", vertical="center")

            ws.row_dimensions[3].height = 10

            tbl_hdr_fill = PatternFill(start_color="203764", end_color="203764", fill_type="solid")
            tbl_hdr_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")

            for col_idx, col_name in enumerate(edited_final_df.columns, 1):
                c = ws.cell(row=4, column=col_idx)
                c.value = col_name
                c.fill = tbl_hdr_fill
                c.font = tbl_hdr_font
                c.alignment = Alignment(horizontal="center", vertical="center")
                    
            thin_border = Border(
                left=Side(style='thin', color='BFBFBF'),
                right=Side(style='thin', color='BFBFBF'),
                top=Side(style='thin', color='BFBFBF'),
                bottom=Side(style='thin', color='BFBFBF')
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
            👥 Bakery Staff Members List (Editable Table)
        </div>
        """, unsafe_allow_html=True)
        employees_df = st.data_editor(st.session_state.manual_employees, num_rows="dynamic", key="edit_employees")
        st.session_state.manual_employees = employees_df
        save_persisted_df(employees_df, "employees.csv")

        # --- INTEGRATED CONFIDENTIAL EMPLOYEE PROFILE VIEWER ---
        st.markdown("<br>", unsafe_allow_html=True)
        st.markdown("""
        <div style="background: linear-gradient(135deg, #081d19 0%, #16443c 100%); padding: 10px 18px; border-radius: 12px 12px 0 0; color: #ffffff !important; font-weight: 800; font-size: 1.1rem; letter-spacing: 0.3px; border: 2px solid #e5a93c; border-bottom: none; margin-top: 20px;">
            📜 Click / Select Employee to Open Confidential Profile Form
        </div>
        """, unsafe_allow_html=True)

        emp_options = {k: v.get("employee_name", k) for k, v in user_profiles.items() if v.get("role") == "Employee"}
        if emp_options:
            selected_user_key = st.selectbox(
                "👇 Select staff member from table above to open their confidential profile page:",
                options=list(emp_options.keys()),
                format_func=lambda x: f"👤 {emp_options[x]} (Username: {x})",
                key="select_emp_profile_in_tab"
            )
            
            if selected_user_key:
                render_confidential_profile_form(selected_user_key, is_admin=True)
                
                with st.expander(f"🔑 Reset Password for {emp_options[selected_user_key]}"):
                    new_pw = st.text_input("New Password", type="password", key=f"tab_reset_pw_{selected_user_key}")
                    if st.button("Update Employee Password", key=f"tab_btn_reset_pw_{selected_user_key}"):
                        if new_pw.strip():
                            user_profiles[selected_user_key]["password"] = new_pw.strip()
                            save_user_profiles(user_profiles)
                            st.success(f"✅ Password for {emp_options[selected_user_key]} updated successfully!")
                        else:
                            st.error("Password cannot be empty.")

    # --- TAB 3: UNAVAILABILITY ---
    with tab_unavail:
        st.subheader("Log Staff Unavailability")
        
        # Whole Team Unavailability Matrix (Real-Time Live Sync)
        render_team_unavailability_matrix()
        
        st.markdown("<br>", unsafe_allow_html=True)
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
