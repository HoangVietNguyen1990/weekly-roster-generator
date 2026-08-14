import streamlit as st
import pandas as pd
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
import io
import os
import copy
import json
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
    
    /* Code tag override for dark emerald/gold theme */
    code, .stMarkdown code {
        background-color: rgba(9, 38, 33, 0.9) !important;
        color: #f7d594 !important;
        border: 1px solid rgba(229, 169, 60, 0.4) !important;
        border-radius: 6px !important;
        padding: 2px 8px !important;
        font-family: inherit !important;
        font-weight: 700 !important;
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
    /* TACTILE 3D TABS STYLING */
    .stTabs [data-baseweb="tab-list"] {
        gap: 16px !important;
        background-color: rgba(6, 24, 20, 0.95) !important;
        padding: 16px 24px !important;
        border-radius: 40px !important;
        border: 2px solid rgba(229, 169, 60, 0.45) !important;
        box-shadow: 0 10px 30px rgba(0, 0, 0, 0.6), inset 0 2px 4px rgba(0,0,0,0.5) !important;
    }
    .stTabs button[role="tab"],
    .stTabs [data-baseweb="tab"] {
        height: 54px !important;
        padding-left: 36px !important;
        padding-right: 36px !important;
        border-radius: 28px !important;
        font-size: 1.05rem !important;
        transition: all 0.15s ease-in-out !important;
        margin: 0 4px !important;
        cursor: pointer !important;
    }
    .stTabs [aria-selected="false"] {
        background: linear-gradient(180deg, #1d574c 0%, #134038 50%, #0b2923 100%) !important;
        border: 2px solid #1f5c50 !important;
        box-shadow: 0 5px 0 #061814, 0 6px 14px rgba(0, 0, 0, 0.4), inset 0 1px 0 rgba(255, 255, 255, 0.2) !important;
    }
    .stTabs [aria-selected="false"]:hover {
        background: linear-gradient(180deg, #24685b 0%, #184c42 50%, #0d332c 100%) !important;
        border-color: #e5a93c !important;
        transform: translateY(-2px) !important;
        box-shadow: 0 7px 0 #061814, 0 8px 18px rgba(0, 0, 0, 0.5), inset 0 1px 0 rgba(255, 255, 255, 0.3) !important;
    }
    .stTabs [aria-selected="false"] * {
        color: #ffffff !important;
        font-weight: 700 !important;
        text-shadow: 0 1px 2px rgba(0,0,0,0.8) !important;
    }
    .stTabs [aria-selected="true"] {
        background: linear-gradient(180deg, #fce4b3 0%, #e5a93c 45%, #b87b1c 100%) !important;
        border: 2px solid #ffe8be !important;
        box-shadow: 0 6px 0 #734c0e, 0 10px 24px rgba(229, 169, 60, 0.6), inset 0 1px 0 rgba(255, 255, 255, 0.7) !important;
        transform: translateY(-2px) !important;
    }
    .stTabs [aria-selected="true"] * {
        color: #081d19 !important;
        font-weight: 900 !important;
        text-shadow: 0 1px 1px rgba(255, 255, 255, 0.4) !important;
    }

    /* TACTILE 3D BUTTONS STYLING (ALL BUTTONS & FORM SUBMIT BUTTONS) */
    .stButton > button,
    .stFormSubmitButton > button,
    button[data-testid="baseButton-secondary"],
    button[data-testid="baseButton-primary"] {
        background: linear-gradient(180deg, #fce4b3 0%, #e5a93c 45%, #b87b1c 100%) !important;
        border: 2px solid #ffe8be !important;
        border-radius: 14px !important;
        color: #081d19 !important;
        font-weight: 900 !important;
        font-size: 1.05rem !important;
        padding: 12px 24px !important;
        box-shadow: 0 5px 0 #734c0e, 0 8px 20px rgba(0, 0, 0, 0.4), inset 0 1px 0 rgba(255, 255, 255, 0.7) !important;
        transition: all 0.15s ease-in-out !important;
        cursor: pointer !important;
        text-shadow: 0 1px 1px rgba(255, 255, 255, 0.4) !important;
    }
    .stButton > button:hover,
    .stFormSubmitButton > button:hover {
        background: linear-gradient(180deg, #fff0d4 0%, #f0b548 45%, #c78822 100%) !important;
        box-shadow: 0 6px 0 #734c0e, 0 12px 24px rgba(229, 169, 60, 0.5), inset 0 1px 0 rgba(255, 255, 255, 0.8) !important;
        transform: translateY(-2px) !important;
    }
    .stButton > button:active,
    .stFormSubmitButton > button:active {
        box-shadow: 0 2px 0 #734c0e, 0 4px 10px rgba(0, 0, 0, 0.4), inset 0 1px 0 rgba(0, 0, 0, 0.2) !important;
        transform: translateY(3px) !important;
    }
    .stButton > button *,
    .stFormSubmitButton > button * {
        color: #081d19 !important;
        font-weight: 900 !important;
    }

    /* TACTILE 3D EXPANDER BUTTON STYLING */
    .stExpander, 
    div[data-testid="stExpander"], 
    details {
        background-color: transparent !important;
        border: none !important;
        box-shadow: none !important;
        margin-bottom: 14px !important;
    }
    .stExpander summary, 
    div[data-testid="stExpander"] summary,
    details summary {
        background: linear-gradient(180deg, #1d574c 0%, #144038 50%, #0b2923 100%) !important;
        border: 2px solid #e5a93c !important;
        border-radius: 12px !important;
        color: #f7d594 !important;
        padding: 12px 18px !important;
        box-shadow: 0 5px 0 #061814, 0 8px 20px rgba(0, 0, 0, 0.5), inset 0 1px 0 rgba(255, 255, 255, 0.25) !important;
        transition: all 0.15s ease-in-out !important;
        cursor: pointer !important;
    }
    .stExpander summary:hover,
    div[data-testid="stExpander"] summary:hover,
    details summary:hover {
        background: linear-gradient(180deg, #23665a 0%, #184c42 50%, #0e332c 100%) !important;
        border-color: #f7d594 !important;
        box-shadow: 0 6px 0 #061814, 0 10px 24px rgba(0, 0, 0, 0.6), inset 0 1px 0 rgba(255, 255, 255, 0.35) !important;
        transform: translateY(-1px) !important;
    }
    .stExpander summary:active,
    div[data-testid="stExpander"] summary:active,
    details summary:active {
        box-shadow: 0 2px 0 #061814, 0 4px 10px rgba(0, 0, 0, 0.4), inset 0 1px 0 rgba(0, 0, 0, 0.2) !important;
        transform: translateY(3px) !important;
    }
    .stExpander summary *, 
    div[data-testid="stExpander"] summary *, 
    div[data-testid="stExpander"] summary p, 
    div[data-testid="stExpander"] summary span, 
    details summary p, 
    details summary span,
    [data-testid="stExpanderToggleIcon"],
    svg[data-testid="stExpanderToggleIcon"] {
        color: #f7d594 !important;
        font-weight: 800 !important;
        font-size: 1.05rem !important;
        fill: #f7d594 !important;
        text-shadow: 0 1px 3px rgba(0, 0, 0, 0.8) !important;
    }
    div[data-testid="stExpanderDetails"],
    details[open] > div {
        background: #0f2e29 !important;
        border: 2px solid #e5a93c !important;
        border-top: none !important;
        border-radius: 0 0 12px 12px !important;
        padding: 16px !important;
        margin-top: -6px !important;
        box-shadow: 0 8px 20px rgba(0, 0, 0, 0.4) !important;
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

    /* GLOBAL INPUT & WIDGET TEXT CONTRAST ENHANCEMENTS */
    div[data-baseweb="input"] input,
    div[data-baseweb="base-input"] input,
    div[data-baseweb="textarea"] textarea,
    div[data-baseweb="select"] div,
    div[data-baseweb="select"] input,
    div[data-baseweb="select"] span,
    .stTextInput input, 
    .stNumberInput input, 
    .stDateInput input, 
    .stSelectbox div[role="combobox"], 
    .stTextArea textarea {
        background-color: #0c2b25 !important;
        color: #ffffff !important;
        -webkit-text-fill-color: #ffffff !important;
        border: 1.5px solid #e5a93c !important;
        font-weight: 700 !important;
        font-size: 1rem !important;
        opacity: 1 !important;
        border-radius: 8px !important;
    }

    /* ULTRA HIGH-CONTRAST PLACEHOLDER TEXT FOR ALL INPUT FIELDS */
    input::placeholder,
    textarea::placeholder,
    div[data-baseweb="input"] input::placeholder,
    div[data-baseweb="base-input"] input::placeholder,
    div[data-baseweb="textarea"] textarea::placeholder,
    .stTextInput input::placeholder,
    .stNumberInput input::placeholder,
    .stDateInput input::placeholder,
    .stTextArea textarea::placeholder {
        color: #f7d594 !important;
        -webkit-text-fill-color: #f7d594 !important;
        opacity: 0.95 !important;
        font-weight: 600 !important;
    }

    /* DROPDOWN OPTIONS & MENU LIST CONTRAST */
    div[data-baseweb="popover"] *,
    div[data-baseweb="menu"] *,
    div[role="option"] {
        background-color: #081d19 !important;
        color: #ffffff !important;
        font-weight: 600 !important;
    }
    div[role="option"]:hover, 
    div[role="option"][aria-selected="true"] {
        background-color: #16443c !important;
        color: #f7d594 !important;
    }

    /* CHECKBOX & RADIO LABELS HIGH CONTRAST */
    .stCheckbox label, .stCheckbox span, .stCheckbox p,
    .stRadio label, .stRadio span, .stRadio p {
        color: #ffffff !important;
        font-weight: 700 !important;
        font-size: 0.98rem !important;
        opacity: 1 !important;
    }

    /* STREAMLIT ALERT / NOTIFICATION BOX CONTRAST */
    .stAlert, div[data-testid="stAlert"] {
        background-color: #0c2b25 !important;
        border: 1px solid #e5a93c !important;
        border-radius: 12px !important;
        color: #ffffff !important;
    }
    .stAlert *, div[data-testid="stAlert"] * {
        color: #ffffff !important;
        font-weight: 700 !important;
    }
</style>
""", unsafe_allow_html=True)

import json

# --- DISK PERSISTENCE ENGINE ---
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DATA_DIR = os.path.join(BASE_DIR, "data")
os.makedirs(DATA_DIR, exist_ok=True)
FINALIZED_DIR = os.path.join(DATA_DIR, "finalized_rosters")
os.makedirs(FINALIZED_DIR, exist_ok=True)

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
  },
  "demo.employee": {
    "username": "demo.employee", "password": "DemoPass123!", "role": "Employee", "employee_name": "Demo Employee (Test Account)",
    "profile": { "full_name": "Demo Employee", "address": "123 Test Street", "home_phone": "0390000000", "mobile": "0400000000", "email": "demo.employee@example.com", "dob": "2005-05-15", "gender": "Female", "tfn": "123456789", "store": "Brumby's Pakenham", "classification": "Casual", "commencement_date": "2024-01-01", "employment_level": "Junior Team Member", "super_fund": "AustralianSuper", "super_policy": "AS123456", "super_address": "", "super_contact": "", "super_abn": "", "bank_name": "ANZ Bank", "bank_branch": "Pakenham", "bank_bsb": "013000", "bank_account": "12345678", "account_name": "Demo Employee" }
  },
  "demo.manager": {
    "username": "demo.manager", "password": "DemoPass123!", "role": "Manager", "employee_name": "Demo Manager (Test Account)",
    "profile": { "full_name": "Demo Manager", "address": "456 Test Ave", "home_phone": "0391111111", "mobile": "0411111111", "email": "demo.manager@example.com", "dob": "1990-08-20", "gender": "Male", "tfn": "987654321", "store": "Brumby's Pakenham", "classification": "Full-Time", "commencement_date": "2023-01-01", "employment_level": "Bakery Manager", "super_fund": "Hostplus", "super_policy": "HP987654", "super_address": "", "super_contact": "", "super_abn": "", "bank_name": "Commonwealth Bank", "bank_branch": "Pakenham", "bank_bsb": "063000", "bank_account": "87654321", "account_name": "Demo Manager" }
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
        profiles = copy.deepcopy(DEFAULT_PROFILES)
        if not st.session_state.get("is_demo", False):
            try:
                with open(USER_PROFILES_FILE, "w", encoding="utf-8") as f:
                    json.dump(profiles, f, indent=2)
            except:
                pass
                
    # Always ensure default demo accounts exist
    for d_key, d_val in DEFAULT_PROFILES.items():
        if d_key not in profiles:
            profiles[d_key] = copy.deepcopy(d_val)
            
    return profiles

def get_active_user_profiles():
    if st.session_state.get("is_demo", False) or (st.session_state.get("logged_in_user") or "").startswith("demo."):
        if "demo_user_profiles" not in st.session_state:
            st.session_state.demo_user_profiles = copy.deepcopy(load_user_profiles())
        return st.session_state.demo_user_profiles
    return load_user_profiles()

def save_user_profiles(profiles):
    if st.session_state.get("is_demo", False) or (st.session_state.get("logged_in_user") or "").startswith("demo."):
        st.session_state.demo_user_profiles = copy.deepcopy(profiles)
        st.toast("🧪 Sandbox Mode: Account profile updates held in memory only.", icon="🧪")
        return
    try:
        with open(USER_PROFILES_FILE, "w", encoding="utf-8") as f:
            json.dump(profiles, f, indent=2)
    except Exception as e:
        st.error(f"Error saving user profiles: {e}")

# --- SMTP EMAIL CONFIGURATION ENGINE ---
SMTP_CONFIG_FILE = os.path.join(DATA_DIR, "smtp_config.json")

def load_smtp_config():
    default_config = {
        "smtp_server": "smtp.gmail.com",
        "smtp_port": 587,
        "sender_email": "",
        "sender_password": "",
        "portal_url": "https://weekly-roster-generator.streamlit.app",
        "sender_name": "Bakery Manager"
    }
    if os.path.exists(SMTP_CONFIG_FILE):
        try:
            with open(SMTP_CONFIG_FILE, "r", encoding="utf-8") as f:
                cfg = json.load(f)
                for k, v in default_config.items():
                    if k not in cfg:
                        cfg[k] = v
                return cfg
        except:
            pass
    return default_config

def save_smtp_config(cfg):
    try:
        with open(SMTP_CONFIG_FILE, "w", encoding="utf-8") as f:
            json.dump(cfg, f, indent=2)
    except:
        pass

def build_welcome_email_content(emp_name, username, temp_password, portal_url="", sender_name=""):
    cfg = load_smtp_config()
    if not portal_url:
        portal_url = cfg.get("portal_url", "https://weekly-roster-generator.streamlit.app")
    if not sender_name:
        sender_name = cfg.get("sender_name", "Bakery Manager")
        
    subject = "Welcome to Brumby's Pakenham! 🥐 — Your Account Setup & Portal Login"
    body = f"""Subject: {subject}

Hi {emp_name},

Welcome to the team at Brumby's Pakenham! 🥐

We use an online portal for managing shift rostering, availability, and confidential employee onboarding. Please set up your account by completing these quick steps:

1. Log In to the Portal

Website: {portal_url}
Username: {username}
Temporary Password: {temp_password}

2. Complete Your Setup Tasks
Once logged in, please complete the following tabs:

📋 Personal Information Form: Fill out your contact details, Tax File Number (TFN), bank account details (for payroll), and superannuation fund.
📅 My Availability & Constraints: Select the days/time windows you are available to work.
🔑 Change Password: Click Change My Password in the left sidebar to update your temporary password to a secure personal one.

If you run into any issues, please feel free to reach out.

Welcome aboard!
{sender_name}
Brumby's Pakenham"""
    return subject, body

import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

def send_welcome_email_smtp(recipient_email, emp_name, username, temp_password, portal_url="", sender_name=""):
    cfg = load_smtp_config()
    smtp_server = cfg.get("smtp_server", "smtp.gmail.com")
    smtp_port = int(cfg.get("smtp_port", 587))
    sender_email = cfg.get("sender_email", "").strip()
    sender_pass = cfg.get("sender_password", "").strip()
    
    if not portal_url:
        portal_url = cfg.get("portal_url", "https://weekly-roster-generator.streamlit.app")
    if not sender_name:
        sender_name = cfg.get("sender_name", "Bakery Manager")

    subject, body = build_welcome_email_content(emp_name, username, temp_password, portal_url, sender_name)
    
    if not recipient_email or not recipient_email.strip():
        return False, "Recipient email is blank."
        
    if not sender_email or not sender_pass:
        return False, "SMTP credentials not yet configured in Email Settings."

    try:
        msg = MIMEMultipart()
        msg["From"] = f"{sender_name} <{sender_email}>"
        msg["To"] = recipient_email.strip()
        msg["Subject"] = subject
        
        email_text_content = body.split("\n\n", 1)[-1] if "\n\n" in body else body
        msg.attach(MIMEText(email_text_content, "plain", "utf-8"))

        server = smtplib.SMTP(smtp_server, smtp_port, timeout=10)
        server.starttls()
        server.login(sender_email, sender_pass)
        server.send_message(msg)
        server.quit()
        return True, f"Welcome Email sent to {recipient_email}!"
    except smtplib.SMTPAuthenticationError:
        return False, "Gmail authentication failed (Error 535). Gmail requires a 16-character 'Google App Password' instead of your normal account password."
    except Exception as e:
        return False, f"SMTP Error: {e}"

def send_test_email_smtp(recipient_email):
    cfg = load_smtp_config()
    smtp_server = cfg.get("smtp_server", "smtp.gmail.com")
    smtp_port = int(cfg.get("smtp_port", 587))
    sender_email = cfg.get("sender_email", "").strip()
    sender_pass = cfg.get("sender_password", "").strip()
    sender_name = cfg.get("sender_name", "Bakery Manager")
    
    if not recipient_email or not recipient_email.strip():
        return False, "Sender / recipient email address is blank."
    if not sender_email or not sender_pass:
        return False, "Please enter both Sender Email Address and App Password."
        
    try:
        msg = MIMEMultipart()
        msg["From"] = f"{sender_name} <{sender_email}>"
        msg["To"] = recipient_email.strip()
        msg["Subject"] = "🧪 Test Email — Brumby's Bakery Portal"
        
        test_body = f"Hello!\n\nThis is a test email sent from Brumby's Bakery Portal to verify your SMTP configuration.\n\nSender: {sender_email}\nSMTP Host: {smtp_server}:{smtp_port}\n\nIf you received this message, your welcome email setup is working perfectly! 🎉"
        msg.attach(MIMEText(test_body, "plain", "utf-8"))

        server = smtplib.SMTP(smtp_server, smtp_port, timeout=10)
        server.starttls()
        server.login(sender_email, sender_pass)
        server.send_message(msg)
        server.quit()
        return True, f"✅ Test email delivered to {recipient_email}!"
    except smtplib.SMTPAuthenticationError:
        return False, "❌ Gmail Authentication Failed (Error 535). Gmail requires a 16-character 'Google App Password' generated in your Google Account Security settings."
    except Exception as e:
        return False, f"❌ Connection Error: {e}"

def load_persisted_df(filename, default_df):
    path = os.path.join(DATA_DIR, filename)
    if os.path.exists(path):
        try:
            # Read all columns as string without converting empty cells to NaN
            df = pd.read_csv(path, dtype=str, keep_default_na=False)
            if df is None or df.empty:
                return default_df
            return df
        except:
            return default_df
    return default_df

def save_persisted_df(df, filename):
    if st.session_state.get("is_demo", False):
        st.toast("🧪 Sandbox Mode: Database updates held in memory only.", icon="🧪")
        return
    path = os.path.join(DATA_DIR, filename)
    try:
        df.astype(str).to_csv(path, index=False)
    except:
        pass

def build_roster_excel_bytes(edited_final_df, start_date):
    if isinstance(start_date, str):
        try:
            start_date = datetime.strptime(start_date, "%Y-%m-%d").date()
        except:
            start_date = datetime.now().date()
            
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

    # Write roster data rows
    row_start = 5
    for r_idx, row_data in enumerate(edited_final_df.itertuples(index=False), start=row_start):
        for c_idx, val in enumerate(row_data, start=1):
            c = ws.cell(row=r_idx, column=c_idx)
            val_str = "" if pd.isna(val) else str(val).strip()
            
            # Don't show "unavailable" or "off" in exported Excel file - leave cells clean & blank
            if val_str.lower() in ["unavailable", " unavailable", "off", "none", "nan"]:
                val_str = ""
            
            c.value = val_str
            c.alignment = Alignment(horizontal="center", vertical="center")
            c.border = thin_border
            
            if val_str:
                if c_idx == 1:
                    c.font = Font(name="Calibri", size=11, bold=True)
                    c.alignment = Alignment(horizontal="left", vertical="center")
                else:
                    c.font = Font(name="Calibri", size=10)

    last_roster_row = row_start + len(edited_final_df) - 1

    # --- EMPLOYEE SHIFT BREAK ENTITLEMENTS GUIDE ---
    guide_row_1 = last_roster_row + 3
    
    ws.merge_cells(start_row=guide_row_1, start_column=1, end_row=guide_row_1, end_column=8)
    g_title = ws.cell(row=guide_row_1, column=1)
    g_title.value = "☕ EMPLOYEE SHIFT BREAK ENTITLEMENTS GUIDE (General Retail Industry Award 2020)"
    g_title.fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    g_title.font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    g_title.alignment = Alignment(horizontal="center", vertical="center")

    hdr_row = guide_row_1 + 1
    headers_break = [
        (1, 2, "Shift Duration"),
        (3, 4, "Paid Rest Break (10 min)"),
        (5, 6, "Unpaid Meal Break (30 min)"),
        (7, 8, "Total Break Entitlement")
    ]
    for start_col, end_col, text in headers_break:
        ws.merge_cells(start_row=hdr_row, start_column=start_col, end_row=hdr_row, end_column=end_col)
        hc = ws.cell(row=hdr_row, column=start_col)
        hc.value = text
        hc.fill = PatternFill(start_color="203764", end_color="203764", fill_type="solid")
        hc.font = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
        hc.alignment = Alignment(horizontal="center", vertical="center")
        for col in range(start_col, end_col + 1):
            ws.cell(row=hdr_row, column=col).border = thin_border

    break_data = [
        ("Less than 4 hours", "No break", "No break", "No breaks required"),
        ("4 hours up to 5 hours", "1 x 10-minute rest break", "No meal break", "10 min Paid break"),
        ("5 hours up to 7 hours", "1 x 10-minute rest break", "1 x 30-minute meal break", "10m Paid + 30m Unpaid meal"),
        ("7 hours up to 10 hours", "2 x 10-minute rest breaks", "1 x 30-minute meal break", "20m Paid + 30m Unpaid meal"),
        ("10 hours or more", "2 x 10-minute rest breaks", "2 x 30-minute meal breaks", "20m Paid + 60m Unpaid meals"),
    ]

    curr_b_row = hdr_row + 1
    for d1, d2, d3, d4 in break_data:
        row_cols = [(1, 2, d1), (3, 4, d2), (5, 6, d3), (7, 8, d4)]
        for start_col, end_col, text in row_cols:
            ws.merge_cells(start_row=curr_b_row, start_column=start_col, end_row=curr_b_row, end_column=end_col)
            cell = ws.cell(row=curr_b_row, column=start_col)
            cell.value = text
            cell.font = Font(name="Calibri", size=9.5)
            cell.alignment = Alignment(horizontal="center", vertical="center")
            for col in range(start_col, end_col + 1):
                ws.cell(row=curr_b_row, column=col).border = thin_border
        curr_b_row += 1

    note_row = curr_b_row + 1
    ws.merge_cells(start_row=note_row, start_column=1, end_row=note_row+2, end_column=8)
    n_cell = ws.cell(row=note_row, column=1)
    n_cell.value = (
        "📌 Key Break Rules for Bakery Staff:\n"
        "• Rest breaks (10 mins) are paid. Meal breaks (30 mins) are unpaid.\n"
        "• Breaks cannot be taken in the first or last hour of work.\n"
        "• An unpaid meal break must be taken no later than after 5 hours of continuous work."
    )
    n_cell.fill = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")
    n_cell.font = Font(name="Calibri", size=9, italic=True, color="333333")
    n_cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
    for r in range(note_row, note_row + 3):
        for c in range(1, 9):
            ws.cell(row=r, column=c).border = thin_border

    # Dedicated 2nd Sheet Tab: Break Entitlements Guide
    ws2 = wb.create_sheet(title="Break Entitlements Guide")
    ws2.merge_cells("A1:D1")
    t2 = ws2.cell(row=1, column=1)
    t2.value = "BRUMBY'S PAKENHAM - EMPLOYEE BREAK ENTITLEMENTS GUIDE"
    t2.fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    t2.font = Font(name="Calibri", size=13, bold=True, color="FFFFFF")
    t2.alignment = Alignment(horizontal="center", vertical="center")

    headers_s2 = ["Shift Duration", "Paid Rest Break (10 min)", "Unpaid Meal Break (30 min)", "Total Break Entitlement"]
    for col_idx, h_text in enumerate(headers_s2, 1):
        c2 = ws2.cell(row=3, column=col_idx)
        c2.value = h_text
        c2.fill = PatternFill(start_color="203764", end_color="203764", fill_type="solid")
        c2.font = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
        c2.alignment = Alignment(horizontal="center", vertical="center")
        c2.border = thin_border

    for r_idx, (d1, d2, d3, d4) in enumerate(break_data, start=4):
        vals = [d1, d2, d3, d4]
        for c_idx, val in enumerate(vals, start=1):
            c = ws2.cell(row=r_idx, column=c_idx)
            c.value = val
            c.font = Font(name="Calibri", size=10)
            c.alignment = Alignment(horizontal="center", vertical="center")
            c.border = thin_border

    ws2.column_dimensions["A"].width = 25
    ws2.column_dimensions["B"].width = 28
    ws2.column_dimensions["C"].width = 28
    ws2.column_dimensions["D"].width = 32

    # Auto-adjust column widths for Sheet 1 (based only on roster table cells, ignoring merged title & note cells)
    for col_idx in range(1, len(edited_final_df.columns) + 1):
        col_letter = openpyxl.utils.get_column_letter(col_idx)
        max_len = max(len(str(ws.cell(row=r, column=col_idx).value or '')) for r in range(4, last_roster_row + 1))
        if col_idx == 1:
            ws.column_dimensions[col_letter].width = max(max_len + 5, 20)
        else:
            ws.column_dimensions[col_letter].width = max(max_len + 4, 16)

    excel_buffer = io.BytesIO()
    wb.save(excel_buffer)
    return excel_buffer.getvalue()

def save_finalized_roster(df, start_date):
    if isinstance(start_date, str):
        try:
            start_date = datetime.strptime(start_date, "%Y-%m-%d").date()
        except:
            start_date = datetime.now().date()
            
    date_str = start_date.strftime("%Y-%m-%d")
    date_label = start_date.strftime("%d.%m.%Y")
    
    csv_filename = f"Roster_{date_str}.csv"
    xlsx_filename = f"Team_Roster_{date_label}.xlsx"
    
    excel_bytes = build_roster_excel_bytes(df, start_date)
    
    if st.session_state.get("is_demo", False):
        st.toast("🧪 Sandbox Mode: Finalized roster saved in session memory only.", icon="🧪")
        return date_str, xlsx_filename, excel_bytes
        
    csv_path = os.path.join(FINALIZED_DIR, csv_filename)
    xlsx_path = os.path.join(FINALIZED_DIR, xlsx_filename)
    
    df.astype(str).to_csv(csv_path, index=False)
    with open(xlsx_path, "wb") as f:
        f.write(excel_bytes)
        
    return date_str, xlsx_filename, excel_bytes

def list_finalized_rosters():
    if not os.path.exists(FINALIZED_DIR):
        return []
    files = [f for f in os.listdir(FINALIZED_DIR) if f.endswith(".csv") and f.startswith("Roster_")]
    files.sort(reverse=True)
    results = []
    for f in files:
        raw_date = f.replace("Roster_", "").replace(".csv", "")
        try:
            dt = datetime.strptime(raw_date, "%Y-%m-%d").date()
            end_dt = dt + timedelta(days=6)
            label = f"Week of {dt.strftime('%d/%m/%Y')} (Mon {dt.strftime('%d/%m')} - Sun {end_dt.strftime('%d/%m')})"
        except:
            label = f"Roster {raw_date}"
        results.append({"csv_filename": f, "date_str": raw_date, "label": label})
    return results

def load_finalized_roster(csv_filename):
    csv_path = os.path.join(FINALIZED_DIR, csv_filename)
    if os.path.exists(csv_path):
        try:
            df = pd.read_csv(csv_path, dtype=str, keep_default_na=False)
            return df
        except:
            return None
    return None

def delete_finalized_roster(date_str):
    if st.session_state.get("is_demo", False):
        st.toast("🧪 Sandbox Mode: Deletion previewed, disk files preserved.", icon="🧪")
        return True
    if not os.path.exists(FINALIZED_DIR):
        return False
    deleted = False
    date_label = date_str
    try:
        dt = datetime.strptime(date_str, "%Y-%m-%d").date()
        date_label = dt.strftime("%d.%m.%Y")
    except:
        pass
        
    for f in os.listdir(FINALIZED_DIR):
        if date_str in f or date_label in f:
            try:
                os.remove(os.path.join(FINALIZED_DIR, f))
                deleted = True
            except:
                pass
    return deleted

def calculate_roster_wages(edited_df):
    if edited_df is None or edited_df.empty:
        return {
            "total_gross": 0.0,
            "total_tax": 0.0,
            "total_net": 0.0,
            "total_super": 0.0,
            "total_hours": 0.0,
            "avg_hourly_rate": 0.0,
            "breakdown_df": pd.DataFrame()
        }

    emp_meta = {}
    if "manual_employees" in st.session_state and isinstance(st.session_state.manual_employees, pd.DataFrame):
        e_df = st.session_state.manual_employees
        n_col = find_column(e_df, ["name", "employee", "staff"], "Name")
        a_col = find_column(e_df, ["age"], "Age")
        s_col = find_column(e_df, ["employment type", "status", "classification", "type"], "Employment Type")
        
        for _, r in e_df.iterrows():
            name = str(r.get(n_col, "")).strip()
            if name:
                age_val = r.get(a_col, 21)
                try:
                    age = int(float(age_val))
                except:
                    age = 21
                status = str(r.get(s_col, "Casual")).strip().lower()
                emp_meta[name.lower()] = {"name": name, "age": age, "status": status}

    for u_key, u_data in user_profiles.items():
        emp_name = u_data.get("employee_name", u_key)
        prof = u_data.get("profile", {})
        status = str(prof.get("classification", "Casual")).strip().lower()
        if emp_name.lower() not in emp_meta:
            emp_meta[emp_name.lower()] = {"name": emp_name, "age": 21, "status": status}

    emp_col = find_column(edited_df, ["employee", "name", "staff"], "Employee")
    days_of_week = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"]

    emp_summary = []

    for _, row in edited_df.iterrows():
        emp_raw_name = str(row.get(emp_col, "")).strip()
        if not emp_raw_name or emp_raw_name.lower() in ["none", "nan", "total", "summary"]:
            continue

        meta = emp_meta.get(emp_raw_name.lower(), {"name": emp_raw_name, "age": 21, "status": "casual"})
        age = meta["age"]
        status_clean = meta["status"]
        is_casual = "casual" in status_clean

        base_adult_rate = 26.10
        if "manager" in emp_raw_name.lower() or "owner" in status_clean:
            base_adult_rate = 30.00
        elif "senior" in status_clean:
            base_adult_rate = 26.80

        if age < 16:
            j_scale = 0.45
        elif age == 16:
            j_scale = 0.50
        elif age == 17:
            j_scale = 0.60
        elif age == 18:
            j_scale = 0.70
        elif age == 19:
            j_scale = 0.80
        elif age == 20:
            j_scale = 0.90
        else:
            j_scale = 1.00

        base_hourly = base_adult_rate * j_scale

        total_emp_hours = 0.0
        total_emp_gross = 0.0

        for day in days_of_week:
            if day in row:
                shift_val = str(row[day]).strip()
                parsed = parse_shift_range(shift_val)
                if parsed:
                    start_t, end_t, duration = parsed
                    paid_hrs = duration - 0.5 if duration >= 5.0 else duration
                    total_emp_hours += paid_hrs

                    if day in ["Saturday"]:
                        multiplier = 1.50 if is_casual else 1.25
                    elif day in ["Sunday"]:
                        multiplier = 1.75 if is_casual else 1.50
                    else:
                        if end_t > 18.0:
                            multiplier = 1.50 if is_casual else 1.25
                        else:
                            multiplier = 1.25 if is_casual else 1.00

                    total_emp_gross += paid_hrs * base_hourly * multiplier

        g = total_emp_gross
        if g <= 359:
            tax = 0.0
        elif g <= 865:
            tax = (g - 359) * 0.19
        elif g <= 2500:
            tax = 96.14 + (g - 865) * 0.325
        else:
            tax = 627.51 + (g - 2500) * 0.37

        net_pay = max(0.0, g - tax)
        super_sg = g * 0.125

        status_label = "Casual" if is_casual else ("Part-Time" if "part" in status_clean else "Full-Time")
        if age < 21:
            status_label += f" ({age}yo)"

        emp_summary.append({
            "Staff Member": emp_raw_name,
            "Status": status_label,
            "Paid Hours": round(total_emp_hours, 1),
            "Gross Pay": round(total_emp_gross, 2),
            "Est. Tax": round(tax, 2),
            "Net Pay": round(net_pay, 2),
            "Super (12.5%)": round(super_sg, 2)
        })

    breakdown_df = pd.DataFrame(emp_summary)
    tot_gross = sum(x["Gross Pay"] for x in emp_summary)
    tot_tax = sum(x["Est. Tax"] for x in emp_summary)
    tot_net = sum(x["Net Pay"] for x in emp_summary)
    tot_super = sum(x["Super (12.5%)"] for x in emp_summary)
    tot_hrs = sum(x["Paid Hours"] for x in emp_summary)
    avg_rate = (tot_gross / tot_hrs) if tot_hrs > 0 else 0.0

    return {
        "total_gross": round(tot_gross, 2),
        "total_tax": round(tot_tax, 2),
        "total_net": round(tot_net, 2),
        "total_super": round(tot_super, 2),
        "total_hours": round(tot_hrs, 1),
        "avg_hourly_rate": round(avg_rate, 2),
        "breakdown_df": breakdown_df
    }

def build_payroll_historical_trend():
    past_rosters = list_finalized_rosters()
    if not past_rosters:
        return pd.DataFrame()
    
    records = []
    for r in past_rosters:
        df = load_finalized_roster(r["csv_filename"])
        if df is not None and not df.empty:
            w_sum = calculate_roster_wages(df)
            try:
                dt = datetime.strptime(r["date_str"], "%Y-%m-%d").date()
                week_label = dt.strftime("%d/%m/%Y")
            except:
                dt = datetime.min.date()
                week_label = r["date_str"]
                
            records.append({
                "date": dt,
                "Roster Week": week_label,
                "Gross Payroll ($)": w_sum["total_gross"],
                "Est. PAYG Tax ($)": w_sum["total_tax"],
                "Net Take-Home ($)": w_sum["total_net"],
                "Super 12.5% ($)": w_sum["total_super"]
            })
            
    if not records:
        return pd.DataFrame()
        
    trend_df = pd.DataFrame(records).sort_values("date").drop(columns=["date"]).reset_index(drop=True)
    return trend_df

def find_column(df, candidates, default=""):
    if df is None or not hasattr(df, "columns") or len(df.columns) == 0:
        return default
    # 1. Exact match
    for c in df.columns:
        if str(c).strip().lower() in [cand.lower() for cand in candidates]:
            return c
    # 2. Substring match
    for c in df.columns:
        c_clean = str(c).strip().lower()
        for cand in candidates:
            cand_clean = cand.lower()
            if cand_clean in c_clean or c_clean in cand_clean:
                return c
    return default

# Initialize Session States with Disk Cache
if 'authenticated' not in st.session_state:
    st.session_state.authenticated = False
if 'logged_in_user' not in st.session_state:
    st.session_state.logged_in_user = None
if 'user_role' not in st.session_state:
    st.session_state.user_role = None
if 'is_demo' not in st.session_state:
    st.session_state.is_demo = False

user_profiles = get_active_user_profiles()

# LOGIN PAGE IF NOT AUTHENTICATED
if not st.session_state.authenticated:
    st.markdown("""
    <style>
        .stApp {
            background-color: #051412 !important;
        }
    </style>
    """, unsafe_allow_html=True)
    
    col1, col2, col3 = st.columns([1, 1.8, 1])
    with col2:
        st.markdown("""
        <div style="background: linear-gradient(135deg, #081d19 0%, #16443c 100%); padding: 35px 30px; border-radius: 20px; border: 2px solid #e5a93c; box-shadow: 0 12px 36px rgba(0,0,0,0.6); margin-top: 40px;">
            <h2 style="color: #e5a93c !important; margin-top: 0; text-align: center; font-size: 1.8rem; font-weight: 900;">🥐 Brumby's Pakenham — Portal</h2>
            <p style="color: #c8e6e0 !important; font-size: 0.95rem; text-align: center; margin-bottom: 25px;">Please log in with your username and password to access your bakery account.</p>
        """, unsafe_allow_html=True)
        
        login_user = st.text_input("Username", key="input_user")
        login_pass = st.text_input("Password", type="password", key="input_pass")
        
        st.markdown("<br>", unsafe_allow_html=True)
        if st.button("🚀 Login to Portal", use_container_width=True, key="btn_login"):
            login_user_clean = login_user.strip().lower()
            current_profiles = load_user_profiles()
            if login_user_clean in current_profiles:
                account = current_profiles[login_user_clean]
                if account.get("password") == login_pass:
                    st.session_state.authenticated = True
                    st.session_state.logged_in_user = login_user_clean
                    st.session_state.user_role = account.get("role", "Employee")
                    st.session_state.is_demo = account.get("is_demo", False) or "demo." in login_user_clean
                    st.success(f"Welcome back, {account.get('employee_name', login_user_clean)}!")
                    st.rerun()
                else:
                    st.error("🔒 Invalid username or password. Please check your credentials.")
            else:
                st.error("🔒 Invalid username or password. Please check your credentials.")
                
        # Quick One-Click Demo Mode Login Buttons
        st.markdown("<br><hr style='border-color: rgba(229, 169, 60, 0.3);'>", unsafe_allow_html=True)
        st.markdown("""
        <div style="text-align: center; margin-bottom: 12px;">
            <span style="color: #e5a93c; font-weight: 800; font-size: 0.95rem;">🧪 ONE-CLICK DEMO / QUICK TEST MODE</span><br>
            <span style="color: #c8e6e0; font-size: 0.85rem;">Test all portal features safely without database side-effects:</span>
        </div>
        """, unsafe_allow_html=True)
        
        col_demo1, col_demo2 = st.columns(2)
        with col_demo1:
            if st.button("👤 Demo Employee Mode", use_container_width=True, key="btn_demo_emp"):
                st.session_state.authenticated = True
                st.session_state.logged_in_user = "demo.employee"
                st.session_state.user_role = "Employee"
                st.session_state.is_demo = True
                st.success("🧪 Logging in as Demo Employee (Sandbox Mode)...")
                st.rerun()
                
        with col_demo2:
            if st.button("👑 Demo Manager Mode", use_container_width=True, key="btn_demo_mgr"):
                st.session_state.authenticated = True
                st.session_state.logged_in_user = "demo.manager"
                st.session_state.user_role = "Manager"
                st.session_state.is_demo = True
                st.success("🧪 Logging in as Demo Manager (Sandbox Mode)...")
                st.rerun()

        st.markdown("""
        </div>
        """, unsafe_allow_html=True)
    st.stop()

# --- SIDEBAR CONFIGURATION (AUTHENTICATED) ---
if st.session_state.is_demo:
    st.sidebar.warning("⚠️ **SANDBOX TEST MODE**\nChanges made here will not be permanently saved.")

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

# Helper function to render Change Password form
def render_change_password_form(user_key, is_admin=False):
    user_data = user_profiles.get(user_key, {})
    emp_name = user_data.get("employee_name", user_key)
    
    with st.expander(f"🔑 Change Password for {emp_name}", expanded=False):
        with st.form(key=f"form_change_pw_{user_key}"):
            if not is_admin:
                curr_pw = st.text_input("Current Password", type="password", key=f"cp_curr_pw_{user_key}")
            else:
                curr_pw = None
                
            new_pw = st.text_input("New Password", type="password", key=f"cp_new_pw_{user_key}")
            confirm_pw = st.text_input("Confirm New Password", type="password", key=f"cp_conf_pw_{user_key}")
            
            submit_pw = st.form_submit_button("🔒 Save New Password")
            
            if submit_pw:
                actual_pw = user_data.get("password", "")
                if not is_admin and curr_pw != actual_pw:
                    st.error("❌ Incorrect current password.")
                elif not new_pw.strip():
                    st.error("❌ New password cannot be empty.")
                elif new_pw != confirm_pw:
                    st.error("❌ New passwords do not match.")
                else:
                    user_profiles[user_key]["password"] = new_pw.strip()
                    save_user_profiles(user_profiles)
                    st.success(f"✅ Password for {emp_name} updated successfully!")

# Sidebar Change Password Expander
with st.sidebar.expander("🔑 Change My Password", expanded=False):
    with st.form(key=f"form_sidebar_pw_{curr_user_key}"):
        sb_curr_pw = st.text_input("Current Password", type="password", key=f"sb_cp_curr_{curr_user_key}")
        sb_new_pw = st.text_input("New Password", type="password", key=f"sb_cp_new_{curr_user_key}")
        sb_conf_pw = st.text_input("Confirm New Password", type="password", key=f"sb_cp_conf_{curr_user_key}")
        
        sb_submit_pw = st.form_submit_button("🔒 Update Password")
        
        if sb_submit_pw:
            actual_pw = curr_user_info.get("password", "")
            if sb_curr_pw != actual_pw:
                st.error("❌ Incorrect current password.")
            elif not sb_new_pw.strip():
                st.error("❌ New password cannot be empty.")
            elif sb_new_pw != sb_conf_pw:
                st.error("❌ New passwords do not match.")
            else:
                user_profiles[curr_user_key]["password"] = sb_new_pw.strip()
                save_user_profiles(user_profiles)
                st.success("✅ Your password has been updated successfully!")

if role_title == "Manager":
    with st.sidebar.expander("📧 Email Settings (SMTP & Portal Link)", expanded=False):
        smtp_cfg = load_smtp_config()
        with st.form(key="form_smtp_config"):
            st.markdown("##### 🌐 Default Portal Web Link")
            cfg_url = st.text_input("Streamlit Cloud URL", value=smtp_cfg.get("portal_url", "https://weekly-roster-generator.streamlit.app"))
            cfg_sname = st.text_input("Sender Display Name", value=smtp_cfg.get("sender_name", "Bakery Manager"))
            
            st.markdown("##### 📮 SMTP Server Credentials")
            cfg_semail = st.text_input("Sender Email Address", value=smtp_cfg.get("sender_email", ""), placeholder="e.g. manager@brumbys.com.au")
            cfg_spass = st.text_input("Sender App Password", value=smtp_cfg.get("sender_password", ""), type="password")
            cfg_host = st.text_input("SMTP Host", value=smtp_cfg.get("smtp_server", "smtp.gmail.com"))
            cfg_port = st.number_input("SMTP Port", value=int(smtp_cfg.get("smtp_port", 587)))
            
            c_save, c_test = st.columns([1.2, 1])
            with c_save:
                btn_save_smtp = st.form_submit_button("💾 Save Settings")
            with c_test:
                btn_test_smtp = st.form_submit_button("🧪 Test Email")

            if btn_save_smtp:
                new_cfg = {
                    "portal_url": cfg_url.strip(),
                    "sender_name": cfg_sname.strip(),
                    "sender_email": cfg_semail.strip(),
                    "sender_password": cfg_spass.strip(),
                    "smtp_server": cfg_host.strip(),
                    "smtp_port": int(cfg_port)
                }
                save_smtp_config(new_cfg)
                st.success("✅ Email settings saved successfully!")

            if btn_test_smtp:
                new_cfg = {
                    "portal_url": cfg_url.strip(),
                    "sender_name": cfg_sname.strip(),
                    "sender_email": cfg_semail.strip(),
                    "sender_password": cfg_spass.strip(),
                    "smtp_server": cfg_host.strip(),
                    "smtp_port": int(cfg_port)
                }
                save_smtp_config(new_cfg)
                ok, msg = send_test_email_smtp(cfg_semail.strip())
                if ok:
                    st.success(msg)
                else:
                    st.error(msg)

def logout_user():
    for k in ["authenticated", "logged_in_user", "user_role", "is_demo", "demo_user_profiles", "demo_state_initialized", "manual_employees", "manual_unavailability", "manual_requirements", "manual_fixed", "final_roster_df", "edit_employees", "edit_unavailability_v4", "edit_requirements", "edit_fixed"]:
        st.session_state.pop(k, None)
    st.session_state.authenticated = False
    st.rerun()

if st.sidebar.button("🚪 Logout", key="btn_logout"):
    logout_user()

if st.session_state.get("is_demo", False):
    st.markdown("""
    <div style="background: rgba(229, 169, 60, 0.2); border: 2px solid #e5a93c; padding: 12px 18px; border-radius: 14px; margin-bottom: 20px;">
        <span style="color: #f7d594; font-weight: 800; font-size: 1.05rem;">🧪 DEMO & SANDBOX TEST MODE ACTIVE</span>
        <span style="color: #ffffff; margin-left: 10px; font-size: 0.95rem;">You are using a test demo account. All features, forms, calculations, and tables are fully interactive, but permanent updates to database files are safely disabled.</span>
    </div>
    """, unsafe_allow_html=True)

col_head1, col_head2 = st.columns([3.5, 1])
with col_head1:
    st.markdown('<h1 class="header-style">🥐 Brumby\'s Pakenham Portal</h1>', unsafe_allow_html=True)
    st.markdown(f'<p class="sub-header-style">Welcome back, <b>{display_name}</b> ({role_title})</p>', unsafe_allow_html=True)
with col_head2:
    st.markdown("<div style='margin-top: 15px;'></div>", unsafe_allow_html=True)
    if st.button("🚪 Logout", key="btn_logout_header", use_container_width=True):
        logout_user()

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

import re

def parse_date_robust(date_str):
    if not date_str or str(date_str).strip().lower() in ["nan", "none", "nat", ""]:
        return None
    date_str = str(date_str).strip()
    try:
        dt = pd.to_datetime(date_str, dayfirst=True, errors='coerce')
        if pd.notna(dt):
            return dt.date()
    except:
        pass
    return None

def extract_date_range(text):
    if not text or str(text).strip().lower() in ["nan", "none", "nat", ""]:
        return None, None
    text = str(text)
    dates_found = []
    date_matches = re.findall(r'\b\d{1,4}[-/\.]\d{1,2}[-/\.]\d{2,4}\b', text)
    for m in date_matches:
        d = parse_date_robust(m)
        if d:
            dates_found.append(d)
    
    if len(dates_found) >= 2:
        return min(dates_found), max(dates_found)
    elif len(dates_found) == 1:
        if any(kw in text.lower() for kw in ['from', 'after', 'since', 'starting']):
            return dates_found[0], None
        elif any(kw in text.lower() for kw in ['until', 'to', 'before', 'ending']):
            return None, dates_found[0]
        return dates_found[0], dates_found[0]
    return None, None

def is_unavail_applicable_to_date(dt_obj, u_day, u_win):
    u_day_clean = str(u_day).strip()
    u_win_clean = str(u_win).strip()
    
    # 1. Date range bounds
    start_d, end_d = extract_date_range(f"{u_day_clean} {u_win_clean}")
    if start_d and dt_obj < start_d:
        return False
    if end_d and dt_obj > end_d:
        return False
        
    # 2. Explicit date check
    explicit_date = parse_date_robust(u_day_clean)
    if explicit_date:
        return dt_obj == explicit_date
        
    # 3. Day-of-week matching
    weekday_names = ["monday", "tuesday", "wednesday", "thursday", "friday", "saturday", "sunday"]
    dt_day_name = weekday_names[dt_obj.weekday()]
    dt_day_short = dt_day_name[:3]
    
    u_day_lower = u_day_clean.lower()
    
    if dt_day_name in u_day_lower or dt_day_short in u_day_lower:
        return True
        
    if any(kw in u_day_lower for kw in ["all week", "everyday", "any day"]):
        return True
    if "weekend" in u_day_lower and dt_obj.weekday() in [5, 6]:
        return True
    if "weekday" in u_day_lower and dt_obj.weekday() in [0, 1, 2, 3, 4]:
        return True
        
def calculate_age_from_dob(dob_str):
    if not dob_str or str(dob_str).strip().lower() in ["nan", "none", "nat", "null", ""]:
        return None
    try:
        dt = pd.to_datetime(dob_str, dayfirst=True, errors='coerce')
        if pd.notna(dt):
            today = datetime.now()
            age = today.year - dt.year - ((today.month, today.day) < (dt.month, dt.day))
            return age, dt.strftime('%d/%m/%Y')
    except:
        pass
    return None

def cleanup_duplicate_employee_columns(df):
    if df is None or not isinstance(df, pd.DataFrame) or df.empty:
        return df
    
    df = df.copy()
    
    concept_mappings = {
        "NAME": ["name", "employee", "staff", "staff name", "employee name"],
        "DOB": ["dob", "date of birth", "birth date"],
        "Commencing Date": ["commencing date", "commence date", "commencement date", "start date", "started", "start"],
        "status": ["status", "employment type", "type", "classification", "employmenttype"],
        "position": ["position", "role", "employment level", "title", "job"]
    }
    
    matched_cols = {}
    for canonical, candidates in concept_mappings.items():
        matched_cols[canonical] = []
        for col in df.columns:
            if str(col).strip().lower() in [cand.lower() for cand in candidates]:
                matched_cols[canonical].append(col)
                
    new_rows = []
    for idx in df.index:
        row_dict = {}
        for canonical, source_cols in matched_cols.items():
            val = ""
            for sc in source_cols:
                v = str(df.at[idx, sc]).strip() if pd.notna(df.at[idx, sc]) else ""
                if v and v.lower() not in ["none", "nan", "nat", "null", ""]:
                    val = v
                    break
            row_dict[canonical] = val
            
        if row_dict["DOB"]:
            try:
                dt = pd.to_datetime(row_dict["DOB"], dayfirst=True, errors='coerce')
                if pd.notna(dt):
                    row_dict["DOB"] = dt.strftime('%d/%m/%Y')
            except:
                pass
                
        if row_dict["Commencing Date"]:
            try:
                dt = pd.to_datetime(row_dict["Commencing Date"], dayfirst=True, errors='coerce')
                if pd.notna(dt):
                    row_dict["Commencing Date"] = dt.strftime('%d/%m/%Y')
            except:
                pass

        new_rows.append(row_dict)

    res_df = pd.DataFrame(new_rows, columns=["NAME", "DOB", "Commencing Date", "status", "position"])
    return res_df

def sync_user_profiles_to_employees(emp_df):
    if emp_df is None:
        emp_df = pd.DataFrame(columns=["NAME", "DOB", "Commencing Date", "status", "position"])
    
    emp_df = cleanup_duplicate_employee_columns(emp_df)
    
    # Exclude demo accounts from employee dataframe
    if "NAME" in emp_df.columns:
        emp_df = emp_df[~emp_df["NAME"].astype(str).str.lower().str.contains("demo")].reset_index(drop=True)
    
    profiles = get_active_user_profiles()
    existing_names = [str(n).strip().lower() for n in emp_df["NAME"].tolist() if pd.notna(n)]
    
    new_rows = []
    for u_key, u_data in profiles.items():
        if u_key.startswith("demo.") or "demo" in u_key.lower():
            continue
        if u_data.get("role") == "Employee":
            emp_name = u_data.get("employee_name", u_key).strip()
            if "demo" in emp_name.lower():
                continue
            if emp_name and emp_name.lower() not in existing_names:
                prof = u_data.get("profile", {})
                comm_date = prof.get("commencement_date", datetime.now().strftime("%d/%m/%Y"))
                try:
                    dt = pd.to_datetime(comm_date, dayfirst=True, errors='coerce')
                    if pd.notna(dt):
                        comm_date = dt.strftime('%d/%m/%Y')
                except:
                    pass
                new_rows.append({
                    "NAME": emp_name,
                    "DOB": prof.get("dob", ""),
                    "Commencing Date": comm_date,
                    "status": str(prof.get("classification", "casual")).lower(),
                    "position": str(prof.get("employment_level", "Service Staff"))
                })
                existing_names.append(emp_name.lower())
                
    if new_rows:
        combined = pd.concat([emp_df, pd.DataFrame(new_rows)], ignore_index=True)
        emp_df = cleanup_duplicate_employee_columns(combined)
        save_persisted_df(emp_df, "employees.csv")
        
    return emp_df

def standardize_unavailability_df(df):
    if df is None or not isinstance(df, pd.DataFrame) or df.empty:
        return df
    
    df = df.copy()
    
    emp_col = find_column(df, ["employee", "name", "staff", "user", "person", "employee name", "staff name"], "Employee")
    day_col = find_column(df, ["day", "date", "weekday", "when", "day of week", "unavailable date"], "Day")
    win_col = find_column(df, ["time window", "window", "time", "unavailability", "reason", "constraint", "time constraint", "notes", "details"], "Time Window")

    weekday_names = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"]

    for idx, row in df.iterrows():
        # 1. Clean Employee Name
        emp_val = str(row.get(emp_col, "")).strip()
        if emp_val and emp_val.lower() not in ["nan", "none", "nat", "null", ""]:
            if emp_val.islower():
                emp_val = emp_val.title()
            df.at[idx, emp_col] = emp_val
            
        # 2. Clean & Standardize Day / Date Column
        day_val = str(row.get(day_col, "")).strip()
        if day_val and day_val.lower() not in ["nan", "none", "nat", "null", ""]:
            dt = parse_date_robust(day_val)
            if dt:
                w_name = weekday_names[dt.weekday()]
                date_formatted = dt.strftime('%d/%m/%Y')
                if w_name.lower() not in day_val.lower():
                    df.at[idx, day_col] = f"{w_name} ({date_formatted})"
                else:
                    df.at[idx, day_col] = day_val
            else:
                matched_w = [w for w in weekday_names if w.lower() == day_val.lower()]
                if matched_w:
                    df.at[idx, day_col] = matched_w[0]
                    
        # 3. Clean Time Window Column
        win_val = str(row.get(win_col, "")).strip()
        if not win_val or win_val.lower() in ["nan", "none", "nat", "null", ""]:
            df.at[idx, win_col] = "All Day"
            
    return df

def clean_win_display(win_str):
    if not win_str or str(win_str).strip().lower() in ["nan", "none", "nat", "null", ""]:
        return "All Day"
    cleaned = re.sub(r'\(From.*?\)', '', str(win_str), flags=re.IGNORECASE).strip()
    if not cleaned or cleaned.lower() in ["nan", "none", "nat", "null", ""]:
        return "All Day"
    return cleaned

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
    st.session_state.manual_employees = sync_user_profiles_to_employees(load_persisted_df("employees.csv", default_emp))

if 'manual_unavailability' not in st.session_state or st.session_state.manual_unavailability is None or st.session_state.manual_unavailability.empty:
    default_unavail = pd.DataFrame([
        {"Employee": "Elizabeth", "Day": "Saturday", "Time Window": "All Day"},
        {"Employee": "Elizabeth", "Day": "Sunday", "Time Window": "All Day"},
        {"Employee": "Stella", "Day": "Monday", "Time Window": "Before 3:30pm"},
        {"Employee": "Stella", "Day": "Tuesday", "Time Window": "Before 3:30pm"},
        {"Employee": "Stella", "Day": "Thursday", "Time Window": "Before 3:30pm"},
        {"Employee": "Stella", "Day": "Friday", "Time Window": "Before 3:30pm"},
        {"Employee": "Aimi", "Day": "Wednesday", "Time Window": "All Day (Uni)"},
        {"Employee": "Ainsley Mactier", "Day": "Monday", "Time Window": "After 5:00pm"},
        {"Employee": "Ainsley Mactier", "Day": "Friday", "Time Window": "After 5:00pm"},
        {"Employee": "Jude", "Day": "Sunday", "Time Window": "Before 12:00pm"},
    ])
    st.session_state.manual_unavailability = standardize_unavailability_df(load_persisted_df("unavailability.csv", default_unavail))
    save_persisted_df(st.session_state.manual_unavailability, "unavailability.csv")

if 'manual_requirements' not in st.session_state:
    default_req = pd.DataFrame([
        {"Shift": "7:30am-12:30pm", "Monday": "2", "Tuesday": "2", "Wednesday": "2", "Thursday": "2", "Friday": "2", "Saturday": "0", "Sunday": "0"},
        {"Shift": "12:30pm-5:30pm", "Monday": "1", "Tuesday": "1", "Wednesday": "1", "Thursday": "1", "Friday": "1", "Saturday": "2", "Sunday": "2"},
    ])
    st.session_state.manual_requirements = load_persisted_df("requirements.csv", default_req)

if 'manual_fixed' not in st.session_state:
    default_fixed = pd.DataFrame([
        {"Employee": "Elizabeth", "Monday": "7:30am-12:30pm", "Tuesday": "", "Wednesday": "", "Thursday": "", "Friday": "", "Saturday": "", "Sunday": ""},
        {"Employee": "Aroha", "Monday": "6:00am-1:00pm", "Tuesday": "6:00am-1:00pm", "Wednesday": "6:00am-1:00pm", "Thursday": "", "Friday": "", "Saturday": "6:00am-2:00pm", "Sunday": "6:00am-11:00am"},
    ])
    st.session_state.manual_fixed = load_persisted_df("fixed.csv", default_fixed)
if st.session_state.manual_fixed is not None:
    st.session_state.manual_fixed = st.session_state.manual_fixed.replace(["off", "Off", "OFF", "None", "none", "nan", "NaN", None], "")

# --- ROLE-BASED TAB NAVIGATION ---
is_manager = (st.session_state.user_role == "Manager")

if is_manager:
    tab_home, tab_gen, tab_emp, tab_unavail, tab_req, tab_fixed = st.tabs([
        "🏠 Home / Executive Dashboard",
        "⚡ Weekly Roster Generator",
        "👥 Staff Members", 
        "🚫 Unavailability", 
        "📋 Daily Requirements", 
        "📌 Fixed Shifts"
    ])
else:
    # Employee sees 3 tabs: Current Roster (1st), Personal Information (2nd), Availability (3rd)
    tab_my_current_roster, tab_my_info, tab_my_avail = st.tabs([
        "📅 Current Roster",
        "📋 Personal Information Form",
        "📅 My Availability & Constraints"
    ])

def render_employee_current_roster_tab(user_key):
    user_info = user_profiles.get(user_key, {})
    emp_name = user_info.get("employee_name", user_key)
    
    today = datetime.now().date()
    
    past_rosters = list_finalized_rosters()
    matched_roster = None
    matched_start_date = None
    
    # Compare current date against roster week periods (Monday to Sunday)
    for r in past_rosters:
        try:
            s_dt = datetime.strptime(r["date_str"], "%Y-%m-%d").date()
            e_dt = s_dt + timedelta(days=6)
            if s_dt <= today <= e_dt:
                matched_roster = r
                matched_start_date = s_dt
                break
        except:
            pass
            
    # Fallback to most recent finalized roster if no exact match for current week
    if not matched_roster and past_rosters:
        matched_roster = past_rosters[0]
        try:
            matched_start_date = datetime.strptime(matched_roster["date_str"], "%Y-%m-%d").date()
        except:
            matched_start_date = today
            
    st.markdown("""
    <div style="background: linear-gradient(135deg, #0e2b26 0%, #1a4d43 100%); padding: 20px; border-radius: 16px; border: 2px solid #e5a93c; box-shadow: 0 8px 24px rgba(0,0,0,0.4); margin-bottom: 20px;">
        <h2 style="color: #f7d594 !important; margin-top: 0; font-size: 1.6rem; font-weight: 800;">📅 Bakery Weekly Staff Roster</h2>
        <p style="color: #ffffff !important; font-size: 0.95rem; margin-bottom: 0;">Welcome! View your scheduled shifts and General Retail Industry Award 2020 break entitlements below.</p>
    </div>
    """, unsafe_allow_html=True)
    
    if matched_roster:
        roster_df = load_finalized_roster(matched_roster["csv_filename"])
        if roster_df is not None and not roster_df.empty:
            end_date = matched_start_date + timedelta(days=6)
            
            # Header week badge
            st.markdown(f"""
            <div style="background: rgba(9, 32, 28, 0.6); padding: 10px 16px; border-radius: 10px; border-left: 4px solid #e5a93c; margin-bottom: 15px;">
                <span style="color: #e5a93c; font-weight: 800; font-size: 1.05rem;">🗓️ Week Period:</span>
                <span style="color: #ffffff; font-weight: 700; font-size: 1.05rem; margin-left: 8px;">Monday {matched_start_date.strftime('%d/%m/%Y')} — Sunday {end_date.strftime('%d/%m/%Y')}</span>
            </div>
            """, unsafe_allow_html=True)
            
            # Personalized Employee Shift Callout Summary
            emp_c = find_column(roster_df, ["employee", "name", "staff"], "Employee")
            if emp_c in roster_df.columns:
                matched_rows = roster_df[roster_df[emp_c].astype(str).str.strip().str.lower() == emp_name.strip().lower()]
                if not matched_rows.empty:
                    emp_row = matched_rows.iloc[0]
                    shifts_list = []
                    days_arr = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"]
                    for d in days_arr:
                        if d in emp_row:
                            val = str(emp_row[d]).strip()
                            if val and val.lower() not in ["off", "unavailable", " unavailable", "none", "nan"]:
                                shifts_list.append(f"<b>{d}:</b> {val}")
                    
                    if shifts_list:
                        shifts_str = " &nbsp;|&nbsp; ".join(shifts_list)
                        st.markdown(f"""
                        <div style="background: rgba(229, 169, 60, 0.15); border: 1px solid #e5a93c; padding: 12px 18px; border-radius: 12px; margin-bottom: 20px;">
                            <div style="color: #f7d594; font-weight: 800; font-size: 1.05rem;">👋 Hello {emp_name}, your shifts for this week:</div>
                            <div style="color: #ffffff; font-size: 0.95rem; margin-top: 6px;">{shifts_str}</div>
                        </div>
                        """, unsafe_allow_html=True)
                    else:
                        st.markdown(f"""
                        <div style="background: rgba(229, 169, 60, 0.1); border: 1px solid rgba(229, 169, 60, 0.4); padding: 12px 18px; border-radius: 12px; margin-bottom: 20px;">
                            <div style="color: #f7d594; font-weight: 700; font-size: 0.95rem;">👋 Hello {emp_name}, you have no shifts scheduled for this week.</div>
                        </div>
                        """, unsafe_allow_html=True)

            # Full Schedule Display
            st.markdown("### 📋 Full Team Schedule")
            st.dataframe(roster_df, use_container_width=True)
            
            # Download XLSX Button
            excel_bytes = build_roster_excel_bytes(roster_df, matched_start_date)
            dl_filename = f"Team_Roster_{matched_start_date.strftime('%d.%m.%Y')}.xlsx"
            st.download_button(
                label=f"📥 Download Roster for {matched_start_date.strftime('%d/%m/%Y')} (.XLSX)",
                data=excel_bytes,
                file_name=dl_filename,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="btn_emp_dl_current_roster",
                use_container_width=True
            )

            # Option to select other past weeks
            if len(past_rosters) > 1:
                st.markdown("<br><hr>", unsafe_allow_html=True)
                st.markdown("**🔍 Select Other Roster Weeks to View:**")
                roster_opts = {r["label"]: r for r in past_rosters}
                sel_label = st.selectbox("View another week's schedule:", list(roster_opts.keys()), key="emp_select_other_roster")
                if sel_label != matched_roster["label"]:
                    sel_r = roster_opts[sel_label]
                    other_df = load_finalized_roster(sel_r["csv_filename"])
                    if other_df is not None:
                        st.markdown(f"**Viewing Schedule for: `{sel_label}`**")
                        st.dataframe(other_df, use_container_width=True)
    else:
        st.info("ℹ️ No finalized rosters have been published yet by management. Please check back soon!")

    # --- GENERAL RETAIL INDUSTRY AWARD 2020 BREAK RULES SECTION ---
    st.markdown("<br><hr>", unsafe_allow_html=True)
    st.markdown("""
    <div style="background: linear-gradient(135deg, #081d19 0%, #16443c 100%); padding: 14px 20px; border-radius: 12px; border: 1px solid rgba(229, 169, 60, 0.4); margin-top: 15px; margin-bottom: 15px;">
        <h3 style="color: #e5a93c !important; margin: 0; font-size: 1.25rem; font-weight: 800;">☕ General Retail Industry Award 2020 — Shift Break Entitlements</h3>
    </div>
    """, unsafe_allow_html=True)
    
    col_b1, col_b2 = st.columns([1.5, 1])
    with col_b1:
        break_table_html = """
        <table style="width:100%; border-collapse: collapse; margin-top: 5px; font-size: 0.9rem; color: #ffffff;">
            <thead>
                <tr style="background-color: #1F4E78; text-align: center; font-weight: bold; color: #ffffff;">
                    <th style="padding: 8px; border: 1px solid #336699;">Shift Duration</th>
                    <th style="padding: 8px; border: 1px solid #336699;">Paid Rest Break (10m)</th>
                    <th style="padding: 8px; border: 1px solid #336699;">Unpaid Meal Break (30m)</th>
                    <th style="padding: 8px; border: 1px solid #336699;">Total Entitlement</th>
                </tr>
            </thead>
            <tbody>
                <tr style="background-color: rgba(255,255,255,0.05); text-align: center;">
                    <td style="padding: 8px; border: 1px solid #2a5a50;">Less than 4 hours</td>
                    <td style="padding: 8px; border: 1px solid #2a5a50;">No break</td>
                    <td style="padding: 8px; border: 1px solid #2a5a50;">No break</td>
                    <td style="padding: 8px; border: 1px solid #2a5a50;">No breaks required</td>
                </tr>
                <tr style="background-color: rgba(255,255,255,0.02); text-align: center;">
                    <td style="padding: 8px; border: 1px solid #2a5a50;">4 hrs up to 5 hrs</td>
                    <td style="padding: 8px; border: 1px solid #2a5a50;">1 x 10 min</td>
                    <td style="padding: 8px; border: 1px solid #2a5a50;">No break</td>
                    <td style="padding: 8px; border: 1px solid #2a5a50;">10 min Paid</td>
                </tr>
                <tr style="background-color: rgba(255,255,255,0.05); text-align: center;">
                    <td style="padding: 8px; border: 1px solid #2a5a50;">5 hrs up to 7 hrs</td>
                    <td style="padding: 8px; border: 1px solid #2a5a50;">1 x 10 min</td>
                    <td style="padding: 8px; border: 1px solid #2a5a50;">1 x 30 min</td>
                    <td style="padding: 8px; border: 1px solid #2a5a50;">10m Paid + 30m Unpaid</td>
                </tr>
                <tr style="background-color: rgba(255,255,255,0.02); text-align: center;">
                    <td style="padding: 8px; border: 1px solid #2a5a50;">7 hrs up to 10 hrs</td>
                    <td style="padding: 8px; border: 1px solid #2a5a50;">2 x 10 min</td>
                    <td style="padding: 8px; border: 1px solid #2a5a50;">1 x 30 min</td>
                    <td style="padding: 8px; border: 1px solid #2a5a50;">20m Paid + 30m Unpaid</td>
                </tr>
                <tr style="background-color: rgba(255,255,255,0.05); text-align: center;">
                    <td style="padding: 8px; border: 1px solid #2a5a50;">10 hours or more</td>
                    <td style="padding: 8px; border: 1px solid #2a5a50;">2 x 10 min</td>
                    <td style="padding: 8px; border: 1px solid #2a5a50;">2 x 30 min</td>
                    <td style="padding: 8px; border: 1px solid #2a5a50;">20m Paid + 60m Unpaid</td>
                </tr>
            </tbody>
        </table>
        """
        st.markdown(break_table_html, unsafe_allow_html=True)
        
    with col_b2:
        st.markdown("""
        <div style="background: rgba(9, 32, 28, 0.6); border: 1px solid rgba(229, 169, 60, 0.4); border-radius: 12px; padding: 15px; height: 100%;">
            <h4 style="color: #e5a93c !important; margin-top: 0;">📌 Key Shift Break Rules</h4>
            <ul style="margin-bottom: 0; padding-left: 18px; font-size: 0.88rem; color: #ffffff !important;">
                <li><b>Paid vs Unpaid:</b> Rest breaks (10 min) are paid by employer. Meal breaks (30 min) are unpaid.</li>
                <li><b>Start & End Buffer:</b> No break may be taken during the first or last hour of work.</li>
                <li><b>5-Hour Limit:</b> An unpaid meal break must be taken no later than after 5 hours of continuous work.</li>
                <li><b>No Combining:</b> Rest breaks and meal breaks cannot be combined together.</li>
            </ul>
        </div>
        """, unsafe_allow_html=True)

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
        <h2 style="color: #e5a93c !important; margin: 0;">📅 My Unavailability & Constraint Manager</h2>
        <p style="color: #c8e6e0 !important; margin-top: 4px; margin-bottom: 0;">Logged in Employee: <b>{emp_name}</b>. Follow the 2-Step form below to log your unavailable days & date range.</p>
    </div>
    """, unsafe_allow_html=True)
    
    from datetime import date
    days_list = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"]
    
    with st.form(key=f"form_2step_unavail_{user_key}"):
        st.markdown("### Step 1: Select Unavailability per Day (7 Day Rows)")
        
        day_inputs = {}
        for d in days_list:
            col_check, col_type, col_detail = st.columns([1.2, 1.2, 2.5])
            with col_check:
                is_checked = st.checkbox(f"🔴 {d}", key=f"chk_{d}_{user_key}")
            with col_type:
                time_type = st.radio(f"Time for {d}", ["All Day", "Specific Time"], key=f"rad_{d}_{user_key}", label_visibility="collapsed")
            with col_detail:
                # Input text box ALWAYS visible for every row
                spec_time_input = st.text_input(
                    f"Window for {d}", 
                    value="", 
                    placeholder="Type window e.g. Before 3:30pm, After 5pm, 7:30am-12:30pm", 
                    key=f"input_spec_{d}_{user_key}", 
                    label_visibility="collapsed"
                ).strip()
                
                if time_type == "Specific Time":
                    spec_time = spec_time_input if spec_time_input else "Specific Time"
                else:
                    spec_time = "All Day"
                    
            day_inputs[d] = {"checked": is_checked, "time_type": time_type, "spec_time": spec_time}
            st.markdown("<hr style='margin: 2px 0; border-color: rgba(255,255,255,0.08);'>", unsafe_allow_html=True)

        st.markdown("<br>", unsafe_allow_html=True)
        st.markdown("### Step 2: Input Date Range")
        col_start, col_end = st.columns(2)
        with col_start:
            start_date = st.date_input("Start Date", date.today(), key=f"step2_start_{user_key}")
        with col_end:
            end_date = st.date_input("End Date", date.today() + timedelta(days=90), key=f"step2_end_{user_key}")

        submit_2step = st.form_submit_button("📌 Save Unavailability Constraints")
        
        if submit_2step:
            checked_days = [d for d in days_list if day_inputs[d]["checked"]]
            if not checked_days:
                st.warning("⚠️ Please check at least one day row in Step 1 before saving.")
            else:
                date_note = f"From {start_date.strftime('%d/%m/%Y')} to {end_date.strftime('%d/%m/%Y')}"
                df_curr = st.session_state.manual_unavailability.copy()
                emp_col = find_column(df_curr, ["employee", "name", "staff"], "Employee")
                day_col = find_column(df_curr, ["day", "date", "weekday"], "Day")
                win_col = find_column(df_curr, ["time window", "window", "time", "unavailability"], "Time Window")
                
                new_entries = []
                for d in checked_days:
                    t_val = day_inputs[d]["spec_time"]
                    final_win_str = f"{t_val} ({date_note})"
                    new_entries.append({emp_col: emp_name, day_col: d, win_col: final_win_str})
                    
                df_updated = pd.concat([df_curr, pd.DataFrame(new_entries)], ignore_index=True)
                st.session_state.manual_unavailability = df_updated
                save_persisted_df(df_updated, "unavailability.csv")
                st.success(f"✅ Saved unavailability for {len(checked_days)} day(s) ({', '.join(checked_days)}) from {start_date.strftime('%d/%m/%Y')} to {end_date.strftime('%d/%m/%Y')}!")
                st.rerun()

    st.markdown("---")
    st.markdown("### 📋 My Active Logged Constraints")
    df_curr = st.session_state.manual_unavailability.copy()
    emp_col = find_column(df_curr, ["employee", "name", "staff"], "Employee")
    day_col = find_column(df_curr, ["day", "date", "weekday"], "Day")
    win_col = find_column(df_curr, ["time window", "window", "time", "unavailability"], "Time Window")
    
    mask = df_curr[emp_col].astype(str).str.strip().str.lower() == emp_name.lower()
    user_rows = df_curr[mask].copy()
    
    if user_rows.empty:
        st.info("🎉 You have no logged unavailability constraints. You are listed as Available all week!")
    else:
        for idx, row in user_rows.iterrows():
            day_val = row.get(day_col, "")
            win_val = row.get(win_col, "")
            c1, c2, c3 = st.columns([2, 3.5, 1])
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

# Helper function to render Visual Monthly Calendar Grid with Color-Coded Event Badges for Manager
def render_team_monthly_calendar_grid():
    import calendar
    from datetime import date
    
    st.markdown("""
    <div style="background: linear-gradient(135deg, #081d19 0%, #16443c 100%); padding: 14px 22px; border-radius: 14px 14px 0 0; color: #ffffff !important; font-weight: 800; font-size: 1.25rem; letter-spacing: 0.3px; border: 2px solid #e5a93c; border-bottom: none; margin-top: 15px;">
        📅 Bakery Team Monthly Calendar & Unavailability Grid
    </div>
    """, unsafe_allow_html=True)
    
    # Month / Year Selection Controls
    col_nav1, col_nav2, col_nav3 = st.columns([1, 2, 1])
    with col_nav2:
        selected_month_str = st.selectbox(
            "Select Month & Year View", 
            ["August 2026", "September 2026", "October 2026", "November 2026", "December 2026", "January 2027"],
            index=0,
            key="cal_grid_month_picker"
        )
        
    month_names = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"]
    parts = selected_month_str.split()
    sel_month_name = parts[0]
    sel_year = int(parts[1])
    sel_month = month_names.index(sel_month_name) + 1
    
    # Render 7 day headers (Sun to Sat)
    days_hdr = ["Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"]
    cols_hdr = st.columns(7)
    for i, h in enumerate(days_hdr):
        with cols_hdr[i]:
            st.markdown(f'<div style="text-align: center; font-weight: 800; color: #e5a93c; background: #0c2b25; padding: 8px; border-radius: 8px; font-size: 0.95rem;">{h}</div>', unsafe_allow_html=True)
            
    # Set calendar to Sunday-first (firstweekday=6)
    cal = calendar.Calendar(firstweekday=6)
    month_days = cal.monthdayscalendar(sel_year, sel_month)
    
    default_unavail = pd.DataFrame([
        {"Employee": "Elizabeth", "Day": "Saturday", "Time Window": "All Day"},
        {"Employee": "Elizabeth", "Day": "Sunday", "Time Window": "All Day"},
        {"Employee": "Stella", "Day": "Monday", "Time Window": "Before 3:30pm"},
        {"Employee": "Stella", "Day": "Tuesday", "Time Window": "Before 3:30pm"},
        {"Employee": "Stella", "Day": "Thursday", "Time Window": "Before 3:30pm"},
        {"Employee": "Stella", "Day": "Friday", "Time Window": "Before 3:30pm"},
        {"Employee": "Aimi", "Day": "Wednesday", "Time Window": "All Day (Uni)"},
        {"Employee": "Ainsley Mactier", "Day": "Monday", "Time Window": "After 5:00pm"},
        {"Employee": "Ainsley Mactier", "Day": "Friday", "Time Window": "After 5:00pm"},
        {"Employee": "Jude", "Day": "Sunday", "Time Window": "Before 12:00pm"},
    ])

    unavail_df = st.session_state.get("manual_unavailability", None)
    if unavail_df is None or not isinstance(unavail_df, pd.DataFrame) or unavail_df.empty:
        unavail_df = default_unavail.copy()
        st.session_state.manual_unavailability = unavail_df
        save_persisted_df(unavail_df, "unavailability.csv")

    emp_col = find_column(unavail_df, ["employee", "name", "staff", "user", "person", "employee name", "staff name"], "Employee")
    day_col = find_column(unavail_df, ["day", "date", "weekday", "when", "day of week", "unavailable date"], "Day")
    win_col = find_column(unavail_df, ["time window", "window", "time", "unavailability", "reason", "constraint", "time constraint", "notes", "details"], "Time Window")

    # Positional fallback if named column matching is empty
    if unavail_df is not None and hasattr(unavail_df, "columns") and len(unavail_df.columns) > 0:
        if not emp_col or emp_col not in unavail_df.columns:
            emp_col = unavail_df.columns[0]
        if (not day_col or day_col not in unavail_df.columns) and len(unavail_df.columns) >= 2:
            day_col = unavail_df.columns[1]
        if (not win_col or win_col not in unavail_df.columns) and len(unavail_df.columns) >= 3:
            win_col = unavail_df.columns[2]

    # Filter out empty or 'nan' rows
    clean_rows = []
    if emp_col and day_col and win_col and not unavail_df.empty:
        for _, r in unavail_df.iterrows():
            e = str(r.get(emp_col, "")).strip()
            d = str(r.get(day_col, "")).strip()
            w = str(r.get(win_col, "")).strip()
            if e and e.lower() not in ["nan", "none", ""] and d and d.lower() not in ["nan", "none", ""]:
                clean_rows.append({emp_col: e, day_col: d, win_col: w})
                
    if not clean_rows:
        clean_rows = default_unavail.to_dict('records')
        emp_col = "Employee"
        day_col = "Day"
        win_col = "Time Window"
        
    clean_unavail_df = pd.DataFrame(clean_rows)

    # Build name_map for matching employee names
    emp_df = st.session_state.get("manual_employees", None)
    name_map = {}
    if emp_df is not None and hasattr(emp_df, "columns") and len(emp_df.columns) > 0:
        name_c = find_column(emp_df, ["name", "employee", "staff"])
        if name_c and name_c in emp_df.columns:
            for n in emp_df[name_c].dropna():
                n_str = str(n).strip()
                if n_str and n_str.lower() not in ["nan", "none", ""]:
                    name_map[n_str.lower()] = n_str

    for week in month_days:
        week_cols = st.columns(7)
        for i, day_num in enumerate(week):
            with week_cols[i]:
                if day_num == 0:
                    st.markdown('<div style="min-height: 105px; background: rgba(255,255,255,0.02); border-radius: 8px; margin-top: 4px;"></div>', unsafe_allow_html=True)
                else:
                    dt_obj = date(sel_year, sel_month, day_num)
                    
                    chips_html = []
                    
                    # Unavailability Chips (Employee Names who are not available)
                    for _, urow in clean_unavail_df.iterrows():
                        u_emp = str(urow.get(emp_col, "")).strip()
                        u_day = str(urow.get(day_col, "")).strip()
                        u_win = str(urow.get(win_col, "")).strip()
                        
                        if u_emp and u_day and is_unavail_applicable_to_date(dt_obj, u_day, u_win):
                            matched = find_matching_employee(u_emp, name_map) if name_map else u_emp
                            display_name = matched if matched else u_emp
                            
                            badge_win = clean_win_display(u_win)
                            tooltip_win = u_win if (u_win and str(u_win).strip().lower() not in ["nan", "none", "nat", "null", ""]) else "All Day"
                            
                            if "all day" in tooltip_win.lower():
                                chips_html.append(f'<div style="background-color: #e53e3e; color: white; padding: 3px 8px; border-radius: 6px; font-size: 0.75rem; font-weight: 700; margin-bottom: 3px; white-space: nowrap; overflow: hidden; text-overflow: ellipsis;" title="{display_name}: {tooltip_win}">🔴 {display_name} ({badge_win})</div>')
                            else:
                                chips_html.append(f'<div style="background-color: #dd6b20; color: white; padding: 3px 8px; border-radius: 6px; font-size: 0.75rem; font-weight: 700; margin-bottom: 3px; white-space: nowrap; overflow: hidden; text-overflow: ellipsis;" title="{display_name}: {tooltip_win}">🟨 {display_name} ({badge_win})</div>')
                                    
                    chips_block = "".join(chips_html) if chips_html else '<div style="color: #718096; font-size: 0.75rem; font-weight: 600; padding: 4px 0;">Clear (All Available)</div>'
                    
                    st.markdown(f"""
                    <div style="min-height: 110px; background: #11362f; border: 1px solid #1f5c50; border-radius: 8px; padding: 6px; margin-top: 4px;">
                        <div style="font-weight: 800; font-size: 0.85rem; color: #e5a93c; border-bottom: 1px solid #1f5c50; margin-bottom: 4px; padding-bottom: 2px;">{day_num}</div>
                        {chips_block}
                    </div>
                    """, unsafe_allow_html=True)

# IF EMPLOYEE, RENDER 3 TABS (CURRENT ROSTER 1ST, PERSONAL INFO 2ND, AVAILABILITY CALENDAR 3RD)
if not is_manager:
    with tab_my_current_roster:
        render_employee_current_roster_tab(st.session_state.logged_in_user)
    with tab_my_info:
        render_confidential_profile_form(st.session_state.logged_in_user)
    with tab_my_avail:
        render_employee_availability_manager(st.session_state.logged_in_user)


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
    if not shift_str or "unavailable" in str(shift_str).strip().lower() or str(shift_str).strip().lower() in ["off", "nan", ""]:
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
        if not raw_name or raw_name.lower() in ["nan", ""] or "demo" in raw_name.lower():
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
    roster_output = {emp["Name"]: {day: "" for day in days_of_week} for emp in active_employees}
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
                        val = str(fix_row.get(day_col, "")).strip()
                        if val.lower() not in ["off", "nan", "none", ""]:
                            roster_output[name][day] = val
                            weekly_shifts_count[name] += 1
                            if name == "Elizabeth" and day not in ["Saturday", "Sunday"]:
                                elizabeth_weekday_shifts += 1

    # 2. Check unavailability
    unavail_name_col = find_column(unavailability, ["employee", "name", "employee name", "staff name", "staff", "user", "person"])
    unavail_day_col = find_column(unavailability, ["day", "date", "weekday", "when", "day of week", "unavailable date"])
    unavail_window_col = find_column(unavailability, ["time window", "window", "time", "unavailability", "reason", "constraint", "time constraint", "notes", "details"])

    unavail_map = {}
    if unavail_name_col and unavail_day_col and unavail_window_col:
        for _, un_row in unavailability.iterrows():
            raw_unavail_name = str(un_row.get(unavail_name_col, "")).strip()
            unavail_day = str(un_row.get(unavail_day_col, "")).strip()
            window = str(un_row.get(unavail_window_col, "All Day")).strip()
            matched_name = find_matching_employee(raw_unavail_name, name_map)
            if matched_name and unavail_day:
                for idx, day in enumerate(days_of_week):
                    day_date = start_dt + timedelta(days=idx)
                    if is_unavail_applicable_to_date(day_date, unavail_day, window):
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
                            if roster_output[name][day] in ["", "off"]:
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
            if val not in ["", "off", " unavailable", "none", "nan"]:
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
                
                if roster_output[name][day] not in ["", "off"]:
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
        norm_name = name.lower()
        row = {"Employee": name}
        for day in days_of_week:
            val = str(sched.get(day, "")).strip()
            key = (norm_name, day.lower())
            if not val or val.lower() in ["off", "none", "nan", "unavailable", ""] or val.lower().startswith("unavailable"):
                if key in unavail_map and unavail_map[key]:
                    clean_win = clean_win_display(unavail_map[key][0])
                    row[day] = f"Unavailable ({clean_win})"
                else:
                    row[day] = ""
            else:
                row[day] = val
        roster_rows.append(row)
    res_df = pd.DataFrame(roster_rows)
    return res_df

if is_manager:
    # --- TAB 1: HOME / EXECUTIVE DASHBOARD ---
    with tab_home:
        st.markdown("""
        <div style="background: linear-gradient(135deg, #0e2b26 0%, #1a4d43 100%); padding: 20px; border-radius: 16px; border: 2px solid #e5a93c; box-shadow: 0 8px 30px rgba(0,0,0,0.4); margin-bottom: 25px;">
            <h2 style="color: #f7d594 !important; margin-top: 0; font-size: 1.8rem; font-weight: 800;">🏠 Executive Admin Command Center</h2>
            <p style="color: #ffffff !important; font-size: 1.05rem; margin-bottom: 0;">Select published weekly rosters to view, edit shifts, review real-time payroll breakdowns, and analyze historical financial trends.</p>
        </div>
        """, unsafe_allow_html=True)

        past_rosters = list_finalized_rosters()
        if past_rosters:
            roster_options = {r["label"]: r for r in past_rosters}
            selected_label = st.selectbox("Select Finalized Roster Week to Display & Edit:", list(roster_options.keys()), key="home_select_past_roster")
            selected_info = roster_options[selected_label]
            
            archived_df = load_finalized_roster(selected_info["csv_filename"])
            if archived_df is not None and not archived_df.empty:
                st.markdown(f"""
                <div style="background: linear-gradient(135deg, #081d19 0%, #16443c 100%); padding: 14px 22px; border-radius: 12px 12px 0 0; border: 2px solid #e5a93c; border-bottom: none; display: flex; align-items: center; justify-content: space-between; flex-wrap: wrap; gap: 10px; margin-top: 15px;">
                    <div style="display: flex; align-items: center; gap: 10px; flex-wrap: wrap;">
                        <span style="font-size: 1.3rem;">📅</span>
                        <span style="color: #ffffff !important; font-size: 1.15rem; font-weight: 800; letter-spacing: 0.3px;">Published Roster Schedule for:</span>
                        <span style="color: #f7d594 !important; font-size: 1.1rem; font-weight: 700;">{selected_label}</span>
                    </div>
                    <span style="background: rgba(229, 169, 60, 0.2); color: #f7d594 !important; font-size: 0.78rem; font-weight: 700; padding: 4px 12px; border-radius: 20px; border: 1px solid #e5a93c; letter-spacing: 0.3px;">✏️ Live & Editable</span>
                </div>
                """, unsafe_allow_html=True)
                
                # Live-editable dataframe for displayed roster
                edited_archived_df = st.data_editor(archived_df, num_rows="dynamic", key=f"edit_home_roster_{selected_info['date_str']}")
                
                try:
                    dt = datetime.strptime(selected_info["date_str"], "%Y-%m-%d").date()
                except:
                    dt = datetime.now().date()
                    
                # Action Buttons Row (Save Edits, Download, Delete)
                col_act1, col_act2, col_act3 = st.columns([1.2, 1.2, 1])
                with col_act1:
                    if st.button("💾 SAVE CHANGES TO ROSTER", key=f"btn_save_home_{selected_info['date_str']}", use_container_width=True):
                        save_finalized_roster(edited_archived_df, dt)
                        st.success(f"🎉 Changes to roster for week {selected_info['date_str']} successfully saved to disk!")
                        st.rerun()
                
                with col_act2:
                    archived_excel_bytes = build_roster_excel_bytes(edited_archived_df, dt)
                    past_filename = f"Team_Roster_{dt.strftime('%d.%m.%Y')}.xlsx"
                    st.download_button(
                        label=f"📥 Download {selected_info['date_str']} Roster (.XLSX)",
                        data=archived_excel_bytes,
                        file_name=past_filename,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        key=f"btn_dl_home_{selected_info['date_str']}",
                        use_container_width=True
                    )
                    
                with col_act3:
                    if st.button(f"🗑️ Delete Roster", key=f"btn_del_home_{selected_info['date_str']}", use_container_width=True):
                        if delete_finalized_roster(selected_info["date_str"]):
                            st.success(f"🗑️ Finalized Roster for {selected_info['date_str']} permanently deleted!")
                            st.rerun()

                # Upload/Replace file for selected past roster week
                st.markdown("<br>", unsafe_allow_html=True)
                up_past_file = st.file_uploader(f"📤 Replace Roster File for Week {selected_info['date_str']} (.xlsx / .csv)", type=["xlsx", "csv"], key=f"up_home_{selected_info['date_str']}")
                if up_past_file is not None:
                    past_up_key = f"home_up_{up_past_file.name}_{up_past_file.size}_{selected_info['date_str']}"
                    if st.session_state.get("last_home_upload_key") != past_up_key:
                        df_p_up = read_excel_robust(up_past_file)
                        if df_p_up is not None and not df_p_up.empty:
                            save_finalized_roster(df_p_up, dt)
                            st.session_state.last_home_upload_key = past_up_key
                            st.success(f"🎉 Finalized roster for week {selected_info['date_str']} updated via file upload!")
                            st.rerun()

                # Real-Time Wage, Tax & Super Breakdown for Displayed Roster
                st.markdown("<br>", unsafe_allow_html=True)
                wages_summary = calculate_roster_wages(edited_archived_df)
                
                st.markdown("""
                <div style="background: linear-gradient(135deg, #0e2b26 0%, #1a4d43 100%); padding: 14px 20px; border-radius: 12px 12px 0 0; color: #e5a93c !important; font-weight: 800; font-size: 1.25rem; border: 2px solid #e5a93c; border-bottom: none;">
                    💰 Real-Time Wage, Tax & Super Summary
                </div>
                """, unsafe_allow_html=True)
                
                summary_cards_html = f"""
                <div style="background: rgba(8, 29, 25, 0.85); border: 2px solid #e5a93c; border-top: none; border-radius: 0 0 12px 12px; padding: 20px; margin-bottom: 20px;">
                    <div style="display: grid; grid-template-columns: repeat(auto-fit, minmax(200px, 1fr)); gap: 15px; margin-bottom: 18px;">
                        <div style="background: #0d332b; border: 1.5px solid #e5a93c; border-radius: 10px; padding: 14px; text-align: center;">
                            <div style="color: #e5a93c; font-size: 0.82rem; font-weight: 800; letter-spacing: 0.5px; text-transform: uppercase;">💵 Total Gross Payroll</div>
                            <div style="color: #ffffff; font-size: 1.75rem; font-weight: 900; margin-top: 6px;">${wages_summary['total_gross']:,.2f}</div>
                        </div>
                        <div style="background: #0d332b; border: 1.5px solid #e5a93c; border-radius: 10px; padding: 14px; text-align: center;">
                            <div style="color: #f7d594; font-size: 0.82rem; font-weight: 800; letter-spacing: 0.5px; text-transform: uppercase;">🏛️ Est. PAYG Tax</div>
                            <div style="color: #ffffff; font-size: 1.75rem; font-weight: 900; margin-top: 6px;">${wages_summary['total_tax']:,.2f}</div>
                        </div>
                        <div style="background: #0d332b; border: 1.5px solid #e5a93c; border-radius: 10px; padding: 14px; text-align: center;">
                            <div style="color: #76eec6; font-size: 0.82rem; font-weight: 800; letter-spacing: 0.5px; text-transform: uppercase;">👛 Total Net Take-Home</div>
                            <div style="color: #76eec6; font-size: 1.75rem; font-weight: 900; margin-top: 6px;">${wages_summary['total_net']:,.2f}</div>
                        </div>
                        <div style="background: #0d332b; border: 1.5px solid #e5a93c; border-radius: 10px; padding: 14px; text-align: center;">
                            <div style="color: #f7d594; font-size: 0.82rem; font-weight: 800; letter-spacing: 0.5px; text-transform: uppercase;">🏦 Super (12.5% SG)</div>
                            <div style="color: #ffffff; font-size: 1.75rem; font-weight: 900; margin-top: 6px;">${wages_summary['total_super']:,.2f}</div>
                        </div>
                    </div>
                    <div style="background: rgba(229, 169, 60, 0.12); border: 1px solid rgba(229, 169, 60, 0.4); border-radius: 8px; padding: 10px 16px; text-align: center; color: #ffffff; font-size: 1.05rem;">
                        ⏱️ <b>Total Paid Hours:</b> <span style="color:#e5a93c; font-weight:800;">{wages_summary['total_hours']} hrs</span> &nbsp;&nbsp;|&nbsp;&nbsp; 📊 <b>Average Hourly Rate:</b> <span style="color:#e5a93c; font-weight:800;">${wages_summary['avg_hourly_rate']:.2f} / hr</span>
                    </div>
                </div>
                """
                st.markdown(summary_cards_html, unsafe_allow_html=True)
                
                st.markdown("#### 👥 Staff Earnings & Super Breakdown Table")
                if not wages_summary["breakdown_df"].empty:
                    st.dataframe(wages_summary["breakdown_df"], use_container_width=True, hide_index=True)
        else:
            st.info("ℹ️ No finalized rosters stored yet. Go to the '⚡ Weekly Roster Generator' tab to generate, upload, and finalize weekly schedules.")

        # --- HISTORICAL PROGRESS LINE GRAPH ---
        st.markdown("<br><hr>", unsafe_allow_html=True)
        st.markdown("""
        <div style="background: linear-gradient(135deg, #0e2b26 0%, #1a4d43 100%); padding: 14px 20px; border-radius: 12px; color: #e5a93c !important; font-weight: 800; font-size: 1.25rem; border: 1px solid rgba(229, 169, 60, 0.4); margin-bottom: 15px;">
            📈 Payroll, Tax & Super Progress Over Time (Historical Trend Graph)
        </div>
        """, unsafe_allow_html=True)
        
        trend_df = build_payroll_historical_trend()
        if not trend_df.empty and len(trend_df) >= 1:
            st.markdown("Historical trend analysis of **Gross Payroll**, **Est. PAYG Tax**, **Net Take-Home**, and **Super (12.5% SG)** across all finalized weekly rosters:")
            chart_df = trend_df.set_index("Roster Week")
            st.line_chart(chart_df, use_container_width=True)
            st.dataframe(trend_df, use_container_width=True, hide_index=True)
        else:
            st.info("ℹ️ Finalize at least one weekly roster to view the historical progress line graph!")

    # --- TAB 2: WEEKLY ROSTER GENERATOR ---
    with tab_gen:
        st.markdown("""
        <div style="background: rgba(9, 32, 28, 0.7); border: 2px solid #e5a93c; border-radius: 16px; padding: 25px; margin-bottom: 25px; box-shadow: 0 8px 30px rgba(0,0,0,0.4);">
            <h2 style="color: #f7d594 !important; margin-top: 0; font-size: 1.8rem; font-weight: 800;">⚡ Weekly Roster Generator</h2>
            <p style="color: #ffffff !important; font-size: 1.05rem; margin-bottom: 0;">Configure your target week period below and hit the <b>Generate Weekly Roster</b> button to instantly build an award-compliant bakery schedule.</p>
        </div>
        """, unsafe_allow_html=True)
        
        col1, col2 = st.columns([1, 1])
        with col1:
            start_date = st.date_input("🗓️ Roster Start Date (Monday)", datetime.now() + timedelta(days=(0 - datetime.now().weekday())), key="gen_start_date")
            
            st.markdown('<div class="hero-generate-btn">', unsafe_allow_html=True)
            if st.button("🚀 GENERATE WEEKLY ROSTER", key="btn_hero_generate"):
                with st.spinner("Calculating optimal bakery roster locally..."):
                    try:
                        emp_data = st.session_state.manual_employees
                        unavail_data = st.session_state.manual_unavailability
                        req_data = st.session_state.manual_requirements
                        fixed_data = st.session_state.manual_fixed
                        
                        roster_out_df = solve_roster(emp_data, unavail_data, req_data, fixed_data, start_date)
                        df_clean = roster_out_df.replace(["off", "Off", "OFF", "None", "none", "nan", "NaN", None], "").fillna("")
                        emp_c = find_column(df_clean, ["employee", "name", "staff"], "Employee")
                        if emp_c in df_clean.columns:
                            df_clean = df_clean[~df_clean[emp_c].astype(str).str.strip().str.lower().isin(["", "none", "nan"])].reset_index(drop=True)
                        st.session_state.final_roster_df = df_clean
                        st.success("🎉 Weekly Roster successfully generated!")
                    except Exception as e:
                        st.error(f"Failed to generate roster: {e}")
            
            st.markdown("<br>", unsafe_allow_html=True)
            upload_roster_file = st.file_uploader("📤 OR Upload Existing Roster File (.xlsx / .csv)", type=["xlsx", "csv"], key="upload_roster_selected_week")
            if upload_roster_file is not None:
                file_key = f"roster_up_{upload_roster_file.name}_{upload_roster_file.size}_{start_date}"
                if st.session_state.get("last_uploaded_roster_key") != file_key:
                    df_up = read_excel_robust(upload_roster_file)
                    if df_up is not None and not df_up.empty:
                        df_clean = df_up.replace(["off", "Off", "OFF", "None", "none", "nan", "NaN", None], "").fillna("")
                        emp_c = find_column(df_clean, ["employee", "name", "staff"], "Employee")
                        if emp_c in df_clean.columns:
                            df_clean = df_clean[~df_clean[emp_c].astype(str).str.strip().str.lower().isin(["", "none", "nan"])].reset_index(drop=True)
                        st.session_state.final_roster_df = df_clean
                        st.session_state.last_uploaded_roster_key = file_key
                        st.success(f"📁 Roster loaded for week starting {start_date.strftime('%d/%m/%Y')}!")

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

        if 'final_roster_df' in st.session_state and st.session_state.final_roster_df is not None and not st.session_state.final_roster_df.empty:
            st.markdown("<br>", unsafe_allow_html=True)
            
            st.markdown("""
            <div style="background: linear-gradient(135deg, #081d19 0%, #16443c 100%); padding: 12px 20px; border-radius: 12px 12px 0 0; color: #ffffff !important; font-weight: 800; font-size: 1.2rem; letter-spacing: 0.5px; border: 2px solid #e5a93c; border-bottom: none;">
                📅 Generated Weekly Roster Schedule (Editable)
            </div>
            """, unsafe_allow_html=True)
            
            edited_final_df = st.data_editor(st.session_state.final_roster_df, num_rows="dynamic", key="edit_generated_roster")

            # Real-Time Financial Breakdown for Generated Roster
            wages_summary_gen = calculate_roster_wages(edited_final_df)
            
            st.markdown("<br>", unsafe_allow_html=True)
            st.markdown("""
            <div style="background: linear-gradient(135deg, #0e2b26 0%, #1a4d43 100%); padding: 14px 20px; border-radius: 12px 12px 0 0; color: #e5a93c !important; font-weight: 800; font-size: 1.25rem; border: 2px solid #e5a93c; border-bottom: none;">
                💰 Real-Time Wage, Tax & Super Summary
            </div>
            """, unsafe_allow_html=True)
            
            gen_summary_cards_html = f"""
            <div style="background: rgba(8, 29, 25, 0.85); border: 2px solid #e5a93c; border-top: none; border-radius: 0 0 12px 12px; padding: 20px; margin-bottom: 20px;">
                <div style="display: grid; grid-template-columns: repeat(auto-fit, minmax(200px, 1fr)); gap: 15px; margin-bottom: 18px;">
                    <div style="background: #0d332b; border: 1.5px solid #e5a93c; border-radius: 10px; padding: 14px; text-align: center;">
                        <div style="color: #e5a93c; font-size: 0.82rem; font-weight: 800; letter-spacing: 0.5px; text-transform: uppercase;">💵 Total Gross Payroll</div>
                        <div style="color: #ffffff; font-size: 1.75rem; font-weight: 900; margin-top: 6px;">${wages_summary_gen['total_gross']:,.2f}</div>
                    </div>
                    <div style="background: #0d332b; border: 1.5px solid #e5a93c; border-radius: 10px; padding: 14px; text-align: center;">
                        <div style="color: #f7d594; font-size: 0.82rem; font-weight: 800; letter-spacing: 0.5px; text-transform: uppercase;">🏛️ Est. PAYG Tax</div>
                        <div style="color: #ffffff; font-size: 1.75rem; font-weight: 900; margin-top: 6px;">${wages_summary_gen['total_tax']:,.2f}</div>
                    </div>
                    <div style="background: #0d332b; border: 1.5px solid #e5a93c; border-radius: 10px; padding: 14px; text-align: center;">
                        <div style="color: #76eec6; font-size: 0.82rem; font-weight: 800; letter-spacing: 0.5px; text-transform: uppercase;">👛 Total Net Take-Home</div>
                        <div style="color: #76eec6; font-size: 1.75rem; font-weight: 900; margin-top: 6px;">${wages_summary_gen['total_net']:,.2f}</div>
                    </div>
                    <div style="background: #0d332b; border: 1.5px solid #e5a93c; border-radius: 10px; padding: 14px; text-align: center;">
                        <div style="color: #f7d594; font-size: 0.82rem; font-weight: 800; letter-spacing: 0.5px; text-transform: uppercase;">🏦 Super (12.5% SG)</div>
                        <div style="color: #ffffff; font-size: 1.75rem; font-weight: 900; margin-top: 6px;">${wages_summary_gen['total_super']:,.2f}</div>
                    </div>
                </div>
                <div style="background: rgba(229, 169, 60, 0.12); border: 1px solid rgba(229, 169, 60, 0.4); border-radius: 8px; padding: 10px 16px; text-align: center; color: #ffffff; font-size: 1.05rem;">
                    ⏱️ <b>Total Paid Hours:</b> <span style="color:#e5a93c; font-weight:800;">{wages_summary_gen['total_hours']} hrs</span> &nbsp;&nbsp;|&nbsp;&nbsp; 📊 <b>Average Hourly Rate:</b> <span style="color:#e5a93c; font-weight:800;">${wages_summary_gen['avg_hourly_rate']:.2f} / hr</span>
                </div>
            </div>
            """
            st.markdown(gen_summary_cards_html, unsafe_allow_html=True)
            
            st.markdown("#### 👥 Staff Earnings & Super Breakdown Table")
            if not wages_summary_gen["breakdown_df"].empty:
                st.dataframe(wages_summary_gen["breakdown_df"], use_container_width=True, hide_index=True)

            # Finalize & Export Section
            st.markdown("<br>", unsafe_allow_html=True)
            col_fin1, col_fin2 = st.columns([1.2, 1])
            with col_fin1:
                if st.button("🔒 FINALIZE WEEKLY ROSTER", key="btn_finalize_roster", use_container_width=True):
                    date_str, xlsx_filename, excel_bytes = save_finalized_roster(edited_final_df, start_date)
                    st.success(f"🎉 Weekly Roster for {start_date.strftime('%d/%m/%Y')} successfully finalized and saved online!")
            
            with col_fin2:
                excel_bytes = build_roster_excel_bytes(edited_final_df, start_date)
                file_name_out = f"Team_Roster_{start_date.strftime('%d.%m.%Y')}.xlsx"
                st.download_button(
                    label="📥 DOWNLOAD CURRENT ROSTER (.XLSX)",
                    data=excel_bytes,
                    file_name=file_name_out,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key="btn_export_excel",
                    use_container_width=True
                )
            
            # --- TAB 2: STAFF MEMBERS ---
    with tab_emp:
        st.subheader("Manage Bakery Employees")
        if "email_notice_success" in st.session_state:
            st.success(st.session_state.pop("email_notice_success"))
        if "email_notice_warning" in st.session_state:
            st.warning(st.session_state.pop("email_notice_warning"))

        emp_mode = st.radio("Upload Mode:", ["Replace current data", "Append to current data"], key="emp_upload_mode", horizontal=True)
        upload_emp = st.file_uploader("Upload EMPLOYEE LIST.xlsx (Optional)", type=["xlsx"], key="emp_upload")
        
        if upload_emp is not None:
            file_key = f"processed_{upload_emp.name}_{upload_emp.size}_{emp_mode}"
            if st.session_state.get("last_emp_file") != file_key:
                loaded = read_excel_robust(upload_emp)
                if loaded is not None:
                    loaded = cleanup_duplicate_employee_columns(loaded)
                    if emp_mode == "Replace current data":
                        st.session_state.manual_employees = loaded
                    else:
                        combined = pd.concat([st.session_state.manual_employees, loaded], ignore_index=True).drop_duplicates()
                        st.session_state.manual_employees = cleanup_duplicate_employee_columns(combined)
                    st.session_state.manual_employees = sync_user_profiles_to_employees(st.session_state.manual_employees)
                    st.session_state.last_emp_file = file_key
                    save_persisted_df(st.session_state.manual_employees, "employees.csv")
                    if "edit_employees" in st.session_state:
                        del st.session_state["edit_employees"]
                    st.rerun()
                    
        st.markdown("""
        <div style="background: linear-gradient(135deg, #081d19 0%, #16443c 100%); padding: 10px 18px; border-radius: 12px 12px 0 0; color: #ffffff !important; font-weight: 800; font-size: 1.1rem; letter-spacing: 0.3px; border: 2px solid #e5a93c; border-bottom: none; margin-top: 15px;">
            👥 Bakery Staff Members List (Editable Table)
        </div>
        """, unsafe_allow_html=True)
        if st.session_state.manual_employees is not None and not st.session_state.manual_employees.empty:
            st.session_state.manual_employees = sync_user_profiles_to_employees(st.session_state.manual_employees)
        employees_df = st.data_editor(st.session_state.manual_employees, num_rows="dynamic", key="edit_employees")
        if employees_df is not None:
            # Check if any row was deleted directly in data_editor table
            if st.session_state.manual_employees is not None and not st.session_state.manual_employees.empty:
                name_col = find_column(st.session_state.manual_employees, ["name", "employee", "staff"], "NAME")
                if name_col in st.session_state.manual_employees.columns:
                    editor_state = st.session_state.get("edit_employees", {})
                    deleted_indices = editor_state.get("deleted_rows", []) if isinstance(editor_state, dict) else []
                    
                    deleted_names = set()
                    if deleted_indices:
                        for idx in deleted_indices:
                            if 0 <= idx < len(st.session_state.manual_employees):
                                val = str(st.session_state.manual_employees.iloc[idx][name_col]).strip()
                                if val:
                                    deleted_names.add(val)
                    else:
                        old_names = [str(n).strip() for n in st.session_state.manual_employees[name_col].dropna().tolist() if str(n).strip()]
                        new_names = [str(n).strip() for n in employees_df[name_col].dropna().tolist() if str(n).strip()] if (employees_df is not None and name_col in employees_df.columns) else []
                        deleted_names = set(old_names) - set(new_names)

                    if deleted_names:
                        profiles_to_update = get_active_user_profiles()
                        profiles_changed = False
                        for del_name in deleted_names:
                            keys_to_del = []
                            for u_k, u_v in profiles_to_update.items():
                                emp_n = u_v.get("employee_name", u_k).strip()
                                full_n = u_v.get("profile", {}).get("full_name", "").strip()
                                if emp_n.lower() == del_name.lower() or full_n.lower() == del_name.lower() or u_k.lower() == del_name.lower():
                                    keys_to_del.append(u_k)
                            for k in keys_to_del:
                                del profiles_to_update[k]
                                profiles_changed = True
                        if profiles_changed:
                            save_user_profiles(profiles_to_update)

            st.session_state.manual_employees = cleanup_duplicate_employee_columns(employees_df)
            save_persisted_df(st.session_state.manual_employees, "employees.csv")

        # --- ➕ NEW EMPLOYEE ACCOUNT CREATION FORM ---
        with st.expander("➕ Add New Staff Account (Create Login & Roster Record)", expanded=False):
            send_email_chk = st.checkbox("☑️ Send Welcome Email automatically", value=True, key="chk_send_welcome_email")
            
            emp_email_val = ""
            if send_email_chk:
                emp_email_val = st.text_input("Employee Email Address", placeholder="e.g. jack.smith@outlook.com", key="input_emp_email").strip()

            with st.form(key="form_create_new_employee_account"):
                st.markdown("#### 👤 New Employee Credentials & Information")
                c1, c2 = st.columns(2)
                with c1:
                    new_name = st.text_input("Employee Full Name", placeholder="e.g. Jack Smith").strip()
                    new_user = st.text_input("Username for Login", placeholder="e.g. jack.smith or jack").strip().lower()
                    new_pass = st.text_input("Initial Password", value="TempPass123!", type="password")
                with c2:
                    new_role_level = st.text_input("Role / Position", value="Junior Team Member")
                    new_emp_type = st.selectbox("Employment Classification", ["Casual", "Part-Time", "Full-Time"], index=0)
                    new_age = st.number_input("Age", min_value=14, max_value=80, value=18)

                submit_new_emp = st.form_submit_button("🚀 Create Employee Account")

                if submit_new_emp:
                    if not new_name:
                        st.error("❌ Employee name cannot be empty.")
                    elif not new_user:
                        st.error("❌ Username cannot be empty.")
                    elif new_user in user_profiles:
                        st.error(f"❌ Username '{new_user}' already exists. Please choose a different username.")
                    elif send_email_chk and not emp_email_val:
                        st.error("❌ Please enter the Employee Email Address or uncheck 'Send Welcome Email automatically'.")
                    else:
                        user_profiles[new_user] = {
                            "username": new_user,
                            "password": new_pass if new_pass else "TempPass123!",
                            "role": "Employee",
                            "employee_name": new_name,
                            "profile": {
                                "full_name": new_name,
                                "email": emp_email_val,
                                "store": "Brumby's Pakenham",
                                "classification": new_emp_type,
                                "employment_level": new_role_level,
                                "commencement_date": datetime.now().strftime("%Y-%m-%d")
                            }
                        }
                        save_user_profiles(user_profiles)

                        st.session_state.manual_employees = sync_user_profiles_to_employees(st.session_state.manual_employees)
                        save_persisted_df(st.session_state.manual_employees, "employees.csv")

                        if "edit_employees" in st.session_state:
                            del st.session_state["edit_employees"]

                        subj, body_text = build_welcome_email_content(new_name, new_user, new_pass)
                        
                        email_sent = False
                        email_msg = ""
                        if send_email_chk and emp_email_val:
                            email_sent, email_msg = send_welcome_email_smtp(emp_email_val, new_name, new_user, new_pass)

                        if email_sent:
                            st.session_state["email_notice_success"] = f"🎉 Account created for **{new_name}**! {email_msg}"
                        else:
                            if send_email_chk and emp_email_val:
                                st.session_state["email_notice_warning"] = f"🎉 Account created for **{new_name}**! Username: `{new_user}` | Initial Password: `{new_pass}`\n\n⚠️ Email Status: {email_msg}"
                            else:
                                st.session_state["email_notice_success"] = f"🎉 Account created for **{new_name}**! Username: `{new_user}` | Initial Password: `{new_pass}`"

                        st.rerun()

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
                
                c_pw, c_del = st.columns(2)
                with c_pw:
                    with st.expander(f"🔑 Reset Password for {emp_options[selected_user_key]}"):
                        new_pw = st.text_input("New Password", type="password", key=f"tab_reset_pw_{selected_user_key}")
                        if st.button("Update Employee Password", key=f"tab_btn_reset_pw_{selected_user_key}"):
                            if new_pw.strip():
                                user_profiles[selected_user_key]["password"] = new_pw.strip()
                                save_user_profiles(user_profiles)
                                st.success(f"✅ Password for {emp_options[selected_user_key]} updated successfully!")
                            else:
                                st.error("Password cannot be empty.")

                with c_del:
                    with st.expander(f"🗑️ Delete Account for {emp_options[selected_user_key]}"):
                        st.warning(f"⚠️ Deleting this account will permanently remove **{emp_options[selected_user_key]}**'s login credentials, confidential profile, and roster records.")
                        confirm_del = st.checkbox(f"I understand, delete account for {emp_options[selected_user_key]}", key=f"chk_del_{selected_user_key}")
                        if st.button(f"🚨 Permanently Delete {emp_options[selected_user_key]}", key=f"btn_del_emp_{selected_user_key}"):
                            if confirm_del:
                                target_name = emp_options[selected_user_key]
                                # 1. Remove from user_profiles.json
                                if selected_user_key in user_profiles:
                                    del user_profiles[selected_user_key]
                                    save_user_profiles(user_profiles)
                                
                                # 2. Remove from manual_employees table
                                if st.session_state.manual_employees is not None and not st.session_state.manual_employees.empty:
                                    emp_df = st.session_state.manual_employees.copy()
                                    name_col = find_column(emp_df, ["name", "employee", "staff"], "NAME")
                                    if name_col in emp_df.columns:
                                        emp_df = emp_df[emp_df[name_col].astype(str).str.strip().str.lower() != target_name.strip().lower()].reset_index(drop=True)
                                        st.session_state.manual_employees = emp_df
                                        save_persisted_df(st.session_state.manual_employees, "employees.csv")
                                
                                # 3. Clear widget cache
                                if "edit_employees" in st.session_state:
                                    del st.session_state["edit_employees"]
                                    
                                st.success(f"✅ Account for **{target_name}** has been permanently deleted.")
                                st.rerun()
                            else:
                                st.error("Please check the confirmation box first.")

    # --- TAB 3: UNAVAILABILITY ---
    with tab_unavail:
        col_hdr_u1, col_hdr_u2 = st.columns([3, 1])
        with col_hdr_u1:
            st.subheader("Log Staff Unavailability")
        with col_hdr_u2:
            if st.button("🔄 Reset Sample Data", key="btn_reset_unavail_sample"):
                default_unavail_reset = pd.DataFrame([
                    {"Employee": "Elizabeth", "Day": "Saturday", "Time Window": "All Day"},
                    {"Employee": "Elizabeth", "Day": "Sunday", "Time Window": "All Day"},
                    {"Employee": "Stella", "Day": "Monday", "Time Window": "Before 3:30pm"},
                    {"Employee": "Stella", "Day": "Tuesday", "Time Window": "Before 3:30pm"},
                    {"Employee": "Stella", "Day": "Thursday", "Time Window": "Before 3:30pm"},
                    {"Employee": "Stella", "Day": "Friday", "Time Window": "Before 3:30pm"},
                    {"Employee": "Aimi", "Day": "Wednesday", "Time Window": "All Day (Uni)"},
                    {"Employee": "Ainsley Mactier", "Day": "Monday", "Time Window": "After 5:00pm"},
                    {"Employee": "Ainsley Mactier", "Day": "Friday", "Time Window": "After 5:00pm"},
                    {"Employee": "Jude", "Day": "Sunday", "Time Window": "Before 12:00pm"},
                ])
                st.session_state.manual_unavailability = default_unavail_reset
                save_persisted_df(default_unavail_reset, "unavailability.csv")
                st.success("Reset data!")
                st.rerun()
        
        unavail_mode = st.radio("Upload Mode:", ["Replace current data", "Append to current data"], key="unavail_upload_mode", horizontal=True)
        upload_unavail = st.file_uploader("Upload unavailability list.xlsx (Optional)", type=["xlsx"], key="unavail_upload")
        
        if upload_unavail is not None:
            file_key = f"processed_{upload_unavail.name}_{upload_unavail.size}_{unavail_mode}"
            if st.session_state.get("last_unavail_file") != file_key:
                loaded = read_excel_robust(upload_unavail)
                if loaded is not None:
                    loaded = standardize_unavailability_df(loaded)
                    if unavail_mode == "Replace current data":
                        st.session_state.manual_unavailability = loaded
                    else:
                        combined = pd.concat([st.session_state.manual_unavailability, loaded], ignore_index=True).drop_duplicates()
                        st.session_state.manual_unavailability = standardize_unavailability_df(combined)
                    st.session_state.last_unavail_file = file_key
                    save_persisted_df(st.session_state.manual_unavailability, "unavailability.csv")
                    st.rerun()
                    
        # Bakery Team Monthly Calendar & Unavailability Grid
        render_team_monthly_calendar_grid()
        
        st.markdown("""
        <div style="background: linear-gradient(135deg, #081d19 0%, #16443c 100%); padding: 10px 18px; border-radius: 12px 12px 0 0; color: #ffffff !important; font-weight: 800; font-size: 1.1rem; letter-spacing: 0.3px; border: 2px solid #e5a93c; border-bottom: none; margin-top: 15px;">
            🚫 Staff Weekly Unavailability Constraints
        </div>
        """, unsafe_allow_html=True)
        if st.session_state.manual_unavailability is not None and not st.session_state.manual_unavailability.empty:
            st.session_state.manual_unavailability = standardize_unavailability_df(st.session_state.manual_unavailability)
            unavailability_df = st.data_editor(st.session_state.manual_unavailability, num_rows="dynamic", key="edit_unavailability_v4")
            if unavailability_df is not None and not unavailability_df.empty:
                st.session_state.manual_unavailability = standardize_unavailability_df(unavailability_df)
                save_persisted_df(st.session_state.manual_unavailability, "unavailability.csv")

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
                    st.rerun()
                    
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
                    st.rerun()
                    
        st.markdown("""
        <div style="background: linear-gradient(135deg, #081d19 0%, #16443c 100%); padding: 10px 18px; border-radius: 12px 12px 0 0; color: #ffffff !important; font-weight: 800; font-size: 1.1rem; letter-spacing: 0.3px; border: 2px solid #e5a93c; border-bottom: none; margin-top: 15px;">
            📌 Fixed Baseline Staff Shifts
        </div>
        """, unsafe_allow_html=True)
        fixed_df = st.data_editor(st.session_state.manual_fixed, num_rows="dynamic", key="edit_fixed")
        st.session_state.manual_fixed = fixed_df
        save_persisted_df(fixed_df, "fixed.csv")
