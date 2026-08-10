# 📅 Weekly Roster Creator App

A shareable, web-based tool built in Python and Streamlit that uses Google AI Studio's Gemini models to automatically generate optimized, award-compliant weekly rosters.

## 🚀 How to Host & Share this App (for FREE)

You can deploy and share this app online without installing Python on your computer by using **GitHub** and **Streamlit Community Cloud**.

### Step 1: Upload the Code to GitHub
1. Sign in to your [GitHub account](https://github.com/) (create one for free if you don't have it).
2. Create a new repository (e.g., `weekly-roster-generator`).
3. Upload the following files from your local folder:
   - `app.py`
   - `requirements.txt`
   - `README.md`

### Step 2: Deploy to Streamlit Community Cloud
1. Go to [Streamlit Community Cloud](https://share.streamlit.io/) and log in with your GitHub account.
2. Click **New app**.
3. Select your repository (`weekly-roster-generator`), branch (`main` or `master`), and main file path (`app.py`).
4. Click **Deploy!**
5. Within 1-2 minutes, your app will be online. You can copy the browser URL and share it with anyone!

---

## 🔑 How to Get a Google AI Studio API Key

To run the generator inside the app, users will need an API key:
1. Go to [Google AI Studio](https://aistudio.google.com/).
2. Click the **Get API Key** button in the top left.
3. Click **Create API Key** (you can generate a free tier key).
4. Copy the API key and paste it directly into the app sidebar!

---

## 📂 Input File Formats

Within the app tabs, you can either enter data manually or upload Excel files matching these layouts:
- **Employees**: Column headers: `Name`, `Role`, `Age`, `Employment Type`, `Start Date`.
- **Unavailability**: Column headers: `Employee`, `Day`, `Time Window`.
- **Daily Requirements**: Column headers: `Day`, `Shift`, `Count Required`.
- **Fixed Shifts**: Column headers: `Employee`, followed by `Monday`, `Tuesday`, `Wednesday`, `Thursday`, `Friday`, `Saturday`, `Sunday`.
