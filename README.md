# 📅 Weekly Roster Creator App

A shareable, web-based tool built in Python and Streamlit that automatically generates optimized, award-compliant weekly rosters completely offline (no API keys required).

## 🚀 How to Host & Share this App (for FREE)

You can deploy and share this app online without installing Python on your computer by using **GitHub** and **Streamlit Community Cloud**.

### Step 1: Upload the Code to GitHub
1. Sign in to your [GitHub account](https://github.com/) (create one for free if you don't have it).
2. Create a new repository (e.g., `weekly-roster-generator`).
3. Upload the following files from your local folder:
   - `app.py`
   - `vic_holidays.py`
   - `holiday_sources.py`
   - `holiday_calendar_ui.py`
   - `requirements.txt`
   - `README.md`

   Keep all six files together in the repository root. Uploading only `app.py`
   will cause a `ModuleNotFoundError` because the holiday feature uses the three
   helper files above. For an existing deployment, add the missing helpers beside
   `app.py` in the branch used by Streamlit. Do not upload local credentials or
   replace existing staff data or Streamlit secrets.

### Step 2: Deploy to Streamlit Community Cloud
1. Go to [Streamlit Community Cloud](https://share.streamlit.io/) and log in with your GitHub account.
2. Click **New app**.
3. Select your repository (`weekly-roster-generator`), branch (`main` or `master`), and main file path (`app.py`).
4. Click **Deploy!**
5. Within 1-2 minutes, your app will be online. You can copy the browser URL and share it with anyone!

---

## 📂 Input File Formats

Within the app tabs, you can either enter data manually or upload Excel files matching these layouts:
- **Employees**: Column headers: `Name`, `Role`, `Age`, `Employment Type`, `Start Date`.
- **Unavailability**: Column headers: `Employee`, `Day`, `Time Window`.
- **Daily Requirements**: Column headers: `Day`, `Shift`, `Count Required`.
- **Fixed Shifts**: Column headers: `Employee`, followed by `Monday`, `Tuesday`, `Wednesday`, `Thursday`, `Friday`, `Saturday`, `Sunday`.

## Official Victorian holiday updates

Managers can open **Home / Dashboard → Holiday calendar · official updates**.
The app discovers the public-holiday and school-term ICS downloads at
https://www.vic.gov.au/ical and checks the current and next calendar year.
It derives school breaks using student start dates, including the following
January for summer. These dates are for Victorian government schools; other
schools and pupil-free days need a separate check.

Choose **Check official dates now**, review the proposed differences and source
warnings, then acknowledge and select **Apply reviewed calendar**. Applying a
calendar updates holiday detection and live roster wage calculations. Existing
shifts and exported files are not rewritten. Affected saved roster weeks appear
in the review panel. Local substitute public holidays must be checked separately.

Automatic checking runs when a manager opens the dashboard: every 30 days,
or every 7 days in October–January. A failed check is retried on a later day.
It never applies changes automatically. No independent background scheduler,
email reminder delivery or birthday greetings are activated by this feature.
Internet access is needed to check sources; saved approved dates work offline
when local storage is used.

Downloads are restricted to the official HTTPS host and validated before review.
Missing coverage, changed formats and failed downloads preserve the active
calendar and clear any stale pending proposal. Provisional source dates are not
imported; an existing matching entry is retained with a visible warning.

Calendar state is saved to Firebase `system/holiday_calendar` when configured,
otherwise to `data/holiday_calendar.json`. It includes the last attempt, last
successful check, pending proposal, approval time and operator, source URLs and
the previous calendar. Concurrent approvals are rejected rather than overwriting
one another. Back up this record with the rest of the app data.

Deploy `holiday_sources.py`, `holiday_calendar_ui.py` and `vic_holidays.py`
alongside `app.py`. No extra packages are needed for the source checker.
Focused tests: `python -m unittest discover -s tests -p test_holiday_sources.py -v`.
The test fixtures are official calendar downloads and the download-page snapshot,
retrieved on 29 September 2026; they are test inputs, not the live app calendar.
