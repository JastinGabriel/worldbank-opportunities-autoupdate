import os
import requests
import pandas as pd
from datetime import datetime
import time
import gspread
from gspread_dataframe import set_with_dataframe
from oauth2client.service_account import ServiceAccountCredentials

EXCEL_FILE = "worldbank_current_opportunities.xlsx"

# --- Step 1: Load previous data ---
if os.path.exists(EXCEL_FILE):
    df_existing = pd.read_excel(EXCEL_FILE)
    if 'id' in df_existing.columns:
        df_existing = df_existing[~df_existing['id'].astype(str).str.lower().isin(['new news', 'existing news'])]
else:
    df_existing = pd.DataFrame()

# --- Step 2: Pull fresh data ---
BASE_URL = "https://search.worldbank.org/api/v2/procnotices"
PARAMS = {
    "format": "json",
    "apilang": "en",
    "fl": "id,submission_deadline_date,bid_description,project_ctry_name,project_name,notice_type,notice_status,notice_lang_name,submission_date,noticedate,procurement_group_desc",
    "rows": 100,
    "os": 0
}

MAX_RECORDS = 1000
MAX_PAGES = 50
notices_data = []
page_count = 0

while page_count < MAX_PAGES and len(notices_data) < MAX_RECORDS:
    response = requests.get(BASE_URL, params=PARAMS)
    if response.status_code != 200:
        print(f"⚠️ Request failed with status {response.status_code}.")
        break

    data = response.json()
    procnotices = data.get("procnotices", [])
    if not procnotices:
        break

    for notice in procnotices:
        deadline_str = notice.get("submission_deadline_date")
        if not deadline_str:
            continue
        if deadline_str.endswith("Z"):
            deadline_str = deadline_str[:-1]
        try:
            deadline = datetime.strptime(deadline_str, "%Y-%m-%dT%H:%M:%S")
        except ValueError:
            continue
        if deadline < datetime.now():
            continue  # skip expired

        new_record = {
            "id": notice.get("id", ""),
            "Notice": notice.get("bid_description", ""),
            "Country": notice.get("project_ctry_name", ""),
            "Project Title": notice.get("project_name", ""),
            "Notice Type": notice.get("notice_type", ""),
            "Procurement Type": notice.get("procurement_group_desc", ""),
            "Language": notice.get("notice_lang_name", ""),
            "Published Date": notice.get("noticedate", ""),
            "Submission Deadline": deadline_str
        }
        notices_data.append(new_record)

        if len(notices_data) >= MAX_RECORDS:
            break

    PARAMS["os"] += PARAMS["rows"]
    page_count += 1
    time.sleep(0.3)

df_new = pd.DataFrame(notices_data)

# --- Step 3: Remove expired deadlines from existing ---
if not df_existing.empty:
    if 'Submission Deadline' in df_existing.columns:
        try:
            df_existing['Submission Deadline'] = pd.to_datetime(df_existing['Submission Deadline'], errors='coerce')
            df_existing = df_existing[df_existing['Submission Deadline'] >= datetime.now()]
        except Exception as e:
            print(f"Failed to filter existing data: {e}")

# --- Step 4: Identify new records ---
if not df_existing.empty and 'id' in df_existing.columns:
    df_filtered_new = df_new[~df_new['id'].isin(df_existing['id'])]
else:
    df_filtered_new = df_new.copy()

# --- Step 5: Combine and export ---
df_combined = pd.concat([df_filtered_new, df_existing], ignore_index=True)
df_combined.to_excel(EXCEL_FILE, index=False)

print(f"✅ Added {len(df_filtered_new)} new records. Total records: {len(df_combined)}")

# --- Step 6: Upload to Google Sheets ---
scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
creds = ServiceAccountCredentials.from_json_keyfile_name("google_service_account.json", scope)
client = gspread.authorize(creds)

spreadsheet = client.open_by_url("https://docs.google.com/spreadsheets/d/1UQ0AXeDLFEAbGohuy9GTaX67LQnQ9ZkFi6mLgKOy7Ec/edit")
worksheet = spreadsheet.worksheet("Sheet1")

worksheet.clear()
set_with_dataframe(worksheet, df_combined)

print("✅ Google Sheet updated successfully!")
