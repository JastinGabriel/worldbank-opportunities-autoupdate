import os
import requests
import pandas as pd
from datetime import datetime
import time

import gspread
from gspread_dataframe import set_with_dataframe, get_as_dataframe
from oauth2client.service_account import ServiceAccountCredentials

# --- Step 1: Load existing data from Google Sheet ---
GOOGLE_SHEET_URL = "https://docs.google.com/spreadsheets/d/1UQ0AXeDLFEAbGohuy9GTaX67LQnQ9ZkFi6mLgKOy7Ec/edit"
SHEET_NAME = "Sheet1"

scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
creds = ServiceAccountCredentials.from_json_keyfile_name("google_service_account.json", scope)
client = gspread.authorize(creds)
spreadsheet = client.open_by_url(GOOGLE_SHEET_URL)
worksheet = spreadsheet.worksheet(SHEET_NAME)

# Load existing Google Sheet data into DataFrame
df_existing = get_as_dataframe(worksheet, evaluate_formulas=True)
df_existing = df_existing.dropna(how="all").dropna(axis=1, how="all")

# Remove section labels if present
if 'id' in df_existing.columns:
    df_existing = df_existing[~df_existing['id'].astype(str).str.lower().isin(['new news', 'existing news'])]

# --- Step 2: Pull fresh data from World Bank API ---
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

# --- Step 3: Identify new records ---
if not df_existing.empty and 'id' in df_existing.columns:
    df_filtered_new = df_new[~df_new['id'].isin(df_existing['id'])]
else:
    df_filtered_new = df_new.copy()

# --- Step 4: Build final DataFrame with section headers ---
rows = []

if not df_filtered_new.empty:
    rows.append(["New News"] + [""] * (df_filtered_new.shape[1] - 1))
    rows.extend(df_filtered_new.values.tolist())

if not df_existing.empty:
    rows.append(["Existing News"] + [""] * (df_existing.shape[1] - 1))
    rows.extend(df_existing.values.tolist())

columns = df_new.columns if not df_new.empty else df_existing.columns
df_final = pd.DataFrame(rows, columns=columns)

# --- Step 5: Export to Google Sheet ---
worksheet.clear()
set_with_dataframe(worksheet, df_final)

# --- Done ---
print(f"✅ Added {len(df_filtered_new)} new records on top of {len(df_existing)} existing records.")
print("✅ Google Sheet successfully updated!")
