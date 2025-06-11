#!/usr/bin/env python
# coding: utf-8

# In[2]:


import os
import requests
import pandas as pd

EXCEL_FILE = "world_bank_procurement.xlsx"

# === Load existing Excel file ===
if os.path.exists(EXCEL_FILE):
    existing_df = pd.read_excel(EXCEL_FILE)
    existing_ids = set(existing_df["Notice"])
else:
    existing_df = pd.DataFrame()
    existing_ids = set()

# === Fetch data from API ===
BASE_URL = "https://search.worldbank.org/api/v2/procnotices?format=json&fct=procurement_group_desc_exact&fq=procurement_group_desc_exact:Opportunities"
PARAMS = {
    "format": "json",
    "rows": 100,
    "os": 0
}

new_records = []

while True:
    response = requests.get(BASE_URL, params=PARAMS)
    data = response.json()
    notices = data.get("procnotices", [])

    if not notices:
        break

    for notice in notices:
        notice_id = notice.get("bid_description", "")
        if notice_id not in existing_ids:
            new_records.append({
                "Notice": notice_id,
                "Country": notice.get("project_ctry_name", ""),
                "Project Title": notice.get("project_name", ""),
                "Notice Type": notice.get("notice_type", ""),
                "Procurement Type": notice.get("procurement_group_desc", ""),
                "Language": notice.get("notice_lang_name", ""),
                "Published Date": notice.get("noticedate", ""),
                "Submission Deadline": notice.get("submission_deadline_date", "")
            })

    PARAMS["os"] += PARAMS["rows"]
    if PARAMS["os"] >= int(data.get("total", "0")):
        break

# === Merge data: New on top ===
if new_records:
    new_df = pd.DataFrame(new_records)

    # Optional: Add separator row
    separator = pd.DataFrame([["--- Existing Data Below ---"] + [""] * (new_df.shape[1]-1)], columns=new_df.columns)

    combined_df = pd.concat([new_df, separator, existing_df], ignore_index=True)
    combined_df.to_excel(EXCEL_FILE, index=False)
    print(f"✅ {len(new_df)} new rows added at the top.")
else:
    print("✅ No new data found. Excel not updated.")


# In[4]:


import requests
import pandas as pd
from datetime import datetime
import time

# Base API setup
BASE_URL = "https://search.worldbank.org/api/v2/procnotices"
PARAMS = {
    "format": "json",
    "apilang": "en",
    "fl": "id,submission_deadline_date,bid_description,project_ctry_name,project_name,notice_type,notice_status,notice_lang_name,submission_date,noticedate,procurement_group_desc",
    "rows": 100,  # Number of records per page
    "os": 0       # Offset for pagination
}

# Limits
MAX_RECORDS = 2000   # Total records to collect
MAX_PAGES = 50       # Maximum pages to fetch
notices_data = []
page_count = 0

# Loop through pages
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
        # Skip if no deadline
        deadline_str = notice.get("submission_deadline_date")
        if not deadline_str:
            continue

        # Parse and clean deadline
        if deadline_str.endswith("Z"):
            deadline_str = deadline_str[:-1]
        try:
            deadline = datetime.strptime(deadline_str, "%Y-%m-%dT%H:%M:%S")
        except ValueError:
            continue

        # ✅ Updated filter: Include only current opportunities (deadline in the future)
        if deadline >= datetime.now():
            notices_data.append({
                "Notice": notice.get("bid_description", ""),
                "Country": notice.get("project_ctry_name", ""),
                "Project Title": notice.get("project_name", ""),
                "Notice Type": notice.get("notice_type", ""),
                "Procurement Type": notice.get("procurement_group_desc", ""),
                "Language": notice.get("notice_lang_name", ""),
                "Published Date": notice.get("noticedate", ""),
                "Submission Deadline": deadline_str
            })

        if len(notices_data) >= MAX_RECORDS:
            break

    # Move to next page
    PARAMS["os"] += PARAMS["rows"]
    page_count += 1
    time.sleep(0.3)  # Slight delay to be kind to the API

# Export to Excel
df = pd.DataFrame(notices_data)
df.to_excel("worldbank_current_opportunities.xlsx", index=False)

print(f"✅ Collected {len(notices_data)} records.")
print("✅ Exported to 'worldbank_current_opportunities.xlsx'")


# In[17]:


pip install gspread gspread_dataframe oauth2client


# In[18]:


pip install gspread oauth2client


# In[23]:





# In[9]:


import pandas as pd
import gspread
from gspread_dataframe import set_with_dataframe
from oauth2client.service_account import ServiceAccountCredentials

# Load Excel file into DataFrame
df = pd.read_excel("worldbank_current_opportunities.xlsx")

# Set up credentials
scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
creds = ServiceAccountCredentials.from_json_keyfile_name("google_service_account.json", scope)
client = gspread.authorize(creds)

# Open an existing Google Sheet by **URL**
spreadsheet = client.open_by_url("https://docs.google.com/spreadsheets/d/1UQ0AXeDLFEAbGohuy9GTaX67LQnQ9ZkFi6mLgKOy7Ec/edit")  # Replace with your actual URL
worksheet = spreadsheet.worksheet("Sheet1")  # Adjust if your sheet name is different

# Clear and update
worksheet.clear()
set_with_dataframe(worksheet, df)

print("✅ Updated existing sheet!")


# In[2]:


import os
import requests
import pandas as pd
from datetime import datetime
import time

EXCEL_FILE = "worldbank_current_opportunities.xlsx"

# --- Step 1: Load previous data ---
if os.path.exists(EXCEL_FILE):
    df_existing = pd.read_excel(EXCEL_FILE)

    # Remove any section labels accidentally saved as rows
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

        # Compose new record
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

# --- Step 4: Create final Excel output with section labels ---
rows = []

# Add "New News" section
if not df_filtered_new.empty:
    rows.append(["New News"] + [""] * (df_filtered_new.shape[1] - 1))
    rows.extend(df_filtered_new.values.tolist())

# Add "Existing News" section
if not df_existing.empty:
    rows.append(["Existing News"] + [""] * (df_existing.shape[1] - 1))
    rows.extend(df_existing.values.tolist())

# Define full column headers (same as df_new or df_existing)
columns = df_new.columns if not df_new.empty else df_existing.columns

# Final DataFrame
df_final = pd.DataFrame(rows, columns=columns)

# --- Step 5: Export ---
df_final.to_excel(EXCEL_FILE, index=False)

print(f"✅ Added {len(df_filtered_new)} new records on top of {len(df_existing)} existing records.")
print(f"✅ Exported updated data to '{EXCEL_FILE}'")


# In[8]:


import os
import requests
import pandas as pd
from datetime import datetime, date
import time

EXCEL_FILE = "worldbank_current_opportunities.xlsx"

# --- Step 1: Load previous data and clean expired ---
if os.path.exists(EXCEL_FILE):
    df_existing = pd.read_excel(EXCEL_FILE)

    # Remove any section labels accidentally saved as rows
    if 'id' in df_existing.columns:
        df_existing = df_existing[~df_existing['id'].astype(str).str.lower().isin(['new news', 'existing news'])]

    # 🧹 Remove expired opportunities regardless of section
    if 'Submission Deadline' in df_existing.columns:
        df_existing['Submission Deadline'] = pd.to_datetime(df_existing['Submission Deadline'], errors='coerce')
        before_cleanup = len(df_existing)
        df_existing = df_existing[df_existing['Submission Deadline'] >= pd.to_datetime(date.today())]
        after_cleanup = len(df_existing)
        print(f"🧹 Removed {before_cleanup - after_cleanup} expired records from previous data.")
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

        # Compose new record
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

# --- Step 4: Create final Excel output with section labels ---
rows = []

# Add "New News" section
if not df_filtered_new.empty:
    rows.append(["New News"] + [""] * (df_filtered_new.shape[1] - 1))
    rows.extend(df_filtered_new.values.tolist())

# Add "Existing News" section
if not df_existing.empty:
    rows.append(["Existing News"] + [""] * (df_existing.shape[1] - 1))
    rows.extend(df_existing.values.tolist())

# Define full column headers
columns = df_new.columns if not df_new.empty else df_existing.columns

# Final DataFrame
df_final = pd.DataFrame(rows, columns=columns)

# --- Step 5: Export ---
df_final.to_excel(EXCEL_FILE, index=False)

print(f"✅ Added {len(df_filtered_new)} new records on top of {len(df_existing)} existing records.")
print(f"✅ Exported updated data to '{EXCEL_FILE}'")


# In[ ]:


import os
import requests
import pandas as pd
from datetime import datetime
import time

EXCEL_FILE = "worldbank_current_opportunities.xlsx"

# --- Step 1: Load previous data ---
if os.path.exists(EXCEL_FILE):
    df_existing = pd.read_excel(EXCEL_FILE)
else:
    df_existing = pd.DataFrame()

# --- Step 2: Pull fresh data from API (your existing code with slight adjustment) ---

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

        # Compose new record
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

# --- Step 3: Find new records that don't exist in previous data ---
if not df_existing.empty:
    # Using 'id' as unique key, or fallback to Published Date if id missing
    if 'id' in df_existing.columns and 'id' in df_new.columns:
        # Keep only new records whose id is not in existing data
        df_filtered_new = df_new[~df_new['id'].isin(df_existing['id'])]
    else:
        # Fallback: use Published Date to filter new records
        df_existing['Published Date Parsed'] = pd.to_datetime(df_existing['Published Date'], errors='coerce')
        max_existing_date = df_existing['Published Date Parsed'].max()
        df_new['Published Date Parsed'] = pd.to_datetime(df_new['Published Date'], errors='coerce')
        df_filtered_new = df_new[df_new['Published Date Parsed'] > max_existing_date]
else:
    df_filtered_new = df_new

# --- Step 4: Combine data with section rows ---
rows = []

# Add new news section
rows.append(["New News"] + [""] * (df_new.shape[1] - 1))
if not df_filtered_new.empty:
    rows.extend(df_filtered_new.drop(columns=['Published Date Parsed'], errors='ignore').values.tolist())

# Add previous news section
rows.append(["Previous News"] + [""] * (df_existing.shape[1] - 1))
if not df_existing.empty:
    rows.extend(df_existing.drop(columns=['Published Date Parsed'], errors='ignore').values.tolist())

# Create combined DataFrame
combined_df = pd.DataFrame(rows, columns=df_new.columns if not df_new.empty else df_existing.columns)

# --- Step 5: Export to Excel ---
combined_df.to_excel(EXCEL_FILE, index=False)

print(f"✅ Added {len(df_filtered_new)} new records on top of {len(df_existing)} existing records.")
print(f"✅ Exported updated data to '{EXCEL_FILE}'")

