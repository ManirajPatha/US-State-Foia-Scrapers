import os
import time
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

download_dir = os.path.join(os.path.expanduser("~"), "Downloads")

options = webdriver.ChromeOptions()
prefs = {"download.default_directory": download_dir}
options.add_experimental_option("prefs", prefs)

driver = webdriver.Chrome(options=options)
driver.maximize_window()

url = "https://www.emarketplace.state.pa.us/Search.aspx"
driver.get(url)
wait = WebDriverWait(driver, 20)

archived_radio = wait.until(
    EC.element_to_be_clickable((By.ID, "ctl00_MainBody_rdoArch_1"))
)
driver.execute_script("arguments[0].click();", archived_radio)
time.sleep(1)

search_btn = wait.until(
    EC.element_to_be_clickable((By.ID, "ctl00_MainBody_btnSearch"))
)
driver.execute_script("arguments[0].click();", search_btn)
time.sleep(5)

status_header = wait.until(
    EC.element_to_be_clickable((By.ID, "ColumnHeader_Status"))
)
driver.execute_script("arguments[0].click();", status_header)
time.sleep(3)

export_btn = wait.until(
    EC.element_to_be_clickable((By.ID, "ctl00_MainBody_btnExport"))
)
driver.execute_script("arguments[0].click();", export_btn)
time.sleep(10)

csv_files = [f for f in os.listdir(download_dir) if f.endswith(".csv")]
if not csv_files:
    print("No CSV found in Downloads folder.")
    driver.quit()
    raise SystemExit()

csv_files.sort(key=lambda x: os.path.getmtime(os.path.join(download_dir, x)), reverse=True)
latest_csv = os.path.join(download_dir, csv_files[0])

df = pd.read_csv(latest_csv, engine='python', on_bad_lines='skip')

if 'Status' not in df.columns:
    print("CSV does not contain 'Status' column.")
    driver.quit()
    raise SystemExit()

df_closed = df[df['Status'].str.strip().str.lower() == 'closed']

def clean_text(x):
    if isinstance(x, str):
        return ''.join(
            c for c in x if c == '\t' or c == '\n' or c == '\r'
            or ('\u0020' <= c <= '\uD7FF')
            or ('\uE000' <= c <= '\uFFFD')
            or ('\U00010000' <= c <= '\U0010FFFF')
        )
    return x

df_cleaned = df_closed.applymap(clean_text)


df_first100 = df_cleaned.head(100)

excel_path = os.path.join(download_dir, "pa_archived_closed_bids_first100.xlsx")
df_first100.to_excel(excel_path, index=False)
print(f"First 100 archived closed bids saved to {excel_path}")

driver.quit()
