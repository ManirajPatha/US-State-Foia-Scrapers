from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import pandas as pd
import time
import re
from zoneinfo import ZoneInfo  

PORTAL_URL = "https://utah.bonfirehub.com/portal/?tab=openOpportunities"
LOCAL_TZ = ZoneInfo("America/Denver")  
TARGET_YEAR = pd.Timestamp.now(tz=LOCAL_TZ).year  
AWARDED_ONLY = True  

OUTPUT_FILE = f"utah_bonfire_{TARGET_YEAR}_{'awarded' if AWARDED_ONLY else 'all_statuses'}.xlsx"


ORDINAL_RE = re.compile(r'(\d{1,2})(st|nd|rd|th)', re.I)
TIME_COMPACT_RE = re.compile(r'\b(\d{1,2})(\d{2})\s*(AM|PM)\b', re.I)

def _normalize_time(s: str) -> str:
    
    return TIME_COMPACT_RE.sub(lambda m: f"{int(m.group(1))}:{m.group(2)} {m.group(3).upper()}", s)

def _clean_dt_string(s: str) -> str:
    
    s = ORDINAL_RE.sub(r"\1", s)           
    s = s.replace("MDT", "").replace("MST", "")  
    s = _normalize_time(s)
    s = re.sub(r'\s{2,}', ' ', s).strip().rstrip(',')
    return s

def parse_close_date_local(text: str) -> pd.Timestamp:
    """
    Parse the Close Date string into a timezone-aware Timestamp in America/Denver.
    Returns pandas.NaT on failure.
    """
    if not text:
        return pd.NaT
    s = _clean_dt_string(text)
    
    dt = pd.to_datetime(s, format="%b %d %Y, %I:%M %p", errors="coerce")
    if pd.isna(dt):
        dt = pd.to_datetime(s, errors="coerce")
    if pd.isna(dt):
        return pd.NaT
    return dt.tz_localize(LOCAL_TZ)


options = webdriver.ChromeOptions()
options.add_argument("--start-maximized")
driver = webdriver.Chrome(options=options)
wait = WebDriverWait(driver, 20)

try:
    
    driver.get(PORTAL_URL)
    past_tab = wait.until(EC.element_to_be_clickable((By.ID, "pastOpportunitiesTab")))
    past_tab.click()

    
    wait.until(EC.presence_of_all_elements_located((By.CSS_SELECTOR, "#DataTables_Table_1 tbody tr")))
    time.sleep(0.8)

    
    headers = wait.until(EC.presence_of_all_elements_located((By.CSS_SELECTOR, "#DataTables_Table_1 thead th")))
    close_idx = None
    for i, th in enumerate(headers):
        if "close" in th.text.strip().lower():
            close_idx = i
            break

    if close_idx is not None:
        
        headers[close_idx].click()
        time.sleep(0.6)
        headers = driver.find_elements(By.CSS_SELECTOR, "#DataTables_Table_1 thead th")
        aria_sort = (headers[close_idx].get_attribute("aria-sort") or "").lower()
        if aria_sort != "descending":
            headers[close_idx].click()
            time.sleep(0.6)

    data = []
    stop_due_to_year = False

    while True:
        
        rows = wait.until(EC.presence_of_all_elements_located((By.CSS_SELECTOR, "#DataTables_Table_1 tbody tr")))
        if not rows:
            break

        
        first_row_ref = rows[0]

        for row in rows:
            cols = row.find_elements(By.TAG_NAME, "td")
            if not cols:
                continue

            
            status = cols[0].text.strip()
            ref_no = cols[1].text.strip() if len(cols) > 1 else ""
            project = cols[2].text.strip() if len(cols) > 2 else ""
            department = cols[3].text.strip() if len(cols) > 3 else ""
            close_date_str = cols[4].text.strip() if len(cols) > 4 else ""

            
            dt_local = parse_close_date_local(close_date_str)
            if pd.isna(dt_local):
                continue

            
            if dt_local.year < TARGET_YEAR:
                stop_due_to_year = True
                break

            
            if dt_local.year == TARGET_YEAR:
                if AWARDED_ONLY and status.lower() != "awarded":
                    continue
                data.append({
                    "Ref #": ref_no,
                    "Project": project,
                    "Department": department,
                    "Close Date": close_date_str,  
                    "Status": status
                })

            

        if stop_due_to_year:
            break

        
        wrapper = driver.find_element(By.CSS_SELECTOR, "#DataTables_Table_1_wrapper")
        try:
            next_li = wrapper.find_element(By.CSS_SELECTOR, "li.paginate_button.next")
            classes = next_li.get_attribute("class") or ""
            if "disabled" in classes:
                break
            next_a = next_li.find_element(By.TAG_NAME, "a")
            next_a.click()
            wait.until(EC.staleness_of(first_row_ref))  
        except Exception:
            break

    
    df = pd.DataFrame(data)
    df.to_excel(OUTPUT_FILE, index=False)
    print(f"Saved {len(df)} rows to {OUTPUT_FILE}")

finally:
    driver.quit()
