from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
import pandas as pd
import time
import re

URL = "https://www.njstart.gov/bso/view/search/external/advancedSearchBid.xhtml"

# Current year and previous year
TARGET_YEAR = pd.Timestamp.now().year
YEARS = {TARGET_YEAR, TARGET_YEAR - 1}

# Regex to detect date-like strings such as 07/29/2025 14:00:59
DATE_RE = re.compile(r"\b\d{2}/\d{2}/\d{4}\s+\d{2}:\d{2}:\d{2}\b")

def find_date_col_idx(columns, sample_rows):
    """
    Determine which column contains the date by:
    1) Preferring header names containing 'date', 'open', or 'due'.
    2) Falling back to first column whose sample cell matches DATE_RE.
    """
    # Prefer header name heuristics
    for i, name in enumerate(columns):
        n = name.strip().lower()
        if "date" in n or "open" in n or "due" in n or "posted" in n or "close" in n:
            return i

    # Fallback: inspect sample row cells for date pattern
    for row in sample_rows:
        cells = row.find_elements(By.XPATH, "./td[not(@style='display:none')]")
        for i, cell in enumerate(cells):
            if DATE_RE.search(cell.text.strip()):
                return i

    # If not found, return None
    return None

def parse_date(s):
    # Example format: 07/29/2025 14:00:59 (mm/dd/YYYY HH:MM:SS)
    # Use exact format for speed and reliability; returns NaT if not parseable
    return pd.to_datetime(s, format="%m/%d/%Y %H:%M:%S", errors="coerce")

# ----------------------------
# Selenium setup and search
# ----------------------------
options = webdriver.ChromeOptions()
options.add_argument("--start-maximized")
driver = webdriver.Chrome(options=options)
wait = WebDriverWait(driver, 30)

try:
    driver.get(URL)

    # Select status in dropdown (example keeps the provided value "2BPO")
    status_dropdown = wait.until(EC.presence_of_element_located((By.ID, "bidSearchForm:status")))
    select = Select(status_dropdown)
    select.select_by_value("2BPO")

    # Click Search
    search_btn = wait.until(EC.element_to_be_clickable((By.XPATH, "//span[text()='Search']")))
    search_btn.click()

    # Wait for results header
    wait.until(EC.presence_of_element_located((By.ID, "bidSearchResultsForm:bidResultId_head")))
    time.sleep(1.5)

    # Capture visible headers (used only to help detect date column)
    headers = driver.find_elements(
        By.XPATH,
        "//thead[@id='bidSearchResultsForm:bidResultId_head']//th[not(@style='display:none')]/span[@class='ui-column-title']"
    )
    columns = [h.text.strip() for h in headers if h.text.strip()]

    # Gather first-page rows for date column detection
    sample_rows = driver.find_elements(By.XPATH, "//tbody[@id='bidSearchResultsForm:bidResultId_data']/tr")
    date_col_idx = find_date_col_idx(columns, sample_rows)

    if date_col_idx is None:
        print("Could not determine date column. Exporting all rows without year filtering.")
        date_filter_enabled = False
    else:
        date_filter_enabled = True

    all_data = []

    while True:
        # Wait for current page rows
        rows = driver.find_elements(By.XPATH, "//tbody[@id='bidSearchResultsForm:bidResultId_data']/tr")
        if not rows:
            break

        # Keep reference to the first row to detect DOM redraw when paginating
        first_row_ref = rows[0]

        # Extract rows, filter for current and previous year if date column known
        for row in rows:
            cells = row.find_elements(By.XPATH, "./td[not(@style='display:none')]")
            row_texts = [cell.text.strip() for cell in cells]

            if not row_texts:
                continue

            if date_filter_enabled and date_col_idx is not None and date_col_idx < len(row_texts):
                dt = parse_date(row_texts[date_col_idx])
                if pd.isna(dt) or dt.year not in YEARS:
                    continue

            all_data.append(row_texts)

        # Attempt to click Next; stop if disabled or no further pages
        try:
            next_btn = driver.find_element(By.XPATH, "//a[contains(@class,'ui-paginator-next')]")
            if "ui-state-disabled" in (next_btn.get_attribute("class") or ""):
                break
            next_btn.click()
            # Wait for table to redraw by staleness of first row
            try:
                WebDriverWait(driver, 20).until(EC.staleness_of(first_row_ref))
            except Exception:
                time.sleep(1.5)
        except Exception:
            break

    # Build DataFrame and rename numeric columns 0..9 to desired headers
    df = pd.DataFrame(all_data)

    target_headers = [
        "Bid Solicitation #",
        "Organization Name",
        "Contract #",
        "Buyer",
        "Description",
        "Bid Opening Date",
        "Bid Holder List",
        "Awarded Vendor(s)",
        "Status",
        "Alternate Id",
    ]
    # Map integer column labels 0..9 -> requested names; others remain unchanged
    df.rename(columns=dict(enumerate(target_headers)), inplace=True)

    # Compose output filename with both years (e.g., 2025_2024)
    y_sorted = sorted(list(YEARS), reverse=True)
    out_name = f"njstart_bid_to_po_{y_sorted[0]}_{y_sorted[1]}.xlsx"
    df.to_excel(out_name, index=False)
    print(f"Scraped {len(df)} rows for years {y_sorted[0]} and {y_sorted[1]} and saved to {out_name}")

finally:
    driver.quit()
