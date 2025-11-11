from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import NoSuchElementException, TimeoutException
import pandas as pd
import time
import re
import os
import zipfile
from pathlib import Path
from datetime import datetime

URL = "https://www.njstart.gov/bso/view/search/external/advancedSearchBid.xhtml"
BASE_DOWNLOAD_DIR = Path("downloads")
BASE_DOWNLOAD_DIR.mkdir(exist_ok=True)

# Current year and previous year
TARGET_YEAR = pd.Timestamp.now().year
YEARS = {TARGET_YEAR, TARGET_YEAR - 1}

# Regex to detect date-like strings
DATE_RE = re.compile(r"\b\d{2}/\d{2}/\d{4}\s+\d{2}:\d{2}:\d{2}\b")

def find_date_col_idx(columns, sample_rows):
    """Determine which column contains the date."""
    for i, name in enumerate(columns):
        n = name.strip().lower()
        if "date" in n or "open" in n or "due" in n or "posted" in n or "close" in n:
            return i
    for row in sample_rows:
        cells = row.find_elements(By.XPATH, "./td[not(@style='display:none')]")
        for i, cell in enumerate(cells):
            if DATE_RE.search(cell.text.strip()):
                return i
    return None
def parse_date(s):
    return pd.to_datetime(s, format="%m/%d/%Y %H:%M:%S", errors="coerce")
def safe_text(element):
    """Safely extract text from an element."""
    try:
        return element.text.strip()
    except:
        return None
def wait_for_downloads(download_dir, timeout=60):
    """Wait until all Chrome downloads are complete."""
    end_time = time.time() + timeout
    while time.time() < end_time:
        if any(f.endswith(".crdownload") for f in os.listdir(download_dir)):
            time.sleep(1)
        else:
            return True
    return False
def scrape_detail_page(driver, wait, bid_number, temp_download_dir):
    """Scrape details from the bid detail page."""
    record = {
        "bid_number": bid_number,
        "description": None,
        "bid_opening_date": None,
        "department": None,
        "location": None,
        "required_date": None,
        "available_date": None,
        "begin_date": None,
        "end_date": None,
    }
    try:
        wait.until(EC.presence_of_element_located((By.CLASS_NAME, "tableStripe-01")))
        time.sleep(1)
        # Extract from first table row
        try:
            desc_cell = driver.find_element(By.XPATH, "//td[contains(text(),'Description:')]/following-sibling::td[1]")
            record["description"] = safe_text(desc_cell)
        except NoSuchElementException:
            pass
        try:
            bid_open_cell = driver.find_element(By.XPATH, "//td[contains(text(),'Bid Opening Date:')]/following-sibling::td[1]")
            record["bid_opening_date"] = safe_text(bid_open_cell)
        except NoSuchElementException:
            pass
        try:
            dept_cell = driver.find_element(By.XPATH, "//td[contains(text(),'Department:')]/following-sibling::td[1]")
            record["department"] = safe_text(dept_cell)
        except NoSuchElementException:
            pass
        try:
            loc_cell = driver.find_element(By.XPATH, "//td[contains(text(),'Location:')]/following-sibling::td[1]")
            record["location"] = safe_text(loc_cell)
        except NoSuchElementException:
            pass
        try:
            req_cell = driver.find_element(By.XPATH, "//td[contains(text(),'Required Date:')]/following-sibling::td[1]")
            record["required_date"] = safe_text(req_cell)
        except NoSuchElementException:
            pass
        try:
            avail_cell = driver.find_element(By.XPATH, "//td[contains(text(),'Available Date')]/following-sibling::td[1]")
            record["available_date"] = safe_text(avail_cell)
        except NoSuchElementException:
            pass
        try:
            begin_cell = driver.find_element(By.XPATH, "//td[contains(text(),'Begin Date:')]/following-sibling::td[1]")
            record["begin_date"] = safe_text(begin_cell)
        except NoSuchElementException:
            pass
        try:
            end_cell = driver.find_element(By.XPATH, "//td[contains(text(),'End Date:')]/following-sibling::td[1]")
            record["end_date"] = safe_text(end_cell)
        except NoSuchElementException:
            pass

        # Download File Attachments
        try:
            file_links = driver.find_elements(By.XPATH, "//td[contains(text(),'File Attachments:')]/following-sibling::td//a[contains(@href,'downloadFile')]")
           
            if file_links:
                print(f" Found {len(file_links)} attachments for {bid_number}")
                before_files = set(os.listdir(temp_download_dir))
                for link in file_links:
                    try:
                        driver.execute_script("arguments[0].click();", link)
                        time.sleep(2)
                    except Exception as e:
                        print(f"Failed to download a file: {e}")
                wait_for_downloads(temp_download_dir, timeout=90)
                after_files = set(os.listdir(temp_download_dir))
                new_files = list(after_files - before_files)
                if new_files:
                    # Create zip file
                    zip_name = f"njstart_{bid_number}.zip"
                    zip_path = BASE_DOWNLOAD_DIR / zip_name
                    with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as zf:
                        for f in new_files:
                            file_path = temp_download_dir / f
                            if file_path.exists():
                                zf.write(file_path, arcname=f)
                    print(f" Zipped {len(new_files)} files to {zip_path}")
                    # Clean up temp files
                    for f in new_files:
                        try:
                            (temp_download_dir / f).unlink()
                        except:
                            pass
        except NoSuchElementException:
            print(f" No file attachments found for {bid_number}")
        except Exception as e:
            print(f" Error downloading attachments: {e}")
    except Exception as e:
        print(f" Error scraping detail page: {e}")
    return record


# Selenium setup

temp_download_dir = Path("/tmp/njstart_downloads")
temp_download_dir.mkdir(parents=True, exist_ok=True)
options = webdriver.ChromeOptions()
options.add_argument("--start-maximized")
prefs = {
    "download.default_directory": str(temp_download_dir.resolve()),
    "download.prompt_for_download": False,
    "download.directory_upgrade": True,
    "plugins.always_open_pdf_externally": True,
    "profile.default_content_setting_values.automatic_downloads": 1,
}
options.add_experimental_option("prefs", prefs)
driver = webdriver.Chrome(options=options)
wait = WebDriverWait(driver, 30)
try:
    driver.get(URL)
    status_dropdown = wait.until(EC.presence_of_element_located((By.ID, "bidSearchForm:status")))
    select = Select(status_dropdown)
    select.select_by_value("2BPO")

    search_btn = wait.until(EC.element_to_be_clickable((By.XPATH, "//span[text()='Search']")))
    search_btn.click()

    wait.until(EC.presence_of_element_located((By.ID, "bidSearchResultsForm:bidResultId_head")))
    time.sleep(1.5)
    # Capture headers for date detection
    headers = driver.find_elements(
        By.XPATH,
        "//thead[@id='bidSearchResultsForm:bidResultId_head']//th[not(@style='display:none')]/span[@class='ui-column-title']"
    )
    columns = [h.text.strip() for h in headers if h.text.strip()]
    # Gather sample rows for date column detection
    sample_rows = driver.find_elements(By.XPATH, "//tbody[@id='bidSearchResultsForm:bidResultId_data']/tr")
    date_col_idx = find_date_col_idx(columns, sample_rows)
    if date_col_idx is None:
        print("Could not determine date column. Processing all rows.")
        date_filter_enabled = False
    else:
        date_filter_enabled = True
    all_data = []
    main_window = driver.current_window_handle
    page_number = 1
    while True:
        print(f"\nProcessing page {page_number}...")
       
        # Get current page rows
        rows = driver.find_elements(By.XPATH, "//tbody[@id='bidSearchResultsForm:bidResultId_data']/tr")
        if not rows:
            print("No rows found on this page")
            break
        print(f"Found {len(rows)} rows on page {page_number}")
       
        # Process each row
        for idx in range(len(rows)):
            try:
                # Re-fetch rows to avoid stale elements
                rows = driver.find_elements(By.XPATH, "//tbody[@id='bidSearchResultsForm:bidResultId_data']/tr")
                if idx >= len(rows):
                    break
                   
                row = rows[idx]
               
                # Get bid number link
                bid_link = row.find_element(By.XPATH, ".//td[@role='gridcell']//a")
                bid_number = bid_link.text.strip()
               
                # Check date filter if enabled
                if date_filter_enabled and date_col_idx is not None:
                    cells = row.find_elements(By.XPATH, "./td[not(@style='display:none')]")
                    if date_col_idx < len(cells):
                        date_text = cells[date_col_idx].text.strip()
                        dt = parse_date(date_text)
                        if pd.isna(dt) or dt.year not in YEARS:
                            print(f"Skipping {bid_number} - date filter")
                            continue
                print(f"\nProcessing {bid_number}...")
                # Click to open in new tab
                bid_link.click()
                time.sleep(2)

                for handle in driver.window_handles:
                    if handle != main_window:
                        driver.switch_to.window(handle)
                        break
                # Scrape detail page
                record = scrape_detail_page(driver, wait, bid_number, temp_download_dir)
                all_data.append(record)
                # Close tab and switch back to main window
                driver.close()
                driver.switch_to.window(main_window)
                time.sleep(1)
            except Exception as e:
                print(f"Error processing row {idx + 1}: {e}")
                # Make sure we're back on main window
                try:
                    driver.switch_to.window(main_window)
                except:
                    pass
                continue
        # Try to click Next button
        try:
            next_btn = driver.find_element(By.XPATH, "//a[contains(@class,'ui-paginator-next')]")
            next_btn_classes = next_btn.get_attribute("class") or ""
           
            if "ui-state-disabled" in next_btn_classes:
                print(f"\nReached last page (page {page_number})")
                break
           
            print(f"\nMoving to page {page_number + 1}...")
           
            # Store a reference element to detect page change
            first_row = driver.find_element(By.XPATH, "//tbody[@id='bidSearchResultsForm:bidResultId_data']/tr[1]")
           
            driver.execute_script("arguments[0].click();", next_btn)
           
            try:
                WebDriverWait(driver, 20).until(EC.staleness_of(first_row))
                print("Page changed successfully")
            except TimeoutException:
                print("Timeout waiting for page change, checking if content updated...")
           
            time.sleep(2)
           
            # Verify new rows are loaded
            new_rows = driver.find_elements(By.XPATH, "//tbody[@id='bidSearchResultsForm:bidResultId_data']/tr")
            if not new_rows:
                print("No rows found after pagination, stopping")
                break
               
            page_number += 1
           
        except NoSuchElementException:
            print(f"\nNo more pages (completed {page_number} pages)")
            break
        except Exception as e:
            print(f"\nError during pagination: {e}")
            break
    # Save to Excel
    df = pd.DataFrame(all_data)
   
    y_sorted = sorted(list(YEARS), reverse=True)
    out_name = f"njstart_bid_details_{y_sorted[0]}_{y_sorted[1]}_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"
    df.to_excel(out_name, index=False)
    print(f"\nScraped {len(df)} records and saved to {out_name}")
finally:
    driver.quit()
