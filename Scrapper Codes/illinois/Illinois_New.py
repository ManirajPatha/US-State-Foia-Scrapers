from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import NoSuchElementException, TimeoutException
import pandas as pd
import json
import time
import re
import os
import zipfile
from pathlib import Path
from datetime import datetime

URL = "https://www.bidbuy.illinois.gov/bso/view/search/external/advancedSearchBid.xhtml?openBids=true"
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

def scrape_blanket_page(driver, wait, blanket_number):
    """Scrape details from the blanket/PO page."""
    blanket_data = {
        "blanket_number": blanket_number,
        "actual_cost": None,
        "actual_contract_begin_date": None,
        "actual_contract_end_date": None,
        "organization": None,
        "vendor_name": None,
        "vendor_address": None,
        "vendor_email": None,
        "vendor_phone": None
    }

    try:
        # Wait for page to load
        wait.until(EC.presence_of_element_located((By.CLASS_NAME, "tableText-01")))
        time.sleep(1)

        # Extract Actual Cost
        try:
            cost_cell = driver.find_element(By.XPATH, "//td[contains(text(),'Actual Cost:')]/following-sibling::td[1]")
            blanket_data["actual_cost"] = safe_text(cost_cell)
        except NoSuchElementException:
            pass

        # Extract Actual Contract Begin Date
        try:
            begin_date_cell = driver.find_element(By.XPATH, "//td[contains(text(),'Actual Contract Begin Date:')]/following-sibling::td[1]")
            blanket_data["actual_contract_begin_date"] = safe_text(begin_date_cell)
        except NoSuchElementException:
            pass

        # Extract Actual Contract End Date
        try:
            end_date_cell = driver.find_element(By.XPATH, "//td[contains(text(),'Actual Contract End Date:')]/following-sibling::td[1]")
            blanket_data["actual_contract_end_date"] = safe_text(end_date_cell)
        except NoSuchElementException:
            pass

        # Extract Organization
        try:
            org_cell = driver.find_element(By.XPATH, "//td[contains(text(),'Organization:')]/following-sibling::td[1]")
            blanket_data["organization"] = safe_text(org_cell)
        except NoSuchElementException:
            pass

        # Scroll down to see vendor details
        driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
        time.sleep(1)

        # Extract Vendor Information
        try:
            # Find the vendor cell - it has class tableText-01 and contains vendor profile link
            vendor_cell = driver.find_element(By.XPATH, "//td[@class='tableText-01'][@valign='top'][.//a[contains(@href, 'viewExternalVendorProfile')]]")
            
            # Get the inner HTML to properly parse the structure
            inner_html = vendor_cell.get_attribute('innerHTML')
            
            # Extract vendor name from the link
            try:
                vendor_link = vendor_cell.find_element(By.XPATH, ".//a[contains(@href, 'viewExternalVendorProfile')]")
                blanket_data["vendor_name"] = safe_text(vendor_link)
            except:
                pass
            
            # Parse the HTML to extract text after <br> tags
            # Split by <br> and clean up HTML tags
            parts = re.split(r'<br\s*/?>', inner_html, flags=re.IGNORECASE)
            
            # Clean each part from HTML tags
            cleaned_parts = []
            for part in parts:
                # Remove HTML tags
                clean = re.sub(r'<[^>]+>', '', part)
                clean = clean.strip()
                if clean:
                    cleaned_parts.append(clean)
            
            # Now extract information from cleaned parts
            # The structure is typically:
            # [0] = Vendor ID - Name (from link, already captured)
            # [1+] = Contact person name, address lines, US, Email:, Phone:, FAX:
            
            address_lines = []
            for part in cleaned_parts:
                # Skip if it contains the vendor ID (already have vendor name)
                if part.startswith('V0') or 'viewExternalVendorProfile' in part:
                    continue
                
                # Extract Email
                if 'Email:' in part:
                    email_match = re.search(r'Email:\s*([\w\.-]+@[\w\.-]+\.\w+)', part)
                    if email_match:
                        blanket_data["vendor_email"] = email_match.group(1)
                    continue
                
                # Extract Phone
                if 'Phone:' in part:
                    phone_match = re.search(r'Phone:\s*(.+)', part)
                    if phone_match:
                        blanket_data["vendor_phone"] = phone_match.group(1).strip()
                    continue
                
                # Skip FAX line
                if 'FAX:' in part:
                    continue
                
                # Everything else is part of the address
                address_lines.append(part)
            
            # Combine address lines
            if address_lines:
                blanket_data["vendor_address"] = ', '.join(address_lines)
            
            print(f"  ✅ Vendor: {blanket_data['vendor_name']}")
            print(f"     Email: {blanket_data['vendor_email']}")
            print(f"     Phone: {blanket_data['vendor_phone']}")
            
        except NoSuchElementException:
            print(f"  ℹ️ No vendor information found")
        except Exception as e:
            print(f"  ⚠️ Error extracting vendor info: {e}")

    except Exception as e:
        print(f"  ❌ Error scraping blanket page: {e}")

    return blanket_data

def scrape_detail_page(driver, wait, bid_number, temp_download_dir, row_element):
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
        "blanket_number": None,
        "actual_cost": None,
        "actual_contract_begin_date": None,
        "actual_contract_end_date": None,
        "organization": None,
        "vendor_name": None,
        "vendor_address": None,
        "vendor_email": None,
        "vendor_phone": None,
        "attachments": None
    }

    try:
        # Wait for page to load
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

        # Department and Location
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

        # Required Date
        try:
            req_cell = driver.find_element(By.XPATH, "//td[contains(text(),'Required Date:')]/following-sibling::td[1]")
            record["required_date"] = safe_text(req_cell)
        except NoSuchElementException:
            pass

        # Available Date
        try:
            avail_cell = driver.find_element(By.XPATH, "//td[contains(text(),'Available Date')]/following-sibling::td[1]")
            record["available_date"] = safe_text(avail_cell)
        except NoSuchElementException:
            pass

        # Begin and End Dates
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

        # Download File Attachments and capture filenames
        try:
            file_links = driver.find_elements(By.XPATH, "//td[contains(text(),'File Attachments:')]/following-sibling::td//a[contains(@href,'downloadFile')]")
            
            if file_links:
                print(f"  📥 Found {len(file_links)} attachments for {bid_number}")
                
                # Capture attachment filenames
                attachment_names = []
                for link in file_links:
                    try:
                        filename = safe_text(link)
                        if filename:
                            attachment_names.append(filename)
                    except:
                        pass
                
                # Store attachment names in record
                if attachment_names:
                    record["attachments"] = "; ".join(attachment_names)
                    print(f"  📎 Attachments: {record['attachments']}")
                
                before_files = set(os.listdir(temp_download_dir))

                for link in file_links:
                    try:
                        driver.execute_script("arguments[0].click();", link)
                        time.sleep(2)
                    except Exception as e:
                        print(f"  ⚠️ Failed to download a file: {e}")

                wait_for_downloads(temp_download_dir, timeout=90)

                after_files = set(os.listdir(temp_download_dir))
                new_files = list(after_files - before_files)

                if new_files:
                    # Create zip file
                    zip_name = f"illinois_{bid_number}.zip"
                    zip_path = BASE_DOWNLOAD_DIR / zip_name
                    with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as zf:
                        for f in new_files:
                            file_path = temp_download_dir / f
                            if file_path.exists():
                                zf.write(file_path, arcname=f)
                    print(f"  ✅ Zipped {len(new_files)} files to {zip_path}")

                    # Clean up temp files
                    for f in new_files:
                        try:
                            (temp_download_dir / f).unlink()
                        except:
                            pass
        except NoSuchElementException:
            print(f"  ℹ️ No file attachments found for {bid_number}")
        except Exception as e:
            print(f"  ⚠️ Error downloading attachments: {e}")

    except Exception as e:
        print(f"  ❌ Error scraping detail page: {e}")

    return record

# ----------------------------
# Selenium setup
# ----------------------------
temp_download_dir = Path("/tmp/illinois_downloads")
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
    print("📍 Loaded Illinois BidBuy website")

    # Click on Advanced Search to expand it
    try:
        advanced_search_legend = wait.until(
            EC.element_to_be_clickable((By.XPATH, "//legend[contains(@class, 'ui-fieldset-legend')]"))
        )
        # Check if it's collapsed (has plusthick icon)
        if "ui-icon-plusthick" in driver.page_source:
            print("🔽 Expanding Advanced Search...")
            advanced_search_legend.click()
            time.sleep(1)
    except Exception as e:
        print(f"⚠️ Could not click Advanced Search (might already be expanded): {e}")

    # Select status in dropdown
    print("🔍 Selecting 'Bid to PO' status...")
    status_dropdown = wait.until(EC.presence_of_element_located((By.ID, "bidSearchForm:status")))
    select = Select(status_dropdown)
    select.select_by_value("2BPO")

    # Click Search
    print("🔍 Clicking Search button...")
    search_btn = wait.until(EC.element_to_be_clickable((By.XPATH, "//span[text()='Search']")))
    search_btn.click()

    # Wait for results
    wait.until(EC.presence_of_element_located((By.ID, "bidSearchResultsForm:bidResultId_head")))
    time.sleep(3)
    print("✅ Search results loaded")

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
        print(f"📅 Date column detected at index {date_col_idx}")
        date_filter_enabled = True

    all_data = []
    main_window = driver.current_window_handle

    while True:
        # Get current page rows
        rows = driver.find_elements(By.XPATH, "//tbody[@id='bidSearchResultsForm:bidResultId_data']/tr")
        if not rows:
            break

        first_row_ref = rows[0]
        
        # Process each row
        for idx, row in enumerate(rows):
            try:
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
                            print(f"⏭️ Skipping {bid_number} (date filter)")
                            continue

                print(f"\n🔄 Processing {bid_number}...")

                # Check if there's a Blanket # link in this row
                blanket_link = None
                blanket_number = None
                try:
                    blanket_link = row.find_element(By.XPATH, ".//td[contains(@role,'gridcell')]//span[contains(text(),'Blanket #')]/following-sibling::a")
                    blanket_number = safe_text(blanket_link)
                    if blanket_number:
                        print(f"  🔗 Found Blanket #: {blanket_number}")
                except NoSuchElementException:
                    pass

                # Click to open bid detail in new tab
                bid_link.click()
                time.sleep(2)

                # Switch to new tab
                for handle in driver.window_handles:
                    if handle != main_window:
                        driver.switch_to.window(handle)
                        break

                # Scrape detail page
                record = scrape_detail_page(driver, wait, bid_number, temp_download_dir, row)

                # Close bid detail tab and switch back to main window
                driver.close()
                driver.switch_to.window(main_window)
                time.sleep(1)

                # If there's a blanket link, open it and scrape
                if blanket_link and blanket_number:
                    try:
                        # Re-find the row and blanket link (in case of stale element)
                        rows = driver.find_elements(By.XPATH, "//tbody[@id='bidSearchResultsForm:bidResultId_data']/tr")
                        for r in rows:
                            try:
                                bid_link_check = r.find_element(By.XPATH, ".//td[@role='gridcell']//a")
                                if safe_text(bid_link_check) == bid_number:
                                    blanket_link = r.find_element(By.XPATH, ".//td[contains(@role,'gridcell')]//span[contains(text(),'Blanket #')]/following-sibling::a")
                                    break
                            except:
                                continue
                        
                        print(f"  🔗 Opening Blanket page: {blanket_number}")
                        blanket_link.click()
                        time.sleep(2)

                        # Switch to blanket tab
                        for handle in driver.window_handles:
                            if handle != main_window:
                                driver.switch_to.window(handle)
                                break

                        # Scrape blanket page
                        blanket_data = scrape_blanket_page(driver, wait, blanket_number)
                        
                        # Merge blanket data into record
                        record.update(blanket_data)

                        # Close blanket tab and switch back
                        driver.close()
                        driver.switch_to.window(main_window)
                        time.sleep(1)

                    except Exception as e:
                        print(f"  ⚠️ Error processing blanket link: {e}")
                        driver.switch_to.window(main_window)

                all_data.append(record)

                # Re-fetch rows to avoid stale element
                rows = driver.find_elements(By.XPATH, "//tbody[@id='bidSearchResultsForm:bidResultId_data']/tr")

            except Exception as e:
                print(f"⚠️ Error processing row: {e}")
                # Make sure we're back on main window
                driver.switch_to.window(main_window)
                continue

        # Try to click Next button
        try:
            next_btn = driver.find_element(By.XPATH, "//a[contains(@class,'ui-paginator-next')]")
            if "ui-state-disabled" in (next_btn.get_attribute("class") or ""):
                print("\n✅ Reached last page")
                break
            print("➡️ Going to next page...")
            next_btn.click()
            try:
                WebDriverWait(driver, 20).until(EC.staleness_of(first_row_ref))
            except Exception:
                time.sleep(1.5)
        except Exception:
            print("\n✅ No more pages")
            break

    # Generate filenames with timestamp
    y_sorted = sorted(list(YEARS), reverse=True)
    timestamp = datetime.now().strftime('%Y%m%d_%H%M')
    base_filename = f"illinois_bid_details_{y_sorted[0]}_{y_sorted[1]}_{timestamp}"
    
    # Save to Excel
    df = pd.DataFrame(all_data)
    excel_filename = f"{base_filename}.xlsx"
    df.to_excel(excel_filename, index=False)
    print(f"\n✅ Scraped {len(df)} records and saved to {excel_filename}")
    
    # Save to JSON
    json_filename = f"{base_filename}.json"
    with open(json_filename, 'w', encoding='utf-8') as f:
        json.dump(all_data, f, indent=2, ensure_ascii=False)
    print(f"✅ JSON file saved to {json_filename}")

finally:
    driver.quit()