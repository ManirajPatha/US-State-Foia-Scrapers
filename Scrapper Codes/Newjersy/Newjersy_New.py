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
import json
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

def scrape_contract_details(driver, wait, contract_link, bid_number, contract_number, temp_download_dir, main_window):
    """Scrape contract details including actual cost, PO number, vendor info, and attachments."""
    contract_data = {
        "contract_number": contract_number,
        "actual_cost": None,
        "purchase_order_number": None,
        "vendor_id": None,
        "vendor_name": None,
        "vendor_full": None,
        "contract_attachments_count": 0
    }
    
    try:
        # Click contract link
        contract_link.click()
        time.sleep(2)
        
        # Switch to new tab
        for handle in driver.window_handles:
            if handle != main_window:
                driver.switch_to.window(handle)
                break
        
        # Wait for page to load
        wait.until(EC.presence_of_element_located((By.CLASS_NAME, "tableText-01")))
        time.sleep(1)
        
        # Scrape Actual Cost
        try:
            cost_cell = driver.find_element(By.XPATH, "//td[contains(text(),'Actual Cost:')]/following-sibling::td[1]")
            cost_text = safe_text(cost_cell)
            if cost_text:
                contract_data["actual_cost"] = cost_text
        except NoSuchElementException:
            pass
        
        # Scrape Purchase Order Number
        try:
            po_cell = driver.find_element(By.XPATH, "//td[contains(text(),'Purchase Order Number:')]/following-sibling::td[1]")
            contract_data["purchase_order_number"] = safe_text(po_cell)
        except NoSuchElementException:
            pass
        
        # Scrape Vendor Information
        try:
            vendor_link = driver.find_element(By.XPATH, "//a[contains(@href,'viewExternalVendorProfile')]")
            vendor_text = safe_text(vendor_link)
            contract_data["vendor_full"] = vendor_text
            
            # Extract vendor ID and name
            if vendor_text and " - " in vendor_text:
                parts = vendor_text.split(" - ", 1)
                contract_data["vendor_id"] = parts[0].strip()
                contract_data["vendor_name"] = parts[1].strip() if len(parts) > 1 else None
        except NoSuchElementException:
            pass
        
        # Download Agency Attachments
        try:
            attachment_links = driver.find_elements(By.XPATH, "//td[contains(text(),'Agency Attachments:')]/following-sibling::td//a[contains(@href,'downloadFile')]")
            
            if attachment_links:
                print(f"Found {len(attachment_links)} contract attachments for {bid_number}")
                contract_data["contract_attachments_count"] = len(attachment_links)
                #attachments names
                contract_data["contract_attachment_names"] = [safe_text(link) for link in attachment_links if safe_text(link)]
                
                ## downloading functionality

                before_files = set(os.listdir(temp_download_dir))
                
                for link in attachment_links:
                    try:
                        driver.execute_script("arguments[0].click();", link)
                        time.sleep(2)
                    except Exception as e:
                        print(f"Failed to download contract attachment: {e}")
                
                wait_for_downloads(temp_download_dir, timeout=90)
                
                after_files = set(os.listdir(temp_download_dir))
                new_files = list(after_files - before_files)
                
                if new_files:
                    # Add to existing zip or create new one
                    zip_name = f"njstart_{bid_number}.zip"
                    zip_path = BASE_DOWNLOAD_DIR / zip_name
                    
                    mode = "a" if zip_path.exists() else "w"
                    with zipfile.ZipFile(zip_path, mode, zipfile.ZIP_DEFLATED) as zf:
                        existing_names = zf.namelist() if mode == "a" else []
                        for f in new_files:
                            file_path = temp_download_dir / f
                            if file_path.exists():
                                # Add contract number to prefix to distinguish multiple contracts
                                arcname = f"contract_{contract_number}_{f}"
                                if arcname not in existing_names:
                                    zf.write(file_path, arcname=arcname)
                    
                    print(f"Added {len(new_files)} contract attachments for {contract_number} to {zip_path}")
                    
                    # Clean up temp files
                    for f in new_files:
                        try:
                            (temp_download_dir / f).unlink()
                        except:
                            pass
            else:
                print(f"No contract attachments found for {bid_number}")
                contract_data["contract_attachment_names"] = []
        except Exception as e:
            print(f"Error handling contract attachments: {e}")
            contract_data["contract_attachment_names"] = []
        except NoSuchElementException:
            print(f"No contract attachments found for {bid_number}")
        except Exception as e:
            print(f"Error downloading contract attachments: {e}")
        
        # Close tab and switch back
        driver.close()
        driver.switch_to.window(main_window)
        time.sleep(1)
        
    except Exception as e:
        print(f"Error scraping contract details: {e}")
        # Make sure we're back on main window
        try:
            driver.switch_to.window(main_window)
        except:
            pass
    
    return contract_data

def scrape_detail_page(driver, wait, bid_number, temp_download_dir, main_window, row_element):
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
        "bid_attachments_count": 0,
        "contracts": []
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

        # Download File Attachments (Bid Attachments)
        try:
            file_links = driver.find_elements(By.XPATH, "//td[contains(text(),'File Attachments:')]/following-sibling::td//a[contains(@href,'downloadFile')]")
            
            if file_links:
                print(f"Found {len(file_links)} bid attachments for {bid_number}")
                record["bid_attachments_count"] = len(file_links)
                #attachments names
                record["bid_attachment_names"] = [safe_text(link) for link in file_links if safe_text(link)]
                
                ## downloading functionality
                
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
                                # Add bid_ prefix to distinguish
                                zf.write(file_path, arcname=f"bid_{f}")
                    print(f"Zipped {len(new_files)} bid files to {zip_path}")

                    # Clean up temp files
                    for f in new_files:
                        try:
                            (temp_download_dir / f).unlink()
                        except:
                            pass
            else:
                print(f"No bid file attachments found for {bid_number}")
                record["bid_attachment_names"] = []
        except Exception as e:
            print(f"Error handling bid attachments: {e}")
            record["bid_attachment_names"] = []
        except NoSuchElementException:
            print(f"No bid file attachments found for {bid_number}")
        except Exception as e:
            print(f"Error downloading bid attachments: {e}")

        # Close detail page and return to main results
        driver.close()
        driver.switch_to.window(main_window)
        time.sleep(1)

        # Now check for Contract # links in the row (can be multiple)
        try:
            # Re-locate the row by bid number to avoid stale element
            contract_links = driver.find_elements(
                By.XPATH, 
                f"//a[contains(text(),'{bid_number}')]/ancestor::tr//td[.//span[contains(text(),'Contract #')]]//a"
            )
            
            if contract_links:
                print(f"Found {len(contract_links)} contract link(s)")
                
                # Process each contract link
                for idx, contract_link in enumerate(contract_links):
                    try:
                        # Re-locate the link to avoid stale element
                        contract_links_refreshed = driver.find_elements(
                            By.XPATH, 
                            f"//a[contains(text(),'{bid_number}')]/ancestor::tr//td[.//span[contains(text(),'Contract #')]]//a"
                        )
                        
                        if idx < len(contract_links_refreshed):
                            current_link = contract_links_refreshed[idx]
                            contract_number = safe_text(current_link)
                            print(f"Processing contract: {contract_number}")
                            
                            # Scrape contract details
                            contract_data = scrape_contract_details(
                                driver, wait, current_link, bid_number, contract_number, temp_download_dir, main_window
                            )
                            
                            # Add contract data to list
                            record["contracts"].append(contract_data)
                    except Exception as e:
                        print(f"Error processing contract link {idx + 1}: {e}")
                        # Ensure we're back on main window
                        try:
                            driver.switch_to.window(main_window)
                        except:
                            pass
                
        except NoSuchElementException:
            print(f"No contract links found for {bid_number}")
        except Exception as e:
            print(f"Error processing contract links: {e}")

    except Exception as e:
        print(f"Error scraping detail page: {e}")

        try:
            driver.switch_to.window(main_window)
        except:
            pass

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

    # Select status in dropdown
    status_dropdown = wait.until(EC.presence_of_element_located((By.ID, "bidSearchForm:status")))
    select = Select(status_dropdown)
    select.select_by_value("2BPO")

    # Click Search
    search_btn = wait.until(EC.element_to_be_clickable((By.XPATH, "//span[text()='Search']")))
    search_btn.click()

    # Wait for results
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

                # Switch to new tab
                for handle in driver.window_handles:
                    if handle != main_window:
                        driver.switch_to.window(handle)
                        break

                # Scrape detail page (which will also handle contract details)
                record = scrape_detail_page(driver, wait, bid_number, temp_download_dir, main_window, row)
                all_data.append(record)

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
            
            # Click next button
            driver.execute_script("arguments[0].click();", next_btn)
            
            # Wait for page to change - either stale element or new content
            try:
                WebDriverWait(driver, 20).until(EC.staleness_of(first_row))
                print("Page changed successfully")
            except TimeoutException:
                print("Timeout waiting for page change, checking if content updated...")
            
            # Additional wait for new content to load
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

    # Save to Excel and JSON
    df = pd.DataFrame(all_data)
    
    # Flatten the data for Excel - create separate rows for each contract
    flattened_data = []
    for record in all_data:
        if record["contracts"]:
            # Create a row for each contract
            for contract in record["contracts"]:
                flat_record = {
                    "bid_number": record["bid_number"],
                    "description": record["description"],
                    "bid_opening_date": record["bid_opening_date"],
                    "department": record["department"],
                    "location": record["location"],
                    "required_date": record["required_date"],
                    "available_date": record["available_date"],
                    "begin_date": record["begin_date"],
                    "end_date": record["end_date"],
                    "bid_attachments_count": record["bid_attachments_count"],
                    "contract_number": contract["contract_number"],
                    "actual_cost": contract["actual_cost"],
                    "purchase_order_number": contract["purchase_order_number"],
                    "vendor_id": contract["vendor_id"],
                    "vendor_name": contract["vendor_name"],
                    "vendor_full": contract["vendor_full"],
                    "contract_attachments_count": contract["contract_attachments_count"],
                    "bid_attachment_names": record.get("bid_attachment_names", []),
                    "contract_attachment_names": contract.get("contract_attachment_names", []),
                }
                flattened_data.append(flat_record)
        else:
            # No contracts - just add the bid info
            flat_record = {
                "bid_number": record["bid_number"],
                "description": record["description"],
                "bid_opening_date": record["bid_opening_date"],
                "department": record["department"],
                "location": record["location"],
                "required_date": record["required_date"],
                "available_date": record["available_date"],
                "begin_date": record["begin_date"],
                "end_date": record["end_date"],
                "bid_attachments_count": record["bid_attachments_count"],
                "contract_number": None,
                "actual_cost": None,
                "purchase_order_number": None,
                "vendor_id": None,
                "vendor_name": None,
                "vendor_full": None,
                "contract_attachments_count": 0,
                "bid_attachment_names": record.get("bid_attachment_names", []),
                "contract_attachment_names": [],
            }
            flattened_data.append(flat_record)
    
    df_flat = pd.DataFrame(flattened_data)
    
    y_sorted = sorted(list(YEARS), reverse=True)
    timestamp = datetime.now().strftime('%Y%m%d_%H%M')
    

    excel_name = f"njstart_bid_details_{y_sorted[0]}_{y_sorted[1]}_{timestamp}.xlsx"
    df_flat.to_excel(excel_name, index=False)
    print(f"\nScraped {len(all_data)} bids across {page_number} pages with {len(flattened_data)} total records (including contracts)")
    print(f"Excel saved to {excel_name}")
    
    json_name = f"njstart_bid_details_{y_sorted[0]}_{y_sorted[1]}_{timestamp}.json"
    with open(json_name, 'w', encoding='utf-8') as f:
        json.dump(all_data, f, indent=2, ensure_ascii=False)
    print(f"JSON data saved to {json_name}")

finally:
    driver.quit()