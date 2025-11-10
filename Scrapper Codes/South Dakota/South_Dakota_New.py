from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
import pandas as pd
import time
import os
import logging
import glob

logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
BASE_URL = "https://postingboard.esmsolutions.com/3444a404-3818-494f-84c5-2a850acd7779/events"
DOWNLOAD_DIR = os.path.join(os.getcwd(), "downloads")
os.makedirs(DOWNLOAD_DIR, exist_ok=True)
chrome_options = webdriver.ChromeOptions()
chrome_options.add_experimental_option("prefs", {
    "download.default_directory": DOWNLOAD_DIR,
    "download.prompt_for_download": False,
    "download.directory_upgrade": True,
    "safebrowsing.enabled": True,
    "profile.default_content_setting_values.automatic_downloads": 1
})

def scrape_contact_info(driver):
    """Scrape contact information from the detail page."""
    contact_data = {"Name": "", "Phone": "", "Email": "", "Address": ""}
   
    try:
        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, "//div[contains(@class, 'panelContainer')]"))
        )
        time.sleep(3)
       
        # Extract Name
        try:
            name_elem = driver.find_element(By.XPATH, "//b[contains(text(), 'Name:')]/parent::p")
            contact_data["Name"] = name_elem.text.replace('Name:', '').strip()
        except:
            logging.debug("Name not found")
       
        # Extract Phone
        try:
            phone_elem = driver.find_element(By.XPATH, "//b[contains(text(), 'Phone:')]/parent::p")
            contact_data["Phone"] = phone_elem.text.replace('Phone:', '').strip()
        except:
            logging.debug("Phone not found")
       
        # Extract Email
        try:
            email_elem = driver.find_element(By.XPATH, "//b[contains(text(), 'Email:')]/parent::p")
            contact_data["Email"] = email_elem.text.replace('Email:', '').strip()
        except:
            logging.debug("Email not found")
       
        # Extract Address
        try:
            address_elem = driver.find_element(By.XPATH, "//div[contains(@class, 'contact-address')]")
            full_address = address_elem.text.replace('Address:', '').strip()
            contact_data["Address"] = full_address
        except:
            logging.debug("Address not found")
           
    except Exception as e:
        logging.warning(f"Error scraping contact info: {e}")
   
    return contact_data
def wait_and_rename_zip(event_id):
    """Wait for the zip download to complete and rename it with event_id."""
    max_wait = 60
    wait_interval = 2
    elapsed = 0
   
    while elapsed < max_wait:
        # Look for .crdownload files (incomplete downloads)
        crdownloads = glob.glob(os.path.join(DOWNLOAD_DIR, "*.crdownload"))
        if crdownloads:
            time.sleep(wait_interval)
            elapsed += wait_interval
            continue
       
        zip_files = glob.glob(os.path.join(DOWNLOAD_DIR, "*.zip"))
        if zip_files:
            # Get the latest one (most recent modification time)
            latest_zip = max(zip_files, key=os.path.getmtime)
            old_name = os.path.basename(latest_zip)
            new_name = f"{event_id}_event_documents.zip"
            new_path = os.path.join(DOWNLOAD_DIR, new_name)
           
            # Avoid overwriting if already exists
            counter = 1
            while os.path.exists(new_path):
                new_name = f"{event_id}_event_documents_{counter}.zip"
                new_path = os.path.join(DOWNLOAD_DIR, new_name)
                counter += 1
           
            os.rename(latest_zip, new_path)
            logging.info(f"Renamed downloaded zip to: {new_name}")
            return True
       
        time.sleep(wait_interval)
        elapsed += wait_interval
   
    logging.warning(f"Timeout waiting for zip download for Event ID: {event_id}")
    return False
def scrape_event_documents(driver, event_id):
    """Scrape file names from the Event Documents tab without downloading."""
    attachments = []
    try:
        # Click on 'Event Documents' tab
        event_docs_tab = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, "//div[contains(@class, 'mat-tab-label') and contains(text(), 'Event Documents')]"))
        )
        driver.execute_script("arguments[0].click();", event_docs_tab)
        time.sleep(3)
        logging.info(f"Switched to Event Documents tab for Event ID: {event_id}")
        
        # Scrape file names
        file_name_tds = driver.find_elements(By.CSS_SELECTOR, "td.cdk-column-fileName")
        attachments = [td.text.strip() for td in file_name_tds if td.text.strip()]
        logging.info(f"Found {len(attachments)} attachment file names for Event ID: {event_id}")
        
        # Download all documents
        try:
            select_all_checkbox = WebDriverWait(driver, 5).until(
                EC.element_to_be_clickable((By.XPATH, "//label[contains(@class, 'mat-checkbox-layout') and .//span[contains(text(), 'Select all')]]//input[@type='checkbox']"))
            )
            driver.execute_script("arguments[0].click();", select_all_checkbox)
            time.sleep(2)
            logging.info(f"Selected 'Select all' for Event ID: {event_id}")
        except Exception:
            logging.warning(f"'Select all' checkbox not found for Event ID: {event_id}. Skipping document download.")
            return False, attachments
        
        # Click 'Download all' button
        download_btn = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.ID, "downloadFiles"))
        )
        driver.execute_script("arguments[0].click();", download_btn)
        logging.info(f"Initiated download for Event ID: {event_id}")
        # Wait and rename the zip
        success = wait_and_rename_zip(event_id)
        if success:
            logging.info(f"Successfully downloaded and renamed zip for Event ID: {event_id}")
        else:
            logging.error(f"Failed to download zip for Event ID: {event_id}")
        return success, attachments
        
        return True, attachments
    except Exception as e:
        logging.warning(f"Error scraping documents for Event ID {event_id}: {e}")
        return False, attachments
def click_back_arrow(driver):
    """Click the back arrow to return to the list page, refresh if needed."""
    try:
        back_button = WebDriverWait(driver, 15).until(
            EC.element_to_be_clickable((By.XPATH, "//mat-icon[text()='keyboard_arrow_left']"))
        )
        driver.execute_script("arguments[0].click();", back_button)
        time.sleep(4)
        logging.info("Clicked back arrow successfully")
        # Check if the table with past opportunities is visible
        try:
            WebDriverWait(driver, 8).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, "tr.mat-row"))
            )
            logging.info("Past Opportunities list is visible after going back")
        except:
            logging.warning("Past Opportunities list not visible, refreshing page...")
            driver.refresh()
            time.sleep(6)
            # Ensure 'Past Opportunities' tab is active again
            try:
                past_tab = WebDriverWait(driver, 15).until(
                    EC.element_to_be_clickable((By.XPATH, "//div[contains(text(), 'Past Opportunities')]"))
                )
                driver.execute_script("arguments[0].click();", past_tab)
                time.sleep(5)
                logging.info("Reopened 'Past Opportunities' tab after refresh")
            except Exception as e:
                logging.error(f"Could not reopen Past Opportunities tab after refresh: {e}")
        return True
    except Exception as e:
        logging.error(f"Failed to click back arrow: {e}")
        return False
def get_current_page_number(driver):
    """Get the current active page number from pagination."""
    try:
        active_page = driver.find_element(By.XPATH, "//a[contains(@class, 'active') and @class[contains(., 'page-link')]]")
        return int(active_page.text.strip())
    except:
        return 1
def navigate_to_page(driver, page_num):
    """Navigate to a specific page number using pagination controls."""
    try:
        current_page = get_current_page_number(driver)
       
        if current_page == page_num:
            logging.info(f"Already on page {page_num}")
            return True
       
        # Try to click the specific page number
        page_link = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, f"//a[contains(@class, 'page-link') and normalize-space(text())='{page_num}']"))
        )
        driver.execute_script("arguments[0].click();", page_link)
        time.sleep(3)
        logging.info(f"Navigated to page {page_num}")
        return True
       
    except Exception as e:
        logging.warning(f"Could not navigate to page {page_num}: {e}")
        return False
def scrape_row_data(row):
    """Extract basic data from a table row."""
    data = {}
    try:
        data["Event ID"] = row.find_element(By.CSS_SELECTOR, "td.cdk-column-id").text.strip()
    except:
        data["Event ID"] = ""
   
    try:
        data["Event Name"] = row.find_element(By.CSS_SELECTOR, "td.cdk-column-eventName").text.strip()
    except:
        data["Event Name"] = ""
   
    try:
        data["Published Date"] = row.find_element(By.CSS_SELECTOR, "td.cdk-column-publishedDate").text.strip()
    except:
        data["Published Date"] = ""
   
    try:
        data["Award Date"] = row.find_element(By.CSS_SELECTOR, "td.cdk-column-awardDate").text.strip()
    except:
        data["Award Date"] = ""
   
    try:
        data["Event Due Date"] = row.find_element(By.CSS_SELECTOR, "td.cdk-column-eventDueDate").text.strip()
    except:
        data["Event Due Date"] = ""
   
    try:
        data["Invitation Type"] = row.find_element(By.CSS_SELECTOR, "td.cdk-column-invitationType").text.strip()
    except:
        data["Invitation Type"] = ""
   
    try:
        data["Status"] = row.find_element(By.CSS_SELECTOR, "td.cdk-column-status div").text.strip()
    except:
        data["Status"] = ""
   
    return data
# Main Scraping Logic
driver = None
try:
    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=chrome_options)
    logging.info("WebDriver initialized successfully")
   
    # Navigate to the website
    driver.get(BASE_URL)
    time.sleep(5)
   
    # Click "Past Opportunities" tab
    logging.info("Clicking 'Past Opportunities' tab")
    past_opportunities_tab = WebDriverWait(driver, 20).until(
        EC.element_to_be_clickable((By.XPATH, "//div[contains(text(), 'Past Opportunities')]"))
    )
    driver.execute_script("arguments[0].click();", past_opportunities_tab)
    time.sleep(5)
   
    all_data = []
    page_count = 0
    
    # Main pagination loop
    while True:
        page_count += 1
        logging.info(f"=" * 60)
        logging.info(f"PROCESSING PAGE {page_count}")
        logging.info(f"=" * 60)

        PAGE_LIMIT = 141
        if page_count > PAGE_LIMIT:
            logging.info(f"Reached page limit ({PAGE_LIMIT}). Stopping further scraping.")
            break
       
        try:
            rows = WebDriverWait(driver, 20).until(
                EC.presence_of_all_elements_located((By.CSS_SELECTOR, "tr.mat-row"))
            )
            logging.info(f"Found {len(rows)} total rows on page {page_count}")
        except Exception as e:
            logging.warning(f"No rows found on page {page_count}: {e}")
            break
       
        # First pass: collect all "Awarded" row data and their indices
        awarded_rows_info = []
       
        for idx, row in enumerate(rows):
            try:
                status = row.find_element(By.CSS_SELECTOR, "td.cdk-column-status div").text.strip()
               
                if status == "Awarded":
                    row_data = scrape_row_data(row)
                    awarded_rows_info.append({
                        "index": idx,
                        "data": row_data
                    })
                    logging.info(f"Found Awarded opportunity at index {idx}: {row_data.get('Event ID')} - {row_data.get('Event Name')}")
            except Exception as e:
                logging.debug(f"Error checking row {idx}: {e}")
                continue
       
        logging.info(f"Total Awarded opportunities on page {page_count}: {len(awarded_rows_info)}")
       
        # Second pass: click each awarded row and scrape contact info + scrape documents
        for awarded_info in awarded_rows_info:
            idx = awarded_info["index"]
            row_data = awarded_info["data"]
           
            try:
                # Re-fetch rows to avoid stale element reference
                rows_fresh = WebDriverWait(driver, 10).until(
                    EC.presence_of_all_elements_located((By.CSS_SELECTOR, "tr.mat-row"))
                )
               
                # Click the specific row
                target_row = rows_fresh[idx]
                driver.execute_script("arguments[0].scrollIntoView(true);", target_row)
                time.sleep(2)
                driver.execute_script("arguments[0].click();", target_row)
                logging.info(f"Clicked on row: {row_data.get('Event ID')}")
                time.sleep(4)
               
                contact_info = scrape_contact_info(driver)

                success, attachments = scrape_event_documents(driver, row_data.get('Event ID'))
               
                attachments_str = "; ".join(attachments) if attachments else ""
                complete_data = {**row_data, **contact_info, "Attachments": attachments_str}
                all_data.append(complete_data)
               
                logging.info(f"Scraped contact and attachments: {contact_info.get('Name')} | {contact_info.get('Email')} | Event ID: {row_data.get('Event ID')} | Attachments: {attachments_str[:100]}...")
               
                click_back_arrow(driver)
                time.sleep(3)
               
                # Navigate back to the correct page
                navigate_to_page(driver, page_count)
                time.sleep(3)
               
            except Exception as e:
                logging.error(f"Error processing awarded row {idx} on page {page_count}: {e}")
               
                # Try to recover by clicking back
                try:
                    click_back_arrow(driver)
                    navigate_to_page(driver, page_count)
                except:
                    logging.error("Could not recover from error")
                continue
       
        # Try to go to next page
        try:
            next_button = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((
                    By.XPATH,
                    "//div[contains(@class, 'page-item') and contains(@class, 'arrow') and not(contains(@class, 'disabled'))]"
                    "//mat-icon[text()='keyboard_arrow_right']/ancestor::a"
                ))
            )
            driver.execute_script("arguments[0].click();", next_button)
            logging.info(f"Moving to page {page_count + 1}")
            time.sleep(5)
        except Exception as e:
            logging.info(f"No more pages available. Scraping complete.")
            break
   
    # Save to Excel and JSON
    output_file = os.path.join(os.getcwd(), "Past_Awarded_Opportunities_With_Contacts.xlsx")
    json_file = os.path.join(os.getcwd(), "Past_Awarded_Opportunities_With_Contacts.json")
   
    if all_data:
        df = pd.DataFrame(all_data)
        df.to_excel(output_file, index=False, sheet_name="Awarded_Opportunities")
        df.to_json(json_file, orient='records', indent=4)
        logging.info(f"")
        logging.info(f"{'='*60}")
        logging.info(f"SUCCESS! Excel saved to: {output_file}")
        logging.info(f"SUCCESS! JSON saved to: {json_file}")
        logging.info(f"Total Awarded opportunities scraped: {len(all_data)}")
        logging.info(f"{'='*60}")
    else:
        logging.warning("No Awarded opportunities found")
        df = pd.DataFrame(columns=[
            "Event ID", "Event Name", "Published Date", "Award Date",
            "Event Due Date", "Invitation Type", "Status",
            "Name", "Phone", "Email", "Address", "Attachments"
        ])
        df.loc[0] = ["No Data", "", "", "", "", "", "Awarded", "", "", "", ""]
        df.to_excel(output_file, index=False)
        df.to_json(json_file, orient='records', indent=4)
        logging.info(f"Empty Excel file saved: {output_file}")
        logging.info(f"Empty JSON file saved: {json_file}")
except Exception as e:
    logging.error(f"CRITICAL ERROR: {e}")
    if driver:
        with open(os.path.join(DOWNLOAD_DIR, "error_page_source.html"), "w") as f:
            f.write(driver.page_source)
        logging.info("Error page source saved to downloads/error_page_source.html")
finally:
    if driver:
        logging.info("Closing browser...")
        driver.quit()
        logging.info("Browser closed successfully")