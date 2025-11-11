import re
import time
import os
from urllib.parse import urljoin
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from datetime import datetime, timedelta
from typing import Optional, List, Dict
from dateutil.relativedelta import relativedelta
import tempfile
import pandas as pd
import zipfile
import shutil

def parse_date(date_str: str) -> Optional[str]:
    try:
        return datetime.strptime(date_str, "%m/%d/%Y").strftime("%Y-%m-%d")
    except:
        return None

def parse_time(time_str: str) -> Optional[str]:
    try:
        return datetime.strptime(time_str, "%I:%M %p").strftime("%H:%M:%S")
    except:
        return None

def split_close(close_str):
    """Return (date, time) from 'MM/DD/YYYY HH:MM AM'"""
    m = re.match(r"(\d{2}/\d{2}/\d{4})\s+(\d{1,2}:\d{2}\s+[AP]M)", close_str)
    return m.groups() if m else (close_str, None)

def js_click(elm): 
    driver.execute_script("arguments[0].click();", elm)

HOME = "https://apps.nd.gov/csd/spo/services/bidder/main.htm"

KEYWORD = "RFP"
DOWNLOAD_DIR = os.path.join(os.getcwd(), "downloads")
ATTACHMENTS_DIR = os.path.join(DOWNLOAD_DIR, "temp_attachments")
os.makedirs(DOWNLOAD_DIR, exist_ok=True)
os.makedirs(ATTACHMENTS_DIR, exist_ok=True)

opts = Options()
opts.add_argument("--disable-gpu")
opts.add_argument("--no-sandbox")
prefs = {
    "download.default_directory": os.path.abspath(ATTACHMENTS_DIR),
    "download.prompt_for_download": False,
    "download.directory_upgrade": True,
    "safebrowsing.enabled": True,
    "profile.default_content_setting_values.automatic_downloads": 1,
}
opts.add_experimental_option("prefs", prefs)
driver = webdriver.Chrome(options=opts)

wait = WebDriverWait(driver, 20)


def scrape_north_dakota():
    """Main function to scrape North Dakota RFPs with Award Status"""
    
    all_data = []

    try:
        driver.get(HOME)
        js_click(wait.until(EC.element_to_be_clickable((By.LINK_TEXT, "Search All Solicitations"))))
        
        # Calculate dates: Begin = 1 year before yesterday, End = yesterday
        yesterday = datetime.now() - timedelta(days=1)
        begin_date = yesterday - relativedelta(years=1)
        
        driver.find_element(By.ID, "x40").clear()
        driver.find_element(By.ID, "x40").send_keys(begin_date.strftime("%m/%d/%Y"))
        
        driver.find_element(By.ID, "x50").clear()
        driver.find_element(By.ID, "x50").send_keys(yesterday.strftime("%m/%d/%Y"))
        
        js_click(driver.find_element(By.NAME, "searchSearchSolicitation"))
        wait.until(lambda d: d.execute_script("return document.readyState") == "complete")

        res_table = wait.until(EC.presence_of_element_located(
            (By.XPATH, "//th[normalize-space()='Closes']/ancestor::table")
        ))
        rows = res_table.find_elements(By.XPATH, ".//tbody/tr[td]")

        print(f"Found {len(rows)} opportunities to process")

        for idx, r in enumerate(rows, 1):
            try:
                tds = r.find_elements(By.TAG_NAME, "td")
                close_full = tds[0].text.strip()
                close_date_raw, close_time_raw = split_close(close_full)
                close_date = parse_date(close_date_raw)
                close_time = parse_time(close_time_raw) if close_time_raw else None

                link = tds[1].find_element(By.TAG_NAME, "a")
                sol_number = link.text.strip()
                agency = tds[3].text.strip()
                officer = tds[4].text.strip()

                print(f"\n[{idx}/{len(rows)}] Processing: {sol_number}")

                js_click(link)
                wait.until(EC.presence_of_element_located(
                    (By.XPATH, "//th[normalize-space()='Number:']/ancestor::table")
                ))

                status_element = driver.find_elements(
                    By.XPATH, "//th[normalize-space()='Status:']/following-sibling::td"
                )
                
                if not status_element:
                    print(f"Status not found. Skipping...")
                    driver.back()
                    wait.until(EC.presence_of_element_located(
                        (By.XPATH, "//th[normalize-space()='Closes']/ancestor::table")
                    ))
                    continue
                
                status = status_element[0].text.strip()
                print(f"  Status: {status}")
                
                if status != "Notice of Award Issued":
                    print(f"Status is not 'Notice of Award Issued'. Skipping...")
                    driver.back()
                    wait.until(EC.presence_of_element_located(
                        (By.XPATH, "//th[normalize-space()='Closes']/ancestor::table")
                    ))
                    continue

                print(f"Status matches! Scraping data...")

                # Detail fields
                fld = {}
                for tr in driver.find_element(By.XPATH, "//th[normalize-space()='Number:']/ancestor::table")\
                                .find_elements(By.XPATH, ".//tr[th and td]"):
                    k = tr.find_element(By.TAG_NAME, "th").text.rstrip(":").strip()
                    v = tr.find_element(By.TAG_NAME, "td").text.strip()
                    fld[k] = v

                award_bidders = []
                award_amounts = []
                try:
                    award_table = driver.find_elements(
                        By.XPATH, "//h3[@class='view' and contains(text(), 'Award Notice')]/following-sibling::table[1]"
                    )
                    if award_table:
                        award_rows = award_table[0].find_elements(By.XPATH, ".//tbody/tr[td]")
                        for award_row in award_rows:
                            award_tds = award_row.find_elements(By.TAG_NAME, "td")
                            if len(award_tds) >= 2:
                                bidder = award_tds[0].text.strip()
                                amount = award_tds[1].text.strip()
                                award_bidders.append(bidder)
                                award_amounts.append(amount)
                        print(f"Found {len(award_bidders)} award(s)")
                except Exception as e:
                    print(f"Could not scrape award info: {e}")

                agency_info = {}
                try:
                    agency_table = driver.find_elements(
                        By.XPATH, "//th[normalize-space()='Procurement Officer:']/ancestor::table[1]"
                    )
                    if agency_table:
                        current_key = None
                        for tr in agency_table[0].find_elements(By.XPATH, ".//tr"):
                            ths = tr.find_elements(By.TAG_NAME, "th")
                            tds = tr.find_elements(By.TAG_NAME, "td")
                            if ths and tds:
                                key_text = ths[0].text.strip().rstrip(":").strip()
                                if key_text:
                                    current_key = key_text
                                    value = tds[0].text.strip()
                                    agency_info[current_key] = value
                                elif current_key:
                                    value = tds[0].text.strip()
                                    if value:
                                        agency_info[current_key] += " " + value
                        print(f"Scraped agency information: {agency_info}")
                except Exception as e:
                    print(f"Could not scrape agency info: {e}")

                opp_folder = os.path.join(ATTACHMENTS_DIR, sol_number.replace("/", "_"))
                os.makedirs(opp_folder, exist_ok=True)

                # Attachments
                attachment_count = 0
                for row in driver.find_elements(
                    By.XPATH, "//th[normalize-space()='Title']/ancestor::table[1]//tbody/tr[td]"):
                    t = row.find_elements(By.TAG_NAME, "td")
                    a = t[-1].find_elements(By.TAG_NAME, "a")
                    if not a:
                        continue
                    filename = t[0].text.strip()
                    try:
                        js_click(a[0])
                        time.sleep(3)
                        attachment_count += 1
                    except Exception as e:
                        print(f"    Warning: Could not download {filename}: {e}")

                time.sleep(2)
                for file in os.listdir(ATTACHMENTS_DIR):
                    file_path = os.path.join(ATTACHMENTS_DIR, file)
                    if os.path.isfile(file_path):
                        shutil.move(file_path, os.path.join(opp_folder, file))

                if attachment_count > 0:
                    zip_filename = f"{sol_number.replace('/', '_')}_attachments.zip"
                    zip_path = os.path.join(DOWNLOAD_DIR, zip_filename)
                    with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
                        for root, dirs, files in os.walk(opp_folder):
                            for file in files:
                                file_path = os.path.join(root, file)
                                arcname = os.path.join(sol_number.replace('/', '_'), file)
                                zipf.write(file_path, arcname)
                    print(f"Created zip: {zip_filename} ({attachment_count} files)")
                else:
                    print(f"No attachments to zip")

                pub_date = parse_date(fld.get("Issued", ""))
                
                record = {
                    "Solicitation_Number": sol_number,
                    "Status": status,
                    "Agency": agency,
                    "Officer": officer,
                    "Close_Date": close_date,
                    "Close_Time": close_time,
                    "Published_Date": pub_date,
                    "Title": fld.get("Title", ""),
                    "Description": fld.get("Description", ""),
                    "Type": fld.get("Type", ""),
                    # Award Information
                    "Awarded_Bidder": " | ".join(award_bidders) if award_bidders else "",
                    "Award_Amount": " | ".join(award_amounts) if award_amounts else "",
                    # Agency Information
                    "Procurement_Officer": agency_info.get("Procurement Officer", ""),
                    "Address": agency_info.get("Address", ""),
                    "Telephone": agency_info.get("Telephone", ""),
                    "Fax": agency_info.get("Fax", ""),
                    "TTY": agency_info.get("TTY", ""),
                    "Email": agency_info.get("Email", ""),
                }
                
                for k, v in fld.items():
                    if k not in record:
                        record[k] = v
                
                all_data.append(record)
                print(f"Successfully scraped {sol_number}")

                driver.back()
                wait.until(EC.presence_of_element_located(
                    (By.XPATH, "//th[normalize-space()='Closes']/ancestor::table")
                ))
                
            except Exception as e:
                print(f"Error processing opportunity: {e}")
                try:
                    driver.back()
                    wait.until(EC.presence_of_element_located(
                        (By.XPATH, "//th[normalize-space()='Closes']/ancestor::table")
                    ))
                except:
                    pass
                continue

        # Save to Excel
        if all_data:
            df = pd.DataFrame(all_data)
            excel_filename = f"ND_Awards_{yesterday.strftime('%Y%m%d')}.xlsx"
            excel_path = os.path.join(DOWNLOAD_DIR, excel_filename)
            df.to_excel(excel_path, index=False, engine='openpyxl')
            print(f"\nExcel file created: {excel_filename}")
            print(f"   Total awarded opportunities scraped: {len(all_data)}")
        else:
            print("\nNo opportunities with 'Notice of Award Issued' status found")

    finally:
        if os.path.exists(ATTACHMENTS_DIR):
            shutil.rmtree(ATTACHMENTS_DIR)
        driver.quit()
        print("\nScraping completed!")

if __name__ == "__main__":
    scrape_north_dakota()