import os
import re
import time
import json
import zipfile
import requests
import pandas as pd
from bs4 import BeautifulSoup
from urllib.parse import urljoin
from datetime import datetime
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

# =========================
# CONFIG
# =========================
BASE_URL = "https://www.ms.gov/dfa/contract_bid_search/Bid?autoloadGrid=False"

timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
BASE_OUTPUT_DIR = "/home/developer/Desktop/US-State-Foia-Scrapers-2/Scrapper Codes/mississippi/mississippi_awarded_data"
OUTPUT_DIR = os.path.join(BASE_OUTPUT_DIR, f"run_{timestamp}")
EXCEL_PATH = os.path.join(OUTPUT_DIR, "mississippi_awarded_data.xlsx")
JSON_PATH = os.path.join(OUTPUT_DIR, "mississippi_awarded_data.json")
os.makedirs(OUTPUT_DIR, exist_ok=True)

# =========================
# SELENIUM SETUP
# =========================
chrome_options = Options()
# chrome_options.add_argument("--headless=new")
chrome_options.add_argument("--no-sandbox")
chrome_options.add_argument("--disable-dev-shm-usage")
driver = webdriver.Chrome(options=chrome_options)
wait = WebDriverWait(driver, 15)

# =========================
# HELPERS
# =========================
def safe_filename(name):
    return re.sub(r'[\\/*?:"<>|]', "_", name)

def download_file(url, dest_folder):
    local_filename = os.path.join(dest_folder, safe_filename(url.split("/")[-1].split("?")[0]))
    for attempt in range(3):
        try:
            with requests.get(url, stream=True, timeout=60) as r:
                r.raise_for_status()
                with open(local_filename, 'wb') as f:
                    for chunk in r.iter_content(chunk_size=8192):
                        f.write(chunk)
            return local_filename
        except Exception as e:
            print(f"⚠️ Download error ({attempt+1}/3): {e}")
            time.sleep(3)
    return None

def extract_zip_if_needed(file_path, dest_folder):
    if file_path and file_path.lower().endswith(".zip"):
        try:
            with zipfile.ZipFile(file_path, 'r') as zip_ref:
                zip_ref.extractall(dest_folder)
            os.remove(file_path)
        except Exception as e:
            print(f"⚠️ Error extracting nested zip: {e}")

def click_element_safe(by, value, timeout=10):
    try:
        elem = WebDriverWait(driver, timeout).until(EC.element_to_be_clickable((by, value)))
        driver.execute_script("arguments[0].click();", elem)
        return True
    except Exception as e:
        print(f"⚠️ Could not click element: {value} -> {e}")
        return False

# =========================
# MAIN SCRAPER
# =========================
def scrape_awarded_data():
    driver.get(BASE_URL)
    print("🌐 Opened Mississippi DFA site")

    # ---- Open Advanced Search ----
    click_element_safe(By.ID, "advanceSearchToggle")
    print("✅ Clicked 'Advanced Search Options'")

    # ---- Select 'Awarded' ----
    try:
        status_dropdown = WebDriverWait(driver, 15).until(
            EC.presence_of_element_located((By.ID, "Status"))
        )
        status_dropdown.find_element(By.XPATH, ".//option[contains(text(),'Awarded')]").click()
        print("✅ Selected 'Awarded'")
    except Exception as e:
        print(f"⚠️ Couldn't select Awarded: {e}")

    # ---- Click Search ----
    click_element_safe(By.ID, "btnSearch")
    print("🔍 Search initiated")
    time.sleep(5)

    all_data = []
    seen = set()
    page_num = 1

    while True:
        WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.CSS_SELECTOR, "table tbody tr")))
        soup = BeautifulSoup(driver.page_source, "html.parser")
        rows = soup.select("table tbody tr")
        print(f"📄 Found {len(rows)} records on page {page_num}")

        for row in rows:
            cells = row.find_all("td")
            if len(cells) < 8:
                continue

            smart_link = cells[1].find("a")
            if not smart_link:
                continue

            smart_number = smart_link.text.strip()
            detail_url = urljoin(BASE_URL, smart_link["href"])

            if smart_number in seen:
                continue
            seen.add(smart_number)

            print(f"🔹 Scraping {smart_number} ...")
            driver.get(detail_url)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "lblBidNumber")))

            detail_soup = BeautifulSoup(driver.page_source, "html.parser")

            # ----------------------------
            # Procurement Details
            # ----------------------------
            def get_text_by_id(id_):
                tag = detail_soup.find(id=id_)
                return tag.text.strip() if tag else ""

            procurement_data = {
                "Smart Number": get_text_by_id("lblBidNumber"),
                "Advertised Date": get_text_by_id("lblAdvertisedDate"),
                "RFx #": get_text_by_id("lblObjectId"),
                "Submission Date": get_text_by_id("lblSubmissionDate"),
                "RFx Status": get_text_by_id("lblBidStatus"),
                "Major Procurement Category": get_text_by_id("lblProcCategory"),
                "Sub Procurement Category": get_text_by_id("lblSubProcCategory"),
                "RFx Type": get_text_by_id("lblBidType"),
                "Agency": get_text_by_id("lblAgency"),
                "RFx Description": get_text_by_id("lblDescription"),
            }

            contact_info = {
                "Name": get_text_by_id("lblContactName"),
                "Email": get_text_by_id("lblContactEmail"),
                "Phone": get_text_by_id("lblContactPhone"),
                "Fax": get_text_by_id("lblContactFax"),
            }

            # ----------------------------
            # Vendor Table Extraction (No Tab Click Needed)
            # ----------------------------
            vendors = []
            award_section = detail_soup.find("h2", string="Awarded")
            if award_section:
                vendor_table = award_section.find_next("table", class_="dataGrid")
                if vendor_table:
                    v_rows = vendor_table.find_all("tr")[1:]  # skip header
                    for vr in v_rows:
                        cols = [td.text.strip() for td in vr.find_all("td")]
                        if len(cols) >= 5:
                            vendors.append({
                                "Vendor Name": cols[0],
                                "Vendor Number": cols[1],
                                "Award Date": cols[2],
                                "Award Amount": cols[3],
                                "Funding Source": cols[4]
                            })
                    print(f"🏷️ Extracted {len(vendors)} vendor(s): {[v['Vendor Name'] for v in vendors]}")

            # ----------------------------
            # Attachments
            # ----------------------------
            attachments = [a["href"] for a in detail_soup.select("a[target='_blank'][href*='SRM.MAGIC.MS.GOV']")]

            # Download and zip attachments
            temp_dir = os.path.join(OUTPUT_DIR, f"{smart_number}_files")
            os.makedirs(temp_dir, exist_ok=True)
            downloaded_files = []
            for link in attachments:
                file_path = download_file(link, temp_dir)
                if file_path:
                    extract_zip_if_needed(file_path, temp_dir)
                    downloaded_files.append(file_path)

            zip_path = os.path.join(OUTPUT_DIR, f"{smart_number}_attachments.zip")
            with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as zipf:
                for root, _, files in os.walk(temp_dir):
                    for file in files:
                        zipf.write(os.path.join(root, file), arcname=file)
            for f in os.listdir(temp_dir):
                os.remove(os.path.join(temp_dir, f))
            os.rmdir(temp_dir)

            record = {
                **procurement_data,
                "Contact Info": contact_info,
                "Awarded Vendors": vendors,
                "Attachments": attachments,
                "Downloaded Files": downloaded_files,
                "Zip File": zip_path,
            }

            all_data.append(record)

            if click_element_safe(By.CSS_SELECTOR, "a.searchReturn"):
                print("⬅️ Returned to results")
            else:
                driver.back()
            time.sleep(2)

        # --- Next page
        next_button = soup.find("a", string="Next")
        if next_button and next_button.get("href"):
            driver.get(urljoin(BASE_URL, next_button["href"]))
            page_num += 1
            time.sleep(5)
        else:
            print("✅ No more pages left.")
            break

    # --- Save Excel + JSON
    pd.DataFrame(all_data).to_excel(EXCEL_PATH, index=False)
    with open(JSON_PATH, "w", encoding="utf-8") as f:
        json.dump(all_data, f, ensure_ascii=False, indent=4)

    print(f"\n✅ All data saved to:\n📘 {EXCEL_PATH}\n📄 {JSON_PATH}")

# =========================
# RUN
# =========================
if __name__ == "__main__":
    scrape_awarded_data()
    driver.quit()
