import os
import time
import urllib.parse
import requests
import pandas as pd
from datetime import datetime
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager

BASE_URL = "https://www.ms.gov/dfa/contract_bid_search/Bid?autoloadGrid=False"
OUTPUT_ROOT = os.path.join(os.getcwd(), "mississippi_search_data")

def download_attachment(link_element, save_dir):
    try:
        file_url = link_element.get_attribute("href")
        if not file_url:
            return None

        parsed = urllib.parse.urlparse(file_url)
        filename = os.path.basename(parsed.path)

        if not filename or filename.lower().startswith("docserver"):
            filename = link_element.text.strip().replace(" ", "_")

        if not filename:
            filename = f"attachment_{int(time.time())}.bin"

        lower_url = file_url.lower()
        if "pdf" in lower_url and not filename.endswith(".pdf"):
            filename += ".pdf"
        elif "xlsx" in lower_url and not filename.endswith(".xlsx"):
            filename += ".xlsx"
        elif "xls" in lower_url and not filename.endswith(".xls"):
            filename += ".xls"
        elif "doc" in lower_url and not filename.endswith(".docx"):
            filename += ".docx"

        file_path = os.path.join(save_dir, filename)

        response = requests.get(file_url, verify=False, timeout=30)
        if response.status_code == 200:
            with open(file_path, "wb") as f:
                f.write(response.content)
            print(f"✅ Downloaded: {filename}")
        else:
            print(f"⚠️ Failed to download {filename} (status {response.status_code})")

        return file_path
    except Exception as e:
        print(f"⚠️ Error downloading attachment: {e}")
        return None


def scrape_mississippi():
    print("🌐 Opening Mississippi Contract Search...")

    chrome_options = Options()
    chrome_options.add_argument("--headless=new")
    chrome_options.add_argument("--no-sandbox")
    chrome_options.add_argument("--disable-gpu")
    chrome_options.add_argument("--disable-dev-shm-usage")

    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=chrome_options)
    wait = WebDriverWait(driver, 30)

    driver.get(BASE_URL)

    # Click “Advanced Search Options”
    wait.until(EC.element_to_be_clickable((By.ID, "advanceSearchToggle"))).click()
    print("✅ Clicked 'Advanced Search Options'")

    # Select “Awarded” from dropdown
    status_dropdown = wait.until(EC.presence_of_element_located((By.ID, "Status")))
    Select(status_dropdown).select_by_visible_text("Awarded")
    print("✅ Selected 'Awarded' status")

    # Click search
    search_btn = wait.until(EC.element_to_be_clickable((By.ID, "btnSubmit")))
    search_btn.click()
    print("🔍 Search initiated...")

    time.sleep(8)
    rows = driver.find_elements(By.CSS_SELECTOR, "#bidTable tbody tr")
    print(f"📄 Found {len(rows)} opportunities.")

    timestamp = datetime.now().strftime("run_%Y-%m-%d_%H-%M-%S")
    run_dir = os.path.join(OUTPUT_ROOT, timestamp)
    os.makedirs(run_dir, exist_ok=True)

    excel_records = []

    for idx, row in enumerate(rows, start=1):
        try:
            cols = row.find_elements(By.TAG_NAME, "td")
            if len(cols) < 8:
                continue

            agency = cols[0].text.strip()
            smart_number = cols[1].text.strip()
            rfx_number = cols[2].text.strip()
            description = cols[3].text.strip()
            status = cols[4].text.strip()
            advertised_date = cols[5].text.strip()
            submission_date = cols[6].text.strip()

            print(f"\n📦 Processing {smart_number}...")

            attachment_links = cols[3].find_elements(By.TAG_NAME, "a")
            attachment_folder = os.path.join(run_dir, f"{smart_number}")
            os.makedirs(attachment_folder, exist_ok=True)

            attachments = []
            for link in attachment_links:
                file_path = download_attachment(link, attachment_folder)
                if file_path:
                    attachments.append(os.path.basename(file_path))

            excel_records.append({
                "Agency": agency,
                "Smart Number": smart_number,
                "RFx Number": rfx_number,
                "Description": description,
                "Status": status,
                "Advertised Date": advertised_date,
                "Submission Date": submission_date,
                "Attachments": ", ".join(attachments) if attachments else "No Attachments",
                "Folder Path": attachment_folder
            })

            print(f"✅ Completed {smart_number} ({len(attachments)} attachments)")

        except Exception as e:
            print(f"⚠️ Error scraping row: {e}")

    df = pd.DataFrame(excel_records)
    excel_path = os.path.join(run_dir, "mississippi_search_data.xlsx")
    df.to_excel(excel_path, index=False)

    print(f"\n✅ Data saved to Excel: {excel_path}")
    print(f"📂 Attachments saved in folders under {run_dir}")

    driver.quit()
    print("🏁 Mississippi DFA Scraping completed successfully.")


if __name__ == "__main__":
    scrape_mississippi()
