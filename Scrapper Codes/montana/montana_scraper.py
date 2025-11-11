from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
from bs4 import BeautifulSoup
import pandas as pd
import time
import requests
import zipfile
from pathlib import Path

CHROME_BINARY_PATH = "/usr/bin/google-chrome"
DOWNLOAD_DIR = Path("downloads_test1")
DOWNLOAD_DIR.mkdir(exist_ok=True)


chrome_options = Options()
chrome_options.add_argument("--headless")
chrome_options.add_argument("--no-sandbox")
chrome_options.add_argument("--disable-dev-shm-usage")
chrome_options.binary_location = CHROME_BINARY_PATH

service = Service(ChromeDriverManager().install())
driver = webdriver.Chrome(service=service, options=chrome_options)
print(" ChromeDriver initialized successfully.\n")


url = "https://bids.sciquest.com/apps/Router/PublicEvent?CustomerOrg=StateOfMontana"
driver.get(url)
time.sleep(3)

try:
    awarded_tab = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.ID, "PhoenixNavLink_PHX_NAV_SourcingAward"))
    )
    awarded_tab.click()
    print(" Clicked the 'Awarded' tab.")
    time.sleep(4)
except Exception as e:
    print(f" Error clicking 'Awarded' tab: {e}")
    driver.quit()
    exit()


page = 1
data = []
successful_downloads = 0

while True:
    print(f"\n==============================")
    print(f" Processing Page {page}")
    print("==============================")

    try:
        WebDriverWait(driver, 15).until(EC.presence_of_element_located((By.TAG_NAME, "table")))
        soup = BeautifulSoup(driver.page_source, "html.parser")
        table = soup.find("table")
        if not table:
            print(" No table found on this page. Stopping.")
            break

        rows = table.find_all("tr")[1:]
        print(f"Found {len(rows)} opportunities on page {page}.")

        for idx, row in enumerate(rows, 1):
            tds = row.find_all("td")
            if len(tds) < 2:
                continue

            status = tds[0].get_text(strip=True)
            details_td = tds[1]
            details_text = details_td.get_text(separator="\n").strip()
            lines = [l.strip() for l in details_text.split("\n") if l.strip()]

            
            pdf_link = ''
            pdf_anchor = row.find('a', string=lambda text: text and 'View as PDF' in text)
            if not pdf_anchor:
                pdf_anchor = row.find('a', {'id': lambda x: x and 'BUTTON_PDF_VIEW' in x})
            if not pdf_anchor:
                pdf_anchor = row.find('a', href=lambda x: x and '.pdf' in x.lower())
            
            if pdf_anchor and pdf_anchor.has_attr('href'):
                pdf_link = pdf_anchor['href']
                if pdf_link.startswith("/"):
                    pdf_link = "https://bids.sciquest.com" + pdf_link
                print(f"  ✓ PDF link found")
            else:
                print(f"  ✗ No PDF link found")

        
            i = 0
            name_parts = []
            while i < len(lines) and lines[i] not in ['Open', 'Close', 'Type', 'Number', 'Contact']:
                name_parts.append(lines[i])
                i += 1

            name = ' '.join(name_parts).strip()
            print(f"  Name: {name[:50]}..." if len(name) > 50 else f"  Name: {name}")

            entry = {
                'Name': name,
                'Open': '',
                'Close': '',
                'Type': '',
                'Number': '',
                'Contact': '',
                'Status': status,
                'Details': '',
                'PDF_Link': pdf_link,
                'Zip_File': ''
            }

            
            extra_details = []
            while i < len(lines):
                key = lines[i]
                i += 1
                value = ''
                if i < len(lines) and lines[i] not in ['Open', 'Close', 'Type', 'Number', 'Contact']:
                    value = lines[i]
                    i += 1

                if key in entry:
                    entry[key] = value
                else:
                    extra_details.append(f"{key}: {value}")

            entry['Details'] = '\n'.join(extra_details)

            
            if pdf_link:
                try:
                    safe_name = "".join(c for c in name if c.isalnum() or c in (" ", "-", "_")).strip()
                    if not safe_name:
                        safe_name = f"page{page}_row{idx}"
                    if len(safe_name) > 80:
                        safe_name = safe_name[:80]

                    pdf_filename = f"{safe_name}.pdf"
                    pdf_path = DOWNLOAD_DIR / pdf_filename

                    print("  Downloading PDF...", end=" ")
                    headers = {"User-Agent": "Mozilla/5.0 (X11; Linux x86_64) Chrome/120.0"}
                    response = requests.get(pdf_link, headers=headers, timeout=30)
                    response.raise_for_status()
                    with open(pdf_path, "wb") as f:
                        f.write(response.content)
                    print("✓")

                    zip_filename = f"{safe_name}.zip"
                    zip_path = DOWNLOAD_DIR / zip_filename
                    with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as zipf:
                        zipf.write(pdf_path, pdf_filename)
                    pdf_path.unlink()
                    print(f"   Created zip: {zip_filename}")

                    entry["Zip_File"] = str(zip_path)
                    successful_downloads += 1
                except Exception as e:
                    entry["Zip_File"] = f"Failed: {str(e)}"
                    print(f"   Download failed: {e}")
            else:
                entry["Zip_File"] = "No PDF link found"

            data.append(entry)
            print()

        if page < MAX_PAGES:
            try:
                next_btn = WebDriverWait(driver, 5).until(
                    EC.element_to_be_clickable((
                        By.XPATH,
                        "//button[contains(@title,'Next') or contains(@aria-label,'Next') or i[contains(@title,'Next')]]"
                    ))
                )
                driver.execute_script("arguments[0].click();", next_btn)
                print(" Moving to next page...")
                page += 1
                time.sleep(5)
            except Exception as e:
                print(f" No Next button found on page {page}: {e}")
                break
        else:
            print(" Reached 2-page test limit.")
            break

    except Exception as e:
        print(f" Error on page {page}: {e}")
        break

driver.quit()


if data:
    df = pd.DataFrame(data)
    columns_order = ['Name', 'Open', 'Close', 'Type', 'Number', 'Contact', 'Status', 'Details', 'PDF_Link', 'Zip_File']
    df = df[columns_order]
    df.to_excel("awarded_bids_test2.xlsx", index=False)

    print("\n==============================")
    print(" TEST SUMMARY")
    print("==============================")
    print(f"Pages processed: {page}")
    print(f"Total records: {len(df)}")
    print(f"PDFs successfully downloaded & zipped: {successful_downloads}")
    print(f"Saved to: awarded_bids_test2.xlsx")
    print(f"Zips stored in: {DOWNLOAD_DIR.resolve()}")
else:
    print("\n No data found.")
