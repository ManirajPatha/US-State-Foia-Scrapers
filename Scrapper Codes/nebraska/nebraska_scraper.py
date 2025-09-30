import time
import sys
from urllib.parse import urljoin
from requests.utils import requote_uri

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options as ChromeOptions
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC


try:
    from webdriver_manager.chrome import ChromeDriverManager
    USE_WDM = True
except ImportError:
    USE_WDM = False

import pandas as pd

BASE_URL = "https://das.nebraska.gov"
TARGET_URL = "https://das.nebraska.gov/materiel/bid-opportunities.html#awarded-bids"


EXCLUDE_VENDOR_STATUSES = [
    "NO INTENT TO AWARD",
    "RFQ CLOSED",
    "REJECTION OF ALL BIDS",
    "NO BIDS RECEIVED",
    "NO PROPOSALS RECEIVED",
    "RFI CLOSED",
    "REJECT ALL PROPOSALS",
    "ITB CLOSED",
    "SOLICITATION WITHDRAWN",
    "RFP CLOSED",
    "BID WITHDRAWN",
]

RED_HEX = "#ff0000"  

def build_driver(headless=True):
    options = ChromeOptions()
    if headless:
        options.add_argument("--headless=new")
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-gpu")
    options.add_argument("--window-size=1366,768")
    options.add_argument("--disable-dev-shm-usage")

    if USE_WDM:
        from selenium.webdriver.chrome.service import Service as ChromeService
        driver = webdriver.Chrome(service=ChromeService(ChromeDriverManager().install()), options=options)
    else:
        driver = webdriver.Chrome(options=options)
    return driver

def should_skip_vendor_cell(vendor_cell):
    
    text = vendor_cell.text.strip().lower()
    for phrase in EXCLUDE_VENDOR_STATUSES:
        if phrase.lower() in text:
            return True

    
    spans = vendor_cell.find_elements(By.TAG_NAME, "span")
    for sp in spans:
        style = (sp.get_attribute("style") or "").lower().replace(" ", "")
        if "color:" in style and RED_HEX in style:
            return True

    return False

def find_awarded_table(driver):
    
    xpath = (
        "//table[.//th[contains(translate(normalize-space(.),'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'),'vendor(s)')]"
        " and .//th[contains(translate(normalize-space(.),'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'),'letter of intent')]]"
    )
    return WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, xpath)))

def scrape_awarded_bids(driver, max_rows=100):
    driver.get(TARGET_URL)

   
    WebDriverWait(driver, 20).until(EC.presence_of_all_elements_located((By.TAG_NAME, "table")))

    table = find_awarded_table(driver)

    
    tbodys = table.find_elements(By.TAG_NAME, "tbody")
    row_scope = tbodys[0] if tbodys else table
    rows = row_scope.find_elements(By.TAG_NAME, "tr")

   
    

    out = []
    for tr in rows:
        tds = tr.find_elements(By.TAG_NAME, "td")
        
        if len(tds) < 9:
            continue

        
        vendor_td = tds[2]
        if should_skip_vendor_cell(vendor_td):
            continue

        
        desc_td = tds[0]
        link_els = desc_td.find_elements(By.TAG_NAME, "a")
        if link_els:
            desc_text = link_els[0].text.strip()
            href = link_els[0].get_attribute("href") or ""
            
            if href and not href.lower().startswith("http"):
                href = urljoin(BASE_URL, href)
            desc_url = requote_uri(href) if href else ""
        else:
            desc_text = desc_td.text.strip()
            desc_url = ""

        row = {
            "Description": desc_text,
            "Description URL": desc_url,
            "Letter of Intent Date": tds[1].text.strip(),
            "Vendor(s)": vendor_td.text.strip(),
            "Category - NIGP Code": tds[3].get_attribute("innerText").strip(),
            "Type": tds[4].text.strip(),
            "PCO/Buyer": tds[5].get_attribute("innerText").strip(),
            "Solicitation Number": tds[6].text.strip(),
            "Agency": tds[7].text.strip(),
            "Last Updated": tds[8].text.strip(),
        }
        out.append(row)

        
        if len(out) >= max_rows:
            break

    return out

def main():
    driver = build_driver(headless=True)
    try:
        
        data = scrape_awarded_bids(driver, max_rows=100)
        if not data:
            print("No eligible rows found after filtering.")
            return
        df = pd.DataFrame(data)
        out_file = "ne_awarded_bids.xlsx"
        df.to_excel(out_file, index=False)
        print(f"Wrote {len(df)} rows to {out_file}")
    finally:
        driver.quit()

if __name__ == "__main__":
    main()
