# nebraska_awarded_bids_json_docs_only.py
# Same as before, except we DO NOT include "Description URL" in the output.
# We still open the Description link to capture the master "Project Documents" page URL.

import json
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
    opts = ChromeOptions()
    if headless:
        opts.add_argument("--headless=new")
    opts.add_argument("--no-sandbox")
    opts.add_argument("--disable-gpu")
    opts.add_argument("--disable-dev-shm-usage")
    opts.add_argument("--window-size=1366,768")
    if USE_WDM:
        from selenium.webdriver.chrome.service import Service as ChromeService
        return webdriver.Chrome(service=ChromeService(ChromeDriverManager().install()), options=opts)
    return webdriver.Chrome(options=opts)

def _absolute_url(href: str) -> str:
    if not href:
        return ""
    url = href if href.lower().startswith("http") else urljoin(BASE_URL, href)
    return requote_uri(url)

def should_skip_vendor_cell(vendor_cell) -> bool:
    text = (vendor_cell.text or "").strip().lower()
    for phrase in EXCLUDE_VENDOR_STATUSES:
        if phrase.lower() in text:
            return True
    for sp in vendor_cell.find_elements(By.TAG_NAME, "span"):
        style = (sp.get_attribute("style") or "").lower().replace(" ", "")
        if "color:" in style and RED_HEX in style:
            return True
    return False

def find_awarded_table(driver):
    xp = (
        "//table[.//th[contains(translate(normalize-space(.),"
        "'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'),'vendor(s)')]"
        " and .//th[contains(translate(normalize-space(.),"
        "'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'),'letter of intent')]]"
    )
    return WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, xp)))

def get_project_documents_master_url(driver, detail_url: str) -> str:
    if not detail_url:
        return ""
    original = driver.current_window_handle
    master_url = ""
    try:
        driver.switch_to.new_window("tab")
        driver.get(detail_url)
        WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.TAG_NAME, "body")))
        master_url = driver.current_url or detail_url
    except Exception:
        master_url = detail_url
    finally:
        try:
            if driver.current_window_handle != original:
                driver.close()
                driver.switch_to.window(original)
        except Exception:
            try:
                driver.switch_to.window(original)
            except Exception:
                pass
    return _absolute_url(master_url)

def scrape_awarded_bids(driver):
    driver.get(TARGET_URL)
    WebDriverWait(driver, 20).until(EC.presence_of_all_elements_located((By.TAG_NAME, "table")))
    table = find_awarded_table(driver)
    tbodys = table.find_elements(By.TAG_NAME, "tbody")
    row_scope = tbodys[0] if tbodys else table
    rows = row_scope.find_elements(By.TAG_NAME, "tr")

    results = []
    for tr in rows:
        tds = tr.find_elements(By.TAG_NAME, "td")
        if len(tds) < 9:
            continue

        vendor_td = tds[2]
        if should_skip_vendor_cell(vendor_td):
            continue

        # Description text + (internal) link just for navigation to master docs page
        desc_td = tds[0]
        a_elems = desc_td.find_elements(By.TAG_NAME, "a")
        if a_elems:
            desc_text = a_elems[0].text.strip()
            desc_href = _absolute_url(a_elems[0].get_attribute("href") or "")
        else:
            desc_text = desc_td.text.strip()
            desc_href = ""

        master_docs_url = get_project_documents_master_url(driver, desc_href)

        # NOTE: "Description URL" intentionally removed from output per request
        row = {
            "Description": desc_text,
            "Letter of Intent Date": tds[1].text.strip(),
            "Vendor(s)": vendor_td.text.strip(),
            "Category - NIGP Code": tds[3].get_attribute("innerText").strip(),
            "Type": tds[4].text.strip(),
            "PCO/Buyer": tds[5].get_attribute("innerText").strip(),
            "Solicitation Number": tds[6].text.strip(),
            "Agency": tds[7].text.strip(),
            "Last Updated": tds[8].text.strip(),
            "Nebraska_Project_Documents_URLs": master_docs_url,  # single master URL only
        }
        results.append(row)

    return results

def main():
    driver = build_driver(headless=True)
    try:
        data = scrape_awarded_bids(driver)
        out_path = "ne_awarded_bids.json"
        with open(out_path, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
        print(f"Wrote {len(data)} records to {out_path}")
    finally:
        driver.quit()

if __name__ == "__main__":
    main()
