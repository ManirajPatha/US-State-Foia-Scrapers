import time
import json
import argparse
from datetime import datetime
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options

BASE = "https://bidopportunities.iowa.gov"
AWARDED_URL = f"{BASE}/Home/AwardedContracts"

# ---------------------------------------------------------------------
# SETUP
# ---------------------------------------------------------------------
def setup_driver(headless=False):
    opts = Options()
    if headless:
        opts.add_argument("--headless=new")
    opts.add_argument("--no-sandbox")
    opts.add_argument("--disable-dev-shm-usage")
    opts.add_argument("--start-maximized")
    driver = webdriver.Chrome(options=opts)
    driver.maximize_window()
    return driver


def split_city_state_zip(raw):
    """Split 'Des Moines, IA 50309' → city, state, zip."""
    if not raw:
        return None, None, None
    raw = raw.strip()
    if "," in raw:
        city_part, rest = [x.strip() for x in raw.split(",", 1)]
    else:
        city_part, rest = raw, ""
    parts = rest.split()
    state = parts[0] if len(parts) >= 1 else None
    zip_code = parts[-1] if len(parts) >= 2 else None
    return city_part, state, zip_code


def scrape_contract_detail(driver, url):
    """Open new tab → scrape contract detail → close tab."""
    driver.switch_to.new_window('tab')
    driver.get(url)

    WebDriverWait(driver, 15).until(
        EC.presence_of_element_located((By.CSS_SELECTOR, "div.panel"))
    )

    def get_value_by_label(label_for):
        """Find <label for='...'> and return its paired value text (handles mailto links)."""
        try:
            label = driver.find_element(By.XPATH, f"//label[contains(@for,'{label_for}')]")
            parent_row = label.find_element(By.XPATH, "./ancestor::div[contains(@class,'row')]")
            val_div = parent_row.find_element(By.CSS_SELECTOR, "div.col-md-8, div.col-md-9")
            try:
                link = val_div.find_element(By.TAG_NAME, "a")
                return link.text.strip() or None
            except:
                return val_div.text.strip() or None
        except:
            return None

    # Contract Information
    contract_number = get_value_by_label("Number")
    product_service = get_value_by_label("ProductService")
    contact_name = get_value_by_label("ContactName")
    contact_email = get_value_by_label("ContactEmail")
    contact_phone = get_value_by_label("ContactPhoneNumber")

    # Vendor Information
    vendor_name = get_value_by_label("VendorName")
    vendor_addr = get_value_by_label("VendorAddress1") or get_value_by_label("AddressLine1")
    vendor_cityzip = get_value_by_label("VendorCityStateZip")
    vendor_contact_name = get_value_by_label("VendorContactName")
    vendor_contact_email = get_value_by_label("VendorContactEmail")
    vendor_contact_phone = get_value_by_label("VendorContactPhoneNumber")

    vendor_city, vendor_state, vendor_zip = split_city_state_zip(vendor_cityzip)

   
    # -----------------------------
    # Attachments
    # -----------------------------
    attachments = []
    try:
        # Target all attachment rows under the sixth panel (Documents/Attachments)
        attach_rows = driver.find_elements(
            By.XPATH,
            "//div[contains(@class,'panel-body')]/div[contains(@class,'row')]"
        )
        for r in attach_rows:
            try:
                name_el = r.find_element(By.XPATH, ".//div[contains(@class,'col-md-8')]/a")
                name = name_el.text.strip()
                # Try to find the download icon link in the same row
                try:
                    dl_el = r.find_element(By.XPATH, ".//div[contains(@class,'col-md-1')]/a[contains(@class,'glyphicon-download')]")
                    href = dl_el.get_attribute("href")
                except:
                    href = name_el.get_attribute("href")
                if href and href.startswith("/"):
                    href = BASE + href
                if name:
                    attachments.append({"name": name, "url": href})
            except Exception as e:
                continue
    except Exception as e:
        print(f"[WARN] Attachment extraction issue: {e}")


    detail = {
        "contract_number": contract_number,
        "product_service": product_service,
        "contact_name": contact_name,
        "contact_email": contact_email,
        "contact_phone": contact_phone,
        "vendor_name": vendor_name,
        "vendor_address": vendor_addr,
        "vendor_city": vendor_city,
        "vendor_state": vendor_state,
        "vendor_zip": vendor_zip,
        "vendor_contact_name": vendor_contact_name,
        "vendor_contact_email": vendor_contact_email,
        "vendor_contact_phone": vendor_contact_phone,
        "attachments": attachments
    }

    driver.close()
    driver.switch_to.window(driver.window_handles[0])
    return detail


# ---------------------------------------------------------------------
# MAIN SCRAPER
# ---------------------------------------------------------------------
def scrape_awarded_contracts(headless=False, max_pages=None, save_every=10):
    driver = setup_driver(headless=headless)
    driver.get(AWARDED_URL)
    WebDriverWait(driver, 15).until(
        EC.presence_of_all_elements_located((By.CSS_SELECTOR, "#awardedContractsTbl tbody tr"))
    )

    ts = datetime.now().strftime("%Y%m%d_%H%M")
    out_file = f"iowa_awarded_contracts_{ts}.json"
    all_rows = []
    print(f"[INFO] Started scraping → {out_file}")

    while True:
        current_page = driver.find_element(By.CSS_SELECTOR, "a.paginate_button.current").text.strip()
        print(f"\n[PAGE] {current_page}")

        rows = driver.find_elements(By.CSS_SELECTOR, "#awardedContractsTbl tbody tr")
        print(f"  Found {len(rows)} rows")

        for i in range(len(rows)):
            try:
                rows = driver.find_elements(By.CSS_SELECTOR, "#awardedContractsTbl tbody tr")
                link = rows[i].find_element(By.CSS_SELECTOR, "td:nth-of-type(4) a").get_attribute("href")
                print(f"    → Scraping: {link}")
                detail = scrape_contract_detail(driver, link)
                all_rows.append(detail)

                # Incremental save
                if len(all_rows) % save_every == 0:
                    with open(out_file, "w", encoding="utf-8") as f:
                        json.dump(all_rows, f, indent=2, ensure_ascii=False)
                    print(f"[INFO] Saved progress ({len(all_rows)} contracts)")
            except Exception as e:
                print(f"[WARN] Row {i+1} failed: {e}")
                continue

        # Find next numbered page
        page_links = driver.find_elements(By.CSS_SELECTOR, "#awardedContractsTbl_paginate span a.paginate_button")
        current_found = False
        next_page_link = None
        for p in page_links:
            label = p.text.strip()
            classes = p.get_attribute("class") or ""
            if "current" in classes:
                current_found = True
                continue
            if current_found and label.isdigit():
                next_page_link = p
                break

        # No next page
        if not next_page_link:
            print("[INFO] Reached last page.")
            break

        # Go to next page
        driver.execute_script("arguments[0].scrollIntoView(true);", next_page_link)
        time.sleep(0.4)
        driver.execute_script("arguments[0].click();", next_page_link)
        WebDriverWait(driver, 15).until(
            EC.presence_of_all_elements_located((By.CSS_SELECTOR, "#awardedContractsTbl tbody tr"))
        )
        time.sleep(1.5)

        if max_pages and int(current_page) >= max_pages:
            print(f"[INFO] Stopped at page {max_pages}.")
            break

    driver.quit()

    # Final save
    with open(out_file, "w", encoding="utf-8") as f:
        json.dump(all_rows, f, indent=2, ensure_ascii=False)
    print(f"\n[DONE] Saved {len(all_rows)} contracts → {out_file}")

    return all_rows


# ---------------------------------------------------------------------
# ENTRYPOINT
# ---------------------------------------------------------------------
def main():
    parser = argparse.ArgumentParser(
        description="Scrape Iowa Awarded Contracts → JSON (pagination + new tab + incremental save)"
    )
    parser.add_argument("--headless", action="store_true", help="Run Chrome in headless mode")
    parser.add_argument("--max-pages", type=int, default=None, help="Limit number of pages (for testing)")
    parser.add_argument("--save-every", type=int, default=10, help="Save progress every N contracts")
    args = parser.parse_args()

    scrape_awarded_contracts(
        headless=args.headless,
        max_pages=args.max_pages,
        save_every=args.save_every
    )


if __name__ == "__main__":
    main()
