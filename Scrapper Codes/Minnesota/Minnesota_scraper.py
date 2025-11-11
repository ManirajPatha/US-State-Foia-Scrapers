# Minnesota QuestCDN "Final" awards → JSON with awarded vendor details
# CHANGE: pagination now continues until there are no more pages.
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import NoSuchElementException, TimeoutException
from webdriver_manager.chrome import ChromeDriverManager
import json
import time

def wait_for(driver, by, value, timeout=15):
    return WebDriverWait(driver, timeout).until(EC.presence_of_element_located((by, value)))

def element_exists(driver, by, value, timeout=5):
    try:
        WebDriverWait(driver, timeout).until(EC.presence_of_element_located((by, value)))
        return True
    except TimeoutException:
        return False

def click_next_if_any(driver):
    try:
        next_btn = driver.find_element(By.XPATH, "//a[contains(@class,'page-link') and normalize-space(text())='Next']")
        # detect disabled state if the UI uses <li class="disabled"> around it
        try:
            li = next_btn.find_element(By.XPATH, "./ancestor::li[1]")
            if "disabled" in (li.get_attribute("class") or "").lower():
                return False
        except Exception:
            pass
        driver.execute_script("arguments[0].scrollIntoView(true);", next_btn)
        time.sleep(0.5)
        driver.execute_script("arguments[0].click();", next_btn)
        return True
    except NoSuchElementException:
        return False

def open_link_in_new_tab(driver, link_el):
    driver.execute_script("window.open(arguments[0].href, '_blank');", link_el)
    time.sleep(0.8)
    driver.switch_to.window(driver.window_handles[-1])

def close_current_tab_and_back(driver, main_handle):
    driver.close()
    driver.switch_to.window(main_handle)

def extract_award_date(driver):
    xpaths = [
        "//table//tr[.//*[normalize-space(text())='Award Date:']]/*[last()]",
        "//*[self::td or self::th][normalize-space(text())='Award Date:']/following-sibling::*[1]",
        "//*[normalize-space(text())='Award Date:']/following-sibling::*[1]",
    ]
    for xp in xpaths:
        try:
            el = driver.find_element(By.XPATH, xp)
            txt = el.text.strip()
            if txt:
                return txt
        except NoSuchElementException:
            pass
    return ""

def find_awarded_vendor_row(driver):
    table_candidates = driver.find_elements(
        By.XPATH,
        "//table[.//th[contains(.,'Company')] and (.//th[contains(.,'Awarded')] or .//td[contains(.,'Awarded')])]"
    )
    if not table_candidates:
        return None, None

    table = table_candidates[0]
    headers = [h.text.strip() for h in table.find_elements(By.XPATH, ".//thead//th")]
    if not headers:
        headers = [h.text.strip() for h in table.find_elements(By.XPATH, ".//tr[1]/*")]

    try:
        awarded_idx = next(i for i, h in enumerate(headers) if h.lower().startswith("awarded"))
    except StopIteration:
        awarded_idx = None

    rows = table.find_elements(By.XPATH, ".//tbody/tr") or table.find_elements(By.XPATH, ".//tr[position()>1]")
    awarded_cells = []
    for r in rows:
        cells = r.find_elements(By.XPATH, "./*")
        if not cells:
            continue
        awarded_cell = cells[awarded_idx] if (awarded_idx is not None and awarded_idx < len(cells)) else cells[-1]
        awarded_text = awarded_cell.text.strip()
        has_check_icon = bool(awarded_cell.find_elements(By.XPATH, ".//i[contains(@class,'check')] | .//*[contains(text(),'✓') or contains(text(),'✔')]"))
        if has_check_icon or awarded_text in ("✓", "✔"):
            awarded_cells = [c.text.strip() for c in cells]
            break

    return awarded_cells, headers

def map_vendor_row_to_dict(awarded_cells, headers):
    data = {}
    for i, val in enumerate(awarded_cells):
        key = headers[i] if i < len(headers) else f"col_{i}"
        data[key] = val
    rename = {
        "Company": "company",
        "Contact": "contact",
        "Phone": "phone",
        "E-mail": "email",
        "Amount": "amount",
        "Awarded": "awarded",
        "Comment": "comment",
    }
    normalized = {}
    for k, v in data.items():
        normalized[rename.get(k, k)] = v
    return normalized

options = webdriver.ChromeOptions()
options.add_argument("--start-maximized")
driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)

url = "https://qcpi.questcdn.com/cdn/results/?group=6506969&provider=6506969&projType=all"
driver.get(url)
wait = WebDriverWait(driver, 15)

# Filter: Bid Award Type -> Final
bid_award_input = wait_for(driver, By.NAME, "col9_search")
bid_award_input.clear()
bid_award_input.send_keys("Final")
bid_award_input.send_keys(Keys.RETURN)
time.sleep(2.5)

results = []
main_handle = driver.current_window_handle

def scrape_listing_page():
    wait_for(driver, By.ID, "table_id")
    rows = driver.find_elements(By.XPATH, "//table[@id='table_id']/tbody/tr")
    print(f"[INFO] Rows on page: {len(rows)}")

    for idx, row in enumerate(rows, start=1):
        tds = row.find_elements(By.TAG_NAME, "td")
        if len(tds) < 10:
            continue

        quest_number = tds[0].text.strip()
        bid_name = tds[1].text.strip()
        try:
            bid_link_el = tds[1].find_element(By.TAG_NAME, "a")
        except NoSuchElementException:
            print(f"[SKIP] No detail link for row {idx}.")
            continue

        listing_record = {
            "Quest Number": quest_number,
            "Bid/Request Name": bid_name,
            "Bid Closing Date": tds[2].text.strip(),
            "City": tds[3].text.strip(),
            "County": tds[4].text.strip(),
            "State": tds[5].text.strip(),
            "Owner": tds[6].text.strip(),
            "Solicitor": tds[7].text.strip(),
            "Posting Type": tds[8].text.strip(),
            "Bid Award Type": tds[9].text.strip(),
        }

        open_link_in_new_tab(driver, bid_link_el)
        try:
            wait_for(driver, By.TAG_NAME, "body", timeout=15)
        except TimeoutException:
            print(f"[WARN] Detail page timeout for {bid_name} ({quest_number}); skipping.")
            close_current_tab_and_back(driver, main_handle)
            continue

        award_date = extract_award_date(driver)
        awarded_cells, headers = find_awarded_vendor_row(driver)
        if awarded_cells is None and headers is None:
            print(f"[SKIP] No vendor table for: {bid_name} ({quest_number}).")
            close_current_tab_and_back(driver, main_handle)
            continue
        if not awarded_cells:
            print(f"[SKIP] No awarded row checked for: {bid_name} ({quest_number}).")
            close_current_tab_and_back(driver, main_handle)
            continue

        vendor_info = map_vendor_row_to_dict(awarded_cells, headers)
        record = {**listing_record, "Award Date": award_date, "Awarded Vendor": vendor_info}
        results.append(record)

        close_current_tab_and_back(driver, main_handle)
        time.sleep(0.4)

# --------- ONLY PAGINATION CHANGED BELOW ---------
page_count = 0
while True:  # keep going until there is no usable "Next"
    page_count += 1
    print(f"\n[PAGE] Scraping page {page_count} …")
    scrape_listing_page()
    if not click_next_if_any(driver):
        print("[INFO] No more pages.")
        break
    time.sleep(2.0)
# -------------------------------------------------

try:
    driver.quit()
except Exception:
    pass

out_path = "mn_awards.json"
with open(out_path, "w", encoding="utf-8") as f:
    json.dump(results, f, ensure_ascii=False, indent=2)

print(f"\n[DONE] Opportunities saved: {len(results)}")
print(f"[OUT] {out_path}")
