# pip install selenium webdriver-manager pandas openpyxl

import math
import time
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager

URL = "https://app.az.gov/page.aspx/en/rfp/request_browse_public"
OUTPUT_XLSX = "az_app_awarded_achieved1.xlsx"


STATUS_DROPDOWN = "div.ui.dropdown.selection[data-selector='body_x_selStatusCode_1']"
STATUS_MENU_VISIBLE = "//ul[@id='body_x_selStatusCode_1_MenuItem' and contains(@class,'visible')]"
STATUS_VALUE_INPUT = (By.ID, "body_x_selStatusCode_1")
STATUS_ITEM_ACHIEVED_ID = "body_x_selStatusCode_1_end"  

AWARDED_DROPDOWN = "div.ui.dropdown.selection[data-selector='body_x__ctl0']"
AWARDED_MENU_VISIBLE = "//ul[@id='body_x_txtRfpAwarded_1_MenuItem' and contains(@class,'visible')]"
AWARDED_VALUE_INPUT = (By.ID, "body_x_txtRfpAwarded_1")
AWARDED_ITEM_YES_ID = "body_x__ctl0_True"  

SEARCH_BTN = (By.ID, "body_x_prxFilterBar_x_cmdSearchBtn")
GRID_ROWS = (By.CSS_SELECTOR, "tr[id^='body_x_grid_grd_tr_']")


PAGER_CONTAINER = (By.XPATH, "//ul[contains(@class,'pager') and contains(@class,'buttons')]")
PAGER_COUNT = (By.XPATH, "//span[@data-role='pager-count']//span")
PAGER_NUMERIC_BTNS = (By.XPATH, "//ul[contains(@class,'pager') and contains(@class,'buttons')]//button[contains(@id,'PagerBtn') and contains(@id,'Page') and @data-page-index]")
PAGER_NUMERIC_BTN_ID_FMT = "body_x_grid_PagerBtn{}Page"
PAGER_NUMERIC_BTN_BY_INDEX_X = "//ul[contains(@class,'pager') and contains(@class,'buttons')]//button[@data-page-index='{}']"

def wait_ready(driver, timeout=180):
    wait = WebDriverWait(driver, timeout)
    wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "div.field[data-iv-role='field']")))
    wait.until(lambda d: d.execute_script("return document.readyState") == "complete")

def open_menu_and_pick(driver, dropdown_css, menu_visible_xpath, item_id, value_input, expected_value, timeout=60):
    wait = WebDriverWait(driver, timeout)
    dd = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, dropdown_css)))
    driver.execute_script("arguments[0].scrollIntoView({block:'center'});", dd)
    try:
        dd.click()
    except Exception:
        driver.execute_script("arguments[0].click();", dd)
    wait.until(EC.visibility_of_element_located((By.XPATH, menu_visible_xpath)))
    item = wait.until(EC.element_to_be_clickable((By.ID, item_id)))
    driver.execute_script("arguments[0].scrollIntoView({block:'center'});", item)
    try:
        item.click()
    except Exception:
        driver.execute_script("arguments[0].click();", item)

    def value_selected(d):
        el = d.find_element(*value_input)
        return el.get_attribute("value") == expected_value

    wait.until(value_selected)
    wait.until(EC.invisibility_of_element_located((By.XPATH, menu_visible_xpath)))

def click_search(driver, timeout=30):
    wait = WebDriverWait(driver, timeout)
    btn = wait.until(EC.element_to_be_clickable(SEARCH_BTN))
    driver.execute_script("arguments[0].scrollIntoView({block:'center'});", btn)
    try:
        btn.click()
    except Exception:
        driver.execute_script("arguments[0].click();", btn)

def wait_for_results(driver, timeout=90):
    WebDriverWait(driver, timeout).until(EC.presence_of_element_located(GRID_ROWS))

def parse_current_page_rows(driver):
    rows = driver.find_elements(*GRID_ROWS)
    records = []
    for r in rows:
        tds = r.find_elements(By.CSS_SELECTOR, "td")
        visible_cells = [td for td in tds if 'hidden' not in (td.get_attribute('class') or '')]
        if len(visible_cells) < 9:
            continue
        try:
            code = visible_cells[1].text.strip()
            label = visible_cells[2].text.strip()
            commodity = visible_cells[3].text.strip()
            agency = visible_cells[4].text.strip()
            status = visible_cells[5].text.strip()

            awarded_cell = visible_cells[6]
            awarded = "Yes"
            try:
                chk = awarded_cell.find_element(By.CSS_SELECTOR, "input[type='checkbox']")
                awarded = "Yes" if chk.get_attribute("checked") else "No"
            except Exception:
                awarded = (awarded_cell.text.strip() or "Yes")

            begin = visible_cells[7].text.strip()
            end = visible_cells[8].text.strip()

            records.append({
                "Code": code,
                "Label": label,
                "Commodity": commodity,
                "Agency": agency,
                "Status": status,
                "RFx Awarded": awarded,
                "Begin (UTC-7)": begin,
                "End (UTC-7)": end,
            })
        except Exception:
            continue
    return records

def first_row_el(driver):
    try:
        return driver.find_element(*GRID_ROWS)
    except Exception:
        return None

def first_row_signature(driver):
    try:
        row = driver.find_element(*GRID_ROWS)
        return (row.get_attribute("id") or "") + "|" + (row.text or "").strip()
    except Exception:
        return ""

def get_total_records(driver):
    try:
        txt = WebDriverWait(driver, 10).until(EC.presence_of_element_located(PAGER_COUNT)).text
        digits = "".join(ch for ch in txt if ch.isdigit())
        return int(digits) if digits else None
    except Exception:
        return None

def get_max_page_index(driver):
    """
    Reads all visible numeric pager buttons to determine the highest page index (0-based).
    """
    max_idx = 0
    try:
        btns = driver.find_elements(*PAGER_NUMERIC_BTNS)
        for b in btns:
            try:
                idx = int(b.get_attribute("data-page-index"))
                if idx > max_idx:
                    max_idx = idx
            except Exception:
                continue
    except Exception:
        pass
    return max_idx

def click_page_index(driver, i, timeout=60):
    """
    Click page i (0-based) by id first, then by data-page-index, waiting for grid refresh.
    """
    wait = WebDriverWait(driver, timeout)
    old_first = first_row_el(driver)
    sig_before = first_row_signature(driver)

   
    try:
        btn = driver.find_element(By.ID, PAGER_NUMERIC_BTN_ID_FMT.format(i))
        cls = btn.get_attribute("class") or ""
        if "active" in cls and "disabled" in cls:
            return
    except Exception:
        pass

    
    try:
        btn = driver.find_element(By.ID, PAGER_NUMERIC_BTN_ID_FMT.format(i))
        driver.execute_script("arguments[0].scrollIntoView({block:'center'});", btn)
        try:
            btn.click()
        except Exception:
            driver.execute_script("arguments[0].click();", btn)
    except Exception:
        
        try:
            xp = PAGER_NUMERIC_BTN_BY_INDEX_X.format(i)
            btn = driver.find_element(By.XPATH, xp)
            driver.execute_script("arguments[0].scrollIntoView({block:'center'});", btn)
            try:
                btn.click()
            except Exception:
                driver.execute_script("arguments[0].click();", btn)
        except Exception:
            raise RuntimeError(f"Pager button for index {i} not found")

    
    try:
        if old_first:
            WebDriverWait(driver, timeout).until(EC.staleness_of(old_first))
    except Exception:
        pass
    WebDriverWait(driver, timeout).until(lambda d: first_row_signature(d) != sig_before)

def paginate_all_by_index(driver):
    """
    Uses the pager’s numeric buttons to iterate every page 0..max_index and collect all rows.
    Also cross-checks with the total-record count when available.
    """
    all_records = []

    
    page_records = parse_current_page_rows(driver)
    all_records.extend(page_records)

    
    total = get_total_records(driver)  
    per_page = max(1, len(page_records))
    
    max_idx = get_max_page_index(driver)
    est_pages_from_total = math.ceil(total / per_page) if total else (max_idx + 1)
    total_pages = max(max_idx + 1, est_pages_from_total)

    
    for i in range(1, total_pages):
        click_page_index(driver, i)
        WebDriverWait(driver, 30).until(EC.presence_of_element_located(GRID_ROWS))
        all_records.extend(parse_current_page_rows(driver))

    return all_records

def main():
    chrome_options = Options()
    chrome_options.add_argument("--start-maximized")
    chrome_options.add_argument("--disable-blink-features=AutomationControlled")

    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=chrome_options)
    try:
        driver.get(URL)

        
        wait_ready(driver, timeout=180)

        
        open_menu_and_pick(driver, STATUS_DROPDOWN, STATUS_MENU_VISIBLE, STATUS_ITEM_ACHIEVED_ID, STATUS_VALUE_INPUT, "end", timeout=60)
        open_menu_and_pick(driver, AWARDED_DROPDOWN, AWARDED_MENU_VISIBLE, AWARDED_ITEM_YES_ID, AWARDED_VALUE_INPUT, "True", timeout=60)

        click_search(driver)
        wait_for_results(driver, timeout=90)

        
        all_records = paginate_all_by_index(driver)

        if all_records:
            df = pd.DataFrame(all_records).drop_duplicates()
            df.to_excel(OUTPUT_XLSX, index=False)
            print(f"Wrote {len(df)} records to {OUTPUT_XLSX}")
        else:
            print("No records found with the selected filters.")
    finally:
        driver.quit()

if __name__ == "__main__":
    main()
