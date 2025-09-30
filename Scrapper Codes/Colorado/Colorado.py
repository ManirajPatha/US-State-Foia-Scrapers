import time
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC

URL = "https://prd.co.cgiadvantage.com/PRDVSS1X1/Advantage4"
OUTPUT_XLSX = "colorado_vss_recent_awards.xlsx"

def wait_overlay_gone(driver, timeout=20):
    """Wait until overlay disappears"""
    try:
        WebDriverWait(driver, timeout).until(
            EC.invisibility_of_element_located((By.CSS_SELECTOR, "div.css-o3hj44"))
        )
    except:
        pass

def scrape_page(driver):
    """Scrape the current page rows into a list of dicts"""
    rows = driver.find_elements(By.CSS_SELECTOR, "tr[id^='tableDataRow']")
    records = []

    for row in rows:
        def safe(css):
            try:
                return row.find_element(By.CSS_SELECTOR, css).text.strip()
            except:
                return ""

        def safe_aria_label(css):
            try:
                element = row.find_element(By.CSS_SELECTOR, css)
                aria_label = element.get_attribute("aria-label")
                return aria_label if aria_label else element.text.strip()
            except:
                return ""

        closing_date = ""
        try:
            date_span = row.find_element(By.CSS_SELECTOR, "span[aria-label*='PM MDT'], span[aria-label*='AM MDT'], span[aria-label*='PM MST'], span[aria-label*='AM MST']")
            closing_date = date_span.get_attribute("aria-label")
        except:
            try:
                date_span = row.find_element(By.CSS_SELECTOR, "span[id*='readOnlyDateTimePicker']")
                closing_date = date_span.text.strip()
            except:
                try:
                    closing_date = safe("[data-qa*='.closeDateTimeSta.CLSE_DT']")
                except:
                    closing_date = ""

        records.append({
            "Description": safe("[data-qa*='.DOC_DSCR']"),
            "Department": safe("[data-qa*='.DeptBuyr.DEPT_NM']"),
            "Buyer": safe("[data-qa*='.DeptBuyr.BUYR_NM']"),
            "Solicitation Number": safe("[data-qa*='.solNumTypCat.DOC_REF']"),
            "Solicitation Type": safe("[data-qa*='.solNumTypCat.DOC_CD_CONCAT']"),
            "Solicitation Category": safe("[data-qa*='.solNumTypCat.SO_CAT_CD']"),
            "Closing Date/Time": closing_date,
            "Status": safe("[data-qa*='.closeDateTimeSta.SO_STA']"),
        })

    return records

def go_next_page(driver, wait):
    """Click next page if available. Return True if moved, else False"""
    try:
        next_btn = driver.find_element(By.CSS_SELECTOR, "button[aria-label='Next Page']")
        if next_btn.get_attribute("disabled"):
            return False
        driver.execute_script("arguments[0].click();", next_btn)
        wait_overlay_gone(driver)
        wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "tr[id^='tableDataRow']")))
        return True
    except:
        return False

def scrape_colorado_vss():
    driver = webdriver.Chrome()
    wait = WebDriverWait(driver, 30)
    all_records = []

    try:
        driver.get(URL)

        published_btn = wait.until(
            EC.presence_of_element_located(
                (By.CSS_SELECTOR, "div[title='View Published Solicitations']")
            )
        )
        wait_overlay_gone(driver)
        driver.execute_script("arguments[0].click();", published_btn)

        show_me = wait.until(
            EC.presence_of_element_located(
                (By.NAME, "vss.page.VVSSX10019.gridView1.group1.cardSearch.search1.SHOW_TXT")
            )
        )
        Select(show_me).select_by_visible_text("Recent Awards")

        search_btn = wait.until(
            EC.element_to_be_clickable(
                (By.CSS_SELECTOR, "button[name='vss.page.VVSSX10019.gridView1.Search']")
            )
        )
        driver.execute_script("arguments[0].click();", search_btn)

        wait_overlay_gone(driver)
        wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "tr[id^='tableDataRow']")))

        while True:
            time.sleep(2)
            all_records.extend(scrape_page(driver))
            if not go_next_page(driver, wait):
                break

        df = pd.DataFrame(all_records)
        df.to_excel(OUTPUT_XLSX, index=False)
        print(f"Saved {len(df)} rows to {OUTPUT_XLSX}")

    finally:
        driver.quit()

if __name__ == "__main__":
    scrape_colorado_vss()