# alaska_vss_step2.py
import time
import logging

from selenium import webdriver
from selenium.webdriver.firefox.service import Service
from selenium.webdriver.firefox.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import (
    TimeoutException,
    ElementClickInterceptedException,
    StaleElementReferenceException,
    NoSuchElementException,
)

START_URL = "https://iris-vss.alaska.gov/"

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s"
)

def build_driver() -> webdriver.Firefox:
    # If geckodriver isn't in PATH, set it explicitly:
    # service = Service(executable_path="/opt/homebrew/bin/geckodriver")
    service = Service()
    opts = Options()
    driver = webdriver.Firefox(service=service, options=opts)
    driver.maximize_window()            # Expand the browser
    driver.set_page_load_timeout(60)
    return driver

def wait_for_page_ready(driver, timeout=30):
    WebDriverWait(driver, timeout).until(
        lambda d: d.execute_script("return document.readyState") == "complete"
    )

def click_view_published_solicitations(driver, timeout=30):
    wait = WebDriverWait(driver, timeout)
    locator = (
        By.CSS_SELECTOR,
        'div[title="View Published Solicitations"][aria-label="View Published Solicitations"]'
    )
    el = wait.until(EC.presence_of_element_located(locator))
    driver.execute_script("arguments[0].scrollIntoView({block:'center'});", el)
    wait.until(EC.element_to_be_clickable(locator))
    try:
        el.click()
        logging.info("Clicked 'View Published Solicitations'.")
    except (ElementClickInterceptedException, StaleElementReferenceException):
        el = driver.find_element(*locator)
        driver.execute_script("arguments[0].click();", el)
        logging.info("Clicked via JS fallback.")

def expand_show_more(driver, timeout=20):
    """
    Click the 'Show More' filters button:
    <button class="css-h3jhcj css-rxeo18 css-f9i1ct" aria-label="Please Enter or Space to expand Show More">
    """
    wait = WebDriverWait(driver, timeout)
    btn_locator = (
        By.CSS_SELECTOR,
        'button[aria-label="Please Enter or Space to expand Show More"]'
    )
    try:
        btn = wait.until(EC.presence_of_element_located(btn_locator))
        driver.execute_script("arguments[0].scrollIntoView({block:'center'});", btn)
        wait.until(EC.element_to_be_clickable(btn_locator))
        try:
            btn.click()
        except (ElementClickInterceptedException, StaleElementReferenceException):
            btn = driver.find_element(*btn_locator)
            driver.execute_script("arguments[0].click();", btn)
        logging.info("Expanded 'Show More' filters.")
        time.sleep(0.5)  # allow controls to render
    except TimeoutException:
        logging.info("Show More button not found (maybe already expanded). Continuing.")

def set_status_awarded(driver, timeout=20):
    """
    Status select:
    <select class="css-n9dqus " name="vss.page.VVSSX10019.gridView1.group1.cardSearch.search1.SO_STA" aria-label="Status">
    Select option value="A" (Awarded)
    """
    wait = WebDriverWait(driver, timeout)
    sel_locator = (
        By.CSS_SELECTOR,
        'select[name="vss.page.VVSSX10019.gridView1.group1.cardSearch.search1.SO_STA"][aria-label="Status"]'
    )
    sel = wait.until(EC.presence_of_element_located(sel_locator))
    driver.execute_script("arguments[0].scrollIntoView({block:'center'});", sel)
    Select(sel).select_by_value("A")
    logging.info("Set Status = Awarded (value A).")
    time.sleep(0.2)

def set_show_me_all(driver, timeout=20):
    """
    Show Me select:
    <select class="css-n9dqus " name="vss.page.VVSSX10019.gridView1.group1.cardSearch.search1.SHOW_TXT" aria-label="Show Me">
    Select option value="1" (All)
    """
    wait = WebDriverWait(driver, timeout)
    sel_locator = (
        By.CSS_SELECTOR,
        'select[name="vss.page.VVSSX10019.gridView1.group1.cardSearch.search1.SHOW_TXT"][aria-label="Show Me"]'
    )
    sel = wait.until(EC.presence_of_element_located(sel_locator))
    # FIXED quoting here ↓↓↓
    driver.execute_script("arguments[0].scrollIntoView({block:'center'});", sel)
    Select(sel).select_by_value("1")
    logging.info("Set Show Me = All (value 1).")
    time.sleep(0.2)

def click_search(driver, timeout=20):
    """
    Search button:
    <button class="css-xzr720" name="vss.page.VVSSX10019.gridView1.Search" aria-label="Search">
    """
    wait = WebDriverWait(driver, timeout)
    btn_locator = (
        By.CSS_SELECTOR,
        'button[name="vss.page.VVSSX10019.gridView1.Search"][aria-label="Search"]'
    )
    btn = wait.until(EC.presence_of_element_located(btn_locator))
    driver.execute_script("arguments[0].scrollIntoView({block:'center'});", btn)
    wait.until(EC.element_to_be_clickable(btn_locator))
    try:
        btn.click()
    except (ElementClickInterceptedException, StaleElementReferenceException):
        btn = driver.find_element(*btn_locator)
        driver.execute_script("arguments[0].click();", btn)
    logging.info("Clicked Search.")

def main():
    driver = build_driver()
    try:
        logging.info("Opening IRIS VSS…")
        driver.get(START_URL)
        wait_for_page_ready(driver)
        time.sleep(1)  # brief visual pause

        logging.info("Navigating to 'View Published Solicitations'…")
        click_view_published_solicitations(driver)
        wait_for_page_ready(driver)
        time.sleep(1)

        logging.info("Expanding filters and applying selections…")
        expand_show_more(driver)
        set_status_awarded(driver)
        set_show_me_all(driver)

        logging.info("Submitting the search…")
        click_search(driver)

        logging.info("Waiting a few seconds so you can view the results…")
        time.sleep(5)

        print("\n✅ Step 2 complete. Press ENTER to close the browser.")
        input()
    except TimeoutException as e:
        logging.error("Timed out waiting for an element: %s", e)
        print("\n❌ Timeout—one of the controls didn’t appear as expected. We can adjust locators if needed.")
        input("Press ENTER to close the browser.")
    finally:
        driver.quit()

if __name__ == "__main__":
    main()