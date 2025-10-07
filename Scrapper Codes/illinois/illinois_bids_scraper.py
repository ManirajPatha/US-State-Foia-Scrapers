# filename: il_bidbuy_advanced_search.py
import time
import logging
from selenium import webdriver
from selenium.webdriver.firefox.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC

URL = "https://www.bidbuy.illinois.gov/bso/view/search/external/advancedSearchBid.xhtml?openBids=true"

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s"
)

def main():
    # If geckodriver isn't on PATH, set Service(executable_path="/full/path/to/geckodriver")
    service = Service()  # uses PATH
    options = webdriver.FirefoxOptions()
    # IMPORTANT: Visible browser (no headless)
    # options.add_argument("-headless")  # DO NOT enable headless

    driver = webdriver.Firefox(service=service, options=options)
    wait = WebDriverWait(driver, 30)

    try:
        logging.info("Opening page…")
        driver.get(URL)

        # Maximize window (plus set an explicit large size for macOS)
        try:
            driver.maximize_window()
        except Exception:
            pass
        driver.set_window_size(1600, 1200)

        # Small pause to let page paint
        time.sleep(1)

        # Scroll a bit to ensure the Advanced Search legend is in view
        driver.execute_script("window.scrollBy(0, 300);")

        # Click the "Advanced Search" legend to expand the panel
        logging.info("Expanding Advanced Search…")
        adv_legend = wait.until(
            EC.element_to_be_clickable(
                (By.XPATH, "//legend[contains(@class,'ui-fieldset-legend')][contains(., 'Advanced Search')]")
            )
        )
        driver.execute_script("arguments[0].scrollIntoView({block:'center'});", adv_legend)
        adv_legend.click()

        # Wait for the advanced search form fields container to become visible
        adv_fields = wait.until(
            EC.visibility_of_element_located((By.ID, "advSearchFormFields"))
        )

        # Within the expanded area, find the Status select and choose "Bid to PO" (value 2BPO)
        logging.info("Selecting Status = 'Bid to PO'…")
        status_select_el = wait.until(
            EC.element_to_be_clickable((By.ID, "bidSearchForm:status"))
        )
        driver.execute_script("arguments[0].scrollIntoView({block:'center'});", status_select_el)
        Select(status_select_el).select_by_value("2BPO")

        # Scroll slightly to ensure the Search button is clickable
        driver.execute_script("window.scrollBy(0, 300);")

        # Click the Search button
        logging.info("Clicking Search…")
        search_btn = wait.until(
            EC.element_to_be_clickable((By.ID, "bidSearchForm:btnBidSearch"))
        )
        driver.execute_script("arguments[0].scrollIntoView({block:'center'});", search_btn)
        search_btn.click()

        # Optional: wait for results area to load (table, messages, etc.)
        # Not strictly required per your request, but this helps confirm the click worked.
        try:
            wait.until(
                EC.any_of(
                    EC.presence_of_element_located((By.CSS_SELECTOR, "table[id*='bidSearchResults'], table.dataTable")),
                    EC.presence_of_element_located((By.XPATH, "//*[contains(., 'No results found') or contains(., 'Result')]"))
                )
            )
            logging.info("Search triggered; results area detected.")
        except Exception:
            logging.info("Search click executed; waiting a bit for page activity.")
            time.sleep(3)

        logging.info("Done with filter + search. Browser will remain open.")
        # Keep the browser open; press Enter in the terminal to close.
        input("Press ENTER here to close the browser…")

    finally:
        try:
            # Only quit after user input above
            driver.quit()
        except Exception:
            pass

if __name__ == "__main__":
    main()