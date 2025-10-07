import os
import time
import shutil
import logging
from pathlib import Path

from selenium import webdriver
from selenium.webdriver.firefox.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import (
    TimeoutException,
    ElementClickInterceptedException,
    StaleElementReferenceException,
)

URL = "https://hands.ehawaii.gov/hands/opportunities"
DOWNLOAD_DIR = "/Users/raajthipparthy/Desktop/Opportunity Scrapers"
FINAL_FILENAME = "hawaii-opportunities.xlsx"

# Common Excel MIME types (some sites mislabel as octet-stream)
EXCEL_MIMES = ",".join([
    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    "application/vnd.ms-excel",
    "application/octet-stream",
])

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s"
)

def setup_firefox(download_dir: str) -> webdriver.Firefox:
    """Create a Firefox driver configured to auto-save Excel to download_dir."""
    options = webdriver.FirefoxOptions()
    # If you want to see the browser, do not set headless.
    # options.add_argument("-headless")

    options.set_preference("browser.download.folderList", 2)  # 2 = custom dir
    options.set_preference("browser.download.dir", download_dir)
    options.set_preference("browser.download.useDownloadDir", True)
    options.set_preference("browser.helperApps.neverAsk.saveToDisk", EXCEL_MIMES)
    options.set_preference("pdfjs.disabled", True)
    options.set_preference("browser.download.manager.showWhenStarting", False)
    options.set_preference("browser.download.alwaysOpenPanel", False)
    options.set_preference("browser.download.forbid_open_with", True)

    driver = webdriver.Firefox(options=options)
    driver.maximize_window()  # expand to full screen
    return driver

def wait_busy_gone(driver, timeout=30):
    """
    HANDS shows <default-busy> overlay while loading.
    Wait until it's not visible (or gone).
    """
    wait = WebDriverWait(driver, timeout)
    try:
        wait.until(EC.invisibility_of_element_located((By.CSS_SELECTOR, "default-busy")))
    except TimeoutException:
        # final small grace period to ensure no visible busy
        end = time.time() + 3
        while time.time() < end:
            els = driver.find_elements(By.CSS_SELECTOR, "default-busy")
            if not els or not any(e.is_displayed() for e in els):
                return
            time.sleep(0.2)
        raise

def safe_click(driver, locator, timeout=30):
    """
    Wait for overlay to be gone, ensure element is clickable,
    scroll into view, try normal click; if intercepted, JS-fallback.
    """
    wait_busy_gone(driver, timeout=timeout)
    wait = WebDriverWait(driver, timeout)
    el = wait.until(EC.element_to_be_clickable(locator))

    # Scroll into center to avoid sticky headers intercepting the click
    driver.execute_script("arguments[0].scrollIntoView({block:'center'});", el)
    time.sleep(0.1)

    try:
        el.click()
        return
    except (ElementClickInterceptedException, StaleElementReferenceException):
        try:
            tag = el.tag_name.lower()
            el_type = (el.get_attribute("type") or "").lower()
            if tag == "input" and el_type == "checkbox":
                driver.execute_script("""
                    const el = arguments[0];
                    // Toggle on (we only need it checked)
                    if (!el.checked) {
                        el.checked = true;
                        el.dispatchEvent(new Event('input', {bubbles:true}));
                        el.dispatchEvent(new Event('change', {bubbles:true}));
                    }
                """, el)
            else:
                driver.execute_script("arguments[0].click();", el)
            return
        except Exception:
            time.sleep(0.25)
            driver.execute_script("arguments[0].scrollIntoView({block:'center'});", el)
            el.click()

def wait_for_download_complete(folder: Path, before_snapshot: set, timeout: int = 180) -> Path:
    """
    Wait for a new .xls/.xlsx file to appear in folder and finish downloading.
    Considers .part files and checks for stable file size before returning.
    """
    end = time.time() + timeout
    last_candidate = None

    while time.time() < end:
        current = set(folder.glob("*.xls*"))  # .xls or .xlsx
        new_files = [p for p in current if p not in before_snapshot]

        # Ignore while any .part exists (active download)
        part_files = list(folder.glob("*.part"))

        if new_files:
            # Choose most recently modified among new files
            candidate = max(new_files, key=lambda p: p.stat().st_mtime)
            # Wait until no .part files remain and size is stable
            if not part_files:
                size1 = candidate.stat().st_size
                time.sleep(0.8)
                size2 = candidate.stat().st_size
                if size1 == size2 and size2 > 0:
                    return candidate
            last_candidate = candidate

        time.sleep(0.5)

    # If we had a candidate but timed out, surface it in the error for debugging
    raise TimeoutException(
        f"Timed out waiting for Excel download to complete. "
        f"Last seen candidate: {last_candidate}"
    )

def main():
    target_dir = Path(DOWNLOAD_DIR).expanduser().resolve()
    target_dir.mkdir(parents=True, exist_ok=True)
    final_path = target_dir / FINAL_FILENAME

    # Clean any pre-existing output file
    if final_path.exists():
        logging.info(f"Removing existing file: {final_path}")
        final_path.unlink()

    driver = setup_firefox(str(target_dir))
    wait = WebDriverWait(driver, 30)

    try:
        logging.info("Opening Hawaii HANDS opportunities page…")
        driver.get(URL)

        logging.info("Waiting for initial busy overlay to finish…")
        wait_busy_gone(driver, timeout=60)

        # Checkbox you specified
        checkbox_locator = (By.CSS_SELECTOR, "input.largerCheckbox[name='showclose'][type='checkbox']")
        checkbox = wait.until(EC.presence_of_element_located(checkbox_locator))

        if not checkbox.is_selected():
            logging.info("Clicking 'Show Closed' checkbox to load the full list…")
            safe_click(driver, checkbox_locator, timeout=60)
        else:
            logging.info("'Show Closed' checkbox already selected.")

        # After clicking, the list refreshes under a busy overlay
        logging.info("Waiting for list refresh to complete…")
        wait_busy_gone(driver, timeout=120)

        # Download button you specified
        download_btn_locator = (By.CSS_SELECTOR, "button.btn.btn-sm.btn-primary.ng-star-inserted[type='submit']")
        logging.info("Waiting for the Excel download button to be ready…")
        WebDriverWait(driver, 60).until(EC.presence_of_element_located(download_btn_locator))
        wait_busy_gone(driver, timeout=60)

        # Snapshot BEFORE clicking, so we can detect the new file
        before = set(target_dir.glob("*.xls*"))

        logging.info("Clicking the Excel download button…")
        safe_click(driver, download_btn_locator, timeout=60)

        logging.info("Waiting for the Excel file to download…")
        downloaded_path = wait_for_download_complete(target_dir, before_snapshot=before, timeout=180)

        logging.info(f"Download complete: {downloaded_path.name}")
        # Move/rename to final filename
        if final_path.exists():
            final_path.unlink()
        shutil.move(str(downloaded_path), str(final_path))
        logging.info(f"Saved as: {final_path}")

        logging.info("All done ✅")

    except TimeoutException as e:
        logging.error(f"Timeout waiting for an element or download: {e}")
        raise
    finally:
        try:
            driver.quit()
        except Exception:
            pass

if __name__ == "__main__":
    main()