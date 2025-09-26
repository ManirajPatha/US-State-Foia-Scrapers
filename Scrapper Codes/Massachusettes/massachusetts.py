# massachusetts_closed_export_xls.py
import argparse, os, time, glob, logging
from datetime import datetime
from typing import Optional

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager

ADV_URL = "https://www.commbuys.com/bso/view/search/external/advancedSearchBid.xhtml"
TIMEOUT = 40
LONG_TIMEOUT = 90

logging.basicConfig(level=logging.INFO, format="%(asctime)s [%(levelname)s] %(message)s")


# -------------------- Browser --------------------

def open_browser(download_dir: str, headless: bool = True) -> webdriver.Chrome:
    os.makedirs(download_dir, exist_ok=True)
    prefs = {
        "download.default_directory": os.path.abspath(download_dir),
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "safebrowsing.enabled": True,
    }
    opts = webdriver.ChromeOptions()
    if headless:
        opts.add_argument("--headless=new")
    opts.add_argument("--disable-gpu")
    opts.add_argument("--no-sandbox")
    opts.add_argument("--disable-dev-shm-usage")
    opts.add_argument("--window-size=1680,1100")
    opts.add_experimental_option("prefs", prefs)
    opts.add_experimental_option("excludeSwitches", ["enable-logging"])
    # table renders via JS after initial load
    opts.set_capability("pageLoadStrategy", "eager")

    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=opts)
    # allow downloads in headless (no-op on newer builds but safe)
    try:
        driver.execute_cdp_cmd("Page.setDownloadBehavior", {
            "behavior": "allow",
            "downloadPath": os.path.abspath(download_dir)
        })
    except Exception:
        pass
    return driver


# -------------------- Page actions --------------------

def ensure_advanced_open(driver: webdriver.Chrome):
    wait = WebDriverWait(driver, LONG_TIMEOUT)
    # If the panel has a toggle, click it; otherwise it’s already open like in your screenshot
    try:
        toggle = wait.until(EC.element_to_be_clickable((
            By.XPATH,
            "//button[contains(.,'Advanced Search') or contains(.,'Show Advanced') or contains(.,'More Filters')]"
            " | //a[contains(.,'Advanced Search') or contains(.,'Show Advanced') or contains(.,'More Filters')]"
        )))
        driver.execute_script("arguments[0].scrollIntoView({block:'center'});", toggle)
        time.sleep(0.2)
        toggle.click()
        time.sleep(0.5)
    except Exception:
        pass


def set_status_closed(driver: webdriver.Chrome):
    wait = WebDriverWait(driver, TIMEOUT)
    # Label → next select (matches your layout)
    try:
        sel = wait.until(EC.presence_of_element_located((
            By.XPATH,
            "//label[contains(.,'Status')]/following::select[1] | //span[contains(.,'Status')]/following::select[1]"
        )))
    except Exception:
        # Fallback by id/name
        sel = wait.until(EC.presence_of_element_located((
            By.XPATH,
            "//select[contains(translate(@id,'STATUS','status'),'status') or contains(translate(@name,'STATUS','status'),'status')]"
        )))
    Select(sel).select_by_visible_text("Closed")


def click_search(driver: webdriver.Chrome):
    wait = WebDriverWait(driver, TIMEOUT)
    btn = wait.until(EC.element_to_be_clickable((
        By.XPATH,
        "//button[normalize-space()='Search' or .//span[normalize-space()='Search']]"
        " | //input[@type='submit' and (@value='Search' or @value='Find It')]"
    )))
    driver.execute_script("arguments[0].scrollIntoView({block:'center'});", btn)
    time.sleep(0.2)
    btn.click()

    # Wait for either rows or for the toolbar (export icons) to appear
    WebDriverWait(driver, LONG_TIMEOUT).until(
        EC.any_of(
            EC.presence_of_element_located((By.XPATH, "//table[.//th[contains(.,'Bid Solicitation')]]//tbody//tr[td]")),
            EC.presence_of_element_located((By.XPATH, "//*[contains(@title,'Export') or contains(@title,'Excel')]"))
        )
    )
    time.sleep(0.3)


def click_export_excel(driver: webdriver.Chrome):
    """
    Click the Excel icon in the grid's toolbar. On this UI the icon can be a
    zero-sized <img>/<span>; we climb to the clickable ancestor (<a>/<button>)
    and JS-click it to avoid ElementNotInteractable.
    """
    wait = WebDriverWait(driver, TIMEOUT)

    # Try anchors/buttons with Excel in title first (what your screenshot shows)
    xp = (
        "//a[@title='Export to Excel' or contains(@title,'Export to Excel') or contains(@title,'Excel')]"
        " | //button[@title='Export to Excel' or contains(@title,'Excel')]"
        " | //img[@title='Export to Excel' or contains(@title,'Excel')]/ancestor::a"
        " | //span[contains(@title,'Excel')]/ancestor::a"
        # fallback: any toolbar button with 'excel' in aria-label or class
        " | //a[contains(translate(@aria-label,'EXCEL','excel'),'excel')]"
        " | //button[contains(translate(@aria-label,'EXCEL','excel'),'excel')]"
    )
    icon = wait.until(EC.presence_of_element_located((By.XPATH, xp)))
    driver.execute_script("""
        function clickable(el){ while(el && !(el.tagName==='A'||el.tagName==='BUTTON')) el = el.parentElement; return el; }
        var t = clickable(arguments[0]) || arguments[0];
        if (t){ t.scrollIntoView({block:'center'}); t.click(); }
    """, icon)


# -------------------- Download handling --------------------

def wait_for_download(download_dir: str, timeout: int = 180) -> Optional[str]:
    """
    Wait for a new .xls (or .xlsx) to arrive and for any .crdownload temp files to finish.
    """
    end = time.time() + timeout
    before = set(glob.glob(os.path.join(download_dir, "*.xls*")))
    while time.time() < end:
        # keep waiting while Chrome is writing
        if glob.glob(os.path.join(download_dir, "*.crdownload")):
            time.sleep(0.5); continue
        after = set(glob.glob(os.path.join(download_dir, "*.xls*")))
        new_files = list(after - before)
        if new_files:
            return max(new_files, key=lambda p: os.path.getmtime(p))
        time.sleep(0.5)
    return None


def is_real_excel(path: str) -> bool:
    """
    Validate that the file is a real Excel:
    - Legacy .xls (OLE2/BIFF): D0 CF 11 E0
    - .xlsx (ZIP): PK\x03\x04
    """
    try:
        if os.path.getsize(path) < 1024:
            return False
        with open(path, "rb") as f:
            sig = f.read(4)
        return sig.startswith(b"\xD0\xCF\x11\xE0") or sig.startswith(b"PK\x03\x04")
    except Exception:
        return False


def finalize_output(downloaded: str, out_arg: Optional[str]) -> str:
    """
    If --out is a directory, create a timestamped .xls inside it.
    If --out is a filename, use that. Preserve .xls by default.
    """
    if not out_arg:
        return downloaded

    out_arg = os.path.abspath(out_arg)
    if os.path.isdir(out_arg):
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        final_path = os.path.join(out_arg, f"commbuys_closed_{ts}.xls")
    else:
        root, ext = os.path.splitext(out_arg)
        if not ext:
            final_path = out_arg + ".xls"
        else:
            final_path = out_arg

    os.makedirs(os.path.dirname(final_path) or ".", exist_ok=True)
    if os.path.abspath(downloaded) != os.path.abspath(final_path):
        try:
            os.replace(downloaded, final_path)
        except Exception:
            import shutil
            shutil.copy2(downloaded, final_path)
            os.remove(downloaded)
    return final_path


# -------------------- Main flow --------------------

def main():
    ap = argparse.ArgumentParser(description="COMMBUYS Advanced Search → Status=Closed → Export to Excel (.xls)")
    ap.add_argument("--download-dir", default="downloads", help="Folder where the browser will save the .xls")
    ap.add_argument("--out", default=None, help="Final path OR directory for the .xls")
    ap.add_argument("--show-browser", action="store_true", help="Show Chrome window (debug)")
    args = ap.parse_args()

    driver = open_browser(args.download_dir, headless=not args.show_browser)
    try:
        logging.info("Opening Advanced Search page…")
        driver.get(ADV_URL)

        logging.info("Making sure Advanced panel is open…")
        ensure_advanced_open(driver)

        logging.info("Selecting Status = Closed…")
        set_status_closed(driver)

        logging.info("Clicking Search…")
        click_search(driver)

        logging.info("Clicking Export to Excel…")
        click_export_excel(driver)

        logging.info("Waiting for .xls download…")
        downloaded = wait_for_download(args.download_dir, timeout=180)
        if not downloaded:
            raise RuntimeError("Timed out waiting for the .xls. Try once with --show-browser to confirm tooltip text contains 'Export to Excel'.")

        # Validate actual Excel (defends against mis-labeled HTML)
        if not is_real_excel(downloaded):
            raise RuntimeError(f"Downloaded file isn’t a real Excel: {downloaded}")

        final_path = finalize_output(downloaded, args.out)
        logging.info(f"Excel saved to: {os.path.abspath(final_path)}")

    finally:
        try:
            driver.quit()
        except Exception:
            pass


if __name__ == "__main__":
    main()
