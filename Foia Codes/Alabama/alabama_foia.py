"""
Automates the Alabama Secretary of State Public Records Request Form
URL: https://www.sos.alabama.gov/public-records-request-form

Behavior:
- Reads an input Excel with columns: notice_id, title
- Fills static requester details (edit CONSTANTS below)
- For each row, fills "Specific Records Requested" with:
    Notice ID: <notice_id>
    Title: <title>
- Clicks reCAPTCHA checkbox (does NOT solve challenges)
- Submits immediately (no console ENTER), then continues to the next row
- Logs per-row status to an output Excel

Tested with: Selenium 4.x, Firefox + geckodriver
"""
from __future__ import annotations
import os
import time
import sys
import traceback
from dataclasses import dataclass
from datetime import datetime
from typing import Optional, List

import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.firefox.service import Service as FirefoxService
from selenium.common.exceptions import (
    TimeoutException,
    NoSuchElementException,
    WebDriverException,
    InvalidSessionIdException,
)

# =========================
# ====== CONSTANTS ========
# =========================
FORM_URL = "https://www.sos.alabama.gov/public-records-request-form"

# Input / Output paths (edit as needed)
INPUT_XLSX = r"C:\Users\MANIRAJ\Downloads\alabama_vss_opportunities.xlsx"  # expects columns: notice_id, title
OUTPUT_XLSX = r"C:\Users\MANIRAJ\Downloads\alabama_requests_output.xlsx"

# Geckodriver path (if not on PATH)
GECKODRIVER_PATH = r"C:\Users\MANIRAJ\Downloads\Firefox Driver\geckodriver.exe"

# Static requester details
FULL_NAME = "Maniraj Patha"
ORGANIZATION = "Southern Arkansas University"
POSITION = "Student"
DAYTIME_PHONE = "+1 6824055734"
EMAIL = "pathamaniraj97@gmail.com"

# Address
STREET_ADDRESS = "8181 Fannin St"
CITY = "Houston"
COUNTY = "Harris County"
STATE_VALUE = "Texas"              # REQUIRED
ZIPCODE = "77054"

# Payment of fees without prior notice
MAX_FEE_USD = 1                    # REQUIRED per your update

# Delivery method sentence (exact)
DELIVERY_EMAIL_TEXT = "I would like to receive responsive records electronically at the email address provided above"

# Timeouts (kept short so the form starts filling quickly)
PAGE_TIMEOUT = 12
CLICK_TIMEOUT = 12

# Retry policy
MAX_ROW_RETRIES = 1  # retry a row once if the browser session dies

# =========================
# ====== UTILITIES ========
# =========================
def _xpath_literal(text: str) -> str:
    """Safely build an XPath string literal for arbitrary text."""
    if "'" not in text:
        return f"'{text}'"
    if '"' not in text:
        return f'"{text}"'
    parts = text.split("'")
    concat_parts = []
    for i, part in enumerate(parts):
        if part:
            concat_parts.append(f"'{part}'")
        if i != len(parts) - 1:
            concat_parts.append('"\'"')  # single quote as a double-quoted string
    return "concat(" + ", ".join(concat_parts) + ")"

@dataclass
class RowResult:
    notice_id: Optional[str]
    title: Optional[str]
    status: str
    message: str
    submitted_at: Optional[str]
    confirmation_text: Optional[str]

def launch_driver() -> webdriver.Firefox:
    service = FirefoxService(executable_path=GECKODRIVER_PATH) if GECKODRIVER_PATH else FirefoxService()
    options = webdriver.FirefoxOptions()
    options.set_preference("dom.webnotifications.enabled", False)
    options.set_preference("marionette.acceptInsecureCerts", True)
    options.set_preference("browser.tabs.warnOnClose", False)
    driver = webdriver.Firefox(service=service, options=options)
    driver.set_page_load_timeout(30)
    driver.maximize_window()
    return driver

def by_label_text(driver: webdriver.Firefox, label_contains: str) -> Optional[str]:
    """Find the input/select/textarea id associated with a label that contains text."""
    lit = _xpath_literal(label_contains)
    labels = driver.find_elements(By.XPATH, f"//label[contains(normalize-space(.), {lit})]")
    for lbl in labels:
        for_attr = lbl.get_attribute("for")
        if for_attr:
            return for_attr
    return None

# ---------- CHECKBOX HELPERS ----------

def ensure_checkbox_checked_by_text(driver: webdriver.Firefox, text: str) -> bool:
    """
    Ensure a checkbox whose label/neighbor contains `text` is checked.
    Uses XPath strategies first, then a JavaScript fallback that searches the DOM
    by innerText and dispatches input/change events.
    """
    # 1) XPath attempts
    lit = _xpath_literal(text)
    xpaths = [
        f"//label[contains(normalize-space(.), {lit})]//input[@type='checkbox']",
        f"//input[@type='checkbox' and (following-sibling::*[contains(normalize-space(.), {lit})] or following-sibling::text()[contains(., {lit})])]",
        f"//*[contains(normalize-space(.), {lit})]//input[@type='checkbox']",
    ]
    for xp in xpaths:
        try:
            cb = WebDriverWait(driver, CLICK_TIMEOUT).until(EC.presence_of_element_located((By.XPATH, xp)))
            driver.execute_script("arguments[0].scrollIntoView({block:'center'});", cb)
            if not cb.is_selected():
                try:
                    cb.click()
                except WebDriverException:
                    driver.execute_script("arguments[0].click();", cb)
            if cb.is_selected():
                return True
        except Exception:
            continue

    # 2) JavaScript fallback (very robust)
    js = r"""
const needle = (arguments[0] || '').toLowerCase();
const boxes = Array.from(document.querySelectorAll('input[type="checkbox"]'));
for (const cb of boxes){
  let txt = '';
  // try own label
  const id = cb.getAttribute('id');
  if (id) {
    const label = document.querySelector(`label[for="${CSS.escape(id)}"]`);
    if (label) txt = (label.innerText || '').trim();
  }
  if (!txt) {
    const label = cb.closest('label');
    if (label) txt = (label.innerText || '').trim();
  }
  if (!txt) {
    const cont = cb.closest('li, p, div, section, fieldset') || cb.parentElement;
    if (cont) txt = (cont.innerText || '').trim();
  }
  if ((txt || '').toLowerCase().includes(needle)) {
    cb.scrollIntoView({block:'center'});
    cb.checked = true;
    cb.dispatchEvent(new Event('input',  {bubbles:true}));
    cb.dispatchEvent(new Event('change', {bubbles:true}));
    return true;
  }
}
return false;
"""
    try:
        ok = driver.execute_script(js, text)
        return bool(ok)
    except Exception:
        return False

def click_first_visible_checkbox(driver: webdriver.Firefox) -> bool:
    """Fallback to click the first visible checkbox in the form (JS if needed)."""
    try:
        cb = WebDriverWait(driver, CLICK_TIMEOUT).until(
            EC.presence_of_element_located((By.XPATH, "(//form//input[@type='checkbox' and not(@disabled)])[1]"))
        )
        driver.execute_script("arguments[0].scrollIntoView({block:'center'});", cb)
        if not cb.is_selected():
            try:
                cb.click()
            except WebDriverException:
                driver.execute_script("arguments[0].click();", cb)
        return cb.is_selected()
    except Exception:
        return False

# ---------- OTHER FIELD HELPERS ----------

def select_dropdown_by_visible_text(driver: webdriver.Firefox, element, text: str) -> bool:
    """Try native <select>. Returns True if set."""
    try:
        from selenium.webdriver.support.ui import Select
        Select(element).select_by_visible_text(text)
        return True
    except Exception:
        return False

def set_state_required(driver: webdriver.Firefox, wanted_text: str = "Texas") -> None:
    """
    REQUIRED: Select State = Texas.
    Prefers native <select>. Falls back to custom dropdown behavior.
    Raises RuntimeError if not set.
    """
    # 1) <select> that has the option by text or TX value
    try:
        sel = driver.find_element(By.XPATH, f"//select[option[normalize-space(.)={_xpath_literal(wanted_text)} or translate(@value,'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz')='tx']]")
        from selenium.webdriver.support.ui import Select
        try:
            Select(sel).select_by_visible_text(wanted_text)
        except Exception:
            Select(sel).select_by_value("TX")
        return
    except NoSuchElementException:
        pass
    except Exception:
        pass

    # 2) Any select with id/name containing 'state'
    try:
        sel = driver.find_element(By.XPATH, "//select[contains(translate(@id,'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'),'state') or contains(translate(@name,'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'),'state')]")
        from selenium.webdriver.support.ui import Select
        try:
            Select(sel).select_by_visible_text(wanted_text)
            return
        except Exception:
            for val in ["TX", "tx", "Texas", "texas"]:
                try:
                    Select(sel).select_by_value(val)
                    return
                except Exception:
                    continue
            for opt in sel.find_elements(By.TAG_NAME, "option"):
                if opt.text.strip().lower() == wanted_text.lower():
                    opt.click()
                    return
    except NoSuchElementException:
        pass

    # 3) Custom dropdowns → open and click option
    elem = None
    ref_id = by_label_text(driver, "State")
    if ref_id:
        try:
            elem = driver.find_element(By.ID, ref_id)
        except NoSuchElementException:
            elem = None
    if not elem:
        for xp in [
            "//*[@role='combobox' and (contains(@id,'state') or contains(@aria-label,'State'))]",
            "//*[self::div or self::span][contains(@class,'select') and contains(translate(@class,'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'),'state')]",
            "//*[self::div or self::span][contains(translate(@id,'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'),'state')]",
        ]:
            try:
                elem = driver.find_element(By.XPATH, xp)
                break
            except NoSuchElementException:
                continue

    if elem:
        try:
            driver.execute_script("arguments[0].scrollIntoView({block:'center'});", elem)
            elem.click()
            time.sleep(0.2)
        except Exception:
            pass
        for xp in [
            f"//*[@role='option' and normalize-space(.)={_xpath_literal(wanted_text)}]",
            f"//li[normalize-space(.)={_xpath_literal(wanted_text)}]",
            f"//*[contains(@class,'option') and normalize-space(.)={_xpath_literal(wanted_text)}]",
        ]:
            try:
                opt = WebDriverWait(driver, 3).until(EC.element_to_be_clickable((By.XPATH, xp)))
                driver.execute_script("arguments[0].scrollIntoView({block:'center'});", opt)
                opt.click()
                return
            except Exception:
                continue
        try:
            if elem.tag_name.lower() in ("input", "textarea"):
                try:
                    elem.clear()
                except Exception:
                    pass
                elem.send_keys(wanted_text)
                return
        except Exception:
            pass

    raise RuntimeError("State (required) could not be set to Texas.")

def fill_specific_records(driver: webdriver.Firefox, notice_id: Optional[str], title: Optional[str]) -> bool:
    """Fill 'Specific Records Requested' with row values."""
    value = f"Notice ID: {notice_id or ''}\nTitle: {title or ''}"
    spec_id = by_label_text(driver, "Specific Records Requested")
    target = None
    if spec_id:
        try:
            target = driver.find_element(By.ID, spec_id)
        except NoSuchElementException:
            target = None
    if not target:
        for xp in [
            "//textarea[contains(@placeholder,'Specific') or contains(@aria-label,'Specific') or contains(@name,'Specific') or contains(@id,'Specific')]",
            "//textarea[contains(@placeholder,'Be as specific') or contains(@aria-label,'Records') or contains(@name,'records') or contains(@id,'records')]",
            "//textarea",
        ]:
            try:
                target = driver.find_element(By.XPATH, xp)
                break
            except NoSuchElementException:
                continue
    if not target:
        return False
    driver.execute_script("arguments[0].scrollIntoView({block:'center'});", target)
    try:
        target.clear()
    except Exception:
        pass
    target.send_keys(value)
    return True

def find_send_button(driver: webdriver.Firefox):
    """Find the Send message button via multiple selectors."""
    xpaths = [
        "//button[normalize-space(.)='Send message']",
        "//button[normalize-space(.)='Send Message']",
        "//button[contains(translate(., 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'),'send message')]",
        "//input[@type='submit' and (translate(@value,'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz')='send message' or contains(translate(@value,'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'),'send'))]",
        "//*[@role='button' and contains(translate(., 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'),'send message')]",
        "//button[@type='submit' or @type='Submit']",
    ]
    for xp in xpaths:
        try:
            el = WebDriverWait(driver, 3).until(EC.element_to_be_clickable((By.XPATH, xp)))
            return el
        except Exception:
            continue
    return None

def safe_accept_alert(driver: webdriver.Firefox):
    try:
        alert = driver.switch_to.alert
        _ = alert.text
        alert.accept()
    except Exception:
        pass

# =========================
# ========= MAIN ==========
# =========================
def process_single_row(notice_id: Optional[str], title: Optional[str]) -> RowResult:
    confirmation_text = None
    status = "STARTED"
    message = ""

    driver = launch_driver()
    wait = WebDriverWait(driver, PAGE_TIMEOUT)

    try:
        # Go straight to the form and start filling as soon as any part of the form is present
        driver.get(FORM_URL)
        wait.until(EC.presence_of_element_located((By.XPATH, "//form|//h1|//label")))

        # Possession checkbox (first on page)
        phrase = "I believe the records I am seeking are in the possession of the Office of the Secretary of State"
        if not ensure_checkbox_checked_by_text(driver, phrase):
            click_first_visible_checkbox(driver)  # lenient fallback

        # Contact info
        full_id = by_label_text(driver, "Full Name")
        if full_id:
            driver.find_element(By.ID, full_id).send_keys(FULL_NAME)

        org_id = by_label_text(driver, "Organization")
        if org_id:
            driver.find_element(By.ID, org_id).send_keys(ORGANIZATION)

        pos_id = by_label_text(driver, "Position")
        if pos_id:
            driver.find_element(By.ID, pos_id).send_keys(POSITION)

        phone_id = by_label_text(driver, "Daytime Phone")
        if phone_id:
            driver.find_element(By.ID, phone_id).send_keys(DAYTIME_PHONE)

        email_id = by_label_text(driver, "Email Address")
        if email_id:
            driver.find_element(By.ID, email_id).send_keys(EMAIL)

        # Address
        street_id = by_label_text(driver, "Current street address")
        if street_id:
            driver.find_element(By.ID, street_id).send_keys(STREET_ADDRESS)

        city_id = by_label_text(driver, "City")
        if city_id:
            driver.find_element(By.ID, city_id).send_keys(CITY)

        county_id = by_label_text(driver, "County")
        if county_id:
            driver.find_element(By.ID, county_id).send_keys(COUNTY)

        # State (REQUIRED)
        set_state_required(driver, STATE_VALUE)

        zip_id = by_label_text(driver, "Zipcode")
        if zip_id:
            driver.find_element(By.ID, zip_id).send_keys(ZIPCODE)

        # Payment of Fees -> 1
        fee_id = by_label_text(driver, "The amount I am willing to pay")
        if fee_id:
            fee_elem = driver.find_element(By.ID, fee_id)
            try:
                fee_elem.clear()
            except Exception:
                pass
            fee_elem.send_keys(str(MAX_FEE_USD))

        # Specific Records Requested (uses row values)
        fill_specific_records(driver, notice_id, title)

        # Are you an Alabama Citizen? (checkbox)
        ensure_checkbox_checked_by_text(driver, "Are you an Alabama Citizen?")

        # Method of Delivery → Email (this is the one you said isn’t sticking)
        # Use robust JS fallback to force-check the box by the exact sentence:
        ok_delivery = ensure_checkbox_checked_by_text(driver, DELIVERY_EMAIL_TEXT)
        if not ok_delivery:
            # shorter variants, just in case the text varies slightly
            for variant in [
                "receive responsive records electronically at the email address provided above",
                "receive responsive records electronically",
                "I would like to receive responsive records electronically",
            ]:
                if ensure_checkbox_checked_by_text(driver, variant):
                    ok_delivery = True
                    break

        # Signature
        sig_id = by_label_text(driver, "Signature")
        if sig_id:
            driver.find_element(By.ID, sig_id).send_keys(FULL_NAME)

        # Today's Date (mm/dd/yyyy)
        today_str = datetime.now().strftime("%m/%d/%Y")
        date_input = None
        date_id = by_label_text(driver, "Today's Date")
        if date_id:
            date_input = driver.find_element(By.ID, date_id)
        else:
            try:
                date_input = driver.find_element(By.XPATH, "//input[@type='date']")
            except NoSuchElementException:
                try:
                    date_input = driver.find_element(
                        By.XPATH,
                        "//input[contains(@placeholder, 'Date') or contains(@aria-label, 'Date') or "
                        "contains(@name, 'date') or contains(@id, 'date')]"
                    )
                except NoSuchElementException:
                    date_input = None
        if date_input:
            try:
                date_input.clear()
            except Exception:
                pass
            date_input.send_keys(today_str)

        # reCAPTCHA (best effort) – no waiting for user
        try:
            iframe = driver.find_element(By.XPATH, "//iframe[@title='reCAPTCHA']")
            driver.switch_to.frame(iframe)
            time.sleep(0.5)
            driver.find_element(By.CSS_SELECTOR, ".recaptcha-checkbox-border").click()
            driver.switch_to.default_content()
        except Exception:
            pass

        # Submit immediately
        btn = find_send_button(driver)
        if not btn:
            safe_accept_alert(driver)
            btn = find_send_button(driver)
        if not btn:
            raise RuntimeError("Could not locate the 'Send message' button.")
        driver.execute_script("arguments[0].scrollIntoView({block:'center'});", btn)
        try:
            btn.click()
        except WebDriverException:
            time.sleep(0.3)
            safe_accept_alert(driver)
            btn = find_send_button(driver)
            if btn:
                btn.click()

        # Confirmation (best effort, short wait)
        try:
            conf = WebDriverWait(driver, 6).until(
                EC.presence_of_element_located(
                    (By.XPATH, "//*[contains(@class,'messages') or contains(., 'Thank') or contains(., 'received') or contains(., 'success')]")
                )
            )
            confirmation_text = conf.text.strip()
            status = "SUBMITTED"
            message = "Submitted successfully"
        except TimeoutException:
            status = "SUBMITTED"
            message = "Submitted (no confirmation message detected)"

    except (InvalidSessionIdException, WebDriverException) as e:
        status = "FAILED"
        message = f"WebDriver error: {e}"
        traceback.print_exc()
    except Exception as e:
        status = "FAILED"
        message = f"Error: {e}"
        traceback.print_exc()
    finally:
        try:
            driver.quit()
        except Exception:
            pass

    return RowResult(
        notice_id=notice_id,
        title=title,
        status=status,
        message=message,
        submitted_at=datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        confirmation_text=confirmation_text,
    )

def main():
    # Read input
    if not os.path.exists(INPUT_XLSX):
        print(f"Input not found: {INPUT_XLSX}")
        sys.exit(1)
    df = pd.read_excel(INPUT_XLSX)
    if "notice_id" not in df.columns or "title" not in df.columns:
        print("Expected columns missing: 'notice_id', 'title'")
        sys.exit(1)

    results: List[RowResult] = []

    for _, row in df.iterrows():
        notice_id = str(row.get("notice_id") or "").strip() or None
        title = str(row.get("title") or "").strip() or None

        res = process_single_row(notice_id, title)
        if res.status == "FAILED" and (
            "InvalidSessionIdException" in (res.message or "")
            or "Failed to decode response from marionette" in (res.message or "")
            or "WebDriver error" in (res.message or "")
        ):
            print(">>> Retrying this row once due to browser session issue...")
            res = process_single_row(notice_id, title)

        results.append(res)

    # Write output
    out_df = pd.DataFrame([r.__dict__ for r in results])
    out_df.to_excel(OUTPUT_XLSX, index=False)
    print(f"\nOutput written: {OUTPUT_XLSX}\n")

if __name__ == "__main__":
    main()