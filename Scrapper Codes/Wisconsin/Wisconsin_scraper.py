# pip install selenium pandas python-dateutil
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import (
    TimeoutException,
    StaleElementReferenceException,
    ElementClickInterceptedException,
)
from dateutil import parser as dateparser
from datetime import datetime
import json
import time

START_URL = "https://vendornet.wi.gov/Contracts.aspx"
OUTPUT_JSON = "wisconsin_bids_awarded.json"   # <-- JSON output

def new_driver():
    opts = webdriver.ChromeOptions()
    opts.add_argument("--start-maximized")
    opts.add_argument("--disable-gpu")
    opts.add_argument("--no-sandbox")
    opts.add_argument("--disable-dev-shm-usage")
    drv = webdriver.Chrome(options=opts)
    drv.set_page_load_timeout(60)
    return drv

def wait_for_grid(drv, timeout=20):
    return WebDriverWait(drv, timeout).until(
        EC.presence_of_element_located((By.CSS_SELECTOR, "table.rgMasterTable"))
    )

def click_bids_tab(drv):
    try:
        link = WebDriverWait(drv, 15).until(
            EC.element_to_be_clickable((By.XPATH, "//a[contains(@href,'Bids.aspx') and normalize-space()='Bids']"))
        )
        drv.execute_script("arguments[0].click();", link)
    except TimeoutException:
        # Fallback by link text
        link = WebDriverWait(drv, 10).until(EC.element_to_be_clickable((By.LINK_TEXT, "Bids")))
        drv.execute_script("arguments[0].click();", link)

def click_with_staleness_wait(
    drv, clickable_locator, container_locator=(By.CSS_SELECTOR, "table.rgMasterTable"), max_attempts=6
):
    for _ in range(max_attempts):
        try:
            old = WebDriverWait(drv, 20).until(EC.presence_of_element_located(container_locator))
            el = WebDriverWait(drv, 20).until(EC.element_to_be_clickable(clickable_locator))
            drv.execute_script("arguments[0].scrollIntoView({block:'center'});", el)
            drv.execute_script("arguments[0].click();", el)
            WebDriverWait(drv, 20).until(EC.staleness_of(old))
            WebDriverWait(drv, 20).until(EC.presence_of_element_located(container_locator))
            return True
        except (StaleElementReferenceException, ElementClickInterceptedException, TimeoutException):
            time.sleep(0.3)
    return False

def ensure_checkbox_state_by_label_text(drv, label_text, should_be_checked: bool):
    label = WebDriverWait(drv, 12).until(
        EC.presence_of_element_located((By.XPATH, f"//label[contains(., '{label_text}')]"))
    )
    cls = (label.get_attribute("class") or "").strip()
    is_checked = "rfdCheckboxChecked" in cls
    if (should_be_checked and not is_checked) or ((not should_be_checked) and is_checked):
        try:
            old_table = drv.find_element(By.CSS_SELECTOR, "table.rgMasterTable")
            drv.execute_script("arguments[0].click();", label)
            WebDriverWait(drv, 15).until(EC.staleness_of(old_table))
            wait_for_grid(drv, 15)
        except Exception:
            drv.execute_script("arguments[0].click();", label)
            time.sleep(0.5)

def apply_filters(drv):
    # Keep your original filter semantics
    try:
        ensure_checkbox_state_by_label_text(drv, "Include eSupplier", False)
    except Exception:
        pass
    try:
        ensure_checkbox_state_by_label_text(drv, "Awarded/ Canceled", True)
    except Exception:
        pass
    # Hit Search
    try:
        search_btn = drv.find_element(By.XPATH, "//*[self::input or self::button][@value='Search' or normalize-space()='Search']")
        old_table = drv.find_element(By.CSS_SELECTOR, "table.rgMasterTable")
        drv.execute_script("arguments[0].click();", search_btn)
        WebDriverWait(drv, 15).until(EC.staleness_of(old_table))
        wait_for_grid(drv, 15)
    except Exception:
        wait_for_grid(drv, 15)

def sort_by_available_date_desc(drv):
    header_locator = (By.XPATH, "//th[.//a[normalize-space()='Available Date']]//a")
    if not click_with_staleness_wait(drv, header_locator):
        raise TimeoutException("Failed to click Available Date header (1)")
    time.sleep(0.2)
    if not click_with_staleness_wait(drv, header_locator):
        raise TimeoutException("Failed to click Available Date header (2)")

def parse_dt(txt):
    txt = (txt or "").strip()
    if not txt:
        return ""
    try:
        return dateparser.parse(txt, dayfirst=False).strftime("%Y-%m-%d %H:%M:%S")
    except Exception:
        return txt  # leave as-is if parsing fails

def parse_row(tr):
    """
    Return table row dict or None.
    Add a flag whether this is canceled/cancelled so we can skip it.
    """
    tds = tr.find_elements(By.TAG_NAME, "td")
    if len(tds) < 6:
        return None

    # Column 1: Solicitation Reference # (with URL)
    reference_text = tds[0].text.strip()
    ref_url = ""
    try:
        a = tds[0].find_element(By.TAG_NAME, "a")
        reference_text = a.text.strip() or reference_text
        ref_url = a.get_attribute("href") or ""
    except Exception:
        pass

    # Skip markers
    is_canceled = "canceled" in reference_text.lower() or "cancelled" in reference_text.lower()

    title = tds[1].text.strip()
    agency = tds[2].text.strip()
    available = parse_dt(tds[3].text)
    due = parse_dt(tds[4].text)

    e_supplier = False
    try:
        label = tds[5].find_element(By.XPATH, ".//label")
        e_supplier = "rfdCheckboxChecked" in ((label.get_attribute("class") or ""))
    except Exception:
        pass

    return {
        "Solicitation Reference #": reference_text,
        "Title": title,
        "Agency": agency,
        "Available Date": available,
        "Due Date": due,
        "Available in eSupplier": e_supplier,
        "Bid URL": ref_url,
        "_is_canceled": is_canceled,
    }

def collect_page_rows(drv):
    data = []
    table = drv.find_element(By.CSS_SELECTOR, "table.rgMasterTable")
    tbody = table.find_element(By.TAG_NAME, "tbody")
    rows = tbody.find_elements(By.XPATH, "./tr[contains(@class,'rgRow') or contains(@class,'rgAltRow')]")
    for tr in rows:
        item = parse_row(tr)
        if item:
            data.append(item)
    return data

def _get_first_ref_text(drv):
    try:
        table = drv.find_element(By.CSS_SELECTOR, "table.rgMasterTable")
        tbody = table.find_element(By.TAG_NAME, "tbody")
        first_anchor = tbody.find_element(
            By.XPATH,
            "./tr[(contains(@class,'rgRow') or contains(@class,'rgAltRow'))][1]/td[1]//a"
        )
        return first_anchor.text.strip()
    except Exception:
        return ""

def _get_current_page_index(drv):
    try:
        span = drv.find_element(By.CSS_SELECTOR, ".rgPager span.rgCurrentPage")
        return int(span.text.strip())
    except Exception:
        pass
    try:
        inp = drv.find_element(By.CSS_SELECTOR, ".rgPager input.rgCurrentPageBox")
        return int(inp.get_attribute("value") or "0")
    except Exception:
        pass
    return None

def _find_enabled_next(drv):
    candidates = drv.find_elements(By.CSS_SELECTOR, ".rgPager a.rgPageNext, .rgPager button.rgPageNext, .rgPager input.rgPageNext")
    for el in candidates:
        cls = (el.get_attribute("class") or "")
        if el.get_attribute("disabled") or "rgDisabled" in cls:
            continue
        return el
    return None

def _click_next(drv):
    el = WebDriverWait(drv, 6).until(lambda d: _find_enabled_next(d))
    drv.execute_script("arguments[0].click();", el)

def _click_numeric_page(drv, target_index):
    try:
        link = drv.find_element(By.XPATH, f"//div[contains(@class,'rgPager')]//div[contains(@class,'rgNumPart')]//a[normalize-space()='{target_index}']")
        drv.execute_script("arguments[0].click();", link)
        return True
    except Exception:
        return False

def _set_page_by_input(drv, target_index, timeout=10):
    try:
        inp = drv.find_element(By.CSS_SELECTOR, ".rgPager input.rgCurrentPageBox")
    except Exception:
        return False
    before_first = _get_first_ref_text(drv)
    drv.execute_script("arguments[0].scrollIntoView({block:'center'});", inp)
    inp.clear()
    inp.send_keys(str(target_index))
    inp.send_keys(Keys.ENTER)

    def page_changed(d):
        try:
            cur_idx = _get_current_page_index(d)
            if cur_idx == target_index:
                return True
            cur_first = _get_first_ref_text(d)
            return cur_first and cur_first != before_first
        except StaleElementReferenceException:
            return False

    WebDriverWait(drv, timeout, poll_frequency=0.2).until(page_changed)
    return True

def go_to_next_page(drv, timeout=10):
    before_first = _get_first_ref_text(drv)
    cur_idx = _get_current_page_index(drv)

    if cur_idx is not None and _click_numeric_page(drv, cur_idx + 1):
        pass
    elif _find_enabled_next(drv) is not None:
        _click_next(drv)
    elif cur_idx is not None:
        if not _set_page_by_input(drv, cur_idx + 1, timeout=timeout):
            return False
        return True
    else:
        return False

    def changed(d):
        try:
            new_idx = _get_current_page_index(d)
            if cur_idx is not None and new_idx is not None and new_idx != cur_idx:
                return True
            first = _get_first_ref_text(d)
            return first not in ("", before_first) and first != before_first
        except StaleElementReferenceException:
            return False

    WebDriverWait(drv, timeout, poll_frequency=0.2).until(changed)
    return True

def try_set_page_size(drv, target_size=100):
    # Best effort; if pager size is not available, just continue
    try:
        select_el = drv.find_element(By.CSS_SELECTOR, ".rgPager select")
        old_first = _get_first_ref_text(drv)
        Select(select_el).select_by_visible_text(str(target_size))
        WebDriverWait(drv, 12).until(lambda d: _get_first_ref_text(d) != old_first and _get_first_ref_text(d) != "")
        return True
    except Exception:
        return False

# ---------------- award detail scraping ---------------- #

def open_in_new_tab_and_scrape(drv, url):
    """
    Open URL in a new tab, scrape:
      • Award Date
      • Vendor Details: ONLY names from Award Vendor(s) with a FILLED checkbox; or 'No Records to Display'
    Also capture the page URL.
    NEW: detect 'This bid has been canceled' banner; return canceled flag so caller can skip the row.

    Returns: (attachments_url, award_date_text, vendor_details_text, canceled_on_detail)
    """
    from selenium.webdriver.common.by import By
    from selenium.webdriver.support import expected_conditions as EC
    from selenium.webdriver.support.ui import WebDriverWait
    import time

    base = drv.current_window_handle
    drv.execute_script("window.open(arguments[0], '_blank');", url)
    WebDriverWait(drv, 10).until(lambda d: len(d.window_handles) > 1)
    drv.switch_to.window(drv.window_handles[-1])

    WebDriverWait(drv, 20).until(EC.presence_of_element_located((By.TAG_NAME, "body")))
    time.sleep(0.2)

    attachments_url = drv.current_url

    # --- CANCEL banner detection (detail page) ---
    canceled_on_detail = False
    try:
        # match 'This bid has been canceled' OR 'cancelled' (case-insensitive, extra spaces ok)
        _banner = drv.find_elements(
            By.XPATH,
            "//*[contains(translate(normalize-space(.),"
            " 'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'),"
            " 'this bid has been cancel')]"
        )
        if _banner:
            canceled_on_detail = True
    except Exception:
        pass

    # If canceled, close tab and return quickly (caller will skip)
    if canceled_on_detail:
        drv.close()
        drv.switch_to.window(base)
        return attachments_url, "", "", True

    # --- Award Date (same as before) ---
    award_date_text = ""
    try:
        val = drv.find_element(By.XPATH, "//td[normalize-space()='Award Date:']/following-sibling::td[1]")
        award_date_text = val.text.strip()
    except Exception:
        try:
            val = drv.find_element(By.XPATH, "//*[contains(normalize-space(),'Award Date')]/following::*[1]")
            award_date_text = val.text.strip()
        except Exception:
            award_date_text = ""

    # --- Vendor Details: ONLY awarded rows (filled checkbox) in Award Vendor(s) box ---
    vendor_details_text = ""
    try:
        section = drv.find_element(By.XPATH, "//*[contains(normalize-space(),'Award Vendor(s)')]")
        table = section.find_element(By.XPATH, ".//following::table[1]")

        # 'No records to display.'?
        try:
            no_rec = table.find_element(
                By.XPATH,
                ".//*[contains(translate(normalize-space(.), 'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'),"
                " 'no records to display.')]"
            )
            if no_rec:
                vendor_details_text = "No Records to Display"
        except Exception:
            pass

        if not vendor_details_text:
            headers = table.find_elements(By.XPATH, ".//th")
            vendor_col_idx = None
            for i, th in enumerate(headers):
                if "vendor" in (th.text or "").strip().lower():
                    vendor_col_idx = i
                    break

            names = []
            rows = table.find_elements(By.XPATH, ".//tr[td]")
            for r in rows:
                tds = r.find_elements(By.XPATH, ".//td")
                if not tds:
                    continue

                # awarded = filled checkbox
                awarded = False
                try:
                    cb_label = r.find_element(By.XPATH, ".//td[1]//label")
                    cls = (cb_label.get_attribute("class") or "").lower()
                    awarded = "rfdcheckboxchecked" in cls
                except Exception:
                    try:
                        cb = r.find_element(By.XPATH, ".//td[1]//input[@type='checkbox' and (@checked='true' or @checked='checked')]")
                        awarded = cb is not None
                    except Exception:
                        awarded = False

                if not awarded:
                    continue

                if vendor_col_idx is not None and vendor_col_idx < len(tds):
                    raw = tds[vendor_col_idx].text.strip()
                else:
                    raw = r.text.strip()

                if raw:
                    candidate = [ln.strip() for ln in raw.splitlines() if ln.strip()]
                    if candidate:
                        name = candidate[-1]
                        if name and "vendor" not in name.lower() and "filled in box" not in name.lower():
                            names.append(name)

            vendor_details_text = ", ".join(dict.fromkeys(names)) if names else "No Records to Display"

    except Exception:
        vendor_details_text = vendor_details_text or "No Records to Display"

    drv.close()
    drv.switch_to.window(base)
    return attachments_url, award_date_text, vendor_details_text, False

# ---------------- main ---------------- #

def main():
    drv = new_driver()
    out_records = []
    pages_processed = 0

    try:
        drv.get(START_URL)
        click_bids_tab(drv)
        wait_for_grid(drv)
        apply_filters(drv)
        sort_by_available_date_desc(drv)
        try_set_page_size(drv, 100)  # best-effort; ok if not present

        while True:
            page_rows = collect_page_rows(drv)
            for row in page_rows:
                # 1) Skip canceled/cancelled rows
                if row.get("_is_canceled"):
                    continue

                # 3 & 4) Follow each non-canceled link and scrape Award Date + Vendors, and keep page URL
                bid_url = row.get("Bid URL") or ""
                attach_url = ""
                award_date = ""
                vendor_details = ""
                if bid_url:
                    try:
                        attach_url, award_date, vendor_details = open_in_new_tab_and_scrape(drv, bid_url)
                    except Exception:
                        # If anything fails on detail page, keep row but leave fields blank
                        attach_url, award_date, vendor_details = (bid_url, "", "")

                # Add requested output fields (and remove helper flags)
                row["Award Date (Detail)"] = award_date
                row["Vendor Details"] = vendor_details
                row["Wisconsion_Attachments_URL's"] = [attach_url] if attach_url else []
                row.pop("_is_canceled", None)

                out_records.append(row)

            pages_processed += 1

            # 5) Go through ALL pages until pager is exhausted
            try:
                if not go_to_next_page(drv):
                    break
                # Wait for new grid to appear
                wait_for_grid(drv, 15)
            except Exception:
                break

        # 7) Write JSON
        with open(OUTPUT_JSON, "w", encoding="utf-8") as f:
            json.dump(out_records, f, ensure_ascii=False, indent=2)

        print(f"[OK] Saved {len(out_records)} rows from {pages_processed} pages to {OUTPUT_JSON}")

    finally:
        drv.quit()

if __name__ == "__main__":
    main()
