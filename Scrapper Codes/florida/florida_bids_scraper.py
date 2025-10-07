import time
from typing import Optional, Tuple, List

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.firefox.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
from selenium.common.exceptions import (
    TimeoutException,
    StaleElementReferenceException,
    ElementClickInterceptedException,
    ElementNotInteractableException,
)

URL = "https://vendor.myfloridamarketplace.com/search/bids"

AD_TYPES_TO_SELECT = [
    "Request for Proposals",
    "Request for Information",
    "Request for Statement of Qualifications",
]
AD_STATUS_TO_SELECT = "CLOSED"


# -------------------- core utils --------------------

def open_firefox():
    opts = Options()
    opts.headless = False
    driver = webdriver.Firefox(options=opts)  # Selenium Manager resolves geckodriver
    try:
        driver.maximize_window()
    except Exception:
        driver.set_window_size(1600, 1200)
    return driver

def js_ready(driver, timeout=40):
    WebDriverWait(driver, timeout).until(
        lambda d: d.execute_script("return document.readyState") == "complete"
    )

def scroll_into_view_smart(driver, el):
    # Scroll element into nearest scrollable container AND adjust container scrollTop
    driver.execute_script("""
    const el = arguments[0];
    try { el.scrollIntoView({block:'nearest', inline:'nearest'}); } catch(e) {}
    let parent = el.parentElement;
    const isScrollable = (node) => {
      if (!node) return false;
      const s = getComputedStyle(node);
      return /(auto|scroll)/.test(s.overflowY) || /(auto|scroll)/.test(s.overflow);
    };
    while (parent) {
      if (isScrollable(parent)) {
        const elTop = el.getBoundingClientRect().top + window.scrollY;
        const parentTop = parent.getBoundingClientRect().top + window.scrollY;
        parent.scrollTop += (elTop - parentTop) - (parent.clientHeight/2);
        break;
      }
      parent = parent.parentElement;
    }
    """, el)

def safe_click(driver, el):
    # Smart scroll + move + click, with JS fallback
    try:
        scroll_into_view_smart(driver, el)
        ActionChains(driver).move_to_element(el).pause(0.05).perform()
    except Exception:
        pass
    time.sleep(0.15)
    try:
        el.click()
    except (ElementClickInterceptedException, StaleElementReferenceException, ElementNotInteractableException):
        driver.execute_script("arguments[0].click();", el)

def maybe_dismiss_consent(driver):
    candidates = [
        "//button[normalize-space()='Accept']",
        "//button[normalize-space()='I Accept']",
        "//button[contains(., 'Accept')]",
        "//button[contains(., 'Got it')]",
        "//button[contains(., 'OK')]",
        "//span[normalize-space()='Accept']/parent::button",
    ]
    for xp in candidates:
        try:
            el = WebDriverWait(driver, 2).until(EC.element_to_be_clickable((By.XPATH, xp)))
            safe_click(driver, el)
            time.sleep(0.2)
            break
        except TimeoutException:
            pass
        except Exception:
            pass

def list_frames(driver):
    try:
        return driver.find_elements(By.CSS_SELECTOR, "iframe, frame")
    except Exception:
        return []

def find_across_frames(driver, locator, timeout_each=3, overall_timeout=25):
    """
    Returns (element, frame_index or None).
    Hunts in top document first, then each iframe by index.
    """
    deadline = time.time() + overall_timeout
    last_err = None
    while time.time() < deadline:
        # top
        try:
            driver.switch_to.default_content()
            el = WebDriverWait(driver, timeout_each).until(EC.presence_of_element_located(locator))
            return el, None
        except Exception as e:
            last_err = e

        # frames
        frames = list_frames(driver)
        for idx, fr in enumerate(frames):
            try:
                driver.switch_to.default_content()
                driver.switch_to.frame(fr)
                el = WebDriverWait(driver, timeout_each).until(EC.presence_of_element_located(locator))
                return el, idx
            except Exception as e:
                last_err = e
                continue

        time.sleep(0.4)

    driver.switch_to.default_content()
    raise TimeoutException(f"Not found: {locator}. Last error: {last_err}")

def switch_to_frame_index(driver, idx: Optional[int]):
    driver.switch_to.default_content()
    if idx is None:
        return
    frames = list_frames(driver)
    if 0 <= idx < len(frames):
        driver.switch_to.frame(frames[idx])

# -------------------- tolerant header expansion --------------------

def click_any_header_by_text(driver, text: str, timeout=20) -> bool:
    """
    Flexible: find any node whose text == text, climb to something clickable,
    and click it. Returns True if clicked.
    """
    text_locator = (By.XPATH, f"//*[normalize-space()='{text}']")
    try:
        node, idx = find_across_frames(driver, text_locator, timeout_each=3, overall_timeout=timeout)
        switch_to_frame_index(driver, idx)
    except TimeoutException:
        return False

    # climb to a clickable ancestor
    header_candidates = [
        ".//ancestor::mat-expansion-panel-header[1]",
        ".//ancestor::*[contains(@class,'mat-expansion-panel-header')][1]",
        ".//ancestor::button[1]",
        ".//ancestor::div[1]",
    ]
    header = None
    for rel in header_candidates:
        try:
            header = node.find_element(By.XPATH, rel)
            break
        except Exception:
            continue
    if header is None:
        return False

    # click if not already expanded
    try:
        aria = header.get_attribute("aria-expanded")
        if aria and aria.lower() == "true":
            return True
    except Exception:
        pass

    safe_click(driver, header)
    time.sleep(0.25)
    return True

def expand_all_expansion_headers(driver):
    """
    Fallback: open all expansion headers so options are visible.
    """
    opened = 0
    for idx in [None] + list(range(len(list_frames(driver)))):
        switch_to_frame_index(driver, idx)
        headers: List = driver.find_elements(By.XPATH, "//mat-expansion-panel-header | //*[(contains(@class,'mat-expansion-panel-header'))]")
        for h in headers:
            try:
                aria = h.get_attribute("aria-expanded")
                if aria != "true":
                    safe_click(driver, h)
                    opened += 1
            except Exception:
                continue
    return opened

# -------------------- option selection & search --------------------

def select_option_by_text(driver, visible_text: str, timeout=25):
    """
    Select a list option robustly:
    - Find the mat-list-option by its visible label
    - Scroll nearest scrollable container
    - Click a reliable child (pseudo-checkbox or label)
    - JS-click fallback
    """
    print(f"   ✓ Selecting option: {visible_text}")
    xp = ("//mat-list-option[@role='option']"
          f"[.//div[contains(@class,'mat-list-text') and normalize-space()='{visible_text}']]")
    opt, idx = find_across_frames(driver, (By.XPATH, xp), timeout_each=4, overall_timeout=timeout)
    switch_to_frame_index(driver, idx)

    # Scroll container smartly
    scroll_into_view_smart(driver, opt)
    time.sleep(0.1)

    # Try clicking a reliable child first
    click_targets = [
        ".//div[contains(@class,'mat-pseudo-checkbox')]",  # checkbox visual
        ".//div[contains(@class,'mat-list-text')]",        # the label area
        ".",                                               # the option itself as last resort
    ]
    for rel in click_targets:
        try:
            tgt = opt if rel == "." else opt.find_element(By.XPATH, rel)
            safe_click(driver, tgt)
            time.sleep(0.1)
            break
        except Exception:
            continue

    # Confirm selected
    WebDriverWait(driver, timeout).until(
        EC.presence_of_element_located((
            By.XPATH,
            ("//mat-list-option[@role='option' and @aria-selected='true']"
             f"[.//div[contains(@class,'mat-list-text') and normalize-space()='{visible_text}']]")
        ))
    )

def click_search(driver, timeout=30):
    print("→ Clicking the Search button…")
    btn_locator = (By.XPATH, "//button[@type='submit' and .//span[normalize-space()='Search']]")
    btn, idx = find_across_frames(driver, btn_locator, timeout_each=4, overall_timeout=timeout)
    switch_to_frame_index(driver, idx)
    scroll_into_view_smart(driver, btn)
    safe_click(driver, btn)

    # best-effort wait for some content change on the right
    try:
        WebDriverWait(driver, 30).until(
            EC.presence_of_element_located((
                By.XPATH,
                "//*[contains(@class,'result') or contains(@class,'list') or contains(@class,'content') or contains(@class,'table')]"
            ))
        )
    except TimeoutException:
        print("   (Note) Proceeding after submitting Search.")

# -------------------- main flow --------------------

def run():
    driver = open_firefox()
    try:
        print("→ Opening page…")
        driver.get(URL)
        js_ready(driver)
        maybe_dismiss_consent(driver)

        # Try to open "Ad Type"; otherwise open everything.
        print("→ Expanding panel: Ad Type")
        opened = click_any_header_by_text(driver, "Ad Type")
        if not opened:
            print("   (fallback) Couldn’t directly hit 'Ad Type'. Expanding all panels…")
            expand_all_expansion_headers(driver)

        # Select required Ad Type items
        for label in AD_TYPES_TO_SELECT:
            for attempt in range(2):
                try:
                    select_option_by_text(driver, label)
                    break
                except (StaleElementReferenceException, ElementNotInteractableException):
                    time.sleep(0.3)

        # Open Ad Status (or ensure it’s open)
        print("→ Expanding panel: Ad Status")
        opened = click_any_header_by_text(driver, "Ad Status")
        if not opened:
            expand_all_expansion_headers(driver)

        # Select CLOSED
        for attempt in range(2):
            try:
                select_option_by_text(driver, AD_STATUS_TO_SELECT)
                break
            except (StaleElementReferenceException, ElementNotInteractableException):
                time.sleep(0.3)

        # Click Search
        click_search(driver)

        print("✓ Step 1 complete: filters selected and Search clicked.")
        time.sleep(6)

    finally:
        # Keep browser open for manual inspection while we iterate.
        # Uncomment to auto-close once everything looks good:
        # driver.quit()
        pass


if __name__ == "__main__":
    run()