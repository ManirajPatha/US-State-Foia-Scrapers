# lla_public_records_submitter.py
# Python 3.12+ | Selenium 4.x | Firefox (geckodriver)
#
# How to run:
#   1) `pip install selenium pandas`
#   2) Ensure geckodriver is installed and on PATH (or set GECKO_DRIVER_PATH below)
#   3) python lla_public_records_submitter.py

import time
import os
import sys
import pandas as pd

from selenium import webdriver
from selenium.webdriver.firefox.service import Service as FirefoxService
from selenium.webdriver.firefox.options import Options as FirefoxOptions
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import (
    TimeoutException,
    ElementClickInterceptedException,
    StaleElementReferenceException,
    WebDriverException
)

# ────────────────────────────────────────────────────────────────────────────────
# USER SETTINGS (edit these to your local paths if needed)
# ────────────────────────────────────────────────────────────────────────────────

FORM_URL = "https://lla.la.gov/public-records-request"

# <-- EDIT THIS: point to your local Excel file path -->
# Example: r"/Users/raajthipparthy/Desktop/Opportunity Scrapers/lla/la_awarded_bids.xlsx"
EXCEL_PATH = r"/Users/raajthipparthy/Desktop/Opportunity Scrapers/louisiana-scraper/la_awarded_bids_150.xlsx"

# Optional: If geckodriver is not on PATH, set the full path here; otherwise leave as None.
GECKO_DRIVER_PATH = None  # e.g., r"/usr/local/bin/geckodriver"

# Static fields you asked to use on every submission
FULL_NAME = "Maniraj Patha"
EMAIL = "pathamaniraj97@gmail.com"
PHONE = "6824055734"

# Visual pacing so you can watch it work (seconds)
TYPE_PAUSE = 0.2   # time between keystrokes (not per char, just small delay after send_keys)
STEP_PAUSE = 0.6   # small pause after each field
POST_SUBMIT_WAIT = 2.0  # wait after submit before next row

# Global timeout for waits (seconds)
TIMEOUT = 20

# ────────────────────────────────────────────────────────────────────────────────
# Helpers
# ────────────────────────────────────────────────────────────────────────────────

def create_driver():
    options = FirefoxOptions()
    # Make sure it is NOT headless so you can watch it work
    # options.add_argument("--headless")  # DO NOT enable; we want to see it

    if GECKO_DRIVER_PATH:
        service = FirefoxService(executable_path=GECKO_DRIVER_PATH)
    else:
        service = FirefoxService()

    driver = webdriver.Firefox(service=service, options=options)

    # Maximize then fullscreen for best visibility
    try:
        driver.maximize_window()
        time.sleep(0.4)
        driver.fullscreen_window()
    except Exception:
        # Some environments may not support fullscreen; ignore if so
        pass

    return driver


def wait_for(driver, locator, timeout=TIMEOUT):
    return WebDriverWait(driver, timeout).until(EC.presence_of_element_located(locator))


def scroll_into_view(driver, element, block="center"):
    try:
        driver.execute_script(
            "arguments[0].scrollIntoView({behavior: 'instant', block: arguments[1]});",
            element, block
        )
        time.sleep(0.2)
    except WebDriverException:
        pass


def safe_type(element, value):
    element.clear()
    element.send_keys(value)
    time.sleep(TYPE_PAUSE)


def click_submit(driver):
    # Submit button by class ".form-submit-button" and type="submit"
    submit = WebDriverWait(driver, TIMEOUT).until(
        EC.element_to_be_clickable(
            (By.CSS_SELECTOR, "button.form-submit-button[type='submit']")
        )
    )
    scroll_into_view(driver, submit, block="center")
    try:
        submit.click()
    except ElementClickInterceptedException:
        # Fallback to JS click if something overlays
        driver.execute_script("arguments[0].click();", submit)


def fill_and_submit(driver, description: str, bid_number: str):
    # We’ll always reload the form fresh for each row to avoid stale state
    driver.get(FORM_URL)

    # Wait for the Name field to exist (id=name1)
    name_input = wait_for(driver, (By.ID, "name1"))
    scroll_into_view(driver, name_input, block="center")
    safe_type(name_input, FULL_NAME)
    time.sleep(STEP_PAUSE)

    # Email (id=email)
    email_input = wait_for(driver, (By.ID, "email"))
    scroll_into_view(driver, email_input, block="center")
    safe_type(email_input, EMAIL)
    time.sleep(STEP_PAUSE)

    # Phone (id=phoneNumber)
    phone_input = wait_for(driver, (By.ID, "phoneNumber"))
    scroll_into_view(driver, phone_input, block="center")
    safe_type(phone_input, PHONE)
    time.sleep(STEP_PAUSE)

    # What type of records... (id=whatTypeOfRecordsAreYouRequesting)
    body = (
        f"I am requesting a copy of the winning proposal for {description} "
        f"contract bearing ID {bid_number}."
    )
    textarea = wait_for(driver, (By.ID, "whatTypeOfRecordsAreYouRequesting"))
    scroll_into_view(driver, textarea, block="center")  # keep the page near the textarea
    safe_type(textarea, body)
    time.sleep(STEP_PAUSE)

    # Keep the viewport near the submit button and click
    click_submit(driver)

    # Optional: wait briefly so you can see the result / confirmation
    time.sleep(POST_SUBMIT_WAIT)


def main():
    # Basic validation on Excel path
    if not os.path.isfile(EXCEL_PATH):
        print(f"[ERROR] Excel file not found at: {EXCEL_PATH}")
        sys.exit(1)

    # Read Excel; expect columns: "Bid Number", "Description"
    try:
        df = pd.read_excel(EXCEL_PATH)
    except Exception as e:
        print(f"[ERROR] Failed to read Excel file: {e}")
        sys.exit(1)

    expected_cols = {"Bid Number", "Description"}
    missing = expected_cols - set(df.columns)
    if missing:
        print(f"[ERROR] Excel missing columns: {', '.join(missing)}")
        sys.exit(1)

    # Drop rows where either value is NaN/empty
    df = df.dropna(subset=["Bid Number", "Description"])
    # Normalize to strings
    df["Bid Number"] = df["Bid Number"].astype(str).str.strip()
    df["Description"] = df["Description"].astype(str).str.strip()

    if df.empty:
        print("[INFO] No rows to process after filtering.")
        sys.exit(0)

    print(f"[INFO] Starting submissions for {len(df)} row(s).")

    driver = create_driver()

    try:
        for idx, row in df.iterrows():
            bid = row["Bid Number"]
            desc = row["Description"]
            print(f"[INFO] Submitting row {idx}: Bid Number={bid} | Description={desc}")
            try:
                fill_and_submit(driver, desc, bid)
            except TimeoutException:
                print(f"[WARN] Timeout while submitting row {idx}. Skipping to next.")
            except StaleElementReferenceException:
                print(f"[WARN] Page changed unexpectedly on row {idx}. Retrying once...")
                # Try once more fresh
                try:
                    fill_and_submit(driver, desc, bid)
                except Exception as e2:
                    print(f"[WARN] Second attempt failed for row {idx}: {e2}")
            except Exception as e:
                print(f"[WARN] Unexpected error on row {idx}: {e}. Continuing.")
    finally:
        print("[INFO] Done. Closing browser.")
        # Give you a second to see final state
        time.sleep(1.0)
        driver.quit()


if __name__ == "__main__":
    main()