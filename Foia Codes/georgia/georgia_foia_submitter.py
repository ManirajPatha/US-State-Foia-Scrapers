#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
Georgia Open Records Request auto-filler (Firefox + geckodriver, macOS).
Flow per row:
  - Open fresh Firefox and load the form
  - Scroll down to each field (visible typing) and fill in order
  - Fill Comments LAST and then DO NOT SCROLL (stay where you are)
  - Wait 10 seconds so you can click Submit manually
  - Close browser and continue with next row
"""

import sys
import time
import pandas as pd
from selenium import webdriver
from selenium.webdriver import FirefoxOptions
from selenium.webdriver.common.by import By
from selenium.webdriver.firefox.service import Service as FirefoxService
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

# 🔹 Set your Excel input path here
EXCEL_PATH = "/Users/raajthipparthy/Desktop/88georgia_input_foia.xlsx"

URL = "https://gov.georgia.gov/contact-us/open-records-request"

# ---- Fixed identity fields ----
FIRST_NAME = "Akhila"
LAST_NAME  = "Chennamaneni"
EMAIL      = "chikitha262@gmail.com"
PHONE      = "2035696249"

# IDs per your tags
ID_FIRST_NAME = "edit-first-name"
ID_LAST_NAME  = "edit-last-name"
ID_EMAIL      = "edit-email"
ID_PHONE      = "edit-phone-number"
ID_COMMENTS   = "edit-comments"
FORM_ID       = "webform-submission-webform-3656-node-15226-add-form"

WAIT_SECONDS = 10  # manual-click window per row

# ---------- Helpers ----------
def build_comment(event_title: str, event_id: str, gov_value: str) -> str:
    return (
        f"I am requesting a copy of the winning proposal for {event_title} "
        f"contract bearing ID {event_id} by {gov_value}."
    )

def normalize_columns(df: pd.DataFrame) -> pd.DataFrame:
    df.columns = [c.strip() for c in df.columns]
    lower_map = {c.lower(): c for c in df.columns}

    if "event id" not in lower_map or "event title" not in lower_map:
        raise ValueError("Excel must contain columns 'Event ID' and 'Event Title'.")

    if "government entity" in lower_map:
        gov_col = lower_map["government entity"]
    elif "government id" in lower_map:
        gov_col = lower_map["government id"]
    else:
        raise ValueError("Excel must contain 'Government Entity' or 'Government ID'.")

    out = pd.DataFrame()
    out["Event ID"]    = df[lower_map["event id"]].astype(str).str.strip()
    out["Event Title"] = df[lower_map["event title"]].astype(str).str.strip()
    out["GovVal"]      = df[gov_col].astype(str).str.strip()

    out = out[(out["Event ID"]!="") & (out["Event Title"]!="") & (out["GovVal"]!="")].reset_index(drop=True)
    return out

def wait_for(driver, by, value, timeout=25):
    return WebDriverWait(driver, timeout).until(EC.presence_of_element_located((by, value)))

def open_browser() -> webdriver.Firefox:
    opts = FirefoxOptions()
    opts.set_preference("dom.webnotifications.enabled", False)
    service = FirefoxService()
    driver = webdriver.Firefox(service=service, options=opts)
    try:
        driver.maximize_window()
    except Exception:
        pass
    return driver

def load_form(driver):
    driver.get(URL)
    wait_for(driver, By.ID, FORM_ID, timeout=25)

def scroll_into_center(driver, element):
    driver.execute_script(
        "arguments[0].scrollIntoView({behavior:'smooth', block:'center'});", element
    )
    time.sleep(0.25)

def clear_and_type(element, text: str):
    try:
        element.clear()
    except Exception:
        element.send_keys("\uE009" + "a")  # Cmd/Ctrl + A
        element.send_keys("\u0008")        # Backspace
    element.send_keys(text)

def fill_fields_scrolling_down(driver, comment_text: str):
    """
    Scroll to each field while filling.
    After filling Comments (last), DO NOT scroll anymore—leave page as-is.
    """

    # First Name
    first_el = wait_for(driver, By.ID, ID_FIRST_NAME, 20)
    scroll_into_center(driver, first_el)
    clear_and_type(first_el, FIRST_NAME)

    # Last Name
    last_el = wait_for(driver, By.ID, ID_LAST_NAME, 20)
    scroll_into_center(driver, last_el)
    clear_and_type(last_el, LAST_NAME)

    # Email
    email_el = wait_for(driver, By.ID, ID_EMAIL, 20)
    scroll_into_center(driver, email_el)
    clear_and_type(email_el, EMAIL)

    # Phone
    phone_el = wait_for(driver, By.ID, ID_PHONE, 20)
    scroll_into_center(driver, phone_el)
    clear_and_type(phone_el, PHONE)

    # Comments (LAST) — after this, DO NOT scroll anywhere
    comments_el = wait_for(driver, By.ID, ID_COMMENTS, 20)
    scroll_into_center(driver, comments_el)
    clear_and_type(comments_el, comment_text)

    # Blur to prevent any focus-induced jumps; do NOT scroll.
    driver.execute_script("if (document.activeElement) document.activeElement.blur();")

def manual_window_then_close(driver, seconds: int):
    # Stay exactly where we are; no scrolling here.
    for s in range(seconds, 0, -1):
        print(f"  You can click SUBMIT now… {s}s", end="\r")
        time.sleep(1)
    print()  # newline
    try:
        driver.quit()
    except Exception:
        pass
    time.sleep(0.5)

def main():
    try:
        df = pd.read_excel(EXCEL_PATH)
        df = normalize_columns(df)
    except Exception as e:
        print(f"Excel read/normalize failed: {e}")
        sys.exit(1)

    for i, row in df.iterrows():
        comment = build_comment(row["Event Title"], row["Event ID"], row["GovVal"])
        print(f"[{i+1}/{len(df)}] Preparing form for Event ID {row['Event ID']} …")

        driver = open_browser()
        try:
            load_form(driver)
            fill_fields_scrolling_down(driver, comment)
            manual_window_then_close(driver, WAIT_SECONDS)
        except Exception as e:
            print(f"  ✗ Error on row {i+1}: {e}")
            try:
                driver.quit()
            except Exception:
                pass
            time.sleep(0.5)

    print("✅ All rows processed (manual-submit mode, no scroll after Comments).")

if __name__ == "__main__":
    main()