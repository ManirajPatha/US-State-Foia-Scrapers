#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
Massachusetts AGO (Hyland/OnBase) Public Records Request auto-submitter

- Opens the AGO PRR form for each row in an Excel sheet of closed opportunities
- Fills required fields, checks Declaration, enters Signature, and submits
- Logs per-row status to an .xlsx file

REQUIREMENTS:
  pip install selenium webdriver-manager pandas openpyxl

NOTE on .xls:
  Modern pandas + xlrd>=2.0.1 cannot read legacy .xls files.
  Save your input as .xlsx, then run this script.
"""

import os
import time
from datetime import datetime
from typing import Dict, Any, List

import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service as ChromeService
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import (
    TimeoutException, NoSuchElementException, ElementClickInterceptedException
)
from webdriver_manager.chrome import ChromeDriverManager

# ================== EDIT THESE ==================
FORM_URL   = "https://massago.hylandcloud.com/231appnet/UnityForm.aspx?d1=AUpLDUslaJ6TqrT%2bXV9GB2AQKvMy13t1lHnLrMAuBjMYNwxrk9NYiiKWGZl8bzVE9sRo8J6ZgPXRVOKWhn9myaHZ%2bZjyGoVKzfnwWX6jf9qPME63EoMx58D0g3X5Kx8wsKjJD58fPwBbjMWKKa35m3xixNXSGO0TgtqVJFsXdFgzeLiCCVsLhLUvv54VV6EtNZwGH4um8ZRsBdS1Vw5cbGxMPObSN9Jxe3EFl49P%2fi0WnaxGF936iFNTERkEJMWI9wbwsrwvMAxsX5wofnBaac4%3d"
INPUT_PATH = r"C:\Users\MANIRAJ\OneDrive\Documents\Scrapper\Scrapper-Request\Massachusettes\commbuys_closed_20250923_040247.xlsx"  # <-- use .xlsx
OUT_LOG    = r"C:\Users\MANIRAJ\OneDrive\Documents\Scrapper\Scrapper-Request\Massachusettes\ago_submit_log.xlsx"
WAIT_SECS  = 25

# Fixed requester info (your details)
REQUESTER = {
    "first_name": "Maniraj",
    "last_name": "Patha",
    "organization": "Southern Arkansas University",
    "address": "8181 Fannin St",
    "city": "Houston",
    "state": "Texas",
    "zip": "77054",
    "email": "pathamaniraj97@gmail.com",
    "phone": "6824055734",
    "signature": "Maniraj Patha",
}
# ================================================


def _normalize_out_path(path: str) -> str:
    """Allow OUT_LOG to be a directory or a full file path."""
    if path.lower().endswith(".xlsx"):
        return path
    # If a folder was provided, append a filename.
    if os.path.isdir(path) or path.endswith(os.sep):
        os.makedirs(path, exist_ok=True)
        return os.path.join(path, "ago_submit_log.xlsx")
    # If extension missing, append .xlsx
    return path + ".xlsx"


# -------------------- Excel ---------------------
def read_rows(path: str) -> List[Dict[str, Any]]:
    if not os.path.exists(path):
        raise FileNotFoundError(f"Input file not found: {path}")
    ext = os.path.splitext(path.lower())[1]
    if ext == ".xlsx":
        df = pd.read_excel(path, engine="openpyxl")
    elif ext == ".xls":
        # Modern pandas + xlrd>=2.0.1 cannot read .xls. Bail out with a clear message.
        raise SystemExit(
            "Your input is '.xls'. Please open it and save as '.xlsx', then update INPUT_PATH.\n"
            "Reason: pandas + xlrd>=2.0.1 no longer supports legacy .xls files."
        )
    else:
        raise SystemExit("Please provide an Excel file with extension .xlsx (recommended).")
    return df.fillna("").to_dict(orient="records")


# ---------------- Selenium helpers --------------
def switch_to_form_context(driver) -> None:
    """Stay in top doc if header is there; else probe iframes."""
    try:
        WebDriverWait(driver, 5).until(
            EC.presence_of_element_located(
                (By.XPATH, "//*[contains(translate(., 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'public records request')]")
            )
        )
        return
    except TimeoutException:
        pass
    iframes = driver.find_elements(By.TAG_NAME, "iframe")
    for frame in iframes:
        driver.switch_to.default_content()
        driver.switch_to.frame(frame)
        try:
            WebDriverWait(driver, 3).until(
                EC.presence_of_element_located(
                    (By.XPATH, "//*[contains(translate(., 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'public records request')]")
                )
            )
            return
        except TimeoutException:
            continue
    driver.switch_to.default_content()
    raise TimeoutException("Could not locate the AGO form in top document or any iframe.")


def get_control_by_label(driver, label_text: str, want: str = "input"):
    """
    Resolve control by label text. Try <label for=id> → element; then nearby element.
    want: 'input' | 'textarea' | 'select' | 'any'
    """
    label = None
    for xp in [
        f"//label[contains(normalize-space(.), '{label_text}')]",
        f"//*[label[contains(normalize-space(.), '{label_text}')]]/label[contains(normalize-space(.), '{label_text}')]",
    ]:
        found = driver.find_elements(By.XPATH, xp)
        if found:
            label = found[0]
            break
    if label is not None:
        for_attr = label.get_attribute("for")
        if for_attr:
            for tag in (["input", "textarea", "select"] if want == "any" else [want]):
                try:
                    el = driver.find_element(By.XPATH, f"//{tag}[@id='{for_attr}']")
                    return el, tag
                except NoSuchElementException:
                    pass
        # fallbacks near label
        for xp in [
            "./following::*[self::input or self::textarea or self::select][1]",
            "../*[self::input or self::textarea or self::select][1]",
            "ancestor::*[self::div or self::td or self::tr][1]//*[self::input or self::textarea or self::select][1]",
        ]:
            try:
                el = label.find_element(By.XPATH, xp)
                return el, el.tag_name.lower()
            except Exception:
                pass
    # catch-all
    for xp in [
        f"//label[contains(normalize-space(.), '{label_text}')]/following::input[1]",
        f"//label[contains(normalize-space(.), '{label_text}')]/following::textarea[1]",
        f"//label[contains(normalize-space(.), '{label_text}')]/following::select[1]",
        f"//*[label[contains(normalize-space(.), '{label_text}')]]//input[1]",
        f"//*[label[contains(normalize-space(.), '{label_text}')]]//textarea[1]",
        f"//*[label[contains(normalize-space(.), '{label_text}')]]//select[1]",
    ]:
        els = driver.find_elements(By.XPATH, xp)
        if els:
            el = els[0]
            return el, el.tag_name.lower()
    raise NoSuchElementException(f"Control for label '{label_text}' not found.")


def set_text(driver, el, value: str) -> None:
    driver.execute_script("arguments[0].scrollIntoView({block:'center'});", el)
    try:
        el.clear()
    except Exception:
        pass
    driver.execute_script("arguments[0].value = arguments[1];", el, value)
    try:
        el.click()
    except Exception:
        pass


def set_state(driver, value: str) -> None:
    el, tag = get_control_by_label(driver, "State", want="any")
    driver.execute_script("arguments[0].scrollIntoView({block:'center'});", el)
    if tag == "select":
        sel = Select(el)
        tried = [value, value.upper()[:2], value[:2].upper()]
        for v in tried:
            try:
                sel.select_by_visible_text(v)
                return
            except Exception:
                for opt in sel.options:
                    t = (opt.text or "").strip()
                    if v.lower() == t.lower() or v.lower() in t.lower():
                        sel.select_by_visible_text(opt.text)
                        return
        driver.execute_script("arguments[0].value = arguments[1];", el, value)
    else:
        set_text(driver, el, value)


def click_checkbox_by_label(driver, label_text: str) -> None:
    for xp in [
        f"//label[contains(normalize-space(.), '{label_text}')]/preceding::input[@type='checkbox'][1]",
        f"//*[label[contains(normalize-space(.), '{label_text}')]]//input[@type='checkbox'][1]",
        f"//input[@type='checkbox'][ancestor::*[label[contains(normalize-space(.), '{label_text}')]]][1]",
    ]:
        boxes = driver.find_elements(By.XPATH, xp)
        if boxes:
            cb = boxes[0]
            driver.execute_script("arguments[0].scrollIntoView({block:'center'});", cb)
            if not cb.is_selected():
                driver.execute_script("arguments[0].click();", cb)
            return
    raise NoSuchElementException(f"Checkbox '{label_text}' not found.")


def click_submit(driver) -> None:
    for xp in [
        "//input[@type='submit' and ( @value='Submit' or @title='Submit')]",
        "//button[normalize-space(.)='Submit']",
        "//*[self::a or self::span or self::div][normalize-space(.)='Submit']",
    ]:
        btns = driver.find_elements(By.XPATH, xp)
        if btns:
            btn = btns[0]
            driver.execute_script("arguments[0].scrollIntoView({block:'center'});", btn)
            try:
                btn.click()
            except ElementClickInterceptedException:
                driver.execute_script("arguments[0].click();", btn)
            return
    raise NoSuchElementException("Submit button not found.")


# --------------- Fill one submission --------------
def build_description(row: Dict[str, Any]) -> str:
    """
    Build a clear, records-focused request for the specific (closed) opportunity
    using whatever columns exist in the input sheet. Keeps under 3000 chars.
    """

    def pick(keys: list[str]) -> str:
        for k in keys:
            v = row.get(k, "")
            if isinstance(v, str):
                v = v.strip()
            if v not in (None, "", "nan"):
                return str(v)
        return ""

    # Try common column names (robust to different exports)
    title        = pick(["Description", "Title", "Name", "Bid Title"])
    bid_no       = pick(["Bid Solicitation #", "Solicitation #", "Bid #", "Number", "Solicitation Number"])
    alt_id       = pick(["Alternate Id", "Alt ID", "Reference #"])
    agency       = pick(["Organization Name", "Agency", "Department", "Org"])
    category     = pick(["Category", "Commodity", "Type"])
    buyer        = pick(["Buyer", "Contact", "Contact Name"])
    open_date    = pick(["Open Date", "Bid Opening Date", "Opening Date"])
    close_date   = pick(["Close Date", "Closing Date", "Bid Closing Date"])
    award_date   = pick(["Award Date", "Date Awarded"])
    awardee      = pick(["Award Vendor", "Awarded Vendor", "Vendor", "Supplier"])
    contract_no  = pick(["Contract #", "Contract Number", "PO #"])

    # Header line
    lines = []
    lines.append("Public Records Request – Closed Opportunity")

    # Identity block (only include if present)
    identity_bits = []
    if title:       identity_bits.append(f"Title: {title}")
    if bid_no:      identity_bits.append(f"Bid/Solicitation #: {bid_no}")
    if alt_id:      identity_bits.append(f"Alternate ID: {alt_id}")
    if contract_no: identity_bits.append(f"Contract #: {contract_no}")
    if agency:      identity_bits.append(f"Agency/Organization: {agency}")
    if category:    identity_bits.append(f"Category: {category}")
    if buyer:       identity_bits.append(f"Buyer/Contact: {buyer}")
    if open_date:   identity_bits.append(f"Open Date: {open_date}")
    if close_date:  identity_bits.append(f"Close/Deadline: {close_date}")
    if award_date:  identity_bits.append(f"Award Date: {award_date}")
    if awardee:     identity_bits.append(f"Awarded Vendor: {awardee}")

    if identity_bits:
        lines.append("\n".join(identity_bits))

    # Timeframe hint (prefer close/award year if present)
    def extract_year(s: str) -> str:
        import re
        m = re.search(r"(20\d{2}|19\d{2})", s or "")
        return m.group(1) if m else ""

    yr = extract_year(close_date) or extract_year(award_date) or extract_year(open_date)
    timeframe = f"calendar year {yr}" if yr else "the full life of the solicitation and award"
    lines.append(f"\nTimeframe requested: {timeframe}.")

    # Exactly what records you want (kept consistent)
    lines.append(
        "\nPlease provide the complete procurement/award file, including:\n"
        "- the solicitation and all addenda\n"
        "- bidders list\n"
        "- technical and cost evaluations\n"
        "- evaluation summary / award recommendation\n"
        "- award notice / memo\n"
        "- executed contract(s) and any amendments or change orders\n"
        "\nElectronic copies preferred (PDF for documents; CSV/XLSX for any tabular data). "
        "If estimated costs exceed $50, please notify me before proceeding."
    )

    # Join and enforce 3000-char limit
    text = "\n".join([ln for ln in lines if ln.strip() != ""]).strip()
    if len(text) > 3000:
        text = text[:2995].rstrip() + "…"
    return text


def submit_one(driver, row: Dict[str, Any]) -> None:
    driver.get(FORM_URL)
    WebDriverWait(driver, WAIT_SECS).until(EC.presence_of_element_located((By.TAG_NAME, "body")))
    switch_to_form_context(driver)

    set_text(driver, get_control_by_label(driver, "First Name", want="input")[0], REQUESTER["first_name"])
    set_text(driver, get_control_by_label(driver, "Last Name", want="input")[0], REQUESTER["last_name"])

    # Organization might be optional
    try:
        set_text(driver, get_control_by_label(driver, "Organization", want="input")[0], REQUESTER["organization"])
    except Exception:
        pass

    set_text(driver, get_control_by_label(driver, "Address", want="input")[0], REQUESTER["address"])
    set_text(driver, get_control_by_label(driver, "City", want="input")[0], REQUESTER["city"])
    set_state(driver, REQUESTER["state"])
    set_text(driver, get_control_by_label(driver, "Zip", want="input")[0], REQUESTER["zip"])
    set_text(driver, get_control_by_label(driver, "Email Address", want="input")[0], REQUESTER["email"])
    set_text(driver, get_control_by_label(driver, "Phone Number", want="input")[0], REQUESTER["phone"])

    # Description (Required)
    desc_el, _ = get_control_by_label(driver, "Please describe the records you are requesting", want="any")
    set_text(driver, desc_el, build_description(row))

    # Declaration (Required)
    click_checkbox_by_label(driver, "Declaration")

    # Signature (Required)
    set_text(driver, get_control_by_label(driver, "Signature", want="input")[0], REQUESTER["signature"])

    # Submit
    click_submit(driver)
    time.sleep(2)


# ---------------------- Main ----------------------
def main():
    if "UFKey" in FORM_URL:
        raise SystemExit("FORM_URL still looks like a placeholder. Paste the FULL link from Mass.gov (long string).")

    out_path = _normalize_out_path(OUT_LOG)
    rows = read_rows(INPUT_PATH)

    options = webdriver.ChromeOptions()
    options.add_argument("--start-maximized")
    # options.add_argument("--headless=new")  # use visible first; headless later if stable
    driver = webdriver.Chrome(service=ChromeService(ChromeDriverManager().install()), options=options)

    results = []
    try:
        for idx, row in enumerate(rows, start=1):
            rec = {"row_index": idx, "sent_at": datetime.now().isoformat(timespec="seconds")}
            try:
                submit_one(driver, row)
                rec["status"] = "SUBMITTED"
                print(f"[OK] Row {idx} submitted")
            except Exception as e:
                rec["status"] = f"ERROR: {e}"
                print(f"[ERR] Row {idx}: {e}")
                # capture debug artifacts
                os.makedirs("debug_shots", exist_ok=True)
                os.makedirs("debug_html", exist_ok=True)
                ts = datetime.now().strftime("%Y%m%d_%H%M%S")
                try:
                    driver.save_screenshot(os.path.join("debug_shots", f"row_{idx}_error_{ts}.png"))
                except Exception:
                    pass
                try:
                    with open(os.path.join("debug_html", f"row_{idx}_error_{ts}.html"), "w", encoding="utf-8") as f:
                        f.write(driver.page_source)
                except Exception:
                    pass
            rec.update(row)
            results.append(rec)
            time.sleep(1.2)
    finally:
        driver.quit()

    os.makedirs(os.path.dirname(out_path) or ".", exist_ok=True)
    pd.DataFrame(results).to_excel(out_path, index=False)
    print(f"\nOutput log written: {os.path.abspath(out_path)}")


if __name__ == "__main__":
    main()
