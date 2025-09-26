#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Iowa NextRequest — robust & faster filler
- Targets only VISIBLE inputs (skips hidden duplicates)
- Fast JS set + input/change/blur + verification + fallback typing
- Handles TinyMCE/Quill/textarea description
- Hard-override Email, Full name, State
- Minimal sleeps; reduced waits
"""

import argparse, io, os, shutil, tempfile, time, logging
from typing import Dict, Any, List, Optional

import pandas as pd
from selenium import webdriver
from selenium.webdriver.firefox.service import Service as FirefoxService
from webdriver_manager.firefox import GeckoDriverManager
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException, ElementClickInterceptedException, WebDriverException

# ====== EASY PATHS ======
INPUT_PATH_DEFAULT  = r"C:\Users\MANIRAJ\OneDrive\Documents\Scrapper\Scrapper-Request\lowa\Iowa_Request_Descriptions.xlsx"
OUTPUT_PATH_DEFAULT = r"C:\Users\MANIRAJ\OneDrive\Documents\Scrapper\Scrapper-Request\lowa\Results.xlsx"
HEADLESS_DEFAULT    = "false"
# ========================

FORM_URL     = "https://iowaopenrecords.nextrequest.com/requests/new"
DEPT_DEFAULT = "No Department - General Information Requests"

# Forced values (always used)
FORCE_EMAIL = "pathamaniraj97@gmail.com"
FORCE_NAME  = "Maniraj Patha"
FORCE_STATE = "Texas"   # fallback "TX" if needed

DEFAULT_PROFILE = {
    "Department": DEPT_DEFAULT,
    "Email": FORCE_EMAIL,
    "Name": FORCE_NAME,
    "Phone": "+1 6824055734",
    "Street address": "8181 Fannin st",
    "City": "Houston",
    "State": FORCE_STATE,
    "Zip": "77054",
    "Company": "Southern Arkansas University",
}

HEADER_ALIASES = {
    "request description":"Request description","description":"Request description","request_details":"Request description",
    "department":"Department","dept":"Department",
    "email":"Email","e-mail":"Email","mail":"Email",
    "name":"Name","full name":"Name","applicant name":"Name",
    "phone":"Phone","phone number":"Phone","contact number":"Phone","mobile":"Phone",
    "street address":"Street address","address":"Street address","address line 1":"Street address","addr1":"Street address",
    "city":"City","state":"State","zip":"Zip","zip code":"Zip","postal code":"Zip",
    "company":"Company","organization":"Company","organisation":"Company","org":"Company",
}
EXPECTED_COLUMNS = ["Request description","Department","Email","Name","Phone","Street address","City","State","Zip","Company"]

# ---------- Excel helpers ----------
def _canonicalize_headers(df: pd.DataFrame) -> pd.DataFrame:
    df.columns = [HEADER_ALIASES.get(str(c).strip().lower(), c) for c in df.columns]
    return df

def _ensure_xlsx_path(path: str) -> str:
    root, ext = os.path.splitext(path)
    return path if ext.strip() else (root + ".xlsx")

def _safe_read_excel(path: str) -> pd.DataFrame:
    # copy-then-read to dodge OneDrive locks
    tmpdir = tempfile.mkdtemp(prefix="iowa_")
    tmp = os.path.join(tmpdir, os.path.basename(path))
    shutil.copyfile(path, tmp)
    return pd.read_excel(tmp, engine="openpyxl")

def read_rows(path: str) -> List[Dict[str, Any]]:
    df = _safe_read_excel(path).fillna("")
    df = _canonicalize_headers(df)
    if "Request description" not in df.columns:
        raise ValueError("Input must include a 'Request description' column.")
    for col in EXPECTED_COLUMNS:
        if col not in df.columns:
            df[col] = DEFAULT_PROFILE.get(col, "")
    # fill empties from defaults
    for k, v in DEFAULT_PROFILE.items():
        df[k] = df[k].apply(lambda x: v if str(x).strip()=="" else x)
    # forced values
    df["Email"] = FORCE_EMAIL
    df["Name"]  = FORCE_NAME
    df["State"] = FORCE_STATE
    df["Department"] = df["Department"].apply(lambda v: v if str(v).strip() else DEPT_DEFAULT)
    remaining = [c for c in df.columns if c not in EXPECTED_COLUMNS]
    df = df[EXPECTED_COLUMNS + remaining]
    return df.to_dict(orient="records")

# ---------- Selenium helpers ----------
def vis_xpath(driver, xpath: str):
    """Return first VISIBLE element matching xpath (skips hidden templates)."""
    els = driver.find_elements(By.XPATH, xpath)
    for el in els:
        try:
            visible = driver.execute_script(
                "const e=arguments[0];return !!(e.offsetParent!==null && getComputedStyle(e).visibility!=='hidden' && getComputedStyle(e).display!=='none');",
                el
            )
            if visible:
                return el
        except Exception:
            pass
    return None

def scroll_into_view(driver, el):
    try: driver.execute_script("arguments[0].scrollIntoView({block:'center'});", el)
    except Exception: pass

def js_set_value(driver, el, value: str):
    driver.execute_script("""
        const el=arguments[0], val=arguments[1];
        const d = Object.getOwnPropertyDescriptor(el.__proto__||HTMLElement.prototype,'value')
                 ||Object.getOwnPropertyDescriptor(HTMLInputElement.prototype,'value')
                 ||Object.getOwnPropertyDescriptor(HTMLTextAreaElement.prototype,'value');
        if(d&&d.set) d.set.call(el,val); else el.value=val;
        el.dispatchEvent(new Event('input',{bubbles:true}));
        el.dispatchEvent(new Event('change',{bubbles:true}));
        el.dispatchEvent(new Event('blur',{bubbles:true}));
    """, el, value)

def set_contenteditable_text(driver, el, text: str):
    driver.execute_script("""
        const el=arguments[0], t=arguments[1];
        el.innerHTML=''; el.textContent=t;
        el.dispatchEvent(new Event('input',{bubbles:true}));
        el.dispatchEvent(new Event('keyup',{bubbles:true}));
        el.dispatchEvent(new Event('change',{bubbles:true}));
        el.dispatchEvent(new Event('blur',{bubbles:true}));
    """, el, text)

def fill_input_verified(driver, wait, label: Optional[str], hints: List[str], value: str, required: bool=True) -> bool:
    """
    Fast, robust fill for text inputs:
    - find the FIRST *visible* candidate by label, then by name/id/placeholder/aria-label
    - set via JS + events, verify value, fallback to focus+type if needed
    """
    xps = []
    if label:
        xps.append(f"//label[contains(normalize-space(.), '{label}')]/following::*[self::input or self::textarea][1]")
    for h in hints:
        hlow = h.lower()
        xps += [
            f"//*[self::input or self::textarea][contains(translate(@name,'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'), '{hlow}')]",
            f"//*[self::input or self::textarea][contains(translate(@id,'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'), '{hlow}')]",
            f"//*[self::input or self::textarea][contains(translate(@placeholder,'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'), '{hlow}')]",
            f"//*[self::input or self::textarea][contains(translate(@aria-label,'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'), '{hlow}')]",
        ]

    el = None
    for xp in xps:
        el = vis_xpath(driver, xp)
        if el: break

    if not el:
        if required: raise RuntimeError(f"Field not found: {label or '/'.join(hints)}")
        return False

    wait.until(EC.visibility_of(el))
    scroll_into_view(driver, el)

    # JS set (fast) + verify
    try:
        js_set_value(driver, el, value)
        actual = driver.execute_script("return arguments[0].value;", el)
        if (actual or "").strip() == str(value).strip():
            return True
    except Exception:
        pass

    # Fallback: focus and type quickly, then blur
    try:
        el.click()
    except Exception:
        driver.execute_script("arguments[0].focus();", el)
    try:
        el.clear()
    except Exception:
        pass
    el.send_keys(value)
    driver.execute_script("arguments[0].dispatchEvent(new Event('blur',{bubbles:true}));", el)
    actual = driver.execute_script("return arguments[0].value;", el)
    if (actual or "").strip() == str(value).strip():
        return True

    if required:
        raise RuntimeError(f"Failed to fill: {label or '/'.join(hints)}")
    return False

# ---------- Description filler (iframe/Quill/textarea) ----------
def fill_request_description(driver, wait, text: str) -> bool:
    # TinyMCE iframe directly under label
    for xp in [
        "//label[contains(normalize-space(.),'Request description')]/following::iframe[1]",
        "//iframe[contains(@class,'tox-edit-area__iframe') or contains(@id,'tinymce') or contains(@title,'Rich Text')]",
    ]:
        iframe = vis_xpath(driver, xp)
        if iframe:
            scroll_into_view(driver, iframe)
            driver.switch_to.frame(iframe)
            try:
                body = vis_xpath(driver, "//body")
                if body:
                    driver.execute_script("arguments[0].innerHTML='';", body)
                    body.click()
                    body.send_keys(text)
                    driver.execute_script("""
                        const b=arguments[0], t=arguments[1];
                        b.innerText=t;
                        b.dispatchEvent(new Event('input',{bubbles:true}));
                        b.dispatchEvent(new Event('change',{bubbles:true}));
                        b.dispatchEvent(new Event('blur',{bubbles:true}));
                    """, body, text)
                    return True
            finally:
                driver.switch_to.default_content()

    # Quill/contenteditable
    for xp in [
        "//label[contains(.,'Request description')]/following::*[contains(@class,'ql-editor') or @contenteditable='true' or @role='textbox'][1]",
        "//*[contains(@class,'ql-editor') or @contenteditable='true' or @role='textbox']"
    ]:
        ed = vis_xpath(driver, xp)
        if ed:
            scroll_into_view(driver, ed)
            try: ed.click()
            except Exception: driver.execute_script("arguments[0].focus();", ed)
            set_contenteditable_text(driver, ed, text)
            return True

    # Plain textarea
    ta = vis_xpath(driver, "//label[contains(normalize-space(.),'Request description')]/following::textarea[1]")
    if ta:
        scroll_into_view(driver, ta)
        js_set_value(driver, ta, text)
        return True

    # Last chance: any description-ish input
    return fill_input_verified(driver, wait, "Request description", ["description","request"], text, required=False)

# ---------- Select helpers ----------
def select_by_label_or_combo(driver, wait, label_text: str, target_text: str, fallback: Optional[str]=None) -> bool:
    # native select
    sel = vis_xpath(driver, f"//label[contains(.,'{label_text}')]/following::select[1]")
    if sel:
        try:
            Select(sel).select_by_visible_text(target_text)
            return True
        except Exception:
            if fallback:
                try:
                    Select(sel).select_by_visible_text(fallback)
                    return True
                except Exception:
                    pass

    # custom combobox (vue-select)
    combo = None
    for xp in [
        f"//label[contains(.,'{label_text}')]/following::*[@role='combobox'][1]",
        f"//*[contains(@aria-label,'{label_text}')][@role='combobox']",
        f"//label[contains(.,'{label_text}')]/following::*[contains(@class,'vs__dropdown-toggle') or contains(@class,'select')][1]"
    ]:
        combo = vis_xpath(driver, xp)
        if combo: break

    if combo:
        scroll_into_view(driver, combo)
        try: combo.click()
        except ElementClickInterceptedException: driver.execute_script("arguments[0].click();", combo)

        # search input inside combo
        si = None
        for sx in [".//input[contains(@class,'vs__search') or @role='searchbox' or @type='search']",
                   ".//input"]:
            try:
                si = combo.find_element(By.XPATH, sx)
                if si.is_displayed():
                    break
            except NoSuchElementException:
                continue
        if si:
            si.clear()
            si.send_keys(target_text)
            time.sleep(0.2)
            si.send_keys(Keys.ENTER)
            return True

        # clickable option fallback
        for text_try in [target_text, fallback] if fallback else [target_text]:
            if not text_try: continue
            for ox in [
                f"//div[@role='option' or @role='listbox']//*[normalize-space(text())='{text_try}']",
                f"//li//*[normalize-space(text())='{text_try}']",
                f"//*[contains(@class,'option')][normalize-space(text())='{text_try}']",
            ]:
                opt = vis_xpath(driver, ox)
                if opt:
                    try: opt.click()
                    except ElementClickInterceptedException: driver.execute_script("arguments[0].click();", opt)
                    return True

    return False

def select_department(driver, wait, desired: str) -> bool:
    if select_by_label_or_combo(driver, wait, "Department", desired):
        return True
    # known label exact text
    return select_by_label_or_combo(driver, wait, "Department", "No Department - General Information Requests")

def click_submit(driver, wait) -> bool:
    for xp in ["//button[normalize-space()='Make a request']",
               "//button[normalize-space()='Send message']",
               "//button[@type='submit']","//input[@type='submit']"]:
        try:
            btn = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, xp)))
            scroll_into_view(driver, btn)
            try: btn.click()
            except ElementClickInterceptedException: driver.execute_script("arguments[0].click();", btn)
            return True
        except TimeoutException:
            continue
    return False

def fill_form(driver, wait, row: Dict[str, Any]) -> None:
    # Description
    if not fill_request_description(driver, wait, row["Request description"]):
        raise RuntimeError("Failed to fill Request description.")

    # Department
    if not select_department(driver, wait, row.get("Department", DEPT_DEFAULT)):
        raise RuntimeError("Failed to set Department.")

    # Email (forced)
    fill_input_verified(driver, wait, "Email", ["email","e-mail"], FORCE_EMAIL)

    # Name (forced)
    if not fill_input_verified(driver, wait, "Name", ["name","fullname","full_name"], FORCE_NAME, required=False):
        # split fallback
        parts = FORCE_NAME.split(" ", 1)
        fill_input_verified(driver, wait, "First Name", ["first"], parts[0], required=False)
        fill_input_verified(driver, wait, "Last Name",  ["last"],  parts[1] if len(parts)>1 else "", required=False)

    # Phone / Address
    fill_input_verified(driver, wait, "Phone", ["phone","mobile"], str(row["Phone"]), required=False)
    fill_input_verified(driver, wait, "Street address", ["street","address"], str(row["Street address"]), required=False)
    fill_input_verified(driver, wait, "City", ["city"], str(row["City"]), required=False)

    # State (forced Texas, fallback TX)
    if not select_by_label_or_combo(driver, wait, "State", FORCE_STATE, fallback="TX"):
        # if it's a text input instead of select
        fill_input_verified(driver, wait, "State", ["state"], FORCE_STATE, required=False)

    fill_input_verified(driver, wait, "Zip", ["zip","postal"], str(row["Zip"]), required=False)
    if not fill_input_verified(driver, wait, "Company", ["company","organization","organisation","org"], str(row["Company"]), required=False):
        pass

def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--input", default=INPUT_PATH_DEFAULT)
    parser.add_argument("--out",   default=OUTPUT_PATH_DEFAULT)
    parser.add_argument("--headless", default=HEADLESS_DEFAULT, choices=["true","false"])
    args = parser.parse_args()
    args.out = _ensure_xlsx_path(args.out)

    logging.basicConfig(level=logging.INFO, format="%(asctime)s [%(levelname)s] %(message)s")
    logging.info(f"Reading: {args.input}")
    rows = read_rows(args.input)
    logging.info(f"Rows: {len(rows)}")

    opts = webdriver.FirefoxOptions()
    if args.headless.lower() == "true": opts.add_argument("-headless")
    service = FirefoxService(GeckoDriverManager().install())
    driver  = webdriver.Firefox(service=service, options=opts)
    wait    = WebDriverWait(driver, 20)

    results = []
    try:
        for i, row in enumerate(rows, 1):
            logging.info(f"[{i}/{len(rows)}] Open form")
            driver.get(FORM_URL)
            try: wait.until(EC.presence_of_element_located((By.XPATH, "//form")))
            except TimeoutException: pass

            try:
                fill_form(driver, wait, row)
                submitted = click_submit(driver, wait)
                status = "submitted" if submitted else "submit_failed_or_captcha"
                confirm_url = driver.current_url
            except Exception as e:
                status = f"error: {e}"
                confirm_url = driver.current_url

            results.append({**row, "status": status, "confirmation_url": confirm_url})
            # small pause so UI can reset
            time.sleep(0.3)
    finally:
        driver.quit()

    pd.DataFrame(results).to_excel(args.out, index=False)
    print(f"Results -> {args.out}")

if __name__ == "__main__":
    main()
