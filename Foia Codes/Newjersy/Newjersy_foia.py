#!/usr/bin/env python3
# New_jersy_foia_fixed.py
# UC-safe options + slow pacing + WAF/403 handling + division AJAX waits + mirror retries
# + clean success detection for "Your confirmation number is ..." and 5-row cap.

import os
import sys
import time
import random
import traceback
import tempfile
import shutil
from pathlib import Path
import pandas as pd

from selenium import webdriver as se_webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver import ChromeOptions as StdChromeOptions
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from selenium.common.exceptions import TimeoutException, NoAlertPresentException

# Prefer undetected-chromedriver for Imperva
USE_UNDETECTED = True
try:
    import undetected_chromedriver as uc  # type: ignore
except Exception:
    uc = None
    USE_UNDETECTED = False

# webdriver_manager if available
try:
    from webdriver_manager.chrome import ChromeDriverManager
    HAVE_WDM = True
except Exception:
    HAVE_WDM = False

# ---------------------------
# Config
# ---------------------------

OPRA_PORTAL_URLS = [
    "https://www-njlib.nj.gov/NJ_OPRA/department.jsp",
    "https://www-njlib.nj.gov/NJ_OPRA/department",
    "https://www16.state.nj.us/NJ_OPRA/department.jsp",
    "https://www16.state.nj.us/NJ_OPRA/department",
]

FIRST_NAME = "Raaj"
LAST_NAME = "Thipparthy"
Email_Address = "raajnrao@gmail.com"
ADDRESS = "8181 Fannin St"
CITY = "Houston"
STATE_VALUE = "Texas"
ZIP_CODE = "77054"
MAX_AUTH_COST = "0.02"

DEPT_VALUE_TREASURY = "82"
DEPT_TEXT_TREASURY = "Treasury"
DIVISION_TEXT_TARGET = "Purchase and Property"
DIVISION_VALUE_SEGMENT = ":PUR:"

INPUT_XLSX = "/home/developer/Desktop/US-State-Foia-Scrapers/Scrapper Codes/Newjersy/njstart_bid_to_po_all.xlsx"
OUTPUT_XLSX = "opra_submit_results4.xlsx"

OUTPUT_COLUMNS = [
    "Bid Solicitation #",
    "Organization Name",
    "Contract #",
    "Buyer",
    "Description",
    "Bid Opening Date",
    "Status",
]

MAX_IFRAME_DEPTH = 6
WINDOW_SIZE = "1366,900"
PAGE_LOAD_TIMEOUT = 95
DRIVER_CREATE_RETRIES = 2

# Row cap (first N rows)
MAX_ROWS = int(os.getenv("MAX_ROWS", "5"))  # uses df.head(MAX_ROWS) [web:125]

# Global pacing (tunable via env)
SLOW_MODE_MULTIPLIER = float(os.getenv("SLOW_MODE_MULTIPLIER", "3.0"))
BASE_PAUSE_MIN = float(os.getenv("BASE_PAUSE_MIN", "0.6"))
BASE_PAUSE_MAX = float(os.getenv("BASE_PAUSE_MAX", "1.4"))
PER_ROW_MIN = float(os.getenv("PER_ROW_MIN", "6.0"))
PER_ROW_MAX = float(os.getenv("PER_ROW_MAX", "12.0"))
LONG_REST_EVERY = int(os.getenv("LONG_REST_EVERY", "3"))
LONG_REST_MIN = float(os.getenv("LONG_REST_MIN", "20.0"))
LONG_REST_MAX = float(os.getenv("LONG_REST_MAX", "40.0"))
WAF_BACKOFF_MIN = float(os.getenv("WAF_BACKOFF_MIN", "35.0"))
WAF_BACKOFF_MAX = float(os.getenv("WAF_BACKOFF_MAX", "75.0"))

# Optional simple proxy rotation via --proxy-server
PROXY_POOL = [p.strip() for p in os.getenv("PROXY_POOL", "").split(",") if p.strip()]

# ---------------------------
# Utility
# ---------------------------

def human_pause(min_s=None, max_s=None):
    if min_s is None: min_s = BASE_PAUSE_MIN
    if max_s is None: max_s = BASE_PAUSE_MAX
    time.sleep(random.uniform(max(min_s, BASE_PAUSE_MIN), max(max_s, BASE_PAUSE_MAX)) * SLOW_MODE_MULTIPLIER)

def debug_save_page(driver, fname="debug_page.html"):
    try:
        with open(fname, "w", encoding="utf-8") as f:
            f.write(driver.page_source or "")
        print(f"[debug] page saved to {fname}")
    except Exception as e:
        print("[debug] failed to save page source:", e)

def is_waf_block(driver):
    try:
        html = (driver.page_source or "").lower()
        return ("error 15" in html) or ("imperva" in html) or ("blocked" in html and "error" in html)
    except Exception:
        return False

def handle_alert_if_present(driver, timeout=5):
    try:
        WebDriverWait(driver, timeout).until(EC.alert_is_present())
        al = driver.switch_to.alert
        txt = (al.text or "").strip()
        print(f"[alert] {txt}")
        al.accept()
        human_pause(0.8, 1.6)
        return txt
    except Exception:
        return ""

def wait_document_ready(driver, timeout=40):
    WebDriverWait(driver, timeout).until(lambda d: d.execute_script("return document.readyState") in ("interactive", "complete"))

def accept_cookies_if_present(driver):
    try:
        xpaths = [
            "//button[contains(translate(., 'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'),'accept')]",
            "//button[contains(translate(., 'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'),'agree')]",
            "//*[contains(@class,'accept') and (self::button or self::a)]",
        ]
        for xp in xpaths:
            els = driver.find_elements(By.XPATH, xp)
            if els:
                driver.execute_script("arguments[0].scrollIntoView({block:'center'});", els[0])
                human_pause()
                try:
                    els[0].click()
                except Exception:
                    pass
                human_pause()
                break
    except Exception:
        pass

# ---------------------------
# Driver/session creation (UC-safe)
# ---------------------------

def random_proxy():
    return random.choice(PROXY_POOL) if PROXY_POOL else None

def make_chrome_options(user_data_dir, for_uc=False, proxy_url=None):
    if for_uc and uc is not None:
        opts = uc.ChromeOptions()
    else:
        opts = StdChromeOptions()

    opts.add_argument(f"--user-data-dir={user_data_dir}")
    opts.add_argument("--profile-directory=Default")
    opts.add_argument("--no-sandbox")
    opts.add_argument("--disable-dev-shm-usage")
    opts.add_argument(f"--window-size={WINDOW_SIZE}")
    opts.add_argument("--disable-blink-features=AutomationControlled")
    opts.add_argument("--user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/123.0.0.0 Safari/537.36")
    opts.page_load_strategy = "normal"
    if proxy_url:
        opts.add_argument(f"--proxy-server={proxy_url}")

    # Do NOT add excludeSwitches/useAutomationExtension for UC to avoid capability parse errors [web:98]
    if not for_uc:
        try:
            opts.add_experimental_option("excludeSwitches", ["enable-automation", "enable-logging"])
            opts.add_experimental_option("useAutomationExtension", False)
        except Exception:
            pass

    return opts

def create_driver(use_headless=False):
    tmp_dir = tempfile.mkdtemp(prefix="chrome_sess_")
    proxy = random_proxy()

    if USE_UNDETECTED and uc is not None:
        opts = make_chrome_options(tmp_dir, for_uc=True, proxy_url=proxy)
        if use_headless:
            opts.add_argument("--headless=new")
        driver = uc.Chrome(options=opts)
        driver.set_page_load_timeout(PAGE_LOAD_TIMEOUT)
        return driver, tmp_dir, proxy

    # vanilla fallback
    opts = make_chrome_options(tmp_dir, for_uc=False, proxy_url=proxy)
    if use_headless:
        opts.add_argument("--headless=new")
    service = None
    if HAVE_WDM:
        try:
            service = Service(ChromeDriverManager().install())
        except Exception:
            service = Service()
    else:
        service = Service()
    driver = se_webdriver.Chrome(service=service, options=opts)
    driver.set_page_load_timeout(PAGE_LOAD_TIMEOUT)
    # small stealth patch
    try:
        driver.execute_cdp_cmd("Page.addScriptToEvaluateOnNewDocument", {
            "source": "Object.defineProperty(navigator,'webdriver',{get:()=>undefined});"
        })
    except Exception:
        pass
    return driver, tmp_dir, proxy

def rebuild_session(old_driver, old_tmp):
    try:
        if old_driver:
            old_driver.quit()
    except Exception:
        pass
    try:
        if old_tmp and os.path.isdir(old_tmp):
            shutil.rmtree(old_tmp, ignore_errors=True)
    except Exception:
        pass
    return create_driver(use_headless=False)

# ---------------------------
# Frame helpers
# ---------------------------

def switch_to_default(driver):
    try:
        driver.switch_to.default_content()
    except Exception:
        pass

def enumerate_frame_paths(driver, max_depth=MAX_IFRAME_DEPTH):
    yield []
    def dfs(path, depth):
        if depth >= max_depth:
            return
        try:
            switch_to_default(driver)
            for idx in path:
                frames = driver.find_elements(By.TAG_NAME, "iframe")
                if idx >= len(frames):
                    return
                driver.switch_to.frame(frames[idx])
        except Exception:
            return
        try:
            frames = driver.find_elements(By.TAG_NAME, "iframe")
        except Exception:
            frames = []
        for i in range(len(frames)):
            new_path = path + [i]
            yield new_path
            for deeper in dfs(new_path, depth + 1):
                yield deeper
    for p in dfs([], 0):
        yield p

def switch_to_path(driver, path):
    switch_to_default(driver)
    for idx in path:
        frames = driver.find_elements(By.TAG_NAME, "iframe")
        driver.switch_to.frame(frames[idx])

def find_anywhere(driver, by, value, timeout=30):
    end = time.time() + timeout
    last_err = None
    while time.time() < end:
        try:
            switch_to_default(driver)
            el = WebDriverWait(driver, 2.0).until(EC.presence_of_element_located((by, value)))
            return el, []
        except Exception as e:
            last_err = e
        for path in enumerate_frame_paths(driver):
            try:
                switch_to_path(driver, path)
                el = WebDriverWait(driver, 1.2).until(EC.presence_of_element_located((by, value)))
                return el, path
            except Exception as e:
                last_err = e
        human_pause(0.5, 1.0)
    raise TimeoutError(f"Element not found: {(by, value)} ; last={last_err}")

# ---------------------------
# Navigation and WAF-aware open
# ---------------------------

def open_portal_try_urls(driver, urls=OPRA_PORTAL_URLS):
    last_exc = None
    for url in urls:
        try:
            driver.get(url)
            wait_document_ready(driver, timeout=35)
            human_pause(1.2, 2.5)
            accept_cookies_if_present(driver)
            human_pause(0.8, 1.6)
            if is_waf_block(driver):
                print("[warn] WAF page detected on load; cool-down then next mirror...")
                human_pause(WAF_BACKOFF_MIN, WAF_BACKOFF_MAX)
                continue
            return
        except Exception as e:
            last_exc = e
            print(f"[warn] opening {url} failed: {e}")
            human_pause(2.0, 4.0)
    if last_exc:
        raise last_exc

# ---------------------------
# Division population waits
# ---------------------------

def wait_select_has_option_css(driver, select_id, css_opt, timeout=30):
    WebDriverWait(driver, timeout).until(EC.presence_of_element_located((By.CSS_SELECTOR, f"#{select_id} {css_opt}")))

def wait_select_count_or_mutation(driver, select_el, min_count=2, timeout_sec=40):
    end = time.time() + timeout_sec
    while time.time() < end:
        try:
            cnt = int(driver.execute_script("return arguments[0].options ? arguments[0].options.length : 0;", select_el))
            if cnt >= min_count:
                return True
        except Exception:
            pass
        txt = handle_alert_if_present(driver, timeout=1)
        if "403" in txt:
            return False
        human_pause(0.6, 1.2)
    try:
        driver.set_script_timeout(timeout_sec + 5)
        return driver.execute_async_script("""
const sel = arguments[0]; const min = arguments[1]; const timeout = Math.max(500,(arguments[2]||35000)); const done = arguments[3];
function count(){try{return sel && sel.options ? sel.options.length : 0;}catch(e){return 0;}}
if (count() >= min) return done(true);
let resolved=false;
const obs = new MutationObserver(()=>{ if(count()>=min&&!resolved){resolved=true; try{obs.disconnect();}catch(e){} done(true);} });
try{ obs.observe(sel,{childList:true,subtree:true}); }catch(e){}
setTimeout(()=>{ if(!resolved){ try{obs.disconnect();}catch(e){} done(count()>=min);} }, timeout);
""", select_el, min_count, int(timeout_sec*1000))
    except Exception:
        return False

# ---------------------------
# Department & Division selection
# ---------------------------

def get_department_select(driver, timeout=45):
    try:
        el, path = find_anywhere(driver, By.ID, "departmentChoice", timeout=timeout//2)
        switch_to_path(driver, path)
        return Select(el), path
    except Exception:
        pass
    end = time.time() + timeout
    while time.time() < end:
        for path in enumerate_frame_paths(driver):
            try:
                switch_to_path(driver, path)
                selects = driver.find_elements(By.TAG_NAME, "select")
                for s in selects:
                    for o in s.find_elements(By.TAG_NAME, "option"):
                        txt = (o.text or "").strip()
                        val = (o.get_attribute("value") or "").strip()
                        if txt == DEPT_TEXT_TREASURY or val == DEPT_VALUE_TREASURY:
                            return Select(s), path
            except Exception:
                continue
        human_pause(0.5, 1.0)
    raise TimeoutError("Department select not found")

def get_division_select(driver, timeout=60):
    try:
        el, path = find_anywhere(driver, By.ID, "divisionChoice", timeout=timeout//2)
        switch_to_path(driver, path)
        return Select(el), path
    except Exception:
        pass
    end = time.time() + timeout
    while time.time() < end:
        for path in enumerate_frame_paths(driver):
            try:
                switch_to_path(driver, path)
                selects = driver.find_elements(By.TAG_NAME, "select")
                for s in selects:
                    for o in s.find_elements(By.TAG_NAME, "option"):
                        txt = (o.text or "").strip()
                        val = (o.get_attribute("value") or "").strip()
                        if (DIVISION_TEXT_TARGET.lower() in txt.lower()) or (DIVISION_VALUE_SEGMENT in val):
                            return Select(s), path
            except Exception:
                continue
        human_pause(0.5, 1.0)
    raise TimeoutError("Division select not found")

def select_department_treasury(driver):
    sel, path = get_department_select(driver, timeout=45)
    switch_to_path(driver, path)
    human_pause(0.8, 1.6)
    try:
        sel.select_by_value(DEPT_VALUE_TREASURY)
    except Exception:
        sel.select_by_visible_text(DEPT_TEXT_TREASURY)
    try:
        dept_el = driver.find_element(By.ID, "departmentChoice")
        driver.execute_script("""
arguments[0].dispatchEvent(new Event('input', {bubbles:true}));
arguments[0].dispatchEvent(new Event('change', {bubbles:true}));
arguments[0].dispatchEvent(new Event('blur', {bubbles:true}));
if (window.jQuery) { window.jQuery(arguments[0]).trigger('change'); }
""", dept_el)
    except Exception:
        pass
    human_pause(1.5, 2.8)

def select_division_purchase_and_property(driver):
    sel, path = get_division_select(driver, timeout=60)
    switch_to_path(driver, path)
    try:
        wait_select_has_option_css(driver, "divisionChoice", "option[value]:not([value=''])", timeout=15)
    except Exception:
        pass
    try:
        base_el = driver.find_element(By.ID, "divisionChoice")
    except Exception:
        base_el = sel._el
    if not wait_select_count_or_mutation(driver, base_el, min_count=2, timeout_sec=40):
        print("[warn] division not populated (placeholder only)")
        for o in sel.options:
            print("  opt:", (o.text or "").strip(), (o.get_attribute("value") or "").strip())
        raise RuntimeError("Division list not populated (likely WAF 403)")

    preferred = ["Division of Purchase and Property", "Purchase & Property", "Purchase and Property"]
    chosen = False
    for t in preferred:
        try:
            sel.select_by_visible_text(t)
            chosen = True
            break
        except Exception:
            continue
    if not chosen:
        for o in sel.options:
            txt = (o.text or "").strip()
            val = (o.get_attribute("value") or "").strip()
            if ("purchase" in txt.lower() and "property" in txt.lower()) or (":PUR:" in val) or ("PUR" in val.upper()) or ("DPP" in val.upper()):
                try:
                    sel.select_by_visible_text(txt)
                except Exception:
                    sel.select_by_value(val)
                chosen = True
                break
    if not chosen:
        for o in sel.options:
            print("  opt:", (o.text or "").strip(), (o.get_attribute("value") or "").strip())
        raise RuntimeError("Could not select Purchase and Property")

    try:
        driver.execute_script("""
arguments[0].dispatchEvent(new Event('input', {bubbles:true}));
arguments[0].dispatchEvent(new Event('change', {bubbles:true}));
arguments[0].dispatchEvent(new Event('blur', {bubbles:true}));
if (window.jQuery) { window.jQuery(arguments[0]).trigger('change'); }
""", base_el)
    except Exception:
        pass
    human_pause(1.0, 2.0)

def wait_for_state_request_form(driver, timeout=55):
    end = time.time() + timeout
    while time.time() < end:
        if is_waf_block(driver):
            raise RuntimeError("WAF blocked while waiting for form")
        try:
            el, path = find_anywhere(driver, By.ID, "first", timeout=4)
            switch_to_path(driver, path)
            human_pause(0.6, 1.2)
            return
        except Exception:
            _ = handle_alert_if_present(driver, timeout=1)
            human_pause(0.6, 1.0)
    raise TimeoutError("Form did not appear (first-name field not found)")

# ---------------------------
# Form & submission
# ---------------------------

def fill_request_form(driver, bid_solicitation_value, timeout=35):
    wait = WebDriverWait(driver, timeout)
    f = wait.until(EC.element_to_be_clickable((By.ID, "first"))); f.clear(); human_pause(); f.send_keys(FIRST_NAME); human_pause()
    l = driver.find_element(By.ID, "last"); l.clear(); human_pause(); l.send_keys(LAST_NAME); human_pause()
    e = driver.find_element(By.ID, "email"); e.clear(); human_pause(); e.send_keys(Email_Address); human_pause()
    a = driver.find_element(By.ID, "address"); a.clear(); human_pause(); a.send_keys(ADDRESS); human_pause()
    c = driver.find_element(By.ID, "city"); c.clear(); human_pause(); c.send_keys(CITY); human_pause()
    try:
        st = Select(driver.find_element(By.ID, "state")); human_pause()
        try:
            st.select_by_value(STATE_VALUE)
        except Exception:
            try:
                st.select_by_visible_text(STATE_VALUE)
            except Exception:
                st.select_by_visible_text("Texas")
        human_pause()
    except Exception:
        human_pause()
    z = driver.find_element(By.ID, "zip"); z.clear(); human_pause(); z.send_keys(ZIP_CODE); human_pause()

    radios = [
        [(By.ID, "convictedRadio2"), (By.CSS_SELECTOR, "input[name='convictedRadioOptions'][type='radio'][value='N']")],
        [(By.ID, "commercialRadio2"), (By.CSS_SELECTOR, "input[name='commercial'][type='radio'][value='N']")],
        [(By.ID, "legal_proceedingRadio2"), (By.CSS_SELECTOR, "input[name='legal_proceeding'][type='radio'][value='N']")],
        [(By.ID, "paymentRadio1"), (By.CSS_SELECTOR, "input[name='paymentRadioOptions'][type='radio'][value='a']")],
    ]
    for group in radios:
        for locator in group:
            try:
                el = WebDriverWait(driver, 8).until(EC.element_to_be_clickable(locator))
                driver.execute_script("arguments[0].scrollIntoView({block:'center'});", el)
                human_pause(); el.click(); human_pause(); break
            except Exception:
                continue

    m = driver.find_element(By.ID, "maxCost"); m.clear(); human_pause(); m.send_keys(MAX_AUTH_COST); human_pause(); m.send_keys(Keys.TAB); human_pause()
    message = ("I am requesting a copy of the winning and shortlisted proposals. " f"The solicitation/contract number is {bid_solicitation_value}.")
    msg = driver.find_element(By.ID, "message"); msg.clear(); human_pause(); msg.send_keys(message); human_pause()

def click_accept_if_present(driver, timeout=15):
    try:
        btn = WebDriverWait(driver, timeout).until(EC.element_to_be_clickable((By.XPATH, "//input[@name='btnSubmit' and (@value='Accept' or contains(., 'Accept'))]")))
        driver.execute_script("arguments[0].scrollIntoView({block:'center'});", btn); human_pause(); btn.click(); human_pause(1.0, 2.0); return True
    except Exception:
        return False

def submit_request(driver, timeout=30):
    btn = WebDriverWait(driver, timeout).until(EC.element_to_be_clickable((By.XPATH, "//input[@name='submit' and @type='submit']")))
    driver.execute_script("arguments[0].scrollIntoView({block:'center'});", btn); human_pause(); btn.click(); human_pause(1.2, 2.4)
    _ = click_accept_if_present(driver, timeout=10); human_pause(1.2, 2.4)
    txt = handle_alert_if_present(driver, timeout=2)
    if "403" in txt:
        return False
    return True

def detect_submission_outcome(driver, timeout=40):
    """
    Clean success signal: prefer the confirmation line, otherwise fall back
    to generic success, then to errors/WAF, else timeout fail.
    """
    if is_waf_block(driver):
        return ("Fail", "WAF blocked (Error 15)")
    end = time.time() + timeout
    while time.time() < end:
        _ = handle_alert_if_present(driver, timeout=1)
        # Primary: confirmation line like "Your confirmation number is W241356"
        try:
            el = driver.find_element(By.XPATH, "//*[contains(text(),'Your confirmation number is')]")
            txt = (el.text or "").strip()
            if txt:
                return ("Success", txt)  # exact, readable success signal [web:151]
        except Exception:
            pass
        # Secondary: generic success banners/messages
        for xp in [
            "//*[contains(@class,'alert') and contains(@class,'success')]",
            "//*[contains(translate(., 'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'),'thank') and contains(translate(., 'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'),'request')]",
            "//*[contains(translate(., 'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'),'confirmation')]",
        ]:
            els = driver.find_elements(By.XPATH, xp)
            if els:
                return ("Success", (els[0].text or "").strip() or "Submitted successfully")
        human_pause(0.6, 1.0)
    # Errors
    for xp in [
        "//*[contains(@class,'alert') and contains(@class,'danger')]",
        "//*[contains(translate(., 'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'),'error')]",
        "//*[contains(translate(., 'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'),'failed')]",
        "//*[contains(translate(., 'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'),'invalid')]",
    ]:
        els = driver.find_elements(By.XPATH, xp)
        if els:
            return ("Fail", (els[0].text or "").strip() or "Submission failed")
    if is_waf_block(driver):
        return ("Fail", "WAF blocked (Error 15)")
    return ("Fail", "No confirmation detected after submit")

# ---------------------------
# Data helpers
# ---------------------------

def ensure_output_columns(df_in):
    out = pd.DataFrame()
    for col in OUTPUT_COLUMNS:
        out[col] = (pd.Series(dtype="string") if col == "Status" else (df_in[col] if col in df_in.columns else ""))
    return out

def get_bid_solicitation_value(row):
    if "Bid Solicitation #" in row and pd.notna(row["Bid Solicitation #"]) and str(row["Bid Solicitation #"]).strip():
        return str(row["Bid Solicitation #"]).strip()
    for alt in ["Bid Solicitation", "Solicitation #", "Solicitation", "Bid #", "Bid ID"]:
        if alt in row and pd.notna(row[alt]) and str(row[alt]).strip():
            return str(row[alt]).strip()
    return ""

# ---------------------------
# Main
# ---------------------------

def main():
    if not Path(INPUT_XLSX).exists():
        print(f"Input file not found: {INPUT_XLSX}", file=sys.stderr)
        sys.exit(1)

    # Limit to first MAX_ROWS rows using DataFrame.head(n) [web:125]
    df = pd.read_excel(INPUT_XLSX).head(MAX_ROWS)
    out_df = ensure_output_columns(df)

    driver = None
    tmp_profile_dir = None

    try:
        driver, tmp_profile_dir, proxy = create_driver(use_headless=False)
        print("[info] session proxy:", proxy or "none")

        open_portal_try_urls(driver)
        if is_waf_block(driver):
            print("[warn] WAF on initial load; cool-down and rotate session...")
            human_pause(WAF_BACKOFF_MIN, WAF_BACKOFF_MAX)
            driver, tmp_profile_dir, proxy = rebuild_session(driver, tmp_profile_dir)
            print("[info] new session proxy:", proxy or "none")
            open_portal_try_urls(driver)

        select_department_treasury(driver)

        # up to 3 attempts to populate division in case of WAF/403
        for attempt in range(1, 4):
            try:
                select_division_purchase_and_property(driver)
                break
            except Exception as e:
                print(f"[warn] division attempt {attempt} failed: {e}")
                if attempt >= 3:
                    raise
                human_pause(WAF_BACKOFF_MIN, WAF_BACKOFF_MAX)
                driver, tmp_profile_dir, proxy = rebuild_session(driver, tmp_profile_dir)
                print("[info] rotated session proxy:", proxy or "none")
                open_portal_try_urls(driver)
                select_department_treasury(driver)

        wait_for_state_request_form(driver)

        statuses = []
        for idx, row in df.iterrows():
            human_pause(PER_ROW_MIN, PER_ROW_MAX)
            if (idx + 1) % LONG_REST_EVERY == 0:
                print("[info] periodic long rest...")
                human_pause(LONG_REST_MIN, LONG_REST_MAX)

            try:
                if is_waf_block(driver):
                    statuses.append("Fail: WAF blocked (Error 15)")
                    human_pause(WAF_BACKOFF_MIN, WAF_BACKOFF_MAX)
                    driver, tmp_profile_dir, proxy = rebuild_session(driver, tmp_profile_dir)
                    print("[info] rotated after WAF, proxy:", proxy or "none")
                    open_portal_try_urls(driver)
                    select_department_treasury(driver)
                    select_division_purchase_and_property(driver)
                    wait_for_state_request_form(driver)
                    continue

                try:
                    if not driver.find_elements(By.ID, "first"):
                        open_portal_try_urls(driver)
                        select_department_treasury(driver)
                        select_division_purchase_and_property(driver)
                        wait_for_state_request_form(driver)
                except Exception:
                    pass

                bid_value = get_bid_solicitation_value(row)
                if not bid_value:
                    statuses.append("Fail: Missing Bid Solicitation #")
                    continue

                fill_request_form(driver, bid_value)
                ok = submit_request(driver)
                if not ok and is_waf_block(driver):
                    statuses.append("Fail: WAF blocked (Error 15)")
                else:
                    outcome, msg = detect_submission_outcome(driver, timeout=40)
                    statuses.append(f"{outcome}: {msg}")

                human_pause(1.2, 2.2)
                try:
                    if not driver.find_elements(By.ID, "first"):
                        open_portal_try_urls(driver)
                        select_department_treasury(driver)
                        select_division_purchase_and_property(driver)
                        wait_for_state_request_form(driver)
                except Exception:
                    pass

            except Exception as e:
                statuses.append(f"Fail: {type(e).__name__} - {str(e)[:180]}")
                human_pause(WAF_BACKOFF_MIN, WAF_BACKOFF_MAX)
                try:
                    driver, tmp_profile_dir, proxy = rebuild_session(driver, tmp_profile_dir)
                    print("[info] rotated after exception, proxy:", proxy or "none")
                    open_portal_try_urls(driver)
                    select_department_treasury(driver)
                    select_division_purchase_and_property(driver)
                    wait_for_state_request_form(driver)
                except Exception:
                    pass

        out_df["Status"] = statuses
        out_df.to_excel(OUTPUT_XLSX, index=False)
        print(f"[ok] Saved results to {OUTPUT_XLSX}")
        return

    except Exception as e:
        print("[error] exception during run:", e)
        traceback.print_exc()
        if driver is not None:
            debug_save_page(driver, "debug_initial_page.html")
    finally:
        try:
            if driver is not None:
                driver.quit()
        except Exception:
            pass
        try:
            if tmp_profile_dir and os.path.isdir(tmp_profile_dir):
                if os.getenv("KEEP_TMP_PROFILE"):
                    print("[info] keeping tmp profile dir:", tmp_profile_dir)
                else:
                    shutil.rmtree(tmp_profile_dir, ignore_errors=True)
        except Exception:
            pass

    print("Run ended with errors.", file=sys.stderr)
    sys.exit(1)

if __name__ == "__main__":
    main()
