# indiana_apra_automation.py — Direct sign-in with returnUrl -> auto-fill form

import argparse, os, time, logging
from typing import Optional, Dict, Any, List
import pandas as pd
from urllib.parse import quote_plus

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.firefox.options import Options as FirefoxOptions
from selenium.common.exceptions import TimeoutException, NoSuchElementException, ElementClickInterceptedException, ElementNotInteractableException

# --------------------------
# MANUAL CONFIG — EDIT HERE
# --------------------------
# This is the ONLY base URL we hit first. We just add ?returnUrl=... to it.
LOGIN_URL_BASE = "https://access.in.gov/client/signin/"

# Where to land after login (the actual APRA form)
TARGET_FORM_URL = "https://in.accessgov.com/indiana/Forms/Edit/oed-apra/9baa8344-c35c-46dd-bd06-5e64ede3075d/2"

# Credentials
MANUAL_LOGIN_EMAIL    = r"pathamaniraj97@gmail.com"   # e.g., "you@example.com"
MANUAL_LOGIN_PASSWORD = r"Indianapassword@12345"   # e.g., "YourPassword"

# Paths (if output is a folder, we create a timestamped xlsx inside)
MANUAL_INPUT_PATH  = r"C:\Users\MANIRAJ\OneDrive\Documents\Scrapper\Scrapper-Request\Indiana\indiana_contracts_20250920_1423.xlsx"
MANUAL_OUTPUT_PATH = r"C:\Users\MANIRAJ\OneDrive\Documents\Scrapper\Scrapper-Request\Indiana"

# Auto text template if a row lacks 'request_text'
DEFAULT_REQUEST_TEMPLATE = (
    "Under APRA, please provide all records related to {title} (Contract ID {contract_id}) "
    "between {start_date} and {end_date}, held by {agency_name}. "
    "Include contract, amendments, invoices, POs, and correspondence with {vendor_name}. "
    "Deliver electronically; non-commercial request."
)

# --------------------------
# Logging
# --------------------------
logging.basicConfig(level=logging.INFO, format="%(asctime)s [%(levelname)s] %(message)s", datefmt="-%Y-%m-%d %H:%M:%S")
log = logging.getLogger("indiana-apra")

class _SafeDict(dict):
    def __missing__(self, key): return ""

# --------------------------
# Selenium helpers
# --------------------------
def new_driver(headless: bool = False) -> webdriver.Firefox:
    opts = FirefoxOptions(); opts.page_load_strategy = "eager"
    if headless: opts.add_argument("-headless")
    d = webdriver.Firefox(options=opts); d.set_window_size(1400, 1000)
    return d

def _scroll_into_view(d, el): d.execute_script("arguments[0].scrollIntoView({block:'center'});", el)
def _safe_click(d, el):
    _scroll_into_view(d, el)
    try: el.click()
    except (ElementClickInterceptedException, ElementNotInteractableException):
        d.execute_script("arguments[0].click();", el)

def find_first_present(d, xps: List[str], timeout=10):
    end = time.time()+timeout; last=None
    while time.time()<end:
        for xp in xps:
            try: return d.find_element(By.XPATH, xp)
            except NoSuchElementException as e: last=e
        time.sleep(0.15)
    if last: raise last
    raise TimeoutException("Element not found")

def input_by_label_contains(d, label_text: str, value: str, timeout=20):
    try:
        label = WebDriverWait(d, timeout).until(EC.presence_of_element_located((By.XPATH, f"//label[contains(normalize-space(.), '{label_text}')]")))
        target_id = label.get_attribute("for")
        if target_id:
            el = d.find_element(By.ID, target_id); _set_value(d, el, value); return
    except TimeoutException: pass
    for tag in ("input","textarea","select"):
        try:
            el = WebDriverWait(d, timeout).until(EC.presence_of_element_located((By.XPATH, f"//label[contains(normalize-space(.), '{label_text}')]/following::*[self::{tag}][1]")))
            _set_value(d, el, value); return
        except TimeoutException: continue
    cont = WebDriverWait(d, timeout).until(EC.presence_of_element_located((By.XPATH, f"(//*[contains(normalize-space(.), '{label_text}')])[1]")))
    for tag in ("input","textarea","select"):
        try:
            el = cont.find_element(By.XPATH, f".//{tag}"); _set_value(d, el, value); return
        except NoSuchElementException: continue
    raise NoSuchElementException(f"Field not found for: {label_text}")

def _set_value(d, el, val: str):
    tag = el.tag_name.lower(); txt = "" if val is None else str(val)
    _scroll_into_view(d, el)
    if tag in ("input","textarea"):
        try: el.clear()
        except Exception: pass
        el.send_keys(txt)
    elif tag=="select":
        try: el.click()
        except Exception: pass
        for xp in (f".//option[normalize-space(.)='{txt}']",
                   f".//option[contains(translate(., 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), '{txt.lower()}')]"):
            try: _safe_click(d, el.find_element(By.XPATH, xp)); return
            except NoSuchElementException: continue
        raise NoSuchElementException(f"Option not found in select: {txt}")
    else:
        el.send_keys(txt)

def ensure_checkbox_checked_by_label_contains(d, label_text: str, timeout=20):
    try:
        label = WebDriverWait(d, timeout).until(EC.presence_of_element_located((By.XPATH, f"//label[contains(normalize-space(.), '{label_text}')]")))
        tid = label.get_attribute("for")
        if tid:
            cb = d.find_element(By.ID, tid); _check(d, cb); return
    except TimeoutException: pass
    try:
        cb = WebDriverWait(d, timeout).until(EC.presence_of_element_located((By.XPATH, f"//label[contains(normalize-space(.), '{label_text}')]/following::input[@type='checkbox'][1]")))
        _check(d, cb); return
    except TimeoutException: pass
    cont = WebDriverWait(d, timeout).until(EC.presence_of_element_located((By.XPATH, f"(//*[contains(normalize-space(.), '{label_text}')])[1]")))
    cb = cont.find_element(By.XPATH, ".//input[@type='checkbox']"); _check(d, cb)

def _check(d, el):
    _scroll_into_view(d, el)
    if not el.is_selected(): _safe_click(d, el)

def click_submit(d, timeout=20):
    for t in ("Submit","Submit Request","Send","Continue","Next"):
        try:
            btn = WebDriverWait(d, timeout).until(EC.element_to_be_clickable((By.XPATH, f"//button[normalize-space(.)='{t}' or contains(., '{t}')] | //input[@type='submit' and (@value='{t}' or contains(@value,'{t}'))]")))
            _safe_click(d, btn); return True
        except TimeoutException: continue
    try:
        btn = WebDriverWait(d, timeout).until(EC.element_to_be_clickable((By.XPATH, "//button[@type='submit' or contains(.,'Submit')] | //input[@type='submit']")))
        _safe_click(d, btn); return True
    except TimeoutException: return False

def capture_confirmation(d, timeout=20) -> Dict[str,Any]:
    info={}
    try:
        WebDriverWait(d, timeout).until(EC.presence_of_element_located((By.XPATH, "//*[contains(., 'Thank') or contains(., 'confirmation') or contains(., 'submitted') or contains(., 'Reference') or contains(., 'received')]")))
        info["status"]="submitted"
    except TimeoutException:
        info["status"]="unknown"
    try:
        ref = d.find_element(By.XPATH, "//*[contains(., 'Reference') or contains(., 'Confirmation') or contains(., 'Request #') or contains(., 'Request ID')]")
        info["confirmation_text"]=ref.text.strip()
    except Exception:
        info["confirmation_text"]=""
    info["final_url"]=d.current_url
    return info

def ensure_output_path(path:str)->str:
    if not path: return "indiana_apra_output.xlsx"
    if os.path.isdir(path):
        return os.path.join(path, f"indiana_apra_output_{time.strftime('%Y%m%d_%H%M%S')}.xlsx")
    if os.path.splitext(path)[1].lower()!=".xlsx": path += ".xlsx"
    return path

# --------------------------
# Detect form visibility / token redirect
# --------------------------
def is_form_visible(d)->bool:
    probes = [
        "//label[contains(.,'Full Name')]",
        "//label[contains(.,'Email Address')]",
        "//label[contains(.,'Request for Public Records')]",
    ]
    for xp in probes:
        try:
            el = d.find_element(By.XPATH, xp)
            if el.is_displayed(): return True
        except NoSuchElementException:
            continue
    return False

def looks_like_token_redirect(url: str)->bool:
    return "/redirect?token=" in url

# --------------------------
# Sign in with returnUrl
# --------------------------
def sign_in_and_land_on_form(d, login_base: str, target_form_url: str, email: str, password: str, wait_secs: int=30):
    # Build login URL with returnUrl so IdP takes us straight to the form
    login_url = f"{login_base}?returnUrl={quote_plus(target_form_url)}"
    d.get(login_url)
    time.sleep(0.8)

    # Email -> Confirm/Continue
    try:
        email_input = find_first_present(d, [
            "//input[@type='email']",
            "//label[contains(.,'Email')]/following::input[1]",
            "//input[contains(translate(@placeholder,'EMAIL','email'),'email')]",
        ], timeout=12)
        _set_value(d, email_input, email)
        confirm_btn = find_first_present(d, [
            "//button[normalize-space()='Confirm']",
            "//button[normalize-space()='Continue']",
            "//button[normalize-space()='Next']",
            "//input[@type='submit' and (@value='Confirm' or @value='Continue' or @value='Next')]",
            "//button[contains(.,'Confirm') or contains(.,'Continue') or contains(.,'Next')]",
        ], timeout=10)
        _safe_click(d, confirm_btn)
        time.sleep(1.0)
    except Exception:
        log.info("Email step not found; possibly already on password screen.")

    # Password -> Sign in
    try:
        pwd_input = find_first_present(d, [
            "//input[@type='password']",
            "//label[contains(.,'Password')]/following::input[1]"
        ], timeout=12)
        _set_value(d, pwd_input, password)
        signin_btn = find_first_present(d, [
            "//button[normalize-space()='Sign In']",
            "//button[normalize-space()='Sign in']",
            "//button[normalize-space()='Log In']",
            "//input[@type='submit' and (@value='Sign In' or @value='Sign in' or @value='Log In')]",
            "//button[contains(.,'Sign In') or contains(.,'Sign in') or contains(.,'Log In')]"
        ], timeout=10)
        _safe_click(d, signin_btn)
        time.sleep(1.2)
    except Exception:
        log.info("Password step not shown; session may already be active.")

    # Wait for redirect to the actual form
    WebDriverWait(d, wait_secs).until(lambda drv: is_form_visible(drv) or looks_like_token_redirect(drv.current_url))

    # If we landed on a token relay, explicitly open the form
    if looks_like_token_redirect(d.current_url) and not is_form_visible(d):
        d.get(target_form_url)
        WebDriverWait(d, 15).until(lambda drv: is_form_visible(drv))

# --------------------------
# Template helper
# --------------------------
def build_request_text(row: pd.Series, template: str) -> str:
    if "request_text" in row and str(row["request_text"]).strip():
        return str(row["request_text"]).strip()
    data = {k: "" if pd.isna(v) else str(v) for k,v in row.to_dict().items()}
    try: return template.format_map(_SafeDict(data))
    except Exception as e:
        logging.warning(f"Template format issue -> {e}")
        for k in ("title","contract_id","start_date","end_date","agency_name","vendor_name"):
            data.setdefault(k,"")
        return template.format_map(_SafeDict(data))

# --------------------------
# Main
# --------------------------
def run(
    login_base: str,
    target_form_url: str,
    login_email: str,
    login_password: str,
    input_excel: str,
    output_excel: str,
    start_row: int = 0,
    max_rows: Optional[int] = None,
    headless: bool = False,
    template: str = DEFAULT_REQUEST_TEMPLATE
):
    df = pd.read_excel(input_excel)
    df = df.iloc[start_row:(start_row+max_rows)] if max_rows is not None else df.iloc[start_row:]
    if df.empty:
        log.warning("No rows to process after start/max filters."); return

    output_excel = ensure_output_path(output_excel)

    results: List[Dict[str,Any]] = []
    d = new_driver(headless=headless)
    try:
        # Sign in and land directly on the form
        sign_in_and_land_on_form(d, login_base, target_form_url, login_email, login_password)

        for idx, row in df.iterrows():
            log.info(f"Submitting row {idx} ({row.get('full_name','') or row.get('vendor_name','')})")

            # Make sure we’re on a fresh form (if we’re on a confirmation page, try Back)
            if not is_form_visible(d):
                try: d.back(); time.sleep(0.6)
                except Exception: pass
                WebDriverWait(d, 10).until(lambda drv: is_form_visible(drv))

            def gv(name, default=""):
                return str(row[name]) if name in row and not pd.isna(row[name]) else default

            try:
                input_by_label_contains(d, "Full Name",     gv("full_name","Maniraj Patha"))
                input_by_label_contains(d, "Phone Number",  gv("phone","+1 6824055734"))
                input_by_label_contains(d, "Email Address", gv("email","pathamaniraj97@gmail.com"))
                input_by_label_contains(d, "Address",       gv("address","8181 Fannin St"))
                input_by_label_contains(d, "City",          gv("city","Houston"))
                input_by_label_contains(d, "State",         gv("state","Texas"))
                input_by_label_contains(d, "Zip",           gv("zip","77054"))

                req_text = build_request_text(row, template)
                input_by_label_contains(d, "Request for Public Records", req_text)

                ensure_checkbox_checked_by_label_contains(d, "I attest")
                time.sleep(0.2)

                if not click_submit(d):
                    raise RuntimeError("Could not find/click a Submit button.")

                info = capture_confirmation(d)
                info.update({
                    "row_index": int(idx),
                    "full_name": gv("full_name"),
                    "email": gv("email"),
                    "timestamp": pd.Timestamp.now().isoformat(timespec="seconds"),
                    "request_text_used": req_text
                })
                results.append(info)
                log.info(f"Row {idx} -> {info['status']}")
            except Exception as e:
                log.exception(f"Error on row {idx}: {e}")
                results.append({
                    "row_index": int(idx),
                    "full_name": row.get("full_name",""),
                    "email": row.get("email",""),
                    "status": "error",
                    "error": str(e),
                    "final_url": d.current_url,
                    "confirmation_text": "",
                    "request_text_used": build_request_text(row, template)
                })

        pd.DataFrame(results).to_excel(output_excel, index=False)
        log.info(f"Saved log to {output_excel}")
    finally:
        d.quit()

def parse_args(argv=None):
    p = argparse.ArgumentParser(description="Indiana APRA automation via direct sign-in with returnUrl")
    p.add_argument("--login-base", default=LOGIN_URL_BASE, help="Base sign-in URL (we append ?returnUrl=...)")
    p.add_argument("--target-form-url", default=TARGET_FORM_URL, help="Where to land after sign-in")

    default_email = MANUAL_LOGIN_EMAIL or os.getenv("LOGIN_EMAIL","")
    default_pass  = MANUAL_LOGIN_PASSWORD or os.getenv("LOGIN_PASSWORD","")
    p.add_argument("--login-email", default=default_email, help="Login email")
    p.add_argument("--login-password", default=default_pass, help="Login password")

    default_input  = MANUAL_INPUT_PATH if MANUAL_INPUT_PATH else None
    default_output = MANUAL_OUTPUT_PATH if MANUAL_OUTPUT_PATH else "indiana_apra_output.xlsx"
    if default_input: p.add_argument("--input", default=default_input, help="Path to input Excel (.xlsx)")
    else:             p.add_argument("--input", required=True, help="Path to input Excel (.xlsx)")
    p.add_argument("--output", default=default_output, help="Path to output Excel (file or folder)")
    p.add_argument("--start-row", type=int, default=0)
    p.add_argument("--max-rows", type=int, default=None)
    p.add_argument("--headless", action="store_true")
    p.add_argument("--template", default=DEFAULT_REQUEST_TEMPLATE)
    return p.parse_args(argv)

if __name__ == "__main__":
    a = parse_args()
    if not a.login_email or not a.login_password:
        log.info("Login email/password not provided; proceeding (session may already be valid).")
    out = ensure_output_path(a.output)
    run(
        login_base=a.login_base,
        target_form_url=a.target_form_url,
        login_email=a.login_email,
        login_password=a.login_password,
        input_excel=a.input,
        output_excel=out,
        start_row=a.start_row,
        max_rows=a.max_rows,
        headless=a.headless,
        template=a.template,
    )