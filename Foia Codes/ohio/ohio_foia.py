"""
Requirements:
- Chrome browser + matching chromedriver on PATH
- pip install selenium pandas openpyxl
"""

import time
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException, ElementClickInterceptedException


INPUT_XLSX = "ohio_buys_awarded.xlsx"   
OUTPUT_XLSX = "ohio_buys_awarded_results2.xlsx"
FORM_URL = "https://das.ohio.gov/help-center/public-records-request-site-area/public-records-requests"
EMAIL_TO_USE = "raajnrao@gmail.com"
DEPARTMENT_VALUE = "78"  

REQUEST_TEMPLATE = (
    "I am requesting a copy of the winning and shortlisted proposals for {name} by {agency}."
    " I am requesting a copy of the winning and shortlisted proposals for the referenced award."
    " The solicitation/contract number is {id}."
)


PAGE_LOAD_TIMEOUT = 30
ELEMENT_TIMEOUT = 20


df = pd.read_excel(INPUT_XLSX, engine="openpyxl")


possible_id_cols = ["Solicitation ID", "SolicitationID", "Solicitation_Id", "Solicitation Num", "id"]
possible_name_cols = ["Solicitation Name", "SolicitationName", "Solicitation_Name", "Title", "name"]

def find_column(df, candidates):
    for c in candidates:
        if c in df.columns:
            return c
    return None

id_col = find_column(df, possible_id_cols)
name_col = find_column(df, possible_name_cols)

if id_col is None or name_col is None:
    raise SystemExit(f"Could not find Solicitation ID or Solicitation Name columns. "
                     f"Found columns: {list(df.columns)}.\nPlease ensure the input file has those columns.")


df["success"] = ""  


options = webdriver.ChromeOptions()

options.add_argument("--start-maximized")
driver = webdriver.Chrome(options=options)
driver.set_page_load_timeout(PAGE_LOAD_TIMEOUT)
wait = WebDriverWait(driver, ELEMENT_TIMEOUT)

def safe_find(by, value, timeout=ELEMENT_TIMEOUT):
    try:
        return WebDriverWait(driver, timeout).until(EC.presence_of_element_located((by, value)))
    except TimeoutException:
        return None

def safe_click(el):
    try:
        el.click()
        return True
    except (ElementClickInterceptedException, Exception) as e:
        try:
            driver.execute_script("arguments[0].click();", el)
            return True
        except Exception:
            return False


for idx, row in df.iterrows():
    solic_id = str(row[id_col]).strip()
    solic_name = str(row[name_col]).strip()
    print(f"\n=== Row {idx} — ID: {solic_id} | Name: {solic_name} ===")

    try:
        driver.get(FORM_URL)
    except Exception as e:
        df.at[idx, "success"] = f"failed: page load error: {e}"
        print(df.at[idx, "success"])
        continue

    
    time.sleep(1)  

    
    form = safe_find(By.ID, "prrform", timeout=5)
    if form is None:
        
        iframe_found = False
        iframes = driver.find_elements(By.TAG_NAME, "iframe")
        for ifr in iframes:
            try:
                driver.switch_to.frame(ifr)
                inside_form = driver.find_elements(By.ID, "prrform")
                if inside_form:
                    form = inside_form[0]
                    iframe_found = True
                    print("Switched into iframe containing the form.")
                    break
                driver.switch_to.default_content()
            except Exception:
                driver.switch_to.default_content()
                continue
        if not iframe_found and form is None:
            
            try:
                form = WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.ID, "prrform")))
            except TimeoutException:
                pass

    if form is None:
        df.at[idx, "success"] = "failed: form not found on page"
        print(df.at[idx, "success"])
        
        driver.switch_to.default_content()
        continue

    
    try:
        sel = Select(driver.find_element(By.ID, "Request_SelectedClient"))
        
        try:
            sel.select_by_value(DEPARTMENT_VALUE)
        except Exception:
            sel.select_by_visible_text("Unknown")
    except Exception as e:
        
        try:
            sel = Select(driver.find_element(By.NAME, "Request.SelectedClient"))
            sel.select_by_value(DEPARTMENT_VALUE)
        except Exception as e2:
            df.at[idx, "success"] = f"failed: department select error: {e} / {e2}"
            print(df.at[idx, "success"])
            driver.switch_to.default_content()
            continue

    
    try:
        email_radio = driver.find_element(By.ID, "method-email")
        safe_click(email_radio)
    except Exception:
        
        try:
            radios = driver.find_elements(By.CSS_SELECTOR, "input[name='Request.DeliveryMethodKey']")
            for r in radios:
                if r.get_attribute("value") == "email":
                    safe_click(r)
                    break
        except Exception as e:
            print("warning: couldn't click email radio:", e)

    
    try:
        email_el = driver.find_element(By.ID, "Request_RequestorEmail")
        email_el.clear()
        email_el.send_keys(EMAIL_TO_USE)
    except Exception as e:
        
        try:
            email_el = driver.find_element(By.NAME, "Request.RequestorEmail")
            email_el.clear()
            email_el.send_keys(EMAIL_TO_USE)
        except Exception as e2:
            df.at[idx, "success"] = f"failed: email field error: {e} / {e2}"
            print(df.at[idx, "success"])
            driver.switch_to.default_content()
            continue

    
    agency_col_candidates = ["Agency", "agency", "Agency Name", "Agency_Name"]
    agency = ""
    for c in agency_col_candidates:
        if c in df.columns:
            agency = str(row[c]).strip()
            break
    
    if agent := (agency if agency else ""):
        pass

    request_text = REQUEST_TEMPLATE.format(id=solic_id, name=solic_name, agency=(agency or "[Agency]"))

    
    try:
        req_el = driver.find_element(By.ID, "Request_Request")
        req_el.clear()
        req_el.send_keys(request_text)
    except Exception as e:
        try:
            req_el = driver.find_element(By.NAME, "Request.Request")
            req_el.clear()
            req_el.send_keys(request_text)
        except Exception as e2:
            df.at[idx, "success"] = f"failed: request textarea error: {e} / {e2}"
            print(df.at[idx, "success"])
            driver.switch_to.default_content()
            continue

    print("Form filled for this row. Now you must complete the reCAPTCHA manually in the browser.")
    print("When you have completed the captcha and are ready to submit this form, return to this terminal and press ENTER.")
    input("Press ENTER here AFTER you finish the captcha in the browser for this row...")

    
    try:
        submit_btn = driver.find_element(By.CSS_SELECTOR, "input[type='submit'], button[type='submit']")
        
        try:
            driver.execute_script("arguments[0].removeAttribute('disabled')", submit_btn)
        except Exception:
            pass
        safe_click(submit_btn)
    except Exception as e:
        df.at[idx, "success"] = f"failed: could not find or click submit button: {e}"
        print(df.at[idx, "success"])
        driver.switch_to.default_content()
        continue

    
    time.sleep(2)
    success = False
    reason = ""

    try:
        
        val_err = driver.find_elements(By.CSS_SELECTOR, ".validation-summary-errors, .field-validation-error, .validation-summary-valid")
        error_detected = False
        for el in val_err:
            txt = el.text.strip()
            if txt:
                
                if "error" in txt.lower() or "required" in txt.lower() or "please" in txt.lower():
                    error_detected = True
                    reason = txt
                    break
        if error_detected:
            success = False
        else:
            
            body_text = driver.find_element(By.TAG_NAME, "body").text.lower()
            if any(w in body_text for w in ["thank you", "thank", "request has been submitted", "submission received", "we have received"]):
                success = True
            else:
                
                if driver.current_url != FORM_URL and driver.current_url.strip() != "":
                    success = True
                else:
                    
                    success = False
                    reason = "no confirmation text detected"
    except Exception as e:
        success = False
        reason = f"error checking result: {e}"

    if success:
        df.at[idx, "success"] = "success: form submitted"
        print("Submission detected as SUCCESS.")
    else:
        df.at[idx, "success"] = f"fail: {reason}"
        print(f"Submission detected as FAIL ({reason}).")

    
    try:
        driver.switch_to.default_content()
    except Exception:
        pass

    
    time.sleep(1)


df.to_excel(OUTPUT_XLSX, index=False)
print("\nAll rows processed. Results saved to:", OUTPUT_XLSX)


driver.quit()
