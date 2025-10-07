import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.firefox.service import Service as FirefoxService
from selenium.webdriver.firefox.options import Options as FirefoxOptions
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC

GPR_URL = "https://ssl.doas.state.ga.us/gpr/"

def scroll_into_view(driver, el):
    driver.execute_script("arguments[0].scrollIntoView({block:'center'});", el)
    # small pause so the scroll completes smoothly
    time.sleep(0.25)

def click_when_clickable(wait, driver, by, selector):
    el = wait.until(EC.element_to_be_clickable((by, selector)))
    scroll_into_view(driver, el)
    el.click()
    return el

def main():
    # --- Firefox setup (visible window so you can see it work) ---
    options = FirefoxOptions()
    # options.add_argument("-headless")  # leave commented to watch it run
    # options.add_argument("-kiosk")     # optional: true fullscreen, ESC to exit
    service = FirefoxService()           # uses geckodriver from PATH

    driver = webdriver.Firefox(service=service, options=options)

    # Maximize window (good default); uncomment kiosk above if you want true fullscreen
    try:
        driver.maximize_window()
    except Exception:
        # Fallback if window manager denies maximize in some environments
        driver.set_window_size(1440, 900)

    wait = WebDriverWait(driver, 30)

    print("[1/6] Opening GPR…")
    driver.get(GPR_URL)

    # Scroll to top to normalize viewport
    driver.execute_script("window.scrollTo(0,0);")

    # --- Ensure the main Search panel is present ---
    print("[2/6] Waiting for the basic search form to be visible…")
    wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "div.page.search")))
    wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "div#search.collapse.show")))
    wait.until(EC.presence_of_element_located((By.ID, "eventSearchForm")))

    # --- Select Event Status = AWARDED ---
    print("[3/6] Selecting Event Status → AWARDED…")
    event_status_select_el = wait.until(
        EC.element_to_be_clickable((By.CSS_SELECTOR, "select#eventStatus.custom-select.input.input-select"))
    )
    scroll_into_view(driver, event_status_select_el)
    Select(event_status_select_el).select_by_value("AWARD")  # value per your page description

    # --- Ensure Advanced Search is visible (eventProcessType lives here) ---
    print("[4/6] Ensuring Advanced Search section is open…")
    try:
        adv = driver.find_element(By.CSS_SELECTOR, "div#advSearch")
        if "show" not in adv.get_attribute("class"):
            # Try common toggles to open it
            for how, sel in [
                (By.CSS_SELECTOR, "[data-target='#advSearch']"),
                (By.CSS_SELECTOR, "[href='#advSearch']"),
                (By.LINK_TEXT, "Advanced Search"),
                (By.PARTIAL_LINK_TEXT, "Advanced"),
            ]:
                try:
                    btn = driver.find_element(how, sel)
                    scroll_into_view(driver, btn)
                    btn.click()
                    # wait for 'show' class to appear
                    wait.until(lambda d: "show" in d.find_element(By.CSS_SELECTOR, "div#advSearch").get_attribute("class"))
                    break
                except Exception:
                    continue
    except Exception:
        # If not found, we’ll proceed — select by ID below still works if already visible.
        pass

    # --- Select Event Process Type = RFP ---
    print("[5/6] Selecting Event Process Type → RFP…")
    event_proc_select_el = wait.until(
        EC.element_to_be_clickable((By.CSS_SELECTOR, "select#eventProcessType.custom-select.input.input-select"))
    )
    scroll_into_view(driver, event_proc_select_el)
    Select(event_proc_select_el).select_by_value("RFP")  # "Request for Proposal"

    # --- Click Search ---
    print("[6/6] Clicking Search…")
    search_btn = wait.until(
        EC.element_to_be_clickable((By.CSS_SELECTOR, "button#eventSearchButton.button.button--primary"))
    )
    scroll_into_view(driver, search_btn)
    search_btn.click()

    # OPTIONAL: wait for results area to change; we’ll add robust scraping later
    time.sleep(2)
    print("✅ Filters applied and Search clicked. Browser will remain open for inspection.")
    print("   Close the Firefox window manually when you’re done.")

    # Keep the session alive indefinitely (or for a long while).
    # Press Ctrl+C in the terminal to stop the script if needed.
    while True:
        time.sleep(60)

if __name__ == "__main__":
    main()