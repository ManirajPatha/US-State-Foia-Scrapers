from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
import pandas as pd
import time
import os
import logging

logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

chrome_options = webdriver.ChromeOptions()
download_dir = os.path.join(os.getcwd(), "downloads")
if not os.path.exists(download_dir):
    os.makedirs(download_dir)
    os.chmod(download_dir, 0o755)

chrome_options.add_experimental_option("prefs", {
    "download.default_directory": download_dir,
    "download.prompt_for_download": False,
    "download.directory_upgrade": True,
    "safebrowsing.enabled": True
})

try:
    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=chrome_options)
    logging.info("WebDriver initialized successfully")
except Exception as e:
    logging.error(f"Failed to initialize WebDriver: {e}")
    raise

try:
    logging.info("Navigating to website")
    driver.get("https://postingboard.esmsolutions.com/3444a404-3818-494f-84c5-2a850acd7779/events")
    
    time.sleep(5)

    with open(os.path.join(download_dir, "initial_page_source.html"), "w") as f:
        f.write(driver.page_source)
    logging.info("Saved initial page source to downloads/initial_page_source.html")

    logging.info("Checking for human verification checkbox")
    try:
        verification_checkbox = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable(
                (By.XPATH, "//input[@type='checkbox' and (contains(@id, 'recaptcha') or contains(@class, 'recaptcha') or contains(@aria-label, 'robot') or contains(@aria-label, 'verify') or contains(@id, 'captcha') or contains(..//text(), 'verify') or contains(..//text(), 'robot'))] | "
                           "//div[contains(@class, 'recaptcha-checkbox')]//input | "
                           "//label[contains(text(), 'not a robot') or contains(text(), 'verify')]/preceding-sibling::input[@type='checkbox']")
            )
        )
        driver.execute_script("arguments[0].scrollIntoView(true);", verification_checkbox)
        driver.execute_script("arguments[0].click();", verification_checkbox)
        logging.info("Human verification checkbox clicked successfully")
        time.sleep(5)
    except:
        logging.info("No human verification checkbox found - proceeding")

    logging.info("Clicking 'Past Opportunities' tab")
    try:
        past_opportunities_tab = WebDriverWait(driver, 20).until(
            EC.element_to_be_clickable((By.XPATH, "//div[contains(text(), 'Past Opportunities')]"))
        )
        driver.execute_script("arguments[0].click();", past_opportunities_tab)
        logging.info("'Past Opportunities' tab clicked successfully")
        time.sleep(5)
    except Exception as e:
        logging.error(f"Failed to click 'Past Opportunities' tab: {e}")
        with open(os.path.join(download_dir, "post_tab_page_source.html"), "w") as f:
            f.write(driver.page_source)
        logging.info("Saved page source to downloads/post_tab_page_source.html")
        raise

    all_data = []
    page_count = 0

    while True:
        page_count += 1
        logging.info(f"Scraping page {page_count}")

        try:
            rows = WebDriverWait(driver, 20).until(
                EC.presence_of_all_elements_located((By.CSS_SELECTOR, "tr.mat-row"))
            )
        except Exception as e:
            logging.warning(f"No rows found on page {page_count}: {e}")
            break

        awarded_count = 0
        for row in rows:
            try:
                status_element = row.find_element(By.CSS_SELECTOR, "td.cdk-column-status div")
                status = status_element.text.strip()
                
                if status != "Awarded":
                    continue
                
                event_id = ""
                try:
                    event_id_element = row.find_element(By.CSS_SELECTOR, "td.cdk-column-id")
                    event_id = event_id_element.text.strip()
                except:
                    logging.debug("Event ID not found")

                event_name = ""
                try:
                    event_name_element = row.find_element(By.CSS_SELECTOR, "td.cdk-column-eventName")
                    event_name = event_name_element.text.strip()
                except:
                    logging.debug("Event Name not found")

                published_date = ""
                try:
                    published_date_element = row.find_element(By.CSS_SELECTOR, "td.cdk-column-publishedDate")
                    published_date = published_date_element.text.strip()
                except:
                    logging.debug("Published Date not found")

                award_date = ""
                try:
                    award_date_element = row.find_element(By.CSS_SELECTOR, "td.cdk-column-awardDate")
                    award_date = award_date_element.text.strip()
                except:
                    logging.debug("Award Date not found")

                event_due_date = ""
                try:
                    event_due_date_element = row.find_element(By.CSS_SELECTOR, "td.cdk-column-eventDueDate")
                    event_due_date = event_due_date_element.text.strip()
                except:
                    logging.debug("Event Due Date not found")

                invitation_type = ""
                try:
                    invitation_type_element = row.find_element(By.CSS_SELECTOR, "td.cdk-column-invitationType")
                    invitation_type = invitation_type_element.text.strip()
                except:
                    logging.debug("Invitation Type not found")

                if event_id and event_name:
                    all_data.append({
                        "Event ID": event_id,
                        "Event Name": event_name,
                        "Published Date": published_date,
                        "Award Date": award_date,
                        "Event Due Date": event_due_date,
                        "Invitation Type": invitation_type,
                        "Status": status
                    })
                    awarded_count += 1
                    logging.info(f"Scraped Awarded opportunity: {event_id} - {event_name}")

            except Exception as e:
                logging.debug(f"Error parsing row: {e}")
                continue

        logging.info(f"Found {awarded_count} Awarded opportunities on page {page_count}")

        try:
            next_button = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable(
                    (By.XPATH, "//div[contains(@class, 'page-item') and contains(@class, 'arrow') and not(@aria-disabled='true')]//mat-icon[contains(text(), 'keyboard_arrow_right')]/ancestor::a")
                )
            )
            driver.execute_script("arguments[0].click();", next_button)
            logging.info(f"Clicked Next button for page {page_count + 1}")
            time.sleep(5)
        except Exception as e:
            logging.info(f"Next button not found or disabled - stopping: {e}")
            break

    df = pd.DataFrame(all_data)
    output_file = os.path.join(os.getcwd(), "Past_Awarded_Opportunities.xlsx")
    if all_data:
        df.to_excel(output_file, index=False, sheet_name="Awarded_Opportunities")
        logging.info(f"Data scraped successfully. Excel file saved: {output_file}")
        logging.info(f"Total Awarded opportunities scraped: {len(all_data)}")
    else:
        df.loc[0] = ["No Awarded opportunities found", "", "", "", "", "", "Awarded"]
        df.to_excel(output_file, index=False, sheet_name="Awarded_Opportunities")
        logging.warning(f"No Awarded opportunities scraped. Empty Excel with note saved: {output_file}")

except Exception as e:
    logging.error(f"An error occurred: {e}")
    with open(os.path.join(download_dir, "error_page_source.html"), "w") as f:
        f.write(driver.page_source)
    logging.info("Saved error page source to downloads/error_page_source.html")
finally:
    logging.info("Closing browser")
    driver.quit()