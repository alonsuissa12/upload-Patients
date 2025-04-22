from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
from datetime import datetime
from selenium.webdriver.common.keys import Keys
import time
from selenium.webdriver.common.action_chains import ActionChains
import pyautogui
import pandas as pd
import os
import random
import functions
from selenium.common.exceptions import TimeoutException
from Clalit_GUI import get_basic_info
import logging
import os
from datetime import datetime
import logger



logger = logger.setup_logger()
alon = False
# base_path = r"C:\Users\alons\Downloads\files_for_clalit"
# base_path = r"C:\Users\alons\Downloads\files_for_clalit"
base_path, XL_path, report, upload_files, login_password = get_basic_info()
logger.info(f"got info from GUI:\n base_path: {base_path}\n XL_path: {XL_path}\n report: {report}\n upload_files: {upload_files}\n login_password: {'*' * len(login_password)}")

if alon:
    open_calander_x = 800
    open_calander_y = 585
else:
    open_calander_x = 570
    open_calander_y = 700

login_name = "sm81471"
login_verification = "123"
site_link = "https://portalsapakim.mushlam.clalit.co.il/Mushlam/Login.aspx?ReturnUrl=%2fMushlam"
did_file_upload_col = 7
did_reported_col = 6
left_over_treatment_col = 8
need_new_approval_col = 9
error_col = 10

# Loop through each row starting from line 2 (index 1 in pandas)
try:
    costumers = functions.process_excel(XL_path)
    logger.info(f"Found {len(costumers)} customers to process from excel.")
except:
    logger.error("error while tried to process excel")
    report = 0
    upload_files = 0
try:
    functions.clear_col(XL_path, did_reported_col, len(costumers))
    functions.clear_col(XL_path, did_file_upload_col, len(costumers))
    functions.clear_col(XL_path, left_over_treatment_col, len(costumers))
    functions.clear_col(XL_path, need_new_approval_col, len(costumers))
    functions.clear_col(XL_path, error_col, len(costumers))
    logger.info("cleared all columns")

except:
    logger.error("error while clearing excel")
    functions.write_to_excel(XL_path, 1, error_col, "error while clearing excel")
    report = 0
    upload_files = 0

# set up driver
driver = 0
if report or upload_files:
    try:
        driver = functions.set_up_full_log_in(site_link, login_name, login_password, login_verification)
        logger.info("driver set up")
    except:
        logger.error("error with opening driver or log-in")
        functions.write_to_excel(XL_path, 1, error_col, "error with opening driver or log-in")
        report = 0
        upload_files = 0

    # Wait for the claims button
    try:
        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "ctl00_MainContent_cmdNewClaim")))
        logger.info("found claims button")
    except:
        logger.error("error with finding claims button probably WRONG PASSWORD")
        driver.quit()
        quit(1)

# Providers to select from
# todo: make it dynamic
provider_names = [
    "אמסלם דוד",
    "גולדפלד דינה חוה",
    "גלינסקי חיה",
    "גנז בני",
    "וילנסקי צוריאל אהרון",
    "זילקוביץ ישראל",
    "כהן רבקה",
    "לוי דניאל",
    "רובינס רבקה"
]

reported = []
if report:
    logger.info("starting report")
    for costumer in costumers:
        did_report = False
        id = costumer["id"]
        day = costumer["day"]
        year = costumer["year"]
        month = costumer["month"]

        logger.info(f"start reporting for: {id} in date:{day}/{month}/{year}")

        try:
            # Click "הגשת תביעות" (Submit Claims)
            try:
                report_filing = WebDriverWait(driver, 20).until(
                    EC.element_to_be_clickable((By.ID, "ctl00_MainContent_cmdNewClaim")))
                logger.info("found report filing button")
                report_filing.click()
                logger.info("clicked report filing button")

            except RuntimeError as e:
                logger.error(f"error with clicking -  הגשת תביעות   not found (run time) { repr(e)}")
                functions.write_to_excel(XL_path, costumer["row"], error_col,
                                         "error with clicking -  הגשת תביעות   not found (run time)")
                raise e

            except Exception as e:
                functions.write_to_excel(XL_path, costumer["row"], error_col, "error with clicking -  הגשת תביעות")
                logger.error(f"error with clicking -  הגשת תביעות   not found {repr(e)}")
                raise e

            # Wait for ID field
            try:
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "ctl00_MainContent_txtID")))
                logger.info("found id field")
            except Exception as e:
                functions.write_to_excel(XL_path, costumer["row"], error_col, "could not find id")
                logger.error(f"error with finding id {repr(e)}")
                raise e

            # Enter ID
            try:
                # driver.find_element(By.ID, "ctl00_MainContent_txtID").send_keys(str(id))
                driver.execute_script("arguments[0].value = arguments[1];",
                                      driver.find_element(By.ID, "ctl00_MainContent_txtID"),
                                      str(id))
                logger.info(f"entered id: {id}")
            except Exception as e:
                functions.write_to_excel(XL_path, costumer["row"], error_col, "error with filling id")
                logger.error(f"error with filling id {repr(e)}")
                raise e
            # Locate the family name element
            family_name_element = driver.find_element(By.ID, "ctl00_MainContent_txtInsuredFamily")
            logger.info("found family name element")
            first_name_element = driver.find_element(By.ID, "ctl00_MainContent_txtInsuredName")
            logger.info("found first name element")

            # Check if the input field is empty
            if family_name_element.get_attribute("value") == "":
                logger.info("family name field is empty. filling it now...")
                family_name_element.send_keys(costumer["last_name"])
                first_name_element.send_keys(costumer["first_name"])

            # Wait for and select provider
            # Wait for dropdown arrow to be clickable
            try:
                time.sleep(1)
                dropdown_arrow = WebDriverWait(driver, 10).until(
                    EC.element_to_be_clickable((By.CLASS_NAME, "css-b3yrsp-indicatorContainer"))
                )
                logger.info("found dropdown arrow")
                time.sleep(1)

                # Scroll into view if necessary
                driver.execute_script("arguments[0].scrollIntoView();", dropdown_arrow)
                logger.info("scrolled into view")

                # Try clicking the dropdown arrow
                try:
                    dropdown_arrow.click()
                    logger.info("clicked dropdown arrow")
                except:
                    logger.error("Selenium click failed. Trying JavaScript click...")
                    driver.execute_script("arguments[0].click();", dropdown_arrow)
                    logger.info("JavaScript click successful")


                # Randomly select a provider
                chosen_provider = random.choice(provider_names)
                logger.info(f"selected provider: {chosen_provider}")
                provider_option = WebDriverWait(driver, 10).until(
                    EC.presence_of_element_located((By.XPATH, f"//div[contains(text(), '{chosen_provider}')]"))
                )
                logger.info("found selected provider")
                time.sleep(0.5)
                try:
                    provider_option.click()
                    logger.info("clicked selected provider")
                except Exception as e:
                    logger.error("failed to click selected provider")
                    raise e
            except Exception as e:
                functions.write_to_excel(XL_path, costumer["row"], error_col, "error with selecting provider")
                raise e

            ###################
            #    DATE         #
            # #################

            pyautogui.moveTo(open_calander_x, open_calander_y)
            pyautogui.click()

            wait = WebDriverWait(driver, 10)

            # choose year:
            try:
                select_year = wait.until(
                    EC.visibility_of_element_located((By.XPATH, f'//*[@id="ui-datepicker-div"]/div/div/select[2]')))
                select_year.click()
                year_option = wait.until(EC.visibility_of_element_located(
                    (By.XPATH, f'//*[@id="ui-datepicker-div"]/div/div/select[2]/option[@value="{str(year)}"]')
                ))
                logger.info(f"found selected year { year_option}")
                year_option.click()
                logger.info("clicked selected year")
            except Exception as e:
                logger.error(f"error with selecting year (or openning calander) {repr(e)}")
                functions.write_to_excel(XL_path, costumer["row"], error_col,
                                         "error with selecting year (or openning calander)")
                raise e

            # choose month:
            try:
                month = int(month) - 1
                select_month = wait.until(
                    EC.visibility_of_element_located((By.XPATH, f'/html/body/div[7]/div/div/select[1]')))
                logger.info(f"found selected month {select_month}")
                select_month.click()
                logger.info("clicked selected month")

                month_option = wait.until(EC.visibility_of_element_located(
                    (By.XPATH, f'//select[@class="ui-datepicker-month"]/option[@value="{str(month)}"]')
                ))

                # Click on the desired month option
                month_option.click()
                logger.info("clicked month option")
            except Exception as e:
                logger.error(f"error with selecting month {repr(e)}")
                functions.write_to_excel(XL_path, costumer["row"], error_col, "error with selecting month")
                raise e

            try:
                calendar_body = wait.until(
                    EC.visibility_of_element_located((By.XPATH, '//*[@id="ui-datepicker-div"]/table/tbody')))
                td_elements = calendar_body.find_elements(By.XPATH, './/td')
                logger.info("found calendar body")
                while True:
                    # Re-locate all 'td' elements to avoid StaleElementReferenceException
                    td_elements = calendar_body.find_elements(By.XPATH, './/td')

                    for td in td_elements:
                        if td.text == str(day):
                            logger.info(f"found matching day: {td.text}")
                            # Click the element with the matching text
                            td.click()
                            logger.info("clicked matching day")
                            break
                    else:
                        # If no match is found, wait for a moment and try again
                        time.sleep(0.7)
                        continue
                    break  # Exit the loop if we found and clicked the element
            except Exception as e:
                logger.error(f"error with selecting day {repr(e)}")
                functions.write_to_excel(XL_path, costumer["row"], error_col, "error with selecting day")
                raise e

            #  send
            try:
                send_report_button = wait.until(
                    EC.element_to_be_clickable(
                        (By.XPATH, "/html/body/center/table/tbody/tr/td/form[1]/div[6]/div/input[1]")))
                send_report_button.click()
            except Exception as e:
                functions.write_to_excel(XL_path, costumer["row"], error_col, "error with sending report")
                raise e

            # process system messages

            try:
                loop_traker = 0
                counter = 0
                while loop_traker < 2:
                    time.sleep(2)

                    # Wait for the element to be present in the DOM
                    message = WebDriverWait(driver, 20).until(
                        EC.presence_of_element_located((By.XPATH,
                                                        "/html/body/center/table/tbody/tr/td/form[1]/table/tbody/tr/td[2]/div/table/tbody/tr[7]/td/div")))
                    logger.info("found system message")
                    # Store the extracted text in a string variable
                    try:
                        extracted_text = message.text
                        logger.info(f"extracted text: {extracted_text}")
                        # Split the text into words
                        words = extracted_text.split()
                        logger.info(f"splited text to words")
                        if len(words) == 0:
                            logger.error("no words found in message")
                            raise ValueError("No words found in message")
                    except:
                        continue

                    # check for new aproval
                    try:
                        logger.info("checking for new aproval need")
                        # Locate the 'סגור' button by XPath

                        close_button = WebDriverWait(driver, 0.5).until(
                            EC.presence_of_element_located((By.XPATH, "//button[text()='סגור']")))
                        alert_content = WebDriverWait(driver, 0.5).until(
                            EC.presence_of_element_located((By.XPATH, '//*[@id="mp_dialog_err"]/div[2]')))
                        logger.info("found alert")
                        alert_date = functions.extract_date(alert_content)
                        logger.info(f"found alert date: {alert_date}")
                        # Click the 'סגור' button
                        close_button.click()


                        logger.info("closing alert of new approval need")
                        functions.write_to_excel(XL_path, costumer["row"], need_new_approval_col, alert_date)
                        loop_traker = 2

                    except Exception as e:
                        logger.info("did not find alert on approval (its ok)")

                    # Check if the first word is "מספר"
                    if words and words[0] == "מספר":
                        counter += 1
                        logger.info("found the word מספר")
                        if counter >= 80:
                            logger.error("timeout error - TOO MUCH TIME")
                            raise TimeoutError("עבר יותר מדי זמן ולא נמצא האלמנט!")
                    elif words[1] == "נדחתה":
                        logger.info("found the word נדחתה")
                        # write X
                        functions.write_to_excel(XL_path, costumer["row"], did_reported_col, "X")
                        if "קיימת" in words and "כבר" in words:
                            logger.info("found the words קיימת כבר")
                            # still try to report
                            reported.append(costumer)
                            functions.write_to_excel(XL_path, costumer["row"], error_col,extracted_text)
                            functions.write_to_excel(XL_path, costumer["row"], did_reported_col, "V (קיימת כבר)")
                        break

                    elif words[1] == "נקלטה,":
                        logger.info("found the word נקלטה")
                        loop_traker += 1
                        if loop_traker >= 2:
                            logger.info("found the word הצלחה")

                            if "למבוטח" in words:
                                index = words.index("למבוטח") + 1
                                left_over_treatments = words[index]

                                logger.info(f"successfully reported for: {id} in date:{day}/{month}/{year}. left over treatments: {left_over_treatments}")

                                # update the Excel
                                functions.write_to_excel(XL_path, costumer["row"], left_over_treatment_col,
                                                         str(left_over_treatments))
                            reported.append(costumer)
                            did_report = True
                            functions.write_to_excel(XL_path, costumer["row"], did_reported_col, "V")


            except Exception as e:
                logger.error(f"error with system messages {repr(e)}")

                functions.write_to_excel(XL_path, costumer["row"], error_col, "error with processing system messages")
                raise e



            try:
                main_button = wait.until(
                    EC.element_to_be_clickable(
                        (By.XPATH, '/html/body/center/table/tbody/tr/td/form[1]/div[6]/div/input[5]')))
                logger.info("found main button")
                main_button.click()
                logger.info("clicked main button")
                time.sleep(2)
            except Exception as e:
                logger.error(f"error with clicking main button {repr(e)}. starting new report")
                raise e


        except Exception as e:
            if not did_report:
                logger.error(f"error with reporting for: {id} in date:{day}/{month}/{year}!!!!!!!!!!!!")
                functions.write_to_excel(XL_path, costumer["row"], error_col, repr(e))
                for r in costumer["rows"]:
                    functions.write_to_excel(XL_path, r, did_reported_col, "X")
            driver.quit()
            driver = functions.set_up_full_log_in(site_link, login_name, login_password, login_verification)







if upload_files:
    logger.info("\n-------------------- UPLOADING FILES --------------------\n")
    if report:
        unique_customers = functions.get_unique_customers(reported)
        for uc in unique_customers:
            for c in costumers:
                if uc["id"] == c["id"] and (
                        uc["day"] != c["day"] and uc["month"] != c["month"] and uc["year"] != c["year"]):
                    uc["rows"].append(c["row"])

        reported = unique_customers

    else:
        reported = functions.get_unique_customers(costumers)
        for uc in reported:
            for c in costumers:
                if uc["id"] == c["id"] and (
                        uc["day"] != c["day"] and uc["month"] != c["month"] and uc["year"] != c["year"]):
                    uc["rows"].append(c["row"])

    unique_customers = reported
    logger.info(f"found {len(unique_customers)} unique customers:")
    for uc in unique_customers:
        logger.info(f"found unique customer {uc['id']} with rows: [ {uc['rows']}]")
    time.sleep(1)



    for costumer in reported:
        file_uploaded = False
        logger.info(f"start uploading for: {costumer['id']}")
        id = costumer["id"]
        day = costumer["day"]
        year = costumer["year"]
        month = costumer["month"]
        file = costumer["file"]
        full_path = os.path.join(base_path, file)
        logger.info(f"looking for file: {full_path}")

        try:
            wait = WebDriverWait(driver, 10)

            driver.refresh()
            try:
                time.sleep(2)
                EC.presence_of_element_located((By.ID, "ctl00_MainContent_txtInshuredID"))
                logger.info("found main")
            except:
                logger.info("already in main")

            paymant_demand = wait.until(
                EC.element_to_be_clickable((By.XPATH,
                                            "/html/body/center/table/tbody/tr/td/form/table/tbody/tr/td[2]/div/center/table/tbody/tr/td/div/table/tbody/tr[2]/td[2]/table/tbody/tr[1]/td[3]/table/tbody/tr[1]/td/table/tbody/tr/td[2]/input")))
            logger.info("found payment demand")
            paymant_demand.click()
            logger.info("clicked payment demand")

            driver.refresh()
            time.sleep(1.5)
            # Wait until the checkbox is clickable

            checkbox = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, '//*[@id="ctl00_MainContent_chkFilter"]')))
            logger.info("found checkbox")
            # Click the checkbox
            if not checkbox.is_selected():
                logger.info("checkbox not clicked yet. clicking now...")
                checkbox.click()
                logger.info("clicked checkbox")

            # Locate element
            input_element = driver.find_element(By.ID, "ctl00_MainContent_txtInshuredID")
            logger.info("found ID element")

            # Scroll into view
            driver.execute_script("arguments[0].scrollIntoView();", input_element)
            logger.info("scrolled into view")
            time.sleep(1)  # Short delay to allow UI updates

            # Try clicking and then sending keys
            input_element.click()
            input_element.clear()
            input_element.send_keys(str(id))
            logger.info(f"entered ID: {id}")

            # filter
            time.sleep(1)
            make_filt = WebDriverWait(driver, 5).until(
                EC.element_to_be_clickable((By.XPATH, '//*[@id="ctl00_MainContent_cmdApplyFilter"]')))
            logger.info("found filter button")
            make_filt.click()
            logger.info("clicked filter button")

            # make a series
            time.sleep(1)
            series = WebDriverWait(driver, 5).until(
                EC.element_to_be_clickable((By.XPATH, '//*[@id="ctl00_MainButtons_cmdReConfirmationSeries"]')))
            logger.info("found series button")
            series.click()
            logger.info("clicked series button")

            # mark checkboxes
            wait = WebDriverWait(driver, 5)
            checkboxes = wait.until(EC.presence_of_all_elements_located((By.XPATH, "//input[@type='checkbox']")))
            logger.info("found checkboxes")
            # Click each checkbox
            for checkbox in checkboxes:
                if not checkbox.is_selected():
                    checkbox.click()
            logger.info("clicked checkboxes")
            # Switch to the iframe
            iframe = wait.until(EC.presence_of_element_located((By.ID, "ifrFiles")))
            driver.switch_to.frame(iframe)
            logger.info("switched to iframe")

            # Wait for iframe content to load
            time.sleep(1)

            # Find the hidden file input field
            file_input = wait.until(EC.presence_of_element_located((By.ID, "fileToUpload1")))
            logger.info("found file input box")

            time.sleep(1)

            file_input.send_keys(str(full_path))
            logger.info(f"sent file")
            time.sleep(1)

            # scroll down
            driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
            logger.info("scrolled down")

            driver.switch_to.default_content()
            logger.info("switched to default content")

            aspent_f = driver.find_element(By.ID, "aspnetForm")
            logger.info("found aspent form")
            btns = aspent_f.find_element(By.CLASS_NAME, "mainButtons")
            logger.info("found buttons")
            send_button = btns.find_element(By.ID, "ctl00_MainButtons_cmdSend")
            logger.info("found send button")
            main_button = btns.find_element(By.ID, "ctl00_MainButtons_cmdExit")
            logger.info("found main button")

            send_button.click()
            logger.info("clicked send button")
            time.sleep(3)
            main_button.click()
            file_uploaded = True
            logger.info("clicked main button")
            time.sleep(2)

            logger.info(f"file uploded seccessfully for: {id}")



        except Exception as e:
            logger.info(f"ERROR UPLOADING FILES FOR  {id} ")
            logger.error(repr(e))
            driver.quit()
            driver = functions.set_up_full_log_in(site_link, login_name, login_password, login_verification)
        finally:
            for repoted_c in reported:
                for r in repoted_c["rows"]:
                    if file_uploaded:
                        logger.info(f"Writing to excel - V ,for {id}")
                        functions.write_to_excel(XL_path, r, did_file_upload_col, "V")
                    else:
                        logger.info(f"Writing to excel - X ,for {id}")
                        functions.write_to_excel(XL_path, r, did_file_upload_col, "X")

print("DONE")
driver.quit()
input("Press Enter to exit...")


