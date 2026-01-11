import os
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
import pyautogui
from datetime import datetime
import functions
from Clalit_GUI import get_basic_info
import logger
from Clalit_Helper_Functions import select_and_click_provider, upload_Referral, upload_file, select_date, \
    select_care_type
from config import Config
from src.functions import copy_headers_by_index, write_to_excel, write_customer_to_excel, \
    write_customer_to_excel_few_rows

# -------------------- config --------------------
config = Config("clalit")
logger = logger.setup_logger("CLALIT")
base_path, config.XL_path, report, upload_files, login_password = get_basic_info()
logger.info(
    f"got info from GUI:\n base_path: {base_path}\n XL_path: {config.XL_path}\n report: {report}\n upload_files: {upload_files}\n login_password: {'*' * len(login_password)}")

# ----------------- config variables -----------------
try:
    clear_xl = False
    XL_path = config.XL_path
    input_XL_path = XL_path
    base, ext = os.path.splitext(XL_path)
    now = datetime.now().strftime("%Y_%m_%d_%H_%M")
    output_XL_path = f"{base}_output_{now}{ext}"
    did_reported_col = config.did_reported_col
    error_col = config.error_col
except Exception as e:
    logger.error(f"error with config variables: \n\t {repr(e)}")
    quit(1)

# ------------------ main code -----------------
# Loop through each row starting from line 2 (index 1 in pandas)
try:
    costumers = functions.process_excel(XL_path, config, base_path)
    logger.info(f"Found {len(costumers)} customers to process from excel.")
except:
    logger.error("error while tried to process excel")
    report = 0
    upload_files = 0

# copy the headlines from the input excel:
try:
    copy_headers_by_index(input_XL_path, output_XL_path, [0, 1, 2, 3, 5, 6, 7, 8, 9, 11, 12])
except:
    logger.error("error while tried to copy headers to output excel")
    report = 0
    upload_files = 0

# set up driver
driver = 0
if report or upload_files:
    try:
        driver = functions.set_up_full_log_in(config.site_link, config.login_name, login_password,
                                              config.login_verification)
        logger.info("driver set up")
        time.sleep(1)
        # click enter to deal with the pop up
        pyautogui.press("enter")

    except:
        logger.error("error with opening driver or log-in")
        functions.write_to_excel(output_XL_path, 2, error_col, "error with opening driver or log-in")
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
                logger.error(f"error with clicking -  הגשת תביעות   not found (run time) {repr(e)}")
                functions.update_customer_writing(costumer, [error_col],
                                                  ["error with clicking -  הגשת תביעות   not found (run time)"])
                raise e

            except Exception as e:
                logger.error(f"error with clicking -  הגשת תביעות   not found {repr(e)}")
                functions.update_customer_writing(costumer, [error_col], ["error with clicking -  הגשת תביעות"])
                raise e

            # Wait for ID field
            try:
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "ctl00_MainContent_txtID")))
                logger.info("found id field")
            except Exception as e:
                functions.update_customer_writing(costumer, [error_col], ["could not find id"])
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
                functions.update_customer_writing(costumer, [error_col], ["error with filling id"])
                logger.error(f"error with filling id {repr(e)}")
                raise e

            select_and_click_provider(logger, driver, output_XL_path, costumer["row"], error_col, id)

            ###################
            #    DATE         #
            # #################

            select_date(logger, driver, output_XL_path, error_col, costumer)

            # change the names
            # Locate the family name element
            family_name_element = driver.find_element(By.ID, "ctl00_MainContent_txtInsuredFamily")
            logger.info("found family name element")
            # locate the first name element
            first_name_element = driver.find_element(By.ID, "ctl00_MainContent_txtInsuredName")
            logger.info("found first name element")

            logger.info("clearing names")
            driver.execute_script("arguments[0].value = '';", family_name_element)
            driver.execute_script("arguments[0].value = '';", first_name_element)
            # Fill the input fields with the values from the Excel file
            logger.info("filling names")
            driver.execute_script("arguments[0].value = arguments[1];", family_name_element, costumer["last_name"])
            driver.execute_script("arguments[0].value = arguments[1];", first_name_element, costumer["first_name"])

            try:
                select_care_type(driver)
            except Exception as e:
                print(repr(e))

            #  send
            try:
                # send the report
                # driver.execute_script(
                #     "__doPostBack('ctl00$MainButtons$cmdSend','');"
                # )
                driver.execute_script(
                    "document.getElementById('ctl00_MainButtons_cmdSend').click();"
                )


            except Exception as e:
                functions.update_customer_writing(costumer, [error_col], ["error with sending report"])
                raise e

            # process system messages

            try:
                loop_traker = 0
                counter = 0
                sleep_time = 2
                while loop_traker < 2:

                    time.sleep(sleep_time)

                    # Wait for the element to be present in the DOM
                    message = WebDriverWait(driver, 20).until(
                        EC.presence_of_element_located((By.XPATH,
                                                        "/html/body/center/table/tbody/tr/td/form[1]/table/tbody/tr/td[2]/div/table/tbody/tr[7]/td/div")))
                    logger.info("found system message")
                    functions.update_customer_writing(costumer, [config.system_message_col], [message])
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
                        functions.update_customer_writing(costumer, [config.need_new_approval_col], [alert_date])
                        loop_traker = 2

                    except Exception as e:
                        functions.update_customer_writing(costumer, [config.need_new_approval_col], ["no"])
                        logger.info("did not find alert on approval (its ok)")

                    # Check if the first word is "מספר"
                    if words and words[0] == "מספר":
                        counter += 1
                        logger.info("found the word מספר")
                        if counter >= config.wait_time_limit / sleep_time:
                            logger.error("timeout error - TOO MUCH TIME")
                            functions.update_customer_writing(costumer, [config.system_message_col], [extracted_text])
                            raise TimeoutError("עבר יותר מדי זמן ולא נמצאה הודעת אישור")
                    elif words[1] == "נדחתה":
                        logger.info("found the word נדחתה")
                        # write X
                        functions.update_customer_writing(costumer, [did_reported_col], ["X (נדחתה)"])
                        functions.update_customer_writing(costumer, [config.system_message_col], [extracted_text])
                        if "קיימת" in words and "כבר" in words:
                            logger.info("found the words קיימת כבר")
                            # still try to report
                            reported.append(costumer)
                            functions.update_customer_writing(costumer, [did_reported_col], ["V (קיימת כבר)"])
                            did_report = True
                        break

                    elif words[1] == "נקלטה,":
                        logger.info("found the word נקלטה")
                        loop_traker += 1
                        if loop_traker >= 2:
                            logger.info("found the word הצלחה")

                            if "למבוטח" in words:
                                index = words.index("למבוטח") + 1
                                left_over_treatments = words[index]

                                logger.info(
                                    f"successfully reported for: {id} in date:{day}/{month}/{year}. left over treatments: {left_over_treatments}")

                                # update the Excel
                                functions.update_customer_writing(costumer, [config.left_over_treatment_col],
                                                                  [str(left_over_treatments)])
                            reported.append(costumer)
                            did_report = True
                            functions.update_customer_writing(costumer, [config.system_message_col], [extracted_text])
                            functions.update_customer_writing(costumer, [did_reported_col], ["V"])

            except Exception as e:
                logger.error(f"error with system messages {repr(e)}")
                functions.update_customer_writing(costumer, [error_col], ["error with processing system messages"])
                raise e

            try:
                # go back to main page
                driver.execute_script(
                    "document.getElementById('ctl00_MainButtons_cmdExit').click();"
                )

                logger.info("clicked main button")
                time.sleep(2)
            except Exception as e:
                logger.error(f"error with clicking main button {repr(e)}. starting new report")
                raise e

            functions.write_customer_to_excel(output_XL_path, costumer)
        except Exception as e:
            if not did_report:
                logger.error(f"error with reporting for: {id} in date:{day}/{month}/{year}! {repr(e)}")
                if costumer["write_to_excel"][config.error_col] != "":
                    functions.update_customer_writing(costumer, [error_col], [str(repr(e))])
                functions.update_customer_writing(costumer, [did_reported_col], ["X"])
            try:
                functions.write_customer_to_excel(output_XL_path, costumer)
            except:
                logger.error("error with writing to excel after failed report, for:", costumer["id"])
            driver.quit()

            driver = functions.set_up_full_log_in(config.site_link, config.login_name, login_password,
                                                  config.login_verification)

if upload_files:
    logger.info("\n-------------------- UPLOADING FILES --------------------\n")
    if report:
        unique_customers = functions.get_unique_customers(reported)
        for uc in unique_customers:
            for c in costumers:
                if uc["id"] == c["id"] and (
                        uc["day"] != c["day"] or uc["month"] != c["month"] or uc["year"] != c["year"]):
                    uc["rows"].append(c["row"])

        reported = unique_customers

    else:
        reported = functions.get_unique_customers(costumers)
        for uc in reported:
            for c in costumers:
                if uc["id"] == c["id"] and (
                        uc["day"] != c["day"] or uc["month"] != c["month"] or uc["year"] != c["year"]):
                    uc["rows"].append(c["row"])

    unique_customers = reported
    logger.info(f"found {len(unique_customers)} unique customers:")
    for uc in unique_customers:
        logger.info(f"found unique customer {uc['id']} with rows: [ {uc['rows']}]")
    time.sleep(0.5)

    current_customer = 0

    for costumer in reported:
        current_customer = costumer
        file_uploaded = False
        logger.info(f"start uploading for: {costumer['id']}")
        id = costumer["id"]
        day = costumer["day"]
        year = costumer["year"]
        month = costumer["month"]
        full_path = costumer["file"]
        full_path = os.path.abspath(str(os.path.join(base_path, full_path)))
        logger.info(f"looking for file: {full_path}")
        full_path_try2 = full_path + ".pdf"
        # full_path = full_path + "_tc.pdf"

        try:
            wait = WebDriverWait(driver, 10)

            driver.refresh()
            try:
                time.sleep(1.5)
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
            time.sleep(0.5)
            # Wait until the checkbox is clickable

            checkbox = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, '//*[@id="ctl00_MainContent_chkFilter"]')))
            logger.info("found checkbox")
            # Click the checkbox
            if not checkbox.is_selected():
                logger.info("checkbox not clicked yet. clicking now...")
                checkbox.click()
                logger.info("clicked checkbox")

            # Locate ID element
            input_ID_element = driver.find_element(By.ID, "ctl00_MainContent_txtInshuredID")
            logger.info("found ID element")

            # Scroll into view
            driver.execute_script("arguments[0].scrollIntoView();", input_ID_element)
            logger.info("scrolled into view")
            time.sleep(1)  # Short delay to allow UI updates

            # Try clicking and then sending keys
            input_ID_element.click()
            input_ID_element.clear()
            driver.execute_script("arguments[0].value = arguments[1];", input_ID_element, str(id))

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

            # upload the file
            logger.info("uploading receipt...")
            if upload_file(driver, current_customer, full_path, full_path_try2, logger, config, 1) == -1:
                logger.info("error with uploading receipt")
                raise Exception("failed uploading receipt")

            # upload the referral (if needed)
            if upload_Referral(current_customer, driver, logger, base_path, config) == -1:
                logger.info("error with uploading referral")
                functions.update_customer_writing(current_customer, [error_col], [
                    current_customer["write_to_excel"][error_col] + "\nerror with uploading referral"])
                raise Exception("failed uploading referral")

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
            functions.update_customer_writing(current_customer, [error_col],
                                              [current_customer["write_to_excel"][error_col] + "\n" + repr(e)])
            logger.error(repr(e))
            driver.quit()
            driver = functions.set_up_full_log_in(config.site_link, config.login_name, login_password,
                                                  config.login_verification)
        finally:
            if current_customer != 0:
                if file_uploaded:
                    logger.info(f"Writing to excel - V ,for {id}")
                    functions.update_customer_writing(current_customer, [config.did_file_upload_col], ["V"])
                else:
                    logger.info(f"Writing to excel - X ,for {id}")
                    functions.update_customer_writing(current_customer, [config.did_file_upload_col], ["X"])
                if report:
                    functions.write_customer_to_excel_few_rows(current_customer, [config.did_file_upload_col,
                                                                                  config.is_referral_uploaded_col,
                                                                                  error_col], output_XL_path)
                else:
                    functions.write_customer_to_excel(output_XL_path, current_customer)
                current_customer = 0

logger.info("DONE")
driver.quit()
