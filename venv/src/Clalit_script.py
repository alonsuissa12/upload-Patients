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

alon = False
debug = False
# base_path = r"C:\Users\alons\Downloads\files_for_clalit"
base_path, XL_path, report, upload_files, login_password = get_basic_info()

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
except:
    functions.write_to_excel(XL_path, 1, error_col, "error while processing excel")
    report = 0
    upload_files = 0

try:
    functions.clear_col(XL_path, did_reported_col, len(costumers))
    functions.clear_col(XL_path, did_file_upload_col, len(costumers))
    functions.clear_col(XL_path, left_over_treatment_col, len(costumers))
    functions.clear_col(XL_path, need_new_approval_col, len(costumers))
    functions.clear_col(XL_path, error_col, len(costumers))

except:
    functions.write_to_excel(XL_path, 1, error_col, "error while clearing excel")
    report = 0
    upload_files = 0

# set up driver
driver = 0
if report or upload_files:
    try:
        driver = functions.set_up_full_log_in(site_link, login_name, login_password, login_verification)
    except:
        functions.write_to_excel(XL_path, 1, error_col, "error with opening driver or log-in")
        report = 0
        upload_files = 0

    # Wait for the claims button
    try:
        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "ctl00_MainContent_cmdNewClaim")))
    except:
        print("WRONG PASSWORD!")
        driver.quit()
        quit(1)

# Select the correct provider using XPath
provider_names = [
    "אמסלם דוד",
    "גולדפלד דינה חוה",
    "גלינסקי חיה",
    "גנז בני",
    "וילנסקי צוריאל אהרון",
    "זילקוביץ ישראל",
    "כהן רבקה",
    "לוי דניאל",
    "קראוס שי-לי",
    "רובינס רבקה",
    "רוט צביקה"
]

reported = []
if report:
    for costumer in costumers:
        id = costumer["id"]
        day = costumer["day"]
        year = costumer["year"]
        month = costumer["month"]

        try:
            # Click "הגשת תביעות" (Submit Claims)
            try:
                report_filing = WebDriverWait(driver, 20).until(
                    EC.element_to_be_clickable((By.ID, "ctl00_MainContent_cmdNewClaim")))
                report_filing.click()
            except RuntimeError as e:
                functions.write_to_excel(XL_path, costumer["row"], error_col,
                                         "error with clicking -  הגשת תביעות   not found (run time)")
                raise e

            except Exception as e:
                functions.write_to_excel(XL_path, costumer["row"], error_col, "error with clicking -  הגשת תביעות")
                raise e

            # Wait for ID field
            try:
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "ctl00_MainContent_txtID")))
            except Exception as e:
                functions.write_to_excel(XL_path, costumer["row"], error_col, "could not find id")
                raise e

            # Enter ID
            try:
                driver.find_element(By.ID, "ctl00_MainContent_txtID").send_keys(str(id))
            except Exception as e:
                functions.write_to_excel(XL_path, costumer["row"], error_col, "error with filling id")
                raise e

            # Wait for and select provider
            # Wait for dropdown arrow to be clickable
            try:
                dropdown_arrow = WebDriverWait(driver, 10).until(
                    EC.element_to_be_clickable((By.CLASS_NAME, "css-b3yrsp-indicatorContainer"))
                )

                # Scroll into view if necessary
                driver.execute_script("arguments[0].scrollIntoView();", dropdown_arrow)

                # Try clicking the dropdown arrow
                try:
                    dropdown_arrow.click()
                except:
                    print("Selenium click failed. Trying JavaScript click...")
                    driver.execute_script("arguments[0].click();", dropdown_arrow)

                # Randomly select a provider
                chosen_provider = random.choice(provider_names)
                print(f"Chosen provider: {chosen_provider}")
                provider_option = WebDriverWait(driver, 10).until(
                    EC.presence_of_element_located((By.XPATH, f"//div[contains(text(), '{chosen_provider}')]"))
                )
                provider_option.click()
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
                year_option.click()
            except Exception as e:
                functions.write_to_excel(XL_path, costumer["row"], error_col,
                                         "error with selecting year (or openning calander)")
                raise e

            # choose month:
            try:
                month = int(month) - 1
                select_month = wait.until(
                    EC.visibility_of_element_located((By.XPATH, f'/html/body/div[7]/div/div/select[1]')))
                select_month.click()

                month_option = wait.until(EC.visibility_of_element_located(
                    (By.XPATH, f'//select[@class="ui-datepicker-month"]/option[@value="{str(month)}"]')
                ))

                # Click on the desired month option
                month_option.click()
            except Exception as e:
                functions.write_to_excel(XL_path, costumer["row"], error_col, "error with selecting month")
                raise e

            try:
                calendar_body = wait.until(
                    EC.visibility_of_element_located((By.XPATH, '//*[@id="ui-datepicker-div"]/table/tbody')))
                td_elements = calendar_body.find_elements(By.XPATH, './/td')
                while True:
                    # Re-locate all 'td' elements to avoid StaleElementReferenceException
                    td_elements = calendar_body.find_elements(By.XPATH, './/td')

                    for td in td_elements:
                        if td.text == str(day):
                            # Click the element with text '6' (if needed)
                            td.click()
                            break
                    else:
                        # If no match is found, wait for a moment and try again
                        time.sleep(1)
                        continue
                    break  # Exit the loop if we found and clicked the element
            except Exception as e:
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
                    if debug:
                        print("is 1 stall", end="")
                    message = WebDriverWait(driver, 20).until(
                        EC.presence_of_element_located((By.XPATH,
                                                        "/html/body/center/table/tbody/tr/td/form[1]/table/tbody/tr/td[2]/div/table/tbody/tr[7]/td/div")))
                    if debug:
                        print(" NO\n")
                    # Store the extracted text in a string variable
                    try:
                        extracted_text = message.text
                        # Split the text into words
                        words = extracted_text.split()
                    except:
                        continue

                    # check for new aproval
                    try:
                        # Locate the 'סגור' button by XPath
                        if debug:
                            print("is 2 stall", end="")
                        close_button = WebDriverWait(driver, 0.5).until(
                            EC.presence_of_element_located((By.XPATH, "//button[text()='סגור']")))
                        alert_content = WebDriverWait(driver, 0.5).until(
                            EC.presence_of_element_located((By.XPATH, '//*[@id="mp_dialog_err"]/div[2]')))

                        alert_date = functions.extract_date(alert_content)

                        if debug:
                            print("alert date = ", alert_date)

                        # Click the 'סגור' button
                        close_button.click()
                        if debug:
                            print(" NO\n")
                        print("NEED NEW APPROVAL!!!!!")
                        functions.write_to_excel(XL_path, costumer["row"], need_new_approval_col, alert_date)
                        loop_traker = 2

                    except Exception as e:
                        print("did not find alert on approval (its ok)", repr(e))

                    # Check if the first word is "מספר"
                    if words and words[0] == "מספר":
                        counter += 1
                        print(f"first word:{words[0]}")
                        print(f"second word:{words[1]}")
                        if counter >= 40:
                            raise TimeoutError("עבר יותר מדי זמן ולא נמצא האלמנט!")
                    elif words[1] == "נדחתה":
                        # write X
                        functions.write_to_excel(XL_path, costumer["row"], did_reported_col, "X")
                        if "קיימת" in words and "כבר" in words:
                            # still try to report
                            reported.append(costumer)
                            functions.write_to_excel(XL_path, costumer["row"], error_col,extracted_text)
                            functions.write_to_excel(XL_path, costumer["row"], did_reported_col, "X (read errors!)")
                        break

                    elif words[1] == "נקלטה,":
                        loop_traker += 1
                        if debug:
                            print("loop tracker ++")
                        if loop_traker >= 2:
                            print("הצלחה")

                            if "למבוטח" in words:
                                index = words.index("למבוטח") + 1
                                left_over_treatments = words[index]

                                print(f"successfully reported for: {id} in date:{day}/{month}/{year}")

                                # update the Excel
                                functions.write_to_excel(XL_path, costumer["row"], left_over_treatment_col,
                                                         str(left_over_treatments))
                            reported.append(costumer)

                            functions.write_to_excel(XL_path, costumer["row"], did_reported_col, "V")


            except Exception as e:
                print("problem with message:")
                print(e)
                functions.write_to_excel(XL_path, costumer["row"], error_col, "error with processing system messages")
                raise e

            if debug:
                print("is 3 stall", end="")

            try:
                main_button = wait.until(
                    EC.element_to_be_clickable(
                        (By.XPATH, '/html/body/center/table/tbody/tr/td/form[1]/div[6]/div/input[5]')))
                main_button.click()
            except Exception as e:
                functions.write_to_excel(XL_path, costumer["row"], error_col, "error with going back to main")
                driver.quit()
                driver = functions.set_up_full_log_in(site_link, login_name, login_password, login_verification)

            if debug:
                print(" NO\n")

        except Exception as e:
            print(f"FAILED to REPORT for: {id} in date:{day}/{month}/{year}!!!!!!!!!!!!")
            print(repr(e))
            for r in costumer["rows"]:
                functions.write_to_excel(XL_path, r, did_reported_col, "X")
            driver.quit()
            driver = functions.set_up_full_log_in(site_link, login_name, login_password, login_verification)

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
time.sleep(1)

# prints
print("\nUPLOADING FILES FOR:")
for costumer in reported:
    print(f"{costumer['id']} win rows: {costumer['rows']} ")

if upload_files:
    for costumer in reported:
        id = costumer["id"]
        day = costumer["day"]
        year = costumer["year"]
        month = costumer["month"]
        file = costumer["file"]
        full_path = os.path.join(base_path, file)

        try:
            wait = WebDriverWait(driver, 10)

            driver.refresh()
            try:
                time.sleep(2)
                EC.presence_of_element_located((By.ID, "ctl00_MainContent_txtInshuredID"))
            except:
                print("already in main")

            paymant_demand = wait.until(
                EC.element_to_be_clickable((By.XPATH,
                                            "/html/body/center/table/tbody/tr/td/form/table/tbody/tr/td[2]/div/center/table/tbody/tr/td/div/table/tbody/tr[2]/td[2]/table/tbody/tr[1]/td[3]/table/tbody/tr[1]/td/table/tbody/tr/td[2]/input")))
            paymant_demand.click()

            driver.refresh()
            time.sleep(1.5)
            # Wait until the checkbox is clickable

            checkbox = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, '//*[@id="ctl00_MainContent_chkFilter"]')))

            # Click the checkbox
            if not checkbox.is_selected():
                checkbox.click()

            # Locate element
            input_element = driver.find_element(By.ID, "ctl00_MainContent_txtInshuredID")

            # Scroll into view
            driver.execute_script("arguments[0].scrollIntoView();", input_element)
            time.sleep(1)  # Short delay to allow UI updates

            # Try clicking and then sending keys
            if debug:
                print("input id click:")
            input_element.click()
            input_element.clear()
            input_element.send_keys(str(id))

            # filter
            if debug:
                print("filter:")
            time.sleep(1)
            make_filt = WebDriverWait(driver, 5).until(
                EC.element_to_be_clickable((By.XPATH, '//*[@id="ctl00_MainContent_cmdApplyFilter"]')))
            make_filt.click()

            # make a series
            if debug:
                print("make a series:")
            time.sleep(1)
            series = WebDriverWait(driver, 5).until(
                EC.element_to_be_clickable((By.XPATH, '//*[@id="ctl00_MainButtons_cmdReConfirmationSeries"]')))
            series.click()

            # mark checkboxes
            wait = WebDriverWait(driver, 5)  # Adjust timeout as needed
            checkboxes = wait.until(EC.presence_of_all_elements_located((By.XPATH, "//input[@type='checkbox']")))
            # Click each checkbox
            for checkbox in checkboxes:
                if not checkbox.is_selected():
                    checkbox.click()
            # Switch to the iframe
            iframe = wait.until(EC.presence_of_element_located((By.ID, "ifrFiles")))
            driver.switch_to.frame(iframe)

            # Wait for iframe content to load
            time.sleep(1)

            # Find the hidden file input field
            file_input = wait.until(EC.presence_of_element_located((By.ID, "fileToUpload1")))

            time.sleep(1)

            file_input.send_keys(str(full_path))
            time.sleep(1)

            # scroll down
            driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")

            driver.switch_to.default_content()
            if debug:
                print("ID =", id)

            aspent_f = driver.find_element(By.ID, "aspnetForm")
            btns = aspent_f.find_element(By.CLASS_NAME, "mainButtons")
            send_button = btns.find_element(By.ID, "ctl00_MainButtons_cmdSend")
            main_button = btns.find_element(By.ID, "ctl00_MainButtons_cmdExit")

            send_button.click()
            time.sleep(3)
            main_button.click()
            time.sleep(2)

            print(f"File uploaded successfully! for {id} ")
            for repoted_c in reported:
                for r in repoted_c["rows"]:
                    functions.write_to_excel(XL_path, r, did_file_upload_col, "V")
        except Exception as e:
            print(f"ERROR UPLOADING FILES FOR  {id} ")
            print(repr(e))
            for r in costumer["rows"]:
                functions.write_to_excel(XL_path, r, did_file_upload_col, "X")
            driver.quit()
            driver = functions.set_up_full_log_in(site_link, login_name, login_password, login_verification)

print("DONE")
driver.quit()
input("Press Enter to exit...")
