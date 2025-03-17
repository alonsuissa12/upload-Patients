from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
from datetime import datetime
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
import time
from selenium.webdriver.common.action_chains import ActionChains
import pyautogui
import pandas as pd
import os
import functions

base_path = r"C:\Users\alons\Downloads\files_for_clalit"
report = True
upload_files = True

open_calander_x = 800
open_calander_y = 585
send_x = 1250
send_y = 900
main_x = 830
main_y = 770
login_name = "sm81471"
login_password = "farm2025"
login_verification = "123"
site_link = "https://portalsapakim.mushlam.clalit.co.il/Mushlam/Login.aspx?ReturnUrl=%2fMushlam"
did_file_upload_col = 7
did_reported_col = 6

# Load the Excel file
XL_path = 'C:/Users/alons/Downloads/testForClalit.xlsx'

# Loop through each row starting from line 2 (index 1 in pandas)
costumers = functions.process_excel(XL_path)

# set up driver
driver = functions.set_up_full_log_in(site_link, login_name, login_password, login_verification)

# Wait for the claims button
WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "ctl00_MainContent_cmdNewClaim")))

functions.clear_col(XL_path,did_reported_col,len(costumers))
functions.clear_col(XL_path,did_file_upload_col,len(costumers))

reported = []
if report:
    for costumer in costumers:
        id = costumer["id"]
        day = costumer["day"]
        year = costumer["year"]
        month = costumer["month"]

        try:

            # Click "הגשת תביעות" (Submit Claims)
            driver.find_element(By.ID, "ctl00_MainContent_cmdNewClaim").click()

            # Wait for ID field
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "ctl00_MainContent_txtID")))

            # Enter ID
            driver.find_element(By.ID, "ctl00_MainContent_txtID").send_keys(str(id))

            # Wait for and select provider
            # Wait for dropdown arrow to be clickable
            dropdown_arrow = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.CLASS_NAME, "css-b3yrsp-indicatorContainer"))  # Adjust if needed
            )

            # Scroll into view if necessary
            driver.execute_script("arguments[0].scrollIntoView();", dropdown_arrow)

            # Try clicking the dropdown arrow
            try:
                dropdown_arrow.click()
            except:
                print("Selenium click failed. Trying JavaScript click...")
                driver.execute_script("arguments[0].click();", dropdown_arrow)

            # Select the correct provider using XPath
            provider_option = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.XPATH, "//div[contains(text(), 'גלינסקי חיה')]"))
            )
            provider_option.click()

            ###################
            #    DATE         #
            # #################

            # open the date picke:
            date_picker_button = driver.find_element(By.CLASS_NAME, "ui-datepicker-trigger")
            pyautogui.moveTo(open_calander_x, open_calander_y)
            pyautogui.click()

            wait = WebDriverWait(driver, 10)

            # choose year:
            # //*[@id="ui-datepicker-div"]/div/div/select[2]
            select_year = wait.until(
                EC.visibility_of_element_located((By.XPATH, f'//*[@id="ui-datepicker-div"]/div/div/select[2]')))
            select_year.click()
            year_option = wait.until(EC.visibility_of_element_located(
                (By.XPATH, f'//*[@id="ui-datepicker-div"]/div/div/select[2]/option[@value="{str(year)}"]')
            ))
            year_option.click()

            # choose month:
            month = int(month) - 1
            select_month = wait.until(
                EC.visibility_of_element_located((By.XPATH, f'/html/body/div[7]/div/div/select[1]')))
            select_month.click()

            month_option = wait.until(EC.visibility_of_element_located(
                (By.XPATH, f'//select[@class="ui-datepicker-month"]/option[@value="{str(month)}"]')
            ))

            # Click on the desired month option
            month_option.click()

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
            #  send
            send_report_button = wait.until(
                EC.element_to_be_clickable(
                    (By.XPATH, "/html/body/center/table/tbody/tr/td/form[1]/div[6]/div/input[1]")))
            send_report_button.click()
            main_button = wait.until(
                EC.element_to_be_clickable(
                    (By.XPATH, '/html/body/center/table/tbody/tr/td/form[1]/div[6]/div/input[5]')))

            # Click on the "main" button
            main_button.click()
            print(f"successfully reported for: {id} in date:{day}/{month}/{year}")
            reported.append(costumer)
            for r in costumer["rows"]:
                functions.write_to_excel(XL_path, r, did_reported_col, "V")

        except:
            print(f"FAILED to REPORT for: {id} in date:{day}/{month}/{year}!!!!!!!!!!!!")
            for r in costumer["rows"]:
                functions.write_to_excel(XL_path, r, did_reported_col, "X")
            driver.quit()
            driver = functions.set_up_full_log_in(site_link, login_name, login_password, login_verification)

time.sleep(2)
unique_customers = functions.get_unique_customers(reported)

# prints
print("\nUPLOADING FILES FOR:")
for costumer in unique_customers:
    print(f"{costumer['id']} win rows: {costumer['rows']} ")

if upload_files:
    for costumer in unique_customers:
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
            time.sleep(2)
            # Wait until the checkbox is clickable
            checkbox = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, '//*[@id="ctl00_MainContent_chkFilter"]')))

            # Click the checkbox
            checkbox.click()

            # Locate element
            input_element = driver.find_element(By.ID, "ctl00_MainContent_txtInshuredID")

            # Scroll into view
            driver.execute_script("arguments[0].scrollIntoView();", input_element)
            time.sleep(1)  # Short delay to allow UI updates

            # Try clicking and then sending keys
            input_element.click()
            input_element.clear()
            input_element.send_keys(str(id))

            # filter
            make_filt = WebDriverWait(driver, 5).until(
                EC.element_to_be_clickable((By.XPATH, '//*[@id="ctl00_MainContent_cmdApplyFilter"]')))
            make_filt.click()

            # make a series
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

            # send files
            pyautogui.moveTo(send_x, send_y)
            pyautogui.click()
            time.sleep(3)
            pyautogui.moveTo(main_x, main_y)
            time.sleep(2)
            pyautogui.click()

            print(f"File uploaded successfully! for {id} for date {day}/{month}/{year}")
            for r in reported["rows"]:
                functions.write_to_excel(XL_path, r, did_file_upload_col, "V")
        except Exception as e:
            print(f"ERROR UPLOADING FILES FOR  {id} ")
            for r in costumer["rows"]:
                functions.write_to_excel(XL_path, r, did_file_upload_col, "X")
            driver.quit()
            driver = functions.set_up_full_log_in(site_link, login_name, login_password, login_verification)

print("DONE")
driver.quit()
