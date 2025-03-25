from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
from datetime import datetime
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import Select
import time
from selenium.webdriver.common.action_chains import ActionChains
import pyautogui
import pandas as pd
import os
import functions
from Macabi_GUI import get_basic_info2


alon = True
debug = True

login_id = "126280"
login_password = "Zahalafarm2024"
provider_type = "5"
provider_code = "24657"
site_link = "https://wmsup.mac.org.il/mbills"
did_reported_col = 6
left_over_treatment_col = 7
need_new_approval_col = 8  # TODO:  check if needed

# python -m PyInstaller --onefile --add-data "C:\Users\alons\PycharmProjects\script for farm\*;script for farm" "C:\Users\alons\PycharmProjects\script for farm\venv\src\Macabi_script.py"


# Load the Excel file

# if alon:
#     XL_path = r"C:\Users\alons\Downloads\דגימות מכבי.xlsx"
# else:
#     XL_path = r"C:\Users\HOME\Desktop\reports\macabi\participants_macabi.xlsx"

XL_path,login_password = get_basic_info2()

costumers = functions.process_excel(XL_path)

driver = functions.set_up_driver(site_link)

try:
    # Wait for the username field to be visible
    WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.ID, "username")))

    # Fill in the username
    username_input = driver.find_element(By.ID, "username")
    username_input.send_keys("126280")

    # Wait until the password input field is present and visible
    password_input = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.ID, "password"))
    )

    # Enter the password
    password_input.send_keys("Zahalafarm2024")

    # Click the login button
    login_button = driver.find_element(By.CSS_SELECTOR, "input[type='submit']")
    login_button.click()

    # Wait until the 'ServiceType' input field is present and visible
    service_type = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.ID, "ServiceType"))
    )
    service_type.send_keys(provider_type)

    # Wait until the 'ServiceCode' input field is present and visible
    service_code = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.ID, "ServiceCode"))
    )
    service_code.send_keys(provider_code)

    enter = driver.find_element(By.ID, "Save")
    enter.click()

    print("Login attempt completed.")

except Exception as e:
    print(f"Error during login: {e}")

# Wait until the element is present and clickable
patient_intake = WebDriverWait(driver, 10).until(
    EC.element_to_be_clickable(
        (By.XPATH, "/html/body/table/tbody/tr/td/table[2]/tbody/tr/td[2]/table/tbody/tr/td[4]/table/tbody/tr/td[2]"))
)

# Click the element
patient_intake.click()
# open extend option
extend = driver.find_element(By.XPATH,
                             "/html/body/center/table/tbody/tr/td/table/tbody/tr/td[6]/a/u")
extend.click()

# insert patient
number_of_inserts = len(costumers)
number_of_groups = number_of_inserts // 8 + 1
last_group_size = number_of_inserts % 8
print(number_of_groups)

functions.clear_col(XL_path, did_reported_col, number_of_inserts)

for i in range(0, number_of_groups):  # todo: change!
    end = 9
    if i == number_of_groups - 1:  # if its the last group
        end = last_group_size + 1
        print("last one")

    try:

        for j in range(1, end):
            current_patient = costumers[i * 8 + j - 1]

            id_element = driver.find_element(By.ID, "ID" + str(j))
            date_element = driver.find_element(By.ID, "treatmentDate" + str(j))
            date_element.clear()
            id_element.clear()

            # fill up id
            id_element.send_keys(current_patient["id"])

            # fill up date
            day = str(current_patient["day"])
            if len(day) == 1:
                day = "0" + day
            month = str(current_patient["month"])
            if len(month) == 1:
                month = "0" + month

            # try:
            #     treatment_option = WebDriverWait(driver, 3).until(
            #         EC.element_to_be_clickable((By.ID, "99701")))
            #     treatment_option.click()
            # except:
            #     treatment_option = WebDriverWait(driver, 3).until(
            #         EC.element_to_be_clickable((By.ID, "99601")))
            #     treatment_option.click
            #

            date_element.send_keys(day + "/" + month + "/" + str(current_patient["year"]))

        for j in range(1, end):
            # fill up treatment
            treatment_picker = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.ID, "treatmentDescr" + str(j))))
            driver.execute_script("arguments[0].click();", treatment_picker)
            try:
                treatment = WebDriverWait(driver, 10).until(
                    EC.any_of(
                        EC.presence_of_element_located(
                            (By.XPATH, "/html/body/center/div[1]/div[1]/div[2]/table/tbody/tr[2]/td")),
                        EC.presence_of_element_located(
                            (By.XPATH, "/html/body/center/div[1]/div[1]/div[2]/table/tbody/tr[3]/td"))

                    )
                )
                treatment.click()
            except:
                treatment = WebDriverWait(driver, 10).until(
                    EC.any_of(
                        EC.presence_of_element_located(
                            (By.XPATH, "/html/body/center/div[1]/div[1]/div[2]/table/tbody/tr[3]/td")),
                        EC.presence_of_element_located(
                            (By.XPATH, "/html/body/center/div[1]/div[1]/div[2]/table/tbody/tr[2]/td"))
                    )
                )
                treatment.click()
        # functions.clear_table(driver, 1, 9)

        # click on save
        save_button = driver.find_element("id", "imgSave")
        save_button.click()

        # click on enter
        time.sleep(1)
        pyautogui.press("enter")

        # wait for the data to be updated
        while True:
            try:
                id_element = driver.find_element(By.ID, "ID1")
                break
            except:
                time.sleep(2)

        pyautogui.press("enter")

        # update the Excel
        for j in range(1, end):
            current_patient = costumers[i * 8 + j - 1]
            for r in current_patient["rows"]:
                functions.write_to_excel(XL_path, r, did_reported_col, "V")


    except:
        # update the Excel
        for j in range(1, end):
            current_patient = costumers[i * 8 + j - 1]
            for r in current_patient["rows"]:
                functions.write_to_excel(XL_path, r, did_reported_col, "X")

# Close the browser
driver.quit()
