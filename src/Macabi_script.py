from selenium.common import NoSuchElementException
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select
import time
import pyautogui
import functions
from Macabi_GUI import get_basic_info2
import logger
from config import Config
from src.Clalit_script import XL_path

config = Config("macabi")

# -------- logger --------
logger = logger.setup_logger("MACABI")

#--------- GUI --------
config.XL_path, config.login_password = get_basic_info2()

# -------- config variables --------
debug = False

login_id = config.login_id
login_password = config.login_password
provider_type = config.provider_type
provider_code = config.provider_code
site_link = config.site_link
did_reported_col = config.did_reported_col
left_over_treatment_col = config.left_over_treatment_col
need_new_approval_col = config.need_new_approval_col
XL_path = config.XL_path

# --------- main code ---------

logger.info("Starting script")



logger.info("Excel path: " + XL_path)

costumers = functions.process_excel(XL_path,config)
logger.info("Number of patients: " + str(len(costumers)))

driver = functions.set_up_driver(site_link)
logger.info("Driver set up")

try:
    logger.info("Trying to login")
    # Wait for the username field to be visible
    WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.ID, "username")))

    # Fill in the username
    username_input = driver.find_element(By.ID, "username")
    username_input.send_keys(login_id)
    logger.info("Username filled")

    # Wait until the password input field is present and visible
    password_input = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.ID, "password"))
    )

    # Enter the password
    password_input.send_keys(login_password)
    logger.info("Password field filled")

    # Click the login button
    login_button = driver.find_element(By.CSS_SELECTOR, "input[type='submit']")
    login_button.click()
    logger.info("Login button clicked")

    # Wait until the 'ServiceType' input field is present and visible
    service_type = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.ID, "ServiceType"))
    )
    service_type.send_keys(provider_type)
    logger.info("Service type filled")

    # Wait until the 'ServiceCode' input field is present and visible
    service_code = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.ID, "ServiceCode"))
    )
    service_code.send_keys(provider_code)
    logger.info("Service code filled")

    enter = driver.find_element(By.ID, "Save")
    enter.click()
    logger.info("Login attempt completed.")

except Exception as e:
    logger.info(f"Error during login: {e}")

try:
    # Wait until the element is present and clickable
    patient_intake = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable(
            (By.XPATH,
             "/html/body/table/tbody/tr/td/table[2]/tbody/tr/td[2]/table/tbody/tr/td[4]/table/tbody/tr/td[2]"))
    )

    # Click the element
    patient_intake.click()
    logger.info("Clicked on patient intake")
except Exception as e:
    logger.error(f"Error clicking on patient intake:\n {e}")

time.sleep(3)
try:
    # check the month reprot
    if len(costumers) > 0:
        month = str(costumers[0]["month"])
        year = str(costumers[0]["year"])
        if len(month) == 1:
            month = "0" + month
        if len(year) == 2:
            year = "20" + year

        print("month = ", month, "year = ", year)
        date = month + "/" + year
        month_report = Select(driver.find_element("id", "month"))
        month_report.select_by_visible_text(date)

    wait = WebDriverWait(driver, 10)
    extend = wait.until(EC.presence_of_element_located((By.XPATH, "//u[text()='הוספה ברצף']")))
    extend.click()
    logger.info("Extended the report page")

    # insert patient
    number_of_inserts = len(costumers)

    functions.clear_col(XL_path, did_reported_col, number_of_inserts)
    logger.info("Cleared the reported columns in the XL")
    if number_of_inserts > 0:

        for j in range(0, number_of_inserts):
            current_patient = costumers[j]
            try:
                time.sleep(1.5)
                logger.info("Current patient: " + str(current_patient["id"]))

                id_element = WebDriverWait(driver, 10).until(
                    EC.presence_of_element_located((By.ID, "ID1")))
                date_element = driver.find_element(By.ID, "treatmentDate1")
                date_element.clear()
                id_element.clear()
                logger.info("id and date cleared")

                # fill up id
                id_element.send_keys(current_patient["id"])
                logger.info("id filled")

                # fill up date
                day = str(current_patient["day"])
                if len(day) == 1:
                    day = "0" + day
                month = str(current_patient["month"])
                if len(month) == 1:
                    month = "0" + month

                date_element.send_keys(day + "/" + month + "/" + str(current_patient["year"]))

                logger.info("date filled")

                # fill up treatment
                treatment_picker = WebDriverWait(driver, 10).until(
                    EC.element_to_be_clickable((By.ID, "treatmentDescr1")))
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
                    logger.info("treatment selected")
                except:
                    logger.info("treatment falid 1 time, tring again...")
                    treatment = WebDriverWait(driver, 10).until(
                        EC.any_of(
                            EC.presence_of_element_located(
                                (By.XPATH, "/html/body/center/div[1]/div[1]/div[2]/table/tbody/tr[3]/td")),
                            EC.presence_of_element_located(
                                (By.XPATH, "/html/body/center/div[1]/div[1]/div[2]/table/tbody/tr[2]/td"))
                        )
                    )
                    treatment.click()
                    logger.info("treatment selected")

                # Find left over treatments
                time.sleep(0.4)
                left_over_element = WebDriverWait(driver, 10).until(
                    EC.presence_of_element_located((By.ID, "fromField1")))
                logger.info("left over treatment element found")
                left_over_treatments = left_over_element.get_attribute("value")

                if debug:
                    print(f"left over treatments for {current_patient['id']}: {left_over_treatments}")

                # check for error:
                logger.info("checking for error message")
                try:
                    error_message_element = driver.find_element(By.ID, "ErrorMessageId")
                    error_message_text = error_message_element.text
                    if "לא" in error_message_text:
                        logger.info(f"Error message found: {error_message_text}")
                        raise Exception(f"Error message: {error_message_text}")
                except NoSuchElementException:
                    pass  # Do nothing if the element is not found

                # click on save
                save_button = driver.find_element("id", "imgSave")
                save_button.click()
                logger.info("Save button clicked")

                # click on enter
                time.sleep(1)
                pyautogui.press("enter")
                logger.info("Enter key pressed (for clicking the pop up)")

                # wait for the data to be updated
                while True:
                    try:
                        id_element = driver.find_element(By.ID, "ID1")
                        break
                    except:
                        time.sleep(2)
                logger.info("Data updated")
                pyautogui.press("enter")
                logger.info("Enter key pressed (for clicking the pop up) #2")

                # update the Excel
                try:
                    logger.info("updating the Excel...")
                    for r in current_patient["rows"]:
                        functions.write_to_excel(XL_path, r, did_reported_col, "V")
                        functions.write_to_excel(XL_path, r, left_over_treatment_col, left_over_treatments)
                        logger.info(f"updated row {r} with V in column did_reported and {left_over_treatments} in column left_over_treatment")
                except Exception as e:
                    logger.error(
                        f'Error updating Excel in rows {", ".join(str(r) for r in current_patient["rows"])}: {e}')




            except Exception as e:
                logger.error(f"Error with patient {current_patient['id']}: {e}")
                # update the Excel
                for r in current_patient["rows"]:
                    functions.write_to_excel(XL_path, r, did_reported_col, "X")
                    driver.refresh()
                    time.sleep(1)
                    pyautogui.press("enter")
finally:
    # Close the browser
    driver.quit()
