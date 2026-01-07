import os
import time

from selenium.common import InvalidArgumentException
import json
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC

from src import functions
from src.functions import write_to_excel


def upload_Referral(patient, driver, logger, base_path, config):
    """
    Uploads a referral document for a patient if required. This function checks whether
    the patient needs a referral and processes the document accordingly by uploading it
    through a provided driver.

    Args:
        patient (dict): A dictionary representing patient information.
        driver: An automation driver used to interact with the web interface or external
            system for the upload process.
        logger: A logging object for capturing logs and providing feedback on the function's
            execution, such as indicating when a referral is needed or logging file search
            details.
        base_path (str): The base file path where referral documents are located. This is
            used as the starting point for locating the specified referral file.
        config: A configuration object containing settings for the function.
    Returns:
        int: Returns 0 if no referral is required or if the upload is successful.
            returns -1 if the upload fails or if an error occurs during the process.
    """

    if patient["need_referral"]:
        logger.info(f"{patient['id']} needs referral")
        file = patient["referral"]
        full_path = os.path.abspath(str(os.path.join(base_path, file)))
        full_path_2 = full_path + ".pdf"
        logger.info(f"looking for file: {full_path}")
        # click on the referral button
        if upload_file(driver, patient, full_path_2, full_path, logger, config, 2) == -1:
            logger.info(f"error with uploading the file: {full_path}")
            try:
                write_to_excel(config.XL_path, patient["row"], config.is_referral_uploaded_col, "X")
            except Exception as e:
                logger.error(f"error with writing to excel X for referral: {e}")
            return -1
        else:
            write_to_excel(config.XL_path, patient["row"], config.is_referral_uploaded_col, "V")
            logger.info(f"uploaded the referral file: {full_path}")
    else:
        logger.info(f"{patient['id']} does not need referral")
    return 0


def upload_file(driver, patient, file_path, file_path_try_2, logger, config, file_box_int=1):
    """
    Uploads a file to a web application using WebDriver within an iframe. The function handles
    hidden file input fields, error scenarios during file upload, and retries with an alternate
    file path if the initial one fails. It also logs the process progress, catches exceptions,
    and returns an appropriate status code indicating success or failure.

    Parameters:
    driver (webdriver): The WebDriver instance used to control the browser.
    patient (dict): Information about the patient, including data used for logging or updating.
    file_path (str): The primary file path to be uploaded.
    file_path_try_2 (str): The alternate file path to be attempted if the primary path fails.
    logger (logging.Logger): The Logger instance to record informational and error messages.
    file_box_int (int): The integer appended to the file input field's ID to identify it.
    config: Configuration object containing settings for the function.
    Returns:
    int: Returns 0 on successful file upload and -1 on failure.
    """
    wait = WebDriverWait(driver, 10)

    # Switch to the iframe
    try:
        iframe = wait.until(EC.presence_of_element_located((By.ID, "ifrFiles")))
        driver.switch_to.frame(iframe)
        logger.info("switched to iframe")
    except Exception as e:
        logger.info(f"error with switching to iframe(not stopping the process): {repr(e)} ")

    # Wait for iframe content to load
    time.sleep(1)

    # Find the hidden file input field
    file_box = f"fileToUpload{file_box_int}"
    file_input = wait.until(EC.presence_of_element_located((By.ID, file_box)))
    logger.info("found file input box")

    time.sleep(1)
    logger.info(f"sending file: {file_path}")
    try:
        file_input.send_keys(str(file_path))
    except InvalidArgumentException as e:
        try:
            logger.info(f"error with sending file: {repr(e)} trying the path {file_path_try_2}")
            file_input.send_keys(str(file_path_try_2))
        except Exception as e:
            logger.info(f"error with sending file(try 2): {repr(e)}")
            return -1
    except Exception as e:
        logger.info(f"error with sending file: {repr(e)}")
        write_to_excel(config.XL_path, patient["row"], config.error_col, "error with sending file")
        return -1
    logger.info(f"sent file")
    time.sleep(0.5)

    return 0

# choose the provider by the digit sum of the id
def choose_provider_index(id, providers_count = 13):
    sum_digits = 0
    id = int(id)
    while id > 0:
        sum_digits += id % 10
        id //= 10
    return sum_digits % providers_count

def select_and_click_provider(logger,driver,output_XL_path,row,error_col,costumer_id):
    # Wait for and select provider
    # Wait for dropdown arrow to be clickable
    try:
        time.sleep(0.2)
        dropdown_arrow = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.CLASS_NAME, "css-b3yrsp-indicatorContainer"))
        )
        logger.info("found dropdown arrow")
        time.sleep(0.2)

        driver.execute_script(
            "arguments[0].scrollIntoView({block: 'center', inline: 'center'});",
            dropdown_arrow
        )

        # Scroll into view if necessary
        #driver.execute_script("arguments[0].scrollIntoView();", dropdown_arrow)
        logger.info("scrolled into view")

        # Try clicking the dropdown arrow
        try:
            dropdown_arrow.click()
            logger.info("clicked dropdown arrow")
        except:
            logger.error("Selenium click failed. Trying JavaScript click...")
            driver.execute_script("arguments[0].click();", dropdown_arrow)
            logger.info("JavaScript click successful")

        #  Grab the hidden‐input’s value (suppliers) and parse it
        hid = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.ID, "ctl00_MainContent_hidSubSuppliers"))
        )
        raw = hid.get_attribute("value")
        # value is like '[{"value":"81471","val01":"…","val04":"…", …}, …]'
        providers = json.loads(raw)
        provider_names = [p["val04"].strip() for p in providers if p.get("val04", "").strip()]

        # Randomly select a provider
        chosen_provider = provider_names[choose_provider_index(costumer_id, len(provider_names))]
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
        functions.write_to_excel(output_XL_path, row, error_col, "error with selecting provider")
        raise e

def select_date(logger,driver,output_XL_path,error_col,current_patient):
   try:
       day = str(current_patient["day"])
       year = current_patient["year"]
       month = current_patient["month"]
       dp = driver.find_element(By.XPATH,
                                "/html/body/center/table/tbody/tr/td/form[1]/table/tbody/tr/td[2]/div/table/tbody/tr[6]/td/div/table/tbody/tr[2]/td[2]/table/tbody/tr[4]/td[2]/span[1]/img")
       driver.execute_script("arguments[0].click();", dp)

   except Exception as e:
       print(repr(e))
       print(e)
       raise e

   wait = WebDriverWait(driver, 10)

   # choose year:
   try:
       select_year = wait.until(
           EC.visibility_of_element_located((By.XPATH, f'//*[@id="ui-datepicker-div"]/div/div/select[2]')))
       select_year.click()
       year_option = wait.until(EC.visibility_of_element_located(
           (By.XPATH, f'//*[@id="ui-datepicker-div"]/div/div/select[2]/option[@value="{str(year)}"]')
       ))
       logger.info(f"found selected year {year_option}")
       year_option.click()
       logger.info("clicked selected year")
   except Exception as e:
       logger.error(f"error with selecting year (or openning calander) {repr(e)}")
       functions.write_to_excel(output_XL_path, current_patient["row"], error_col,
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
       functions.write_to_excel(output_XL_path, current_patient["row"], error_col, "error with selecting month")
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
               if td.text == day:
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
       functions.write_to_excel(output_XL_path, current_patient["row"], error_col, "error with selecting day")
       raise e


def select_care_type(driver, value="6"):
    select_elem = driver.find_element(By.ID, "ctl00_MainContent_lstCareType")
    select = Select(select_elem)
    select.select_by_value(value)
