import os
import time

from selenium.common import InvalidArgumentException

from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from functions import write_to_excel


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

# choose the provider by the last number in the id
def choose_provider_index(id):
    return int(str(id)[-1])
