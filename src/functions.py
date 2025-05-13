from multiprocessing.reduction import duplicate

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
from datetime import datetime, timedelta
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
import pyautogui
import pandas as pd
import re
import os
from openpyxl import load_workbook
from openpyxl.styles import Alignment
from selenium.webdriver.support.ui import Select

debug = False


def set_up_driver(link):
    # Set up WebDriver
    driver = webdriver.Chrome()
    time.sleep(1)
    driver.maximize_window()
    driver.get(link)
    return driver


def set_up_full_log_in(link, name, password, verification):
    driver = set_up_driver(link)

    # Wait for login fields
    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "ctl00_MainContent_txtUser")))

    # Enter credentials
    driver.find_element(By.ID, "ctl00_MainContent_txtUser").send_keys(name)
    driver.find_element(By.ID, "ctl00_MainContent_txtPass").send_keys(password)
    driver.find_element(By.ID, "ctl00_MainContent_ExtraIdAnswer").send_keys(verification)

    # Click login button
    driver.find_element(By.ID, "ctl00_MainButtons_cmdOK").click()

    return driver


def find_file_with_number(base_path, extracted_number):
    for root, dirs, files in os.walk(base_path):
        for file in files:
            file_path = os.path.join(root, file)
            try:
                with open(file_path, 'r') as f:
                    content = f.read()
                    if str(extracted_number) in content:
                        return os.path.join(base_path, file)  # Return the path of the file containing the number
            except Exception as e:
                print(f"Error reading {file}: {e}")
    return None  # If no file is found


def process_excel(file_path, base_path="/"):
    customers = []
    print("working on:")

    try:
        # Load the Excel file with specific column types
        with pd.ExcelFile(file_path, engine='openpyxl') as xls:
            df = pd.read_excel(xls, dtype={2: str, 4: str})  # Read with column types

        for index, row in df.iterrows():
            # Check if column A (index 0) is empty
            if pd.isna(row[2]):  # Check if column A (first column) is empty
                break  # Exit the loop

            # Get values from the row
            id_value = row[2]  # ID from column C (index 2)
            date_value = row[3]  # Date from column D (index 3)
            file_name = row[4]  # File name from column E (index 4)
            first_name = row[0]
            last_name = row[1]
            if base_path != "/":
                match = re.search(r"\d{4,}", file_name)
                if match:
                    extracted_number = match.group()
                    print("match", extracted_number)
                    file_name = find_file_with_number(base_path, extracted_number)

            # Convert date string to datetime object if needed
            if isinstance(date_value, str):
                date_value = datetime.strptime(date_value, '%Y-%m-%d')  # Adjust format if needed

            # Extract day, month, and year
            customers.append({
                "row": index + 2,  # Adding row number (1-based)
                "id": id_value,
                "day": date_value.day,
                "month": date_value.month,
                "year": date_value.year,
                "date": date_value,
                "file": file_name,
                "rows": [index + 2],  # Initialize rows list with the first occurrence
                "first_name": first_name,
                "last_name": last_name

            })

            print(
                f"           Row: {index + 2}, ID: {id_value}, Date: {date_value.day}-{date_value.month}-{date_value.year}, file: {file_name}")
    except FileNotFoundError:
        print(f"Error: The file '{file_path}' was not found.")
    except PermissionError:
        print(f"Error: Permission denied. Close '{file_path}' if it's open.")
    except Exception as e:
        print(f"An error occurred: {e}")

    try:
        # find if there r duplicates base of id and name
        for i in range(len(customers)):
            for j in range(i + 1, len(customers)):
                # if there is the same customer with the same id and date
                if customers[i]["id"] == customers[j]["id"] and customers[i]["date"] == customers[j]["date"]:
                    # push the dup one day later
                    new_date = customers[j]["date"] + timedelta(days=1)
                    # if the month is different, push it to the previous day
                    if customers[i]["month"] != new_date.month :
                        new_date =  customers[j]["date"] - timedelta(days=1)
                    customers[j]["date"] = new_date
                    customers[j]["day"] = new_date.day
    except Exception as e:
        print(f"An error occurred while looking for duplicates: {e}")

    return customers


def get_unique_customers(customer_list):
    unique_customers = {}

    for customer in customer_list:
        customer_id = customer["id"]

        if customer_id in unique_customers:
            unique_customers[customer_id]["rows"].append(customer["row"])  # Append row number
        else:
            unique_customers[customer_id] = customer  # First occurrence

    return list(unique_customers.values())


def write_to_excel(file_path, row, col, txt):
    try:
        # Load the existing Excel file
        if debug:
            print("STEP 1")
        wb = load_workbook(file_path)
        if debug:
            print("STEP 2")
        sheet = wb.active  # Get the active sheet
        if debug:
            print("STEP 3")
        # Write the text to the specified cell
        cell = sheet.cell(row=row, column=col)
        if debug:
            print("STEP 4")
        cell.value = txt
        if debug:
            print("STEP 5")

        # Center align the text in the cell
        cell.alignment = Alignment(horizontal="center", vertical="center")
        if debug:
            print("STEP 6")
        # Save the modified file
        wb.save(file_path)
        wb.close()  # Close the workbook when you're done with it

        print(f"Text '{txt}' written to row {row}, column {col} (centered) in '{file_path}'")

    except FileNotFoundError:
        print(f"Error: The file '{file_path}' was not found.")
    except Exception as e:
        print(f"An error occurred: {repr(e)}")


def clear_col(file_path, col, end_of_col):
    for i in range(2, end_of_col + 2):
        write_to_excel(file_path, i, col, "")


def extract_date(alert_content):
    """
    Extracts the first date in the format XX/XX/XXXX from a given string.

    :param alert_content: The input element to search for a date.
    :return: The extracted date as a string if found, otherwise None.
    """

    alert_txt = alert_content.text
    print("alert txt = ", alert_txt)

    print(alert_txt)
    date_pattern = r"\b\d{2}/\d{2}/\d{4}\b"
    match = re.search(date_pattern, alert_txt)
    return match.group() if match else None
