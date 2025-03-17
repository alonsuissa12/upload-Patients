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
from openpyxl import load_workbook
from openpyxl.styles import Alignment
from selenium.webdriver.support.ui import Select



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


import pandas as pd
from datetime import datetime


def process_excel(file_path):
    # Load the Excel file with specific column types
    df = pd.read_excel(file_path, engine='openpyxl', dtype={2: str, 4: str})

    customers = []
    print("working on:")

    for index, row in df.iterrows():
        # Check if column A (index 0) is empty
        if pd.isna(row[0]):  # Check if column A (first column) is empty
            break  # Exit the loop

        # Get values from the row
        id_value = row[2]  # ID from column C (index 2)
        date_value = row[3]  # Date from column D (index 3)
        file_name = row[4]  # File name from column E (index 4)

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
            "file": file_name,
            "rows": [index + 2]  # Initialize rows list with the first occurrence
        })

        print(
            f"           Row: {index + 2}, ID: {id_value}, Date: {date_value.day}-{date_value.month}-{date_value.year}, file: {file_name}")

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
    # Load the existing Excel file
    wb = load_workbook(file_path)
    sheet = wb.active  # Get the active sheet

    # Write the text to the specified cell
    cell = sheet.cell(row=row, column=col, value=txt)

    # Center align the text in the cell
    cell.alignment = Alignment(horizontal="center", vertical="center")

    # Save the modified file
    wb.save(file_path)
    # print(f"Text '{txt}' written to row {row}, column {col} (centered) in '{file_path}'")


def clear_col(file_path,col,end_of_col):
    for i in range(2,end_of_col+2):
        write_to_excel(file_path,i,col,"")


def clear_table(driver, start, end):
    # for j in range(start, end):
    #     id_element = driver.find_element(By.ID, "ID" + str(j))
    #     date_element = driver.find_element(By.ID, "treatmentDate" + str(j))
    #     department = driver.find_element(By.ID, "department" + str(j))
    #
    #     date_element.clear()
    #
    #     id_element.clear()
    #
    #     department = Select(department)
    #     department.select_by_value("999999999")
    #
    #     treatment_picker = WebDriverWait(driver, 10).until(
    #         EC.element_to_be_clickable((By.ID, "treatmentDescr" + str(j))))
    #     treatment_picker.click()
    #     treatment_option = WebDriverWait(driver, 2).until(
    #         EC.element_to_be_clickable((By.ID, "9999999999")))
    #     treatment_option.click()

        driver.refresh()
