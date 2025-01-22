import time
import datetime
import os
from dotenv import load_dotenv
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.common.exceptions import StaleElementReferenceException, NoSuchWindowException, WebDriverException
from selenium.webdriver.chrome.service import Service as ChromeService
from webdriver_manager.chrome import ChromeDriverManager
import re

# Load the environment variables
load_dotenv()

# Set the billing date
billing_year = '2025'
billing_month = '01'
billing_day = '22'

# standard_appointment_length is 5 minutes
standard_appointment_length = 5

# counseling_appointment_length is 20 minutes
counseling_appointment_length = 20

# set delay times in seconds
short_delay = 4
long_delay = 6

# Set the number of runs
runs = 10

# set safe_mode (default = True)
safe_mode = False

# set headless mode (default = False)
headless_mode = True

# webdriver options
options = webdriver.ChromeOptions()
if headless_mode:
    options.add_argument("--headless")


# Set up the Chrome driver
driver = webdriver.Chrome(service=ChromeService(ChromeDriverManager().install()), options=options)


def login_to_oscar(driver):
    driver.get("https://well-kerrisdale.kai-oscar.com/oscar/")
    username = driver.find_element(By.ID, "username")
    username.send_keys(os.getenv("USERNAME"))
    password = driver.find_element(By.ID, "password")
    password.send_keys(os.getenv("PASSWORD"))
    pin = driver.find_element(By.ID, "pin")
    pin.send_keys(os.getenv("PIN"))
    pin.send_keys(Keys.RETURN)
    time.sleep(short_delay)

def navigate_to_billing_date(driver):
    billing_date = f"https://well-kerrisdale.kai-oscar.com/oscar/provider/providercontrol.jsp?year={billing_year}&month={billing_month}&day={billing_day}&view=0&displaymode=day&dboperation=searchappointmentday&viewall=0"
    driver.get(billing_date)

def get_appointments(driver):
    return driver.find_elements(By.CLASS_NAME, "appt")

def process_appointment(driver, appointment, day_sheet_window):
    global cumulative_end_time
    global counseling_appointment_count
    # print(f"Processing appointment: {appointment.text}")
    appointment_status = appointment.find_element(By.XPATH, ".//img[1]").get_attribute("title")
    # print(f"Appointment status: {appointment_status}")

    if appointment_status in ["Billed/Signed", "Billed", "Cancelled"]:
        print("Appointment already billed or cancelled.")
        return

    if "Track,Fast" in appointment.text:
        print("Fast track appointment. Skipping.")
        return

    counseling = False
    e_chart = appointment.find_element(By.XPATH, ".//a[contains(@title, 'Encounter')]")
    billing_button = appointment.find_element(By.XPATH, ".//a[contains(@title, 'Billing')]")
    master_record_button = appointment.find_element(By.XPATH, ".//a[contains(@title, 'Master Record')]")
    prescriptions_button = appointment.find_element(By.XPATH, ".//a[contains(@title, 'Prescriptions')]")

    start_time = appointment.find_element(By.XPATH, "./preceding-sibling::td[@align='RIGHT'][1]/a[1]").get_attribute("title").split(" - ")[0]
    start_time = datetime.datetime.strptime(start_time, "%I:%M%p").strftime("%H:%M")
    current_start_time = datetime.datetime.strptime(start_time, "%H:%M")
    print(f"Start time: {start_time}")

    if cumulative_end_time and cumulative_end_time > current_start_time:
        current_start_time = cumulative_end_time
        print(f"Adjusted start time: {current_start_time.strftime('%H:%M')}")

    e_chart.click()
    time.sleep(long_delay)
    encounter = driver.window_handles[1]
    driver.switch_to.window(encounter)
    print("Switched to encounter window")

    show_all_notes = driver.find_element(By.XPATH, "//*[text()='Show All Notes']")
    show_all_notes.click()
    time.sleep(short_delay)
    print("Clicked Show All Notes button")

    all_notes = driver.window_handles[2]
    driver.switch_to.window(all_notes)
    print("Switched to all_notes window")

    note_to_bill = driver.find_element(By.XPATH, "/html/body/div[last()]")
    note_content = note_to_bill.text

    if "A:" in note_content and "P:" in note_content:
        diagnosis = note_content.split("A:")[1].split("P:")[0]
    else:
        diagnosis = "No diagnosis found"

    def extract_diagnostic_code(diagnosis):
        match = re.search(r'ICD-?9: (V?\d+\.?\d)\d?', diagnosis)
        return match.group(1) if match else None

    icd9_code = extract_diagnostic_code(diagnosis)
    icd9_code = icd9_code.replace(".", "") if icd9_code else "No ICD9 code found"
    print(f"Extracted ICD-9 code: {icd9_code}")

    if "#C" in note_content:
        appointment_length = counseling_appointment_length
        counseling_appointment_count += 1
        counseling = True
    else:
        appointment_length = standard_appointment_length

    end_time = (current_start_time + datetime.timedelta(minutes=appointment_length)).strftime("%H:%M")
    cumulative_end_time = current_start_time + datetime.timedelta(minutes=appointment_length)
    print(f"End time: {end_time}")
    print(f"Cumulative end time: {cumulative_end_time.strftime('%H:%M')}")

    driver.close()
    driver.switch_to.window(encounter)
    driver.close()
    driver.switch_to.window(day_sheet_window)
    print("Switched back to day_sheet_window")

    billing_button.click()
    time.sleep(short_delay)
    billing_window = driver.window_handles[1]
    driver.switch_to.window(billing_window)
    print("Switched to billing_window")

    select_billing_form = driver.find_element(By.ID, "selectBillingForm")
    select_billing_form.click()
    print("Clicked selectBillingForm")

    if counseling:
        service_code = driver.find_element(By.XPATH, "//*[@id='billingFormTable']/tbody/tr[1]/td[1]/table[1]/tbody/tr[3]/td[1]/label/input")
    else:
        service_code = driver.find_element(By.XPATH, "//*[@id='billingFormTable']/tbody/tr[1]/td[1]/table[1]/tbody/tr[2]/td[1]/label/input")
    service_code.click()
    print("Selected service code")

    start_time_input = driver.find_element(By.ID, "serviceStartTime")
    start_time_input.send_keys(current_start_time.strftime("%H:%M"))
    print("Set start time")

    end_time_input = driver.find_element(By.ID, "serviceEndTime")
    end_time_input.send_keys(end_time)
    print("Set end time")

    diagnosis_input = driver.find_element(By.ID, "billing_1_fee_dx1")
    diagnosis_input.send_keys(icd9_code)
    print("Entered ICD-9 code")

    continue_button = driver.find_element(By.XPATH, "//*[@value='Continue']")
    continue_button.click()
    time.sleep(short_delay)
    print("Clicked continue button")

    save_bill = driver.find_element(By.XPATH, "//*[@value='Save Bill']")
    if not safe_mode:
        save_bill.click()
        print("Clicked save bill button")
    else:
        print("Safe mode is on. Not saving the bill.")
        time.sleep(999)

    driver.switch_to.window(day_sheet_window)
    time.sleep(short_delay)
    print("Switched back to day_sheet_window after saving the bill")

def process_appointments(driver, day_sheet_window):
    global cumulative_end_time
    global runs
    global counseling_appointment_count

    cumulative_end_time = None
    counseling_appointment_count = 0

    for _ in range(runs):
        appointments = get_appointments(driver)
        print(f"Found {len(appointments)} appointments.")

        for appointment in appointments:
            try:
                process_appointment(driver, appointment, day_sheet_window)
            except StaleElementReferenceException:
                # print("StaleElementReferenceException caught. Refetching appointments and retrying...")
                appointments = get_appointments(driver)
                continue

def main():
    login_to_oscar(driver)
    navigate_to_billing_date(driver)
    day_sheet_window = driver.current_window_handle
    process_appointments(driver, day_sheet_window)
    driver.quit()
    print("Script completed successfully.")

if __name__ == "__main__":
    main()