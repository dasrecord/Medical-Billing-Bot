import time
import datetime
import os
from dotenv import load_dotenv
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import StaleElementReferenceException, NoSuchWindowException, WebDriverException, ElementClickInterceptedException
from selenium.webdriver.chrome.service import Service as ChromeService
from webdriver_manager.chrome import ChromeDriverManager
import platform
import requests
import re
import datetime

# Load the environment variables
load_dotenv()

# Set the billing date
billing_year = str(datetime.date.today().year)
billing_month = str(datetime.date.today().month)
billing_day = str(datetime.date.today().day)

# standard_appointment_length is 5 minutes
standard_appointment_length = 5

# counseling_appointment_length is 20 minutes
counseling_appointment_length = 20

# set delay times in seconds
short_delay = 3
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

# Specify the path to the Chrome binary if running on a Mac
if platform.system() == "Darwin":
    options.binary_location = "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome"

# Set up the Chrome driver
driver = webdriver.Chrome(service=ChromeService(ChromeDriverManager().install()), options=options)

def ping_dasrecord(message):
    try:
        payload = {'text': message}
        response = requests.post("https://relayproxy.vercel.app/das_record_slack", json=payload)
        if response.status_code == 200:
            print("Successfully sent message to DAS Record.")
        else:
            print("Failed to send message to DAS Record.")
    except requests.exceptions.RequestException as e:
        print(f"An error occurred while sending message to DAS Record: {e}")

def login_to_oscar(driver):
    driver.get("https://well-kerrisdale.kai-oscar.com/kaiemr/#/")
    WebDriverWait(driver, long_delay).until(
        EC.presence_of_element_located((By.ID, "okta-signin-username"))
    )
    username = driver.find_element(By.ID, "okta-signin-username")
    username.send_keys(os.getenv("USERNAME"))
    password = driver.find_element(By.ID, "okta-signin-password")
    password.send_keys(os.getenv("PASSWORD"))
    password.send_keys(Keys.RETURN)
    
    # Wait for login to complete and page to fully load
    WebDriverWait(driver, long_delay * 2).until(
        lambda driver: "signin" not in driver.current_url.lower()
    )
    
    # Additional wait for any redirects or page loads to complete
    time.sleep(long_delay)
    
    # Check if we need to navigate to the main oscar interface
    current_url = driver.current_url
    print(f"Current URL after login: {current_url}")
    
    # If we're still on the kaiemr interface, try to navigate to the oscar interface
    if "kaiemr" in current_url and "oscar/provider" not in current_url:
        print("Attempting to navigate to Oscar provider interface...")
        # Try to find and click a link to the provider interface or navigate directly
        try:
            # Look for common navigation elements that might lead to the provider interface
            provider_link = WebDriverWait(driver, short_delay).until(
                EC.element_to_be_clickable((By.PARTIAL_LINK_TEXT, "Provider"))
            )
            provider_link.click()
            print("Clicked provider link")
        except:
            # If no provider link found, try direct navigation
            print("No provider link found, attempting direct navigation...")
            pass

def navigate_to_billing_date(driver):
    billing_date = f"https://well-kerrisdale.kai-oscar.com/oscar/provider/providercontrol.jsp?year={billing_year}&month={billing_month}&day={billing_day}&view=0&displaymode=day&dboperation=searchappointmentday&viewall=0"
    print(f"Navigating to: {billing_date}")
    
    # Navigate to the billing date URL
    driver.get(billing_date)
    
    # Wait for the page to load and check if we're on the right page
    WebDriverWait(driver, long_delay * 2).until(
        lambda driver: "providercontrol.jsp" in driver.current_url
    )
    
    print(f"Successfully navigated to: {driver.current_url}")
    
    # Wait for the appointment schedule to load
    WebDriverWait(driver, long_delay).until(
        EC.presence_of_element_located((By.TAG_NAME, "body"))
    )
    
    # Additional wait for dynamic content to load
    time.sleep(short_delay)

def get_appointments(driver):
    # Wait for appointments to load, then return them
    try:
        WebDriverWait(driver, long_delay).until(
            EC.presence_of_element_located((By.CLASS_NAME, "appt"))
        )
        return driver.find_elements(By.CLASS_NAME, "appt")
    except:
        # If no appointments found, return empty list
        print("No appointments found or appointments not loaded yet.")
        return []

def process_appointment(driver, appointment, day_sheet_window):
    global cumulative_end_time
    global counseling_appointment_count
    # print(f"Processing appointment: {appointment.text}")
    appointment_status = appointment.find_element(By.XPATH, ".//img[1]").get_attribute("title")
    # print(f"Appointment status: {appointment_status}")

    if appointment_status in ["Billed/Verified","Billed/Signed", "Billed", "Cancelled"]:
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

    # Scroll to ensure the element is visible and click using JavaScript if needed
    driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", e_chart)
    time.sleep(short_delay)
    
    try:
        # Try normal click first
        WebDriverWait(driver, short_delay).until(EC.element_to_be_clickable(e_chart))
        e_chart.click()
    except (WebDriverException, ElementClickInterceptedException):
        # If normal click fails, use JavaScript click
        print("Normal click failed, trying JavaScript click...")
        driver.execute_script("arguments[0].click();", e_chart)
    time.sleep(long_delay)
    # Wait for new window to open
    WebDriverWait(driver, long_delay).until(lambda d: len(d.window_handles) > 1)
    encounter = driver.window_handles[1]
    driver.switch_to.window(encounter)
    print("Switched to encounter window")

    # Wait for Show All Notes button to be present and clickable
    show_all_notes = WebDriverWait(driver, long_delay).until(
        EC.element_to_be_clickable((By.XPATH, "//*[text()='Show All Notes']"))
    )
    show_all_notes.click()
    time.sleep(short_delay)
    print("Clicked Show All Notes button")

    # Wait for new notes window to open
    WebDriverWait(driver, long_delay).until(lambda d: len(d.window_handles) > 2)
    all_notes = driver.window_handles[2]
    driver.switch_to.window(all_notes)
    print("Switched to all_notes window")

    # Wait for note content to load
    note_to_bill = WebDriverWait(driver, long_delay).until(
        EC.presence_of_element_located((By.XPATH, "/html/body/div[last()]"))
    )
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

    # Scroll to ensure billing button is visible and click using JavaScript if needed
    driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", billing_button)
    time.sleep(short_delay)
    
    try:
        # Try normal click first
        WebDriverWait(driver, short_delay).until(EC.element_to_be_clickable(billing_button))
        billing_button.click()
    except (WebDriverException, ElementClickInterceptedException):
        # If normal click fails, use JavaScript click
        print("Normal billing button click failed, trying JavaScript click...")
        driver.execute_script("arguments[0].click();", billing_button)
    time.sleep(short_delay)
    # Wait for billing window to open
    WebDriverWait(driver, long_delay).until(lambda d: len(d.window_handles) > 1)
    billing_window = driver.window_handles[1]
    driver.switch_to.window(billing_window)
    print("Switched to billing_window")

    # Wait for billing form to load
    select_billing_form = WebDriverWait(driver, long_delay).until(
        EC.element_to_be_clickable((By.ID, "selectBillingForm"))
    )
    select_billing_form.click()
    print("Clicked selectBillingForm")

    if counseling:
        service_code = WebDriverWait(driver, long_delay).until(
            EC.element_to_be_clickable((By.XPATH, "//*[@id='billingFormTable']/tbody/tr[1]/td[1]/table[1]/tbody/tr[3]/td[1]/label/input"))
        )
    else:
        service_code = WebDriverWait(driver, long_delay).until(
            EC.element_to_be_clickable((By.XPATH, "//*[@id='billingFormTable']/tbody/tr[1]/td[1]/table[1]/tbody/tr[2]/td[1]/label/input"))
        )
    service_code.click()
    print("Selected service code")

    start_time_input = WebDriverWait(driver, long_delay).until(
        EC.presence_of_element_located((By.ID, "serviceStartTime"))
    )
    start_time_input.send_keys(current_start_time.strftime("%H:%M"))
    print("Set start time")

    end_time_input = WebDriverWait(driver, long_delay).until(
        EC.presence_of_element_located((By.ID, "serviceEndTime"))
    )
    end_time_input.send_keys(end_time)
    print("Set end time")

    diagnosis_input = WebDriverWait(driver, long_delay).until(
        EC.presence_of_element_located((By.ID, "billing_1_fee_dx1"))
    )
    diagnosis_input.send_keys(icd9_code)
    print("Entered ICD-9 code")

    continue_button = WebDriverWait(driver, long_delay).until(
        EC.element_to_be_clickable((By.XPATH, "//*[@value='Continue']"))
    )
    continue_button.click()
    time.sleep(short_delay)
    print("Clicked continue button")

    try:
        save_bill = WebDriverWait(driver, long_delay).until(
            EC.element_to_be_clickable((By.XPATH, "//*[@value='Save Bill']"))
        )
        
        if not safe_mode:
            save_bill.click()
            print("Clicked save bill button")
        else:
            print("Safe mode is on. Not saving the bill.")
            time.sleep(999)
    except WebDriverException:
        print("Save bill button not found. Skipping...")
        pass

    driver.switch_to.window(day_sheet_window)
    time.sleep(short_delay)
    print("Switched back to day sheet window")

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
    # ping_dasrecord("Billing bot started.")
    login_to_oscar(driver)
    navigate_to_billing_date(driver)
    day_sheet_window = driver.current_window_handle
    process_appointments(driver, day_sheet_window)
    driver.quit()
    ping_dasrecord("Billing bot completed successfully.")

if __name__ == "__main__":
    main()