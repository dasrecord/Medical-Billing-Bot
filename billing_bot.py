import time
import datetime
import os
import logging
from dotenv import load_dotenv
import undetected_chromedriver as uc
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import StaleElementReferenceException, NoSuchWindowException, WebDriverException, ElementClickInterceptedException, NoSuchElementException
from selenium.webdriver.chrome.service import Service as ChromeService
from webdriver_manager.chrome import ChromeDriverManager
import platform
import requests
import re
import datetime
import random
import json

# ICD9 Code Substitution Dictionary
# Maps invalid codes to valid substitute codes
icd9_substitutes = {
    "V586": "V68",
    "5589": "558",
    "7029": "702",
    "6499": "V724",
    "4860": "486",
    "6000":"600",
    "6061":"606"
    # Add more substitutions as needed based on failed_icd9_codes.log
    # Format: "invalid_code": "valid_substitute"
}

# Advanced Anti-Detection Configuration
class BrowserFingerprint:
    """Generate realistic browser fingerprints to avoid detection"""
    
    @staticmethod
    def get_random_user_agent():
        """Return a random realistic user agent"""
        user_agents = [
            "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/118.0.0.0 Safari/537.36",
            "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/117.0.0.0 Safari/537.36", 
            "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/116.0.0.0 Safari/537.36",
            "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/118.0.0.0 Safari/537.36",
            "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/117.0.0.0 Safari/537.36"
        ]
        return random.choice(user_agents)
    
    @staticmethod
    def get_viewport_size():
        """Return a random realistic viewport size"""
        viewports = [
            (1920, 1080), (1366, 768), (1440, 900), (1536, 864), (1680, 1050)
        ]
        return random.choice(viewports)
    
    @staticmethod
    def get_timezone():
        """Return a realistic timezone"""
        timezones = ["America/New_York", "America/Chicago", "America/Denver", "America/Los_Angeles", "America/Vancouver"]
        return random.choice(timezones)

# Load the environment variables
load_dotenv()

# Set up logging for failed ICD9 codes only
# Create a custom logger to avoid capturing other library logs
icd9_logger = logging.getLogger('failed_icd9')
icd9_logger.setLevel(logging.INFO)

# Create file handler
file_handler = logging.FileHandler('failed_icd9_codes.log')
file_handler.setLevel(logging.INFO)

# Create formatter
formatter = logging.Formatter('%(asctime)s - %(message)s', datefmt='%Y-%m-%d %H:%M:%S')
file_handler.setFormatter(formatter)

# Add handler to logger
icd9_logger.addHandler(file_handler)

# Prevent propagation to root logger to avoid other logs
icd9_logger.propagate = False

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

# Set the number of runs - will be dynamically set based on appointments found
runs = None  # Will be set automatically based on appointment count

# set safe_mode (default = True)
safe_mode = False

# set headless mode (default = False) - MUST be False for EMR compatibility
headless_mode = False

# webdriver options
options = webdriver.ChromeOptions()
if headless_mode:
    options.add_argument("--headless")

# Add additional Chrome options for better stability and compatibility
options.add_argument("--no-sandbox")
options.add_argument("--disable-dev-shm-usage")
options.add_argument("--disable-gpu")
options.add_argument("--disable-web-security")
options.add_argument("--disable-features=VizDisplayCompositor")
options.add_argument("--allow-running-insecure-content")
options.add_argument("--disable-features=VizDisplayCompositor,VizServiceDisplay")
options.add_argument("--disable-cors")
options.add_argument("--disable-site-isolation-trials")
options.add_argument("--disable-features=BlockInsecurePrivateNetworkRequests")
options.add_argument("--ignore-certificate-errors")
options.add_argument("--ignore-ssl-errors")
options.add_argument("--ignore-certificate-errors-spki-list")
options.add_argument("--ignore-ssl-errors-ignore-cert-validity")
options.add_argument("--allow-running-insecure-content")
options.add_argument("--disable-extensions")
options.add_argument("--disable-plugins")
options.add_argument("--disable-images")

# Specify the path to the Chrome binary if running on a Mac
if platform.system() == "Darwin":
    options.binary_location = "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome"

def setup_chrome_driver():
    """Automatically start browser with remote debugging and connect"""
    
    print("🤖 AUTOMATED BROWSER SETUP")
    print("=" * 50)
    
    try:
        # Step 1: Kill any existing browser processes
        print("1️⃣ Closing any existing browsers...")
        os.system("pkill -f 'Google Chrome' 2>/dev/null")
        os.system("pkill -f 'Brave Browser' 2>/dev/null")
        os.system("pkill -f 'chrome' 2>/dev/null")
        time.sleep(2)
        
        # Step 2: Choose and start browser with remote debugging
        brave_path = "/Applications/Brave Browser.app/Contents/MacOS/Brave Browser"
        chrome_path = "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome"
        
        browser_command = None
        if os.path.exists(brave_path):
            browser_command = f'"{brave_path}" --remote-debugging-port=9222 > /dev/null 2>&1 &'
            browser_name = "Brave"
        elif os.path.exists(chrome_path):
            browser_command = f'"{chrome_path}" --remote-debugging-port=9222 > /dev/null 2>&1 &'
            browser_name = "Chrome"
        else:
            raise Exception("Neither Brave nor Chrome found in Applications")
        
        print(f"2️⃣ Starting {browser_name} with remote debugging...")
        os.system(browser_command)
        time.sleep(3)  # Give browser time to start
        
        # Step 3: Connect and check if already logged in
        print(f"3️⃣ {browser_name} started! Connecting and checking login status...")
        
        # Connect to browser first to control it
        chrome_options = webdriver.ChromeOptions()
        chrome_options.add_experimental_option("debuggerAddress", "127.0.0.1:9222")
        
        service = ChromeService(ChromeDriverManager().install())
        driver = webdriver.Chrome(service=service, options=chrome_options)
        
        print("🔍 Checking if already logged into EMR...")
        driver.get("https://well-kerrisdale.kai-oscar.com/oscar")
        time.sleep(3)
        
        # Check if already logged in by looking for login form vs main interface
        current_url = driver.current_url
        page_source = driver.page_source.lower()
        
        if "username" in page_source and "password" in page_source and "login" in current_url.lower():
            # Not logged in - need to login
            print("🔐 Not logged in. Attempting automatic login...")
            
            username = os.getenv('OSCAR_USERNAME')
            password = os.getenv('OSCAR_PASSWORD')
            
            if username and password:
                try:
                    # Auto-fill credentials
                    username_field = WebDriverWait(driver, 5).until(
                        EC.presence_of_element_located((By.NAME, "username"))
                    )
                    password_field = driver.find_element(By.NAME, "password")
                    
                    username_field.clear()
                    username_field.send_keys(username)
                    password_field.clear()
                    password_field.send_keys(password)
                    
                    # Submit login
                    login_button = driver.find_element(By.XPATH, "//input[@type='submit' or @value='Sign In' or @value='Login']")
                    login_button.click()
                    
                    print("✅ Login submitted! Waiting for login to complete...")
                    time.sleep(5)
                    
                except Exception as login_error:
                    print(f"⚠️ Auto-login failed: {login_error}")
                    print("🔐 Please complete login manually")
                    input("Press ENTER when logged in...")
            else:
                print("⚠️ No credentials in .env file - please login manually")
                input("Press ENTER when logged in...")
        else:
            print("✅ Already logged into EMR!")
        
        # Navigate to provider schedule
        print("📅 Navigating to provider schedule...")
        billing_date = f"https://well-kerrisdale.kai-oscar.com/oscar/provider/providercontrol.jsp?year={billing_year}&month={billing_month}&day={billing_day}&view=0&displaymode=day&dboperation=searchappointmentday&viewall=0"
        driver.get(billing_date)
        time.sleep(3)
        
        print("🎯 Ready to process appointments!")
        return driver
        
    except Exception as e:
        print(f"❌ Automated setup failed: {str(e)}")
        print("\n🔧 MANUAL FALLBACK:")
        print("If automated startup failed, you can manually run:")
        print("1. Close all browsers")
        print("2. Run: /Applications/Brave\\ Browser.app/Contents/MacOS/Brave\\ Browser --remote-debugging-port=9222")
        print("3. Login to EMR")
        print("4. Restart this bot")
        
        raise WebDriverException(f"Browser setup failed: {str(e)}")

# Set up the driver
driver = setup_chrome_driver()

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
    """Skip login - we're connecting to existing authenticated session"""
    
    try:
        current_url = driver.current_url
        print(f"📍 Current URL: {current_url}")
        
        # Check if we're already on the right page
        if "oscar" in current_url and "provider" in current_url:
            print("✅ Perfect! Already on provider schedule page")
            return True
        elif "oscar" in current_url:
            print("✅ Connected to EMR - navigating to schedule...")
            billing_date = f"https://well-kerrisdale.kai-oscar.com/oscar/provider/providercontrol.jsp?year={billing_year}&month={billing_month}&day={billing_day}&view=0&displaymode=day&dboperation=searchappointmentday&viewall=0"
            driver.get(billing_date)
            return True
        else:
            print("❌ Not connected to EMR system")
            print("Please make sure you're logged into the EMR in the browser")
            print("and on the provider schedule page before running the bot")
            return False
            
    except Exception as e:
        print(f"❌ Session check failed: {str(e)}")
        return False

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

    if appointment_status in ["No Show","Billed/Verified","Billed/Signed", "Billed", "Cancelled"]:
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

    # Debug: Check what's actually in the encounter window
    print("🔍 Debugging encounter window content...")
    try:
        # Wait a bit for the page to load
        time.sleep(3)
        
        # Check if we got a 403 or other error
        page_source_snippet = driver.page_source[:500]
        if "403" in page_source_snippet or "Forbidden" in page_source_snippet:
            print("❌ 403 Forbidden error detected in encounter window")
            print("🔄 The existing session approach didn't fully bypass detection")
            return
        
        # Check current URL
        current_url = driver.current_url
        print(f"📍 Encounter window URL: {current_url}")
        
        # Look for Show All Notes button with multiple strategies
        print("🔍 Looking for 'Show All Notes' button...")
        
        # Strategy 1: Exact text match
        show_all_notes = None
        try:
            show_all_notes = WebDriverWait(driver, 5).until(
                EC.element_to_be_clickable((By.XPATH, "//*[text()='Show All Notes']"))
            )
            print("✅ Found 'Show All Notes' button (exact text)")
        except:
            print("❌ Exact text 'Show All Notes' not found")
        
        # Strategy 2: Partial text match
        if not show_all_notes:
            try:
                show_all_notes = WebDriverWait(driver, 5).until(
                    EC.element_to_be_clickable((By.XPATH, "//*[contains(text(), 'Show All')]"))
                )
                print("✅ Found button with 'Show All' text")
            except:
                print("❌ 'Show All' partial text not found")
        
        # Strategy 3: Look for any note-related buttons
        if not show_all_notes:
            try:
                show_all_notes = WebDriverWait(driver, 5).until(
                    EC.element_to_be_clickable((By.XPATH, "//*[contains(text(), 'Notes') or contains(text(), 'notes')]"))
                )
                print("✅ Found note-related button")
            except:
                print("❌ No note-related buttons found")
        
        # Strategy 4: Debug - show all clickable elements
        if not show_all_notes:
            print("🔍 Debugging: Looking for all clickable elements...")
            try:
                clickable_elements = driver.find_elements(By.XPATH, "//a | //button | //input[@type='button'] | //input[@type='submit']")
                print(f"Found {len(clickable_elements)} clickable elements:")
                for i, element in enumerate(clickable_elements[:10]):  # Show first 10
                    try:
                        text = element.text.strip()
                        if text:
                            print(f"  {i+1}: '{text}'")
                    except:
                        pass
                        
                # Try to find by common patterns
                for pattern in ["Notes", "Show", "All", "Clinical", "Chart"]:
                    try:
                        element = driver.find_element(By.XPATH, f"//*[contains(text(), '{pattern}')]")
                        print(f"📍 Found element with '{pattern}': {element.text}")
                    except:
                        pass
                        
            except Exception as debug_error:
                print(f"Debug error: {debug_error}")
        
        if show_all_notes:
            show_all_notes.click()
            time.sleep(short_delay)
            print("Clicked Show All Notes button")
        else:
            print("❌ Could not find 'Show All Notes' button")
            print("🔄 The encounter window interface may have changed")
            # Take a screenshot for debugging if possible
            try:
                driver.save_screenshot("/tmp/encounter_debug.png")
                print("📸 Screenshot saved to /tmp/encounter_debug.png")
            except:
                pass
            return
            
    except Exception as encounter_error:
        print(f"❌ Error in encounter window: {str(encounter_error)}")
        return

    # Give time for the Show All Notes action to process
    time.sleep(short_delay)

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
    
    # Apply ICD9 code substitution if needed
    original_icd9 = icd9_code
    if icd9_code in icd9_substitutes:
        icd9_code = icd9_substitutes[icd9_code]
        print(f"Substituted {original_icd9} with {icd9_code}")

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
        print("Save bill button not found. Invalid billing code. Skipping...")
        # Log the failed ICD9 code and diagnosis for creating substitute codes
        icd9_logger.info(f"Failed ICD9 code: {icd9_code} - Diagnosis: {diagnosis.strip()}")
        pass

    driver.switch_to.window(day_sheet_window)
    time.sleep(short_delay)
    print("Switched back to day sheet window")

def process_appointments(driver, day_sheet_window):
    global cumulative_end_time
    global counseling_appointment_count

    cumulative_end_time = None
    counseling_appointment_count = 0

    # Get initial appointments and set runs based on count
    appointments = get_appointments(driver)
    total_appointments = len(appointments)
    print(f"Found {total_appointments} appointments. Processing each one...")
    
    processed_count = 0
    
    # Process each appointment, handling stale element exceptions
    while processed_count < total_appointments:
        # Refresh appointment list in case of stale elements
        current_appointments = get_appointments(driver)
        
        if processed_count >= len(current_appointments):
            # No more appointments to process
            break
            
        appointment = current_appointments[processed_count]
        
        try:
            print(f"\nProcessing appointment {processed_count + 1} of {total_appointments}")
            process_appointment(driver, appointment, day_sheet_window)
            processed_count += 1
            
        except StaleElementReferenceException:
            print("StaleElementReferenceException caught. Refreshing appointments and retrying...")
            # Don't increment processed_count, try the same appointment again
            continue
        except Exception as e:
            print(f"Error processing appointment {processed_count + 1}: {str(e)}")
            processed_count += 1  # Skip this appointment and continue
            continue
    
    print(f"\n✅ Completed processing {processed_count} appointments!")
    if counseling_appointment_count > 0:
        print(f"📋 Counseling appointments processed: {counseling_appointment_count}")

def main():
    # ping_dasrecord("Billing bot started.")
    print("Starting Medical Billing Bot...")
    print("The bot will open a browser for you to manually login and bypass any security checks.")
    
    login_to_oscar(driver)
    
    # Skip the separate navigate function since user already navigated manually
    print("Bot taking control for automated billing processing...")
    
    day_sheet_window = driver.current_window_handle
    process_appointments(driver, day_sheet_window)
    driver.quit()
    ping_dasrecord("Billing bot completed successfully.")

if __name__ == "__main__":
    main()