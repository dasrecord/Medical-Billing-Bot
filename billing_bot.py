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

def debug_page_state(driver, reason="Debug"):
    """Save screenshot and page info for debugging"""
    try:
        timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        
        # Check browser connection first
        if not check_browser_connection(driver):
            print("Browser connection lost, skipping debug capture")
            return
        
        # Take screenshot safely
        screenshot_path = f"debug_screenshot_{timestamp}_{reason}.png"
        if safe_screenshot(driver, screenshot_path):
            print(f"Debug screenshot saved: {screenshot_path}")
        
        # Save page source snippet
        try:
            page_source = driver.page_source[:2000] + "..." if len(driver.page_source) > 2000 else driver.page_source
            source_path = f"debug_source_{timestamp}_{reason}.txt"
            with open(source_path, 'w', encoding='utf-8') as f:
                f.write(f"URL: {driver.current_url}\n")
                f.write(f"Title: {driver.title}\n")
                f.write(f"Window handles: {len(driver.window_handles)}\n")
                f.write("="*50 + "\n")
                f.write(page_source)
            print(f"Debug source saved: {source_path}")
        except Exception as e:
            print(f"Failed to save page source: {e}")
        
    except Exception as e:
        print(f"Debug save failed: {e}")

def check_browser_connection(driver):
    """Check if browser connection is still valid"""
    try:
        # Try to get current URL to test connection
        current_url = driver.current_url
        return True
    except Exception as e:
        print(f"Browser connection lost: {e}")
        return False

def safe_screenshot(driver, filename):
    """Take screenshot with error handling"""
    try:
        driver.save_screenshot(filename)
        return True
    except Exception as e:
        print(f"Screenshot failed: {e}")
        return False

def safe_close_extra_windows(driver, main_window):
    """Fastest possible window cleanup - minimal operations"""
    try:
        # Get current handles once
        handles = driver.window_handles
        
        # If only one window, we're done
        if len(handles) <= 1:
            return True
            
        # Close extra windows with minimal error checking
        for handle in handles:
            if handle != main_window:
                try:
                    driver.switch_to.window(handle)
                    driver.close()
                except:
                    pass
        
        # Switch to main window once at the end
        if main_window in driver.window_handles:
            driver.switch_to.window(main_window)
        elif driver.window_handles:
            driver.switch_to.window(driver.window_handles[0])
        
        return True
    except:
        return False

# ICD9 Code Substitution Dictionary
# Maps invalid codes to valid substitute codes
icd9_substitutes = {
    "V586": "V68",
    "5589": "558",
    "7029": "702",
    "6499": "V724",
    "4860": "486",
    "6000":"600",
    "6061":"606",
    "4900":"490",
    "6901":"690",
    "6840":"684",
    "6901":"690",
    "3391":"7840",
    "4600":"460",
    "4620":"462",
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

# set delay times in seconds - optimized for speed
short_delay = 1
long_delay = 3

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
    
    print("AUTOMATED BROWSER SETUP")
    print("=" * 50)
    
    try:
        # Step 1: Kill any existing browser processes
        print("Step 1: Closing any existing browsers...")
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
        
        print(f"Step 2: Starting {browser_name} with remote debugging...")
        os.system(browser_command)
        time.sleep(2)  # Reduced startup wait
        
        # Step 3: Connect and check if already logged in
        print(f"Step 3: {browser_name} started! Connecting...")
        
        # Connect to browser first to control it
        chrome_options = webdriver.ChromeOptions()
        chrome_options.add_experimental_option("debuggerAddress", "127.0.0.1:9222")
        
        service = ChromeService(ChromeDriverManager().install())
        driver = webdriver.Chrome(service=service, options=chrome_options)
        
        print("Checking login status...")
        driver.get("https://well-kerrisdale.kai-oscar.com/oscar")
        time.sleep(3)  # Give more time for page to load
        
        # More robust login check
        current_url = driver.current_url
        page_source = driver.page_source.lower()
        
        print(f"Current URL: {current_url}")
        
        # Check if we're already on the provider schedule page (means we're logged in)
        if "providercontrol" in current_url and "provider" in current_url:
            print("Already on provider schedule page - logged in!")
            is_login_page = False
        else:
            # Check for login page indicators
            login_indicators = [
                "username" in page_source and "password" in page_source,
                "sign in" in page_source,
                "login" in page_source,
                "log in" in page_source,
                "authentication" in page_source,
                "email" in page_source and "password" in page_source  # Some systems use email instead of username
            ]
            
            # Check if we're redirected to the new EMR system or other non-OSCAR URLs
            not_in_main_system = (
                "kaiemr" in current_url.lower() and "#/" in current_url or 
                "login" in current_url.lower()
            )
            
            # Check for actual error pages (but be more specific)
            has_actual_error = (
                ("404" in page_source and "not found" in page_source) or
                ("error code" in page_source and "404" in page_source) or
                ("page not found" in page_source)
            )
            
            is_login_page = any(login_indicators) or not_in_main_system or has_actual_error
            
            if has_actual_error:
                print("Detected actual error page (404 or similar)")
        
        if is_login_page:
            print("Not logged in. Attempting automatic login...")
            
            # Try different possible login URLs since the system may have changed
            login_urls_to_try = [
                "https://well-kerrisdale.kai-oscar.com/kaiemr/#/",
                "https://well-kerrisdale.kai-oscar.com/oscar/login.jsp",
                "https://well-kerrisdale.kai-oscar.com/oscar/",
                "https://well-kerrisdale.kai-oscar.com/login",
                "https://well-kerrisdale.kai-oscar.com/"
            ]
            
            # Try to find a working login page
            found_login_page = False
            for login_url in login_urls_to_try:
                try:
                    print(f"Trying login URL: {login_url}")
                    driver.get(login_url)
                    time.sleep(3)
                    
                    # Check if this page has login fields
                    current_page = driver.page_source.lower()
                    if ("username" in current_page or "email" in current_page) and "password" in current_page:
                        print(f"Found login page at: {login_url}")
                        found_login_page = True
                        break
                    elif "404" in current_page or "not found" in current_page:
                        print(f"404 error at {login_url}, trying next URL...")
                        continue
                        
                except Exception as url_error:
                    print(f"Error accessing {login_url}: {url_error}")
                    continue
            
            if not found_login_page:
                print("Could not find a working login page")
                print("Please login manually in the browser and then press ENTER")
                input("Press ENTER when logged in...")
                return driver
            
            username = os.getenv('OSCAR_USERNAME')
            password = os.getenv('OSCAR_PASSWORD')
            
            if username and password:
                try:
                    print("Attempting to fill login form...")
                    
                    # Try different possible username field names
                    username_field = None
                    username_selectors = [
                        (By.NAME, "username"),
                        (By.NAME, "user"),
                        (By.NAME, "userName"),
                        (By.NAME, "email"),
                        (By.NAME, "userEmail"),
                        (By.ID, "username"),
                        (By.ID, "user"),
                        (By.ID, "email"),
                        (By.XPATH, "//input[@type='text']"),
                        (By.XPATH, "//input[@type='email']"),
                        (By.XPATH, "//input[contains(@placeholder, 'username')]"),
                        (By.XPATH, "//input[contains(@placeholder, 'email')]"),
                        (By.XPATH, "//input[contains(@placeholder, 'User')]")
                    ]
                    
                    for selector in username_selectors:
                        try:
                            username_field = WebDriverWait(driver, 2).until(
                                EC.presence_of_element_located(selector)
                            )
                            print(f"Found username field using: {selector}")
                            break
                        except:
                            continue
                    
                    if not username_field:
                        print("Could not find username field")
                        input("Please login manually and press ENTER when done...")
                        return driver
                    
                    # Try different possible password field names
                    password_field = None
                    password_selectors = [
                        (By.NAME, "password"),
                        (By.NAME, "pass"),
                        (By.ID, "password"),
                        (By.ID, "pass"),
                        (By.XPATH, "//input[@type='password']")
                    ]
                    
                    for selector in password_selectors:
                        try:
                            password_field = driver.find_element(*selector)
                            print(f"Found password field using: {selector}")
                            break
                        except:
                            continue
                    
                    if not password_field:
                        print("Could not find password field")
                        input("Please login manually and press ENTER when done...")
                        return driver
                    
                    # Fill in the credentials
                    username_field.clear()
                    username_field.send_keys(username)
                    password_field.clear()
                    password_field.send_keys(password)
                    
                    # Try different login button selectors
                    login_button = None
                    button_selectors = [
                        (By.XPATH, "//input[@type='submit']"),
                        (By.XPATH, "//button[@type='submit']"),
                        (By.XPATH, "//*[@value='Sign In']"),
                        (By.XPATH, "//*[@value='Login']"),
                        (By.XPATH, "//*[@value='Log In']"),
                        (By.XPATH, "//button[contains(text(), 'Sign In')]"),
                        (By.XPATH, "//button[contains(text(), 'Login')]"),
                        (By.XPATH, "//input[contains(@value, 'Sign')]"),
                        (By.XPATH, "//input[contains(@value, 'Log')]")
                    ]
                    
                    for selector in button_selectors:
                        try:
                            login_button = driver.find_element(*selector)
                            print(f"Found login button using: {selector}")
                            break
                        except:
                            continue
                    
                    if login_button:
                        login_button.click()
                        print("Login submitted!")
                        time.sleep(5)  # Wait longer for login to process
                        
                        # Check if login was successful
                        new_url = driver.current_url
                        if "login" not in new_url.lower():
                            print("Login appears successful!")
                        else:
                            print("Login may have failed - still on login page")
                            
                    else:
                        print("Could not find login button")
                        input("Please complete login manually and press ENTER when done...")
                    
                except Exception as login_error:
                    print(f"Auto-login failed: {login_error}")
                    input("Please login manually and press ENTER when done...")
            else:
                print("No credentials found in environment variables")
                input("Please login manually and press ENTER when done...")
        else:
            print("Already logged in!")
        
        print("Navigating to provider schedule...")
        billing_date = f"https://well-kerrisdale.kai-oscar.com/oscar/provider/providercontrol.jsp?year={billing_year}&month={billing_month}&day={billing_day}&view=0&displaymode=day&dboperation=searchappointmentday&viewall=0"
        driver.get(billing_date)
        time.sleep(3)
        
        print("Ready to process appointments!")
        return driver
        
    except Exception as e:
        print(f"Automated setup failed: {str(e)}")
        print("\nMANUAL FALLBACK:")
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
        print(f"Current URL: {current_url}")
        
        # Check if we're already on the right page
        if "oscar" in current_url and "provider" in current_url:
            print("Perfect! Already on provider schedule page")
            return True
        elif "oscar" in current_url:
            print("Connected to EMR - navigating to schedule...")
            billing_date = f"https://well-kerrisdale.kai-oscar.com/oscar/provider/providercontrol.jsp?year={billing_year}&month={billing_month}&day={billing_day}&view=0&displaymode=day&dboperation=searchappointmentday&viewall=0"
            driver.get(billing_date)
            return True
        else:
            print("Not connected to EMR system")
            print("Please make sure you're logged into the EMR in the browser")
            print("and on the provider schedule page before running the bot")
            return False
            
    except Exception as e:
        print(f"Session check failed: {str(e)}")
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
    # Check browser connection first
    if not check_browser_connection(driver):
        print("❌ Browser connection lost in get_appointments")
        return []
    
    # Wait for appointments to load, then return them
    try:
        WebDriverWait(driver, long_delay).until(
            EC.presence_of_element_located((By.CLASS_NAME, "appt"))
        )
        appointments = driver.find_elements(By.CLASS_NAME, "appt")
        print(f"✅ Found {len(appointments)} appointments")
        return appointments
    except Exception as e:
        # If no appointments found, return empty list
        print(f"⚠️  No appointments found or appointments not loaded yet: {e}")
        return []

def process_appointment(driver, appointment, day_sheet_window):
    global cumulative_end_time
    global counseling_appointment_count
    encounter_window = None  # Initialize at function level
    billing_window = None   # Initialize at function level
    
    # Check browser connection before proceeding
    if not check_browser_connection(driver):
        print("❌ Browser connection lost before processing appointment")
        raise WebDriverException("Browser connection lost")
    
    try:
        # print(f"Processing appointment: {appointment.text}")
        appointment_status = appointment.find_element(By.XPATH, ".//img[1]").get_attribute("title")
    except Exception as e:
        print(f"❌ Failed to read appointment status: {e}")
        raise
    # print(f"Appointment status: {appointment_status}")

    if appointment_status in ["No Show","Billed/Verified","Billed/Signed", "Billed", "Cancelled"]:
        print(f"⏭️  Skipping - Status: {appointment_status}")
        return

    if "Track,Fast" in appointment.text:
        print("🏃 Fast track appointment. Skipping.")
        return

    print(f"📊 Processing appointment - Status: {appointment_status}")

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

    # Check browser connection before clicking encounter
    if not check_browser_connection(driver):
        print("❌ Browser connection lost before clicking encounter")
        raise WebDriverException("Browser connection lost")
    
    # Scroll to ensure the element is visible and click using JavaScript if needed
    driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", e_chart)
    time.sleep(short_delay)
    
    print(f"🎡 Opening encounter window...")
    
    try:
        # Try normal click first
        WebDriverWait(driver, short_delay).until(EC.element_to_be_clickable(e_chart))
        e_chart.click()
        print("✅ Encounter button clicked")
    except (WebDriverException, ElementClickInterceptedException):
        # Check connection before trying JavaScript click
        if not check_browser_connection(driver):
            print("❌ Browser connection lost during click attempt")
            raise WebDriverException("Browser connection lost")
        # If normal click fails, use JavaScript click
        print("🔄 Normal click failed, trying JavaScript click...")
        driver.execute_script("arguments[0].click();", e_chart)
        print("✅ JavaScript click successful")
    time.sleep(long_delay)
    
    # Check browser connection before window operations
    if not check_browser_connection(driver):
        print("❌ Browser connection lost before window switching")
        raise WebDriverException("Browser connection lost")
    
    # Wait for new window to open
    WebDriverWait(driver, long_delay).until(lambda d: len(d.window_handles) > 1)
    
    # Find the correct encounter window by checking URLs/content
    original_window = driver.current_window_handle
    encounter_window = None
    
    print(f"Total windows after click: {len(driver.window_handles)}")
    
    for handle in driver.window_handles:
        if handle != original_window:
            driver.switch_to.window(handle)
            current_url = driver.current_url
            print(f"Checking window with URL: {current_url}")
            
            # Skip old notes windows from previous patients
            if "method=allNotes" in current_url:
                print("Skipping old notes window")
                continue
            
            # Check if this is the encounter window by looking for encounter-specific keywords
            if any(keyword in current_url.lower() for keyword in ["encounter", "echart", "oscarencounter", "casemanagemententry"]):
                # Additional check: make sure it's NOT an old notes window
                if "method=setupmainencounter" in current_url.lower() or "setupmainencounter" in current_url.lower():
                    encounter_window = handle
                    print("Found encounter window by URL")
                    break
            
            # If URL doesn't help, check page content
            try:
                page_title = driver.title.lower()
                page_source_snippet = driver.page_source[:1000].lower()
                
                # Make sure it's not an old notes window
                if "encounter notes" in page_title:
                    print("Skipping old encounter notes window")
                    continue
                
                if any(keyword in page_title for keyword in ["encounter", "chart", "assessment"]) and "notes" not in page_title:
                    encounter_window = handle
                    print("Found encounter window by title")
                    break
                elif any(keyword in page_source_snippet for keyword in ["show all notes", "assessment", "plan", "progress note"]):
                    encounter_window = handle
                    print("Found encounter window by content")
                    break
                    
            except:
                continue
    
    if not encounter_window:
        # Fallback: use the last opened window
        print("Could not identify encounter window specifically, using last window")
        encounter_window = driver.window_handles[-1]
    
    driver.switch_to.window(encounter_window)
    print(f"Switched to encounter window: {driver.current_url}")

    # Quick encounter window processing
    try:
        # Wait for encounter window to fully load
        time.sleep(1.5)  # Reduced wait time for faster processing
        
        # Wait for the page to be ready
        WebDriverWait(driver, 10).until(
            lambda d: d.execute_script("return document.readyState") == "complete"
        )
        
        # Debug: Print page information
        print(f"Encounter window URL: {driver.current_url}")
        print(f"Encounter window title: {driver.title}")
        
        # Verify we're in the correct encounter window
        current_url = driver.current_url
        if "day" in current_url.lower() or "schedule" in current_url.lower():
            print("WARNING: Still appears to be in day sheet/schedule window!")
            debug_page_state(driver, "wrong_window_detected")
            return
        
        # Check for iframes that might contain the notes
        iframes = driver.find_elements(By.TAG_NAME, "iframe")
        print(f"Found {len(iframes)} iframes in encounter window")
        
        # Quick 403 check
        if "403" in driver.page_source[:300] or "Forbidden" in driver.page_source[:300]:
            print("403 error detected")
            return
        
        # Find Show All Notes button with improved detection
        show_all_notes = None
        max_retries = 3
        
        # First check browser connection
        if not check_browser_connection(driver):
            print("Browser connection lost before finding Show All Notes button")
            return
        
        for retry in range(max_retries):
            try:
                print(f"Attempting to find Show All Notes button (attempt {retry + 1}/{max_retries})")
                
                # Wait for page to stabilize
                time.sleep(1)
                
                # Check connection on each retry
                if not check_browser_connection(driver):
                    print("Browser connection lost during retry")
                    return
                
                # Try multiple selectors for the Show All Notes button
                selectors = [
                    "//*[text()='Show All Notes']",
                    "//*[contains(text(), 'Show All Notes')]",
                    "//*[contains(text(), 'Show All')]",
                    "//input[@value='Show All Notes']",
                    "//button[contains(text(), 'Show All')]",
                    "//a[contains(text(), 'Show All')]"
                ]
                
                for selector in selectors:
                    try:
                        show_all_notes = WebDriverWait(driver, 5).until(
                            EC.element_to_be_clickable((By.XPATH, selector))
                        )
                        print(f"Found Show All Notes button using selector: {selector}")
                        break
                    except:
                        continue
                
                if show_all_notes:
                    break
                    
            except Exception as e:
                print(f"Retry {retry + 1} failed: {str(e)}")
                if retry < max_retries - 1:
                    time.sleep(2)  # Wait before retrying
        
        if not show_all_notes:
            print("Show All Notes button not found after all attempts")
            
            # Try searching in iframes
            iframes = driver.find_elements(By.TAG_NAME, "iframe")
            if iframes:
                print(f"Searching in {len(iframes)} iframes for Show All Notes button...")
                for i, iframe in enumerate(iframes):
                    try:
                        print(f"Switching to iframe {i+1}")
                        driver.switch_to.frame(iframe)
                        
                        # Try to find the button in this iframe
                        show_all_notes = driver.find_element(By.XPATH, "//*[contains(text(), 'Show All')]")
                        if show_all_notes:
                            print(f"Found Show All Notes button in iframe {i+1}")
                            break
                    except:
                        print(f"Button not found in iframe {i+1}")
                    finally:
                        # Switch back to main content
                        driver.switch_to.default_content()
                        driver.switch_to.window(driver.window_handles[1])  # Back to encounter window
            
            if not show_all_notes:
                debug_page_state(driver, "show_notes_not_found")
                # Try to find any button with 'Notes' in the text as a last resort
                try:
                    print("Searching for any button containing 'Notes'...")
                    show_all_notes = WebDriverWait(driver, 3).until(
                        EC.element_to_be_clickable((By.XPATH, "//*[contains(text(), 'Notes')]"))
                    )
                    print("Found alternative Notes button")
                except:
                    print("No Notes-related button found")
                    debug_page_state(driver, "no_notes_button_found")
                    return
        
        # Click the button
        try:
            # Scroll to button and click
            driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", show_all_notes)
            time.sleep(0.5)
            show_all_notes.click()
            print("Successfully clicked Show All Notes button")
        except Exception as click_error:
            print(f"Failed to click Show All Notes button: {click_error}")
            # Try JavaScript click as backup
            try:
                driver.execute_script("arguments[0].click();", show_all_notes)
                print("Successfully clicked using JavaScript")
            except:
                print("Both normal and JavaScript clicks failed")
                return
        
    except Exception as encounter_error:
        print(f"Encounter error: {str(encounter_error)}")
        return

    # Give time for action to process
    time.sleep(1)

    # Check browser connection before window operations
    if not check_browser_connection(driver):
        print("❌ Browser connection lost before notes window processing")
        return

    try:
        # Wait for new notes window to open
        WebDriverWait(driver, long_delay).until(lambda d: len(d.window_handles) > 2)
        all_notes = driver.window_handles[2]
        driver.switch_to.window(all_notes)
        print("📄 Switched to all_notes window")

        # Check connection after window switch
        if not check_browser_connection(driver):
            print("❌ Browser connection lost after switching to notes window")
            return

        # Wait for note content to load
        print("🔍 Reading patient notes...")
        note_to_bill = WebDriverWait(driver, long_delay).until(
            EC.presence_of_element_located((By.XPATH, "/html/body/div[last()]"))
        )
        note_content = note_to_bill.text
    except Exception as notes_error:
        print(f"❌ Error accessing notes window: {notes_error}")
        # Try to recover by going back to day sheet
        try:
            driver.switch_to.window(day_sheet_window)
        except:
            pass
        return

    print("📝 Extracting diagnosis from notes...")
    # print(f"📋 Full note content (first 500 chars): {note_content[:500]}...")
    
    if "A:" in note_content and "P:" in note_content:
        diagnosis = note_content.split("A:")[1].split("P:")[0]
        print(f"🔍 Extracted diagnosis section: {diagnosis.strip()}")
    else:
        print("❌ No A: and P: sections found, using full note content")
        diagnosis = note_content
    
    # Also search for ICD codes in the full note content as backup
    full_note_diagnosis = note_content

    def extract_diagnostic_code(diagnosis):
        # More comprehensive regex patterns for ICD-9 codes
        patterns = [
            r'ICD-?9:\s*([V]?\d{1,3}\.?\d{0,2})',  # Standard format: ICD9: 034.0 or ICD-9: V123.45
            r'ICD-?9\s*([V]?\d{1,3}\.?\d{0,2})',   # Without colon: ICD9 034.0
            r'\(ICD-?9:\s*([V]?\d{1,3}\.?\d{0,2})\)', # In parentheses: (ICD9: 034.0)
            r'ICD-?9\s*code:\s*([V]?\d{1,3}\.?\d{0,2})', # With "code": ICD9 code: 034.0
        ]
        
        for pattern in patterns:
            match = re.search(pattern, diagnosis, re.IGNORECASE)
            if match:
                print(f"🎯 ICD-9 pattern matched: {pattern} -> {match.group(1)}")
                return match.group(1)
        
        print(f"❌ No ICD-9 pattern matched in: {diagnosis[:100]}...")
        return None

    icd9_code = extract_diagnostic_code(diagnosis)
    
    # If no ICD code found in diagnosis section, try the full note content
    if not icd9_code:
        print("🔍 No ICD code in diagnosis section, searching full note content...")
        icd9_code = extract_diagnostic_code(full_note_diagnosis)
    
    icd9_code = icd9_code.replace(".", "") if icd9_code else "No ICD9 code found"
    print(f"🔍 Final extracted ICD-9 code: {icd9_code}")
    
    # Apply ICD9 code substitution if needed
    original_icd9 = icd9_code
    if icd9_code in icd9_substitutes:
        icd9_code = icd9_substitutes[icd9_code]
        print(f"🔄 Substituted {original_icd9} → {icd9_code}")

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

    # Safely close windows and return to day sheet
    try:
        extra_window_count = len(driver.window_handles) - 1
        if extra_window_count > 0:
            print(f"🧹 Cleaning up {extra_window_count} extra windows...")
            
        success = safe_close_extra_windows(driver, day_sheet_window)
        if success:
            print("✅ Windows cleaned")
        else:
            print("⚠️ Window cleanup had issues")
                
    except Exception as window_error:
        print(f"Window cleanup error: {window_error}")
        # Try to get back to a valid window
        if driver.window_handles:
            driver.switch_to.window(driver.window_handles[0])

    # Check browser connection before billing operations
    if not check_browser_connection(driver):
        print("❌ Browser connection lost before billing window operations")
        return
    
    # Scroll to ensure billing button is visible and click using JavaScript if needed
    print(f"💵 Opening billing window...")
    driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", billing_button)
    time.sleep(short_delay)
    
    try:
        # Try normal click first
        WebDriverWait(driver, short_delay).until(EC.element_to_be_clickable(billing_button))
        billing_button.click()
        print("✅ Billing button clicked")
    except (WebDriverException, ElementClickInterceptedException):
        # If normal click fails, use JavaScript click
        print("🔄 Normal billing click failed, trying JavaScript...")
        driver.execute_script("arguments[0].click();", billing_button)
    time.sleep(short_delay)
    # Wait for billing window to open
    WebDriverWait(driver, long_delay).until(lambda d: len(d.window_handles) > 1)
    
    # Find the correct billing window by checking URLs/content
    current_window = driver.current_window_handle
    billing_window = None
    
    print(f"Total windows after billing click: {len(driver.window_handles)}")
    
    for handle in driver.window_handles:
        if handle != current_window:
            driver.switch_to.window(handle)
            current_url = driver.current_url
            print(f"Checking billing window with URL: {current_url}")
            
            # Check if this is the billing window by looking for billing-specific keywords
            if any(keyword in current_url.lower() for keyword in ["billing", "createbilling", "billingserviceform"]):
                billing_window = handle
                print("Found billing window by URL")
                break
            
            # If URL doesn't help, check page content
            try:
                page_title = driver.title.lower()
                page_source_snippet = driver.page_source[:1000].lower()
                
                if any(keyword in page_title for keyword in ["billing", "service"]):
                    billing_window = handle
                    print("Found billing window by title")
                    break
                elif any(keyword in page_source_snippet for keyword in ["billing form", "service code", "selectbillingform"]):
                    billing_window = handle
                    print("Found billing window by content")
                    break
                    
            except:
                continue
    
    if not billing_window:
        # Fallback: use the last opened window
        print("Could not identify billing window specifically, using last window")
        billing_window = driver.window_handles[-1]
    
    driver.switch_to.window(billing_window)
    print(f"Switched to billing window: {driver.current_url}")

    # Quick billing form processing
    try:
        # Wait for page to fully load
        print("⏳ Waiting for billing page to load...")
        WebDriverWait(driver, 10).until(
            lambda d: d.execute_script("return document.readyState") == "complete"
        )
        
        print("🔍 Looking for billing form elements...")
        select_billing_form = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.ID, "selectBillingForm"))
        )
        select_billing_form.click()
        print("✅ Clicked selectBillingForm")

        # Select service code based on counseling
        service_code_type = "counseling" if counseling else "regular"
        service_code_xpath = "//*[@id='billingFormTable']/tbody/tr[1]/td[1]/table[1]/tbody/tr[3]/td[1]/label/input" if counseling else "//*[@id='billingFormTable']/tbody/tr[1]/td[1]/table[1]/tbody/tr[2]/td[1]/label/input"
        print(f"🎯 Selecting {service_code_type} service code...")
        service_code = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, service_code_xpath))
        )
        service_code.click()
        print(f"✅ Selected {service_code_type} service code")

        # Set times and diagnosis quickly
        print(f"⏰ Setting start time to {current_start_time.strftime('%H:%M')}...")
        start_time_input = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.ID, "serviceStartTime"))
        )
        start_time_input.send_keys(current_start_time.strftime("%H:%M"))
        print(f"✅ Start time set to: {current_start_time.strftime('%H:%M')}")
        
    except Exception as billing_form_error:
        print(f"Error in billing form processing: {billing_form_error}")
        debug_page_state(driver, "billing_form_error")
        return

    end_time_input = WebDriverWait(driver, 3).until(
        EC.presence_of_element_located((By.ID, "serviceEndTime"))
    )
    end_time_input.send_keys(end_time)

    diagnosis_input = WebDriverWait(driver, 3).until(
        EC.presence_of_element_located((By.ID, "billing_1_fee_dx1"))
    )
    diagnosis_input.send_keys(icd9_code)

    continue_button = WebDriverWait(driver, 3).until(
        EC.element_to_be_clickable((By.XPATH, "//*[@value='Continue']"))
    )
    continue_button.click()

    try:
        save_bill = WebDriverWait(driver, 5).until(
            EC.element_to_be_clickable((By.XPATH, "//*[@value='Save Bill']"))
        )
        
        if not safe_mode:
            save_bill.click()
        else:
            print("Safe mode: Not saving")
            time.sleep(2)  # Reduced delay for safe mode
    except WebDriverException:
        print("Invalid billing code - logging for review")
        icd9_logger.info(f"Failed ICD9 code: {icd9_code} - Diagnosis: {diagnosis.strip()}")
        pass

    driver.switch_to.window(day_sheet_window)
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
        # Check browser connection before each appointment
        if not check_browser_connection(driver):
            print("❌ Browser connection lost during appointment processing")
            print(f"Successfully processed {processed_count} appointments before connection loss")
            break
        
        # Refresh appointment list in case of stale elements
        try:
            current_appointments = get_appointments(driver)
        except Exception as e:
            print(f"❌ Failed to refresh appointment list: {e}")
            break
        
        if processed_count >= len(current_appointments):
            # No more appointments to process
            break
            
        appointment = current_appointments[processed_count]
        
        try:
            print(f"\nProcessing appointment {processed_count + 1} of {total_appointments}")
            
            # Switch back to day sheet window before each appointment
            driver.switch_to.window(day_sheet_window)
            
            process_appointment(driver, appointment, day_sheet_window)
            processed_count += 1
            
        except StaleElementReferenceException:
            print("⚠️  StaleElementReferenceException caught. Refreshing appointments and retrying...")
            # Don't increment processed_count, try the same appointment again
            continue
        except WebDriverException as we:
            print(f"❌ WebDriver error processing appointment {processed_count + 1}: {str(we)}")
            
            # Check if it's a connection issue
            if not check_browser_connection(driver):
                print("🔌 Browser connection confirmed lost. Stopping processing.")
                break
            else:
                print("🔄 Browser still connected, skipping this appointment and continuing")
                processed_count += 1
                continue
        except Exception as e:
            print(f"❌ Error processing appointment {processed_count + 1}: {str(e)}")
            processed_count += 1  # Skip this appointment and continue
            continue
    
    print(f"\nCompleted processing {processed_count} appointments!")
    if counseling_appointment_count > 0:
        print(f"Counseling appointments processed: {counseling_appointment_count}")

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