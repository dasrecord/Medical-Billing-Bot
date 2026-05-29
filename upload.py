"""
upload.py — Standalone uploader for dr-bill.ca

Finds the most recent .xlsx billing export in the working directory and
uploads it to dr-bill.ca via Selenium.

Run directly:
    python upload.py

Or pass a specific file:
    python upload.py path/to/file.xlsx
"""

import os
import sys
import glob
import time
import shutil
import socket
import subprocess
import tempfile

from dotenv import load_dotenv
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import WebDriverException
from selenium.webdriver.chrome.service import Service as ChromeService
from webdriver_manager.chrome import ChromeDriverManager

load_dotenv()


def _detect_brave_version():
    """Return the installed Brave version string, or None."""
    import subprocess, sys
    candidates = (
        ["/usr/bin/brave-browser", "/usr/bin/brave", "/snap/bin/brave"]
        if sys.platform != "win32" else []
    )
    for p in candidates:
        if __import__("os").path.exists(p):
            try:
                out = subprocess.check_output([p, "--version"], stderr=subprocess.DEVNULL, text=True)
                parts = out.strip().split()
                return parts[-1].split(".")[0] if parts else None
            except Exception:
                return None
    return None

_brave_process = None


def cleanup_brave():
    global _brave_process
    if _brave_process is not None:
        try:
            _brave_process.terminate()
            _brave_process.wait(timeout=5)
        except Exception:
            try:
                _brave_process.kill()
            except Exception:
                pass
        _brave_process = None


def setup_driver(headless=False):
    """Start a clean Brave instance and connect via ChromeDriver."""
    global _brave_process

    if sys.platform == "win32":
        brave_candidates = [
            os.path.join(os.environ.get("LOCALAPPDATA", ""), "BraveSoftware", "Brave-Browser", "Application", "brave.exe"),
            r"C:\Program Files\BraveSoftware\Brave-Browser\Application\brave.exe",
            r"C:\Program Files (x86)\BraveSoftware\Brave-Browser\Application\brave.exe",
        ]
        brave_path = next((p for p in brave_candidates if os.path.exists(p)), None)
    else:
        brave_path = "/Applications/Brave Browser.app/Contents/MacOS/Brave Browser"
        if not os.path.exists(brave_path):
            brave_path = None

    if not brave_path:
        raise RuntimeError("Brave Browser not found. Please install Brave Browser.")

    brave_clean_dir = os.path.join(tempfile.gettempdir(), "brave_clean_upload")
    if os.path.exists(brave_clean_dir):
        shutil.rmtree(brave_clean_dir, ignore_errors=True)

    # Kill any previous upload bot instance
    if sys.platform != "win32":
        os.system(f"pkill -f 'user-data-dir={brave_clean_dir}' 2>/dev/null")
    time.sleep(0.5)

    brave_args = [
        brave_path,
        "--remote-debugging-port=9223",  # separate port so it doesn't clash with billing_bot
        f"--user-data-dir={brave_clean_dir}",
        "--window-size=1280,900",
    ]
    if headless:
        brave_args += ["--headless=new", "--disable-gpu"]

    _brave_process = subprocess.Popen(brave_args, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)

    # Wait for the debugging port to be ready
    for attempt in range(10):
        try:
            sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
            sock.settimeout(1)
            if sock.connect_ex(("127.0.0.1", 9223)) == 0:
                sock.close()
                break
        except Exception:
            pass
        time.sleep(1)
    else:
        raise RuntimeError("Brave failed to open the remote debugging port.")

    options = webdriver.ChromeOptions()
    options.add_experimental_option("debuggerAddress", "127.0.0.1:9223")
    options.add_argument("--disable-blink-features=AutomationControlled")

    brave_ver   = _detect_brave_version()
    mgr_kwargs  = {"driver_version": brave_ver} if brave_ver else {}
    driver_path = ChromeDriverManager(**mgr_kwargs).install()
    if "THIRD_PARTY_NOTICES" in driver_path:
        driver_dir = os.path.dirname(driver_path)
        driver_name = "chromedriver.exe" if sys.platform == "win32" else "chromedriver"
        driver_path = os.path.join(driver_dir, driver_name)

    if sys.platform != "win32" and not os.access(driver_path, os.X_OK):
        os.chmod(driver_path, 0o755)

    service = ChromeService(driver_path)
    driver = webdriver.Chrome(service=service, options=options)
    driver.set_page_load_timeout(30)
    driver.execute_script("Object.defineProperty(navigator, 'webdriver', {get: () => undefined})")
    return driver


def upload(driver, xlsx_path=None):
    """
    Upload a billing .xlsx to dr-bill.ca.

    Args:
        driver:     A live Selenium WebDriver instance.
        xlsx_path:  Absolute path to the .xlsx file.  If None, the most
                    recently modified .xlsx in the current directory is used.

    Returns:
        True on success, False on failure.
    """
    # Resolve the file to upload
    if xlsx_path is None:
        xlsx_files = glob.glob(os.path.join(os.getcwd(), "*.xlsx"))
        if not xlsx_files:
            print("❌ No .xlsx files found to upload")
            return False
        #newest_file
        xlsx_path = os.path.abspath(max(xlsx_files, key=os.path.getmtime))
        #oldest_file
        #xlsx_path = os.path.abspath(min(xlsx_files, key=os.path.getmtime))

    xlsx_path = os.path.abspath(xlsx_path)
    if not os.path.exists(xlsx_path):
        print(f"❌ File not found: {xlsx_path}")
        return False

    print(f"📂 Will upload: {xlsx_path}")

    drbill_username = os.getenv("DRBILL_USERNAME")
    drbill_password = os.getenv("DRBILL_PASSWORD")
    if not drbill_username or not drbill_password:
        print("❌ DRBILL_USERNAME or DRBILL_PASSWORD not set in .env")
        return False

    # Reach the login page (primary URL → fallback)
    login_base = None
    for base_url in ["https://app.dr-bill.ca", "https://secure.dr-bill.ca"]:
        try:
            print(f"🌐 Navigating to {base_url}...")
            driver.get(base_url)
            WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.ID, "input27"))
            )
            login_base = base_url
            print("✅ Reached dr-bill.ca login page")
            break
        except Exception:
            print(f"⚠️ Could not reach {base_url}, trying fallback...")

    if not login_base:
        print("❌ Could not reach dr-bill.ca login page via any URL")
        return False

    # Log in
    try:
        username_field = driver.find_element(By.ID, "input27")
        username_field.clear()
        username_field.send_keys(drbill_username)

        password_field = driver.find_element(By.ID, "input35")
        password_field.clear()
        password_field.send_keys(drbill_password)

        sign_in_btn = driver.find_element(
            By.CSS_SELECTOR, "input.button-primary[type='submit'][value='Sign in']"
        )
        pre_login_url = driver.current_url
        sign_in_btn.click()
        print("✅ Sign-in submitted")
    except Exception as e:
        print(f"❌ dr-bill.ca login failed: {e}")
        return False

    # Wait for the URL to actually change away from the login page
    try:
        WebDriverWait(driver, 15).until(lambda d: d.current_url != pre_login_url)
        post_login_url = driver.current_url
        print(f"🔀 Redirected to: {post_login_url}")
    except Exception:
        print("❌ URL never changed after sign-in — login may have failed silently")
        return False

    # Detect a redirect back to a login/auth page (failed login)
    login_indicators = ["login", "signin", "sign-in", "auth", "input27"]
    if any(ind in post_login_url.lower() for ind in login_indicators):
        print(f"❌ Login failed — redirected back to auth page: {post_login_url}")
        print("   Check DRBILL_USERNAME / DRBILL_PASSWORD in .env")
        return False

    # If redirected to the bare root, confirm we're actually logged in by
    # looking for a known post-login element before proceeding
    try:
        WebDriverWait(driver, 5).until(
            lambda d: d.find_elements(By.XPATH, "//*[contains(@href,'/files') or contains(@href,'/dashboard')]")
        )
        print("✅ Logged in to dr-bill.ca")
    except Exception:
        # Not a hard failure — just log and continue, the /files navigate will confirm
        print("⚠️ Could not verify post-login element, continuing anyway...")

    # Navigate to /files
    driver.get(f"{login_base}/files")
    print("📁 Navigated to /files")

    # Confirm we're not back on a login page
    current = driver.current_url
    if any(ind in current.lower() for ind in login_indicators):
        print(f"❌ Redirected to login when accessing /files — session not established: {current}")
        return False

    # Upload the file
    try:
        file_input = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.ID, "upload_file"))
        )
        file_input.send_keys(xlsx_path)
        print(f"✅ File selected: {os.path.basename(xlsx_path)}")

        upload_btn = WebDriverWait(driver, 5).until(
            EC.element_to_be_clickable((
                By.CSS_SELECTOR,
                "button.btn-primary[type='submit']"
            ))
        )
        upload_btn.click()
        print("✅ Upload button clicked")

        time.sleep(2)
        print(f"✅ Upload complete: {os.path.basename(xlsx_path)}")

        # Move the uploaded file to the archives folder
        archives_dir = os.path.join(os.path.dirname(xlsx_path), "archives")
        os.makedirs(archives_dir, exist_ok=True)
        dest = os.path.join(archives_dir, os.path.basename(xlsx_path))
        shutil.move(xlsx_path, dest)
        print(f"📦 Moved to archives: {dest}")

        return True

    except Exception as e:
        print(f"❌ File upload failed: {e}")
        return False


if __name__ == "__main__":
    xlsx_arg = sys.argv[1] if len(sys.argv) > 1 else None

    driver = setup_driver(headless=False)
    try:
        success = upload(driver, xlsx_path=xlsx_arg)
        sys.exit(0 if success else 1)
    finally:
        try:
            driver.quit()
        except Exception:
            pass
        cleanup_brave()
