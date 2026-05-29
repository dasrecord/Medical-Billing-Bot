#!/usr/bin/env python3
"""
oscar_scraper.py — Lightweight Oscar EMR day-sheet scraper.

Fetches the provider appointment list for a given date and writes a JSON
cache file that the family-dashboard can consume.

Usage:
    python3 oscar_scraper.py [--date YYYY-MM-DD] [--output /path/to/file.json]
    python3 oscar_scraper.py --date 2026-05-28 --no-headless
"""

import argparse
import datetime
import json
import os
import shutil
import socket
import subprocess
import sys
import tempfile
import time

from dotenv import load_dotenv
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import NoSuchElementException, WebDriverException
from selenium.webdriver.chrome.service import Service as ChromeService
from webdriver_manager.chrome import ChromeDriverManager

load_dotenv()

OSCAR_BASE    = "https://well-kerrisdale.kai-oscar.com/oscar"
BRAVE_VERSION = os.environ.get("BRAVE_VERSION")   # override e.g. "148.0.7778.179"


def _detect_brave_version():
    """Return the major.minor version string for the installed Brave, or None."""
    brave_bin = None
    for p in ("/usr/bin/brave-browser", "/usr/bin/brave", "/snap/bin/brave"):
        if os.path.exists(p):
            brave_bin = p
            break
    if not brave_bin:
        return None
    try:
        out = subprocess.check_output([brave_bin, "--version"], stderr=subprocess.DEVNULL, text=True)
        # e.g. "Brave Browser 148.0.7778.179" → "148"
        parts = out.strip().split()
        return parts[-1].split(".")[0] if parts else None
    except Exception:
        return None

# Port 9224 is dedicated to the scraper so it does not clash with:
#   9222 — billing_bot.py
#   9223 — upload.py
SCRAPER_DEBUG_PORT = 9224
SCRAPER_USER_DATA  = os.path.join(tempfile.gettempdir(), "brave_scraper")

_brave_process = None


# ---------------------------------------------------------------------------
# Browser lifecycle
# ---------------------------------------------------------------------------

def cleanup_brave():
    """Terminate the scraper-owned Brave subprocess if it is still running."""
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


def _find_brave_path():
    """Return the path to the Brave binary or raise RuntimeError."""
    if sys.platform == "win32":
        candidates = [
            os.path.join(
                os.environ.get("LOCALAPPDATA", ""),
                "BraveSoftware", "Brave-Browser", "Application", "brave.exe",
            ),
            r"C:\Program Files\BraveSoftware\Brave-Browser\Application\brave.exe",
            r"C:\Program Files (x86)\BraveSoftware\Brave-Browser\Application\brave.exe",
        ]
    else:
        candidates = [
            "/Applications/Brave Browser.app/Contents/MacOS/Brave Browser",
            "/usr/bin/brave-browser",
            "/usr/bin/brave",
            "/snap/bin/brave",
        ]
    for path in candidates:
        if os.path.exists(path):
            return path
    raise RuntimeError(
        "Brave Browser not found. Install it or set the path manually in oscar_scraper.py."
    )


def setup_driver(headless=True):
    """Start a fresh Brave instance and connect via ChromeDriver."""
    global _brave_process

    brave_path = _find_brave_path()

    # Kill any leftover scraper instance scoped to the same user-data-dir
    if sys.platform != "win32":
        os.system(f"pkill -f 'user-data-dir={SCRAPER_USER_DATA}' 2>/dev/null")
    time.sleep(0.5)

    if os.path.exists(SCRAPER_USER_DATA):
        shutil.rmtree(SCRAPER_USER_DATA, ignore_errors=True)

    brave_args = [
        brave_path,
        f"--remote-debugging-port={SCRAPER_DEBUG_PORT}",
        f"--user-data-dir={SCRAPER_USER_DATA}",
        "--window-size=1920,1080",
        "--disable-extensions",
        "--disable-plugins",
    ]
    if headless:
        brave_args += ["--headless=new", "--disable-gpu"]

    print(f"🌐 Starting Brave ({'headless' if headless else 'visible'}) on port {SCRAPER_DEBUG_PORT}…")
    _brave_process = subprocess.Popen(
        brave_args, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL
    )

    # Wait for the remote debugging port to become available
    for attempt in range(15):
        try:
            sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
            sock.settimeout(1)
            if sock.connect_ex(("127.0.0.1", SCRAPER_DEBUG_PORT)) == 0:
                sock.close()
                break
        except Exception:
            pass
        time.sleep(1)
    else:
        raise RuntimeError(
            f"Brave failed to open remote debugging port {SCRAPER_DEBUG_PORT} within 15 s."
        )

    options = webdriver.ChromeOptions()
    options.add_experimental_option("debuggerAddress", f"127.0.0.1:{SCRAPER_DEBUG_PORT}")
    options.add_argument("--disable-blink-features=AutomationControlled")

    brave_ver   = BRAVE_VERSION or _detect_brave_version()
    mgr_kwargs  = {"driver_version": brave_ver} if brave_ver else {}
    driver_path = ChromeDriverManager(**mgr_kwargs).install()
    if "THIRD_PARTY_NOTICES" in driver_path:
        driver_dir  = os.path.dirname(driver_path)
        driver_name = "chromedriver.exe" if sys.platform == "win32" else "chromedriver"
        driver_path = os.path.join(driver_dir, driver_name)

    if sys.platform != "win32" and not os.access(driver_path, os.X_OK):
        os.chmod(driver_path, 0o755)

    service = ChromeService(driver_path)
    driver  = webdriver.Chrome(service=service, options=options)
    driver.set_page_load_timeout(30)
    driver.execute_script(
        "Object.defineProperty(navigator, 'webdriver', {get: () => undefined})"
    )
    print("✓ Connected to Brave via ChromeDriver")
    return driver


# ---------------------------------------------------------------------------
# Oscar login
# ---------------------------------------------------------------------------

def login_oscar(driver):
    """Log into Oscar EMR. Returns True on success."""
    username = os.getenv("OSCA_USERNAME") or os.getenv("OSCAR_USERNAME")
    password = os.getenv("OSCA_PASSWORD") or os.getenv("OSCAR_PASSWORD")
    if not username or not password:
        print("❌ OSCA_USERNAME / OSCA_PASSWORD not set in .env")
        return False

    try:
        print(f"🌐 Navigating to {OSCAR_BASE}…")
        driver.get(OSCAR_BASE)
        WebDriverWait(driver, 15).until(
            lambda d: d.execute_script("return document.readyState") == "complete"
        )
    except Exception as exc:
        print(f"❌ Could not reach Oscar EMR: {exc}")
        return False

    current_url = driver.current_url
    if "providercontrol" in current_url:
        print("✓ Already logged in")
        return True

    page_lower = driver.page_source.lower()
    if not any(k in page_lower for k in ["username", "password", "login", "sign in"]):
        print("⚠️  No login form detected — continuing anyway")
        return True

    print("🔐 Login form detected — filling credentials…")

    # Username field
    username_field = None
    for sel in [
        "input[name='username']", "input[name='user']",
        "input[id='username']",   "input[type='text']",
    ]:
        try:
            username_field = driver.find_element(By.CSS_SELECTOR, sel)
            break
        except NoSuchElementException:
            continue
    if not username_field:
        print("❌ Could not find username field")
        return False
    username_field.clear()
    username_field.send_keys(username)

    # Password field
    password_field = None
    for sel in [
        "input[name='password']", "input[id='password']", "input[type='password']",
    ]:
        try:
            password_field = driver.find_element(By.CSS_SELECTOR, sel)
            break
        except NoSuchElementException:
            continue
    if not password_field:
        print("❌ Could not find password field")
        return False
    password_field.clear()
    password_field.send_keys(password)

    # Submit
    submitted = False
    for sel in [
        "input[type='submit']", "button[type='submit']",
        "input[value*='Login']", "input[value*='Sign']",
    ]:
        try:
            driver.find_element(By.CSS_SELECTOR, sel).click()
            submitted = True
            break
        except NoSuchElementException:
            continue
    if not submitted:
        from selenium.webdriver.common.keys import Keys
        password_field.send_keys(Keys.RETURN)

    try:
        WebDriverWait(driver, 12).until(lambda d: d.current_url != current_url)
    except Exception:
        print("⚠️  URL did not change after login — may have timed out")
        return False

    if "login" in driver.current_url.lower():
        print("❌ Still on login page after submit — check credentials")
        return False

    print("✓ Login successful")
    return True


# ---------------------------------------------------------------------------
# Day-sheet scraping
# ---------------------------------------------------------------------------

def _oscar_day_url(date_obj):
    y, m, d = date_obj.year, date_obj.month, date_obj.day
    return (
        f"{OSCAR_BASE}/provider/providercontrol.jsp"
        f"?year={y}&month={m}&day={d}"
        f"&view=0&displaymode=day&dboperation=searchappointmentday&viewall=0"
    )

# Button labels that appear in appointment.text but are NOT patient names
_STRIP_LABELS = frozenset([
    "Encounter", "Billing", "Master Record", "Prescriptions", "Rx",
    "Preview", "Track", "Track,Fast", "Fast",
])


def scrape_day_sheet(driver, date_obj):
    """
    Navigate to the Oscar day sheet for *date_obj* and extract appointments.
    Returns a list of dicts — one per appointment element found on the page.
    """
    url = _oscar_day_url(date_obj)
    print(f"📅 Navigating to day sheet: {url}")
    driver.get(url)

    try:
        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.TAG_NAME, "body"))
        )
    except Exception:
        print("⚠️  Page-body wait timed out")

    time.sleep(1)  # allow dynamic content to settle

    appt_elements = []
    try:
        WebDriverWait(driver, 6).until(
            EC.presence_of_element_located((By.CLASS_NAME, "appt"))
        )
        appt_elements = driver.find_elements(By.CLASS_NAME, "appt")
        print(f"✓ Found {len(appt_elements)} appointment elements")
    except Exception:
        print("⚠️  No .appt elements found on page — schedule may be empty")
        return []

    results = []
    for idx, appt in enumerate(appt_elements):
        try:
            # --- Status (from the first <img> in the element) ---
            try:
                status = (
                    appt.find_element(By.XPATH, ".//img[1]")
                    .get_attribute("title")
                    or "Unknown"
                )
            except Exception:
                status = "Unknown"

            # --- Start time ---
            try:
                raw_time = (
                    appt.find_element(
                        By.XPATH, "./preceding-sibling::td[@align='RIGHT'][1]/a[1]"
                    )
                    .get_attribute("title")
                    .split(" - ")[0]
                )
                start_time = datetime.datetime.strptime(raw_time, "%I:%M%p").strftime("%H:%M")
            except Exception:
                start_time = None

            # --- Patient name + appointment URL ---
            appointment_url = None
            patient_name_raw = ""
            try:
                appt_link = appt.find_element(By.CSS_SELECTOR, "a.apptLink")
                patient_name_raw = appt_link.text.strip()
            except Exception:
                # Fallback: use full appt text, strip known button labels
                raw_text = appt.text or ""
                lines = [ln.strip() for ln in raw_text.splitlines() if ln.strip()]
                name_parts = [
                    ln for ln in lines
                    if ln not in _STRIP_LABELS and not ln.startswith("Track")
                ]
                patient_name_raw = " ".join(name_parts).strip()

            # Encounter link opens via onclick popup — extract URL from it
            try:
                enc_link = appt.find_element(By.XPATH, ".//a[contains(@title, 'Encounter')]")
                onclick = enc_link.get_attribute("onclick") or ""
                import re as _re
                from urllib.parse import urlparse, urljoin
                url_match = _re.search(
                    r"['\"]([^'\"]*(?:oscarEncounter|echart|encounter)[^'\"]*)['\"]",
                    onclick, _re.IGNORECASE
                )
                if url_match:
                    path = url_match.group(1)
                    # Resolve relative to the day-sheet page URL
                    day_url = _oscar_day_url(date_obj)
                    appointment_url = urljoin(day_url, path)
            except Exception:
                pass

            if "," in patient_name_raw:
                last_raw, first_raw = patient_name_raw.split(",", 1)
                last_name  = last_raw.strip().title()
                first_name = first_raw.strip().title()
            else:
                last_name  = patient_name_raw.title()
                first_name = ""

            results.append(
                {
                    "index":          idx,
                    "time":           start_time,
                    "lastName":       last_name,
                    "firstName":      first_name,
                    "status":         status,
                    "appointmentUrl": appointment_url,
                }
            )

        except Exception as exc:
            print(f"⚠️  Could not parse appointment {idx}: {exc}")
            results.append(
                {
                    "index":          idx,
                    "time":           None,
                    "lastName":       "",
                    "firstName":      "",
                    "status":         "Error",
                    "appointmentUrl": None,
                }
            )

    return results


def _build_summary(appointments):
    billed_statuses    = {"Billed", "Billed/Verified", "Billed/Signed"}
    no_show_statuses   = {"No Show"}
    cancelled_statuses = {"Cancelled"}

    summary = {
        "total":     len(appointments),
        "billed":    0,
        "noShow":    0,
        "cancelled": 0,
        "pending":   0,
    }
    for a in appointments:
        s = a.get("status", "")
        if s in billed_statuses:
            summary["billed"] += 1
        elif s in no_show_statuses:
            summary["noShow"] += 1
        elif s in cancelled_statuses:
            summary["cancelled"] += 1
        else:
            summary["pending"] += 1
    return summary


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------

def main():
    parser = argparse.ArgumentParser(description="Oscar EMR day-sheet scraper")
    parser.add_argument(
        "--date",
        default=datetime.date.today().isoformat(),
        help="Date to scrape (YYYY-MM-DD). Defaults to today.",
    )
    parser.add_argument(
        "--output",
        default=None,
        help=(
            "Output JSON file path. Defaults to "
            "../family-dashboard-app/data/household/oscar-YYYYMMDD.json "
            "relative to this script."
        ),
    )
    headless_group = parser.add_mutually_exclusive_group()
    headless_group.add_argument(
        "--headless",   dest="headless", action="store_true",  default=None,
        help="Run browser in headless mode.",
    )
    headless_group.add_argument(
        "--no-headless", dest="headless", action="store_false",
        help="Run browser in a visible window.",
    )
    args = parser.parse_args()

    # Fall back to config.py headless_mode if not specified on CLI
    if args.headless is None:
        try:
            from config import headless_mode as _cfg_headless
            args.headless = bool(_cfg_headless)
        except Exception:
            args.headless = True

    # Validate date
    try:
        date_obj = datetime.date.fromisoformat(args.date)
    except ValueError:
        print(f"❌ Invalid date '{args.date}'. Use YYYY-MM-DD.")
        sys.exit(1)

    date_str = date_obj.strftime("%Y%m%d")

    # Resolve output path
    if args.output:
        output_path = os.path.abspath(args.output)
    else:
        script_dir  = os.path.dirname(os.path.abspath(__file__))
        output_path = os.path.normpath(
            os.path.join(
                script_dir, "..", "family-dashboard-app",
                "data", "household", f"oscar-{date_str}.json",
            )
        )
    os.makedirs(os.path.dirname(output_path), exist_ok=True)

    now_iso = datetime.datetime.now(datetime.timezone.utc).isoformat()
    payload = {
        "date":        date_obj.isoformat(),
        "lastUpdated": now_iso,
        "status":      "error",
        "appointments": [],
        "summary": {
            "total": 0, "billed": 0, "noShow": 0, "cancelled": 0, "pending": 0,
        },
    }

    driver = None
    try:
        print(f"\n{'='*52}")
        print(f"  OSCAR DAY-SHEET SCRAPER  —  {date_obj.isoformat()}")
        print(f"{'='*52}\n")

        driver = setup_driver(headless=args.headless)

        if not login_oscar(driver):
            payload["error"] = "Login failed"
            with open(output_path, "w", encoding="utf-8") as fh:
                json.dump(payload, fh, indent=2)
            sys.exit(1)

        appointments = scrape_day_sheet(driver, date_obj)
        summary      = _build_summary(appointments)

        payload.update(
            {
                "status":       "ok",
                "appointments": appointments,
                "summary":      summary,
            }
        )
        payload.pop("error", None)

        with open(output_path, "w", encoding="utf-8") as fh:
            json.dump(payload, fh, indent=2)

        print(f"\n✅  Wrote {len(appointments)} appointment(s)  →  {output_path}")
        print(f"    Summary: {summary}")

    except Exception as exc:
        payload["error"] = str(exc)
        try:
            with open(output_path, "w", encoding="utf-8") as fh:
                json.dump(payload, fh, indent=2)
        except Exception:
            pass
        print(f"\n❌  Scraper failed: {exc}")
        sys.exit(1)

    finally:
        if driver:
            try:
                driver.quit()
            except Exception:
                pass
        cleanup_brave()


if __name__ == "__main__":
    main()
