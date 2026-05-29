import datetime
import os

# Set the billing date (env vars override hardcoded values)
billing_year  = os.environ.get("BILLING_DATE_YEAR",  str(datetime.date.today().year))
billing_month = os.environ.get("BILLING_DATE_MONTH", str(datetime.date.today().month))
billing_day   = os.environ.get("BILLING_DATE_DAY",   str(datetime.date.today().day))

# standard_appointment_length is 5 minutes
standard_appointment_length = 5

# counseling_appointment_length is 20 minutes
counseling_appointment_length = 20

# set delay times in seconds - optimized for speed
short_delay = 1
long_delay = 3

# Set the number of runs - default is None (process all appointments)
_runs_env = os.environ.get("BILLING_RUNS")
runs = int(_runs_env) if _runs_env and _runs_env.isdigit() else None

# set safe_mode — env var "1"/"true" overrides
safe_mode = os.environ.get("BILLING_SAFE_MODE", "0").lower() in ("1", "true")

# set headless mode: True = background (headless=new), False = visible window
headless_mode = os.environ.get("BILLING_HEADLESS", "1").lower() not in ("0", "false")

# EXPORT MODE FLAG
export_mode = os.environ.get("BILLING_EXPORT_MODE", "1").lower() not in ("0", "false")

# UPLOAD MODE FLAG
upload_mode = os.environ.get("BILLING_UPLOAD_MODE", "1").lower() not in ("0", "false")