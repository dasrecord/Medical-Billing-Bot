import datetime
import os


def _env_bool(name: str, default: bool) -> bool:
    val = os.environ.get(name)
    if val is None:
        return default
    return val.lower() in ("1", "true", "yes")


def _env_int(name: str, default=None):
    val = os.environ.get(name)
    return int(val) if val and val.isdigit() else default


# Billing date — defaults to today; override via env vars
_today = datetime.date.today()
billing_year  = os.environ.get("BILLING_DATE_YEAR",  str(_today.year))
billing_month = os.environ.get("BILLING_DATE_MONTH", str(_today.month))
billing_day   = os.environ.get("BILLING_DATE_DAY",   str(_today.day))

# Appointment durations (minutes)
standard_appointment_length  = 5
counseling_appointment_length = 20

# Browser delays (seconds)
short_delay = 1
long_delay  = 3

# Number of appointments to process (None = all)
runs = _env_int("BILLING_RUNS")

# Behaviour flags
safe_mode     = _env_bool("BILLING_SAFE_MODE",   False)
headless_mode = _env_bool("BILLING_HEADLESS",    True)
export_mode   = _env_bool("BILLING_EXPORT_MODE", False)
upload_mode   = _env_bool("BILLING_UPLOAD_MODE", False)