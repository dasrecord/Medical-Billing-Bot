import datetime
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

# Set the number of runs - default is None (process all appointments)
runs = None

# set safe_mode (default = True)
safe_mode = False

# set headless mode: True = background (headless=new), False = visible window
headless_mode = True

# EXPORT MODE FLAG
export_mode = False

# UPLOAD MODE FLAG
upload_mode = False