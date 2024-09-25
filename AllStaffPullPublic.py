#Kyle Brogan

# Brought to you by MASC
# Maintain All Staff Coilition
#Sent to git

###Notes### 

###TO DO###
##EASY##
#   open .ics once complete
#   make year,month,and employee all inputs for distribution
#   make all file paths relative solved 9/25/24
 
##requires testing##
#   look into adding tkinter or basic ui for inputing name and date
#   create functions for portions of code

##  DeprecationWarning: datetime.datetime.utcnow() is deprecated and scheduled for removal in a future version. Use timezone-aware objects to represent datetimes in UTC: datetime.datetime.now(datetime.UTC).
##  DTSTAMP:{format_date(datetime.datetime.utcnow())}
#
#   make headless #solved 9/23 options variable

##Wishlist/ pipe dream##
#   Automate the rest of the days, maybe ask about flex time if possible. probably not fessable

#For distribution#
#   check .exe compatibility. 
#   package will need to include chrome of current version and ChromeDriver in relative path


from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
import time
from bs4 import BeautifulSoup
import os
import datetime
#import random
import pytz  # Import pytz for timezone handling
#import lxml # slenium HTML parser
import sys


# Manually set the month and year 
year = int(input('Enter a year (numerically ie 2024):\n'))
month = int(input('Enter a month (numerically):\n'))
employee = str(input('Enter your name in the following format ie "R. Ford"\n'))

""" # multi month setup
usermonths = str(input('Write out months seperated by comas ie 9,08,10\n'))
months = user.strip().split(',')
"""


# Timezone (America/New_York) to handle DST correctly
tz = pytz.timezone('America/New_York')


script_dir = str(os.path.dirname(os.path.abspath(__file__)))

driver_path = os.path.join(script_dir,'drivers','chromedriver-win64','chromedriver.exe')
# Initialize driver options
options = webdriver.ChromeOptions() #chat gpt shows options() with import options
# options.add_argument("--headless=new") removes chrome ui


#this is for chrome path in directory for final package it is too large as is
'''
#chrome_path = os.path.join(script_dir, 'chrome-win64', 'chrome.exe')
# sends chromedirectory to options
#options.binary_location = chrome_path 
'''
# Initialize the WebDriver
service = Service(driver_path)
driver = webdriver.Chrome(service=service, options=options) #changed

# Open the login page URL
login_url = 'http://allstaff.chess.cornell.edu/index.php'
driver.get(login_url)

# Wait for the page to load
print("Navigating to Login Page")
time.sleep(2)

# Enter login credentials
username = str(input('Enter CLASSE username:\n'))
password = str(input('Enter CLASSE password:\n'))

# Find the username and password input fields and submit the form
driver.find_element(By.NAME, 'uname').send_keys(username)
driver.find_element(By.NAME, 'pass').send_keys(password)
driver.find_element(By.XPATH, '//input[@value="User Login"]').click()
print("Successful Login")

time.sleep(5)

#for month in months..........

print(f"Navigating to {month} {year} Calander")
# Navigate to the calendar page after logging in
calendar_url = f'http://allstaff.chess.cornell.edu/modules/piCal2/index.php?cid=0&smode=Monthly&caldate={year}-{month}-1'
driver.get(calendar_url)

# Wait for the page to fully load
print(f"Scraping HTML of {month} {year} Calander")
time.sleep(5)

# Get the full page HTML
full_html = driver.page_source

# Save the HTML to a file
with open(f'{script_dir}/full_calendar_page.html', 'w', encoding='utf-8') as file:
    file.write(full_html,)
print("Full HTML content saved to 'full_calendar_page.html'")

# Close the browser
driver.quit()


with open(f'{script_dir}/full_calendar_page.html', 'r', encoding='ISO-8859-1') as f:
    soup = BeautifulSoup(f, 'lxml')

# Function to format the date for the .ics file (timezone aware)
def format_date(dt):
    return dt.strftime('%Y%m%dT%H%M%S')

# Dictionary to map shifts to times
shift_times = {
    'day': (8, 0, 16, 0),  # 8 AM to 4 PM
    'eve': (16, 0, 0, 0),  # 4 PM to 12 AM
    'owl': (0, 0, 8, 0)    # 12 AM to 8 AM
}

# Find shifts in HTML
shifts = []
for a_tag in soup.find_all('a'):
    if employee in a_tag.text:
        day_text = a_tag.text.lower()
        if any(shift in day_text for shift in ['day', 'eve', 'owl']):
            # Extract the day from the nearest calendar element
            day_element = a_tag.find_previous('span', class_='calbody')
            if day_element:
                shifts.append((day_element.text.strip(), day_text.strip()))

# Start building the .ics content with a time zone definition (Eastern Time)
ics_content = f"""BEGIN:VCALENDAR\nVERSION:2.0\nPRODID:-//Your Product//Your App//EN
CALSCALE:GREGORIAN
BEGIN:VTIMEZONE
TZID:America/New_York
X-LIC-LOCATION:America/New_York
BEGIN:DAYLIGHT
TZOFFSETFROM:-0500
TZOFFSETTO:-0400
TZNAME:EDT
DTSTART:19700308T020000
RRULE:FREQ=YEARLY;BYMONTH=3;BYDAY=2SU
END:DAYLIGHT
BEGIN:STANDARD
TZOFFSETFROM:-0400
TZOFFSETTO:-0500
TZNAME:EST
DTSTART:19701101T020000
RRULE:FREQ=YEARLY;BYMONTH=11;BYDAY=1SU
END:STANDARD
END:VTIMEZONE
"""

# Add events for each shift
for day, shift in shifts:
    day_num = int(day)  # Extract day number
    
    if 'day' in shift:
        start_hour, start_min, end_hour, end_min = shift_times['day']
    elif 'eve' in shift:
        start_hour, start_min, end_hour, end_min = shift_times['eve']
    elif 'owl' in shift:
        start_hour, start_min, end_hour, end_min = shift_times['owl']
    
    # Create event start and end times as timezone-aware
    event_start = tz.localize(datetime.datetime(year, month, day_num, start_hour, start_min))
    event_end = tz.localize(datetime.datetime(year, month, day_num, end_hour, end_min))
    
    if end_hour == 0:
        # Handle shifts that end past midnight
        event_end += datetime.timedelta(days=1)

    unique_uid = f"{datetime.datetime.now().strftime('%Y%m%dT%H%M%S')}-{day}@example.com"
    
    # Add the event to the .ics file
    ics_content += f"""BEGIN:VEVENT
UID:{unique_uid}
DTSTAMP:{format_date(datetime.datetime.now(tz))}
DTSTART:{format_date(event_start)}
DTEND:{format_date(event_end)}
SUMMARY:{shift.title()} Shift
DESCRIPTION:Shift for {shift} on {year}-{month:02d}-{day_num:02d}
END:VEVENT
"""

# Close the .ics file correctly
ics_content += "END:VCALENDAR\n"

# Save the .ics file
with open(f'{script_dir}/shifts_calendar.ics', 'w') as file:
    file.write(ics_content)
print("ICS file created with shifts and saved as 'shifts_calendar.ics'")

# Clean up
os.remove(f'{script_dir}/full_calendar_page.html')

# Open location with .ics file



'''
# Function to open created .ics
calimport = str(input('Would you like to import now? (y/n)\n'))

if calimport == 'y':
    os.system(f'open shifts_calendar.ics')
elif calimport == 'n':
    exit()
else:
    print('Invalid Response')
'''