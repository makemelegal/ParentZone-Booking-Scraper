"""
ParentZone Calendar Scraper for Windows
Generates separate iCal (.ics) files per month for import into Google Calendar

Requirements:
- Python 3.7+
- Chrome browser installed
- selenium package

Install dependencies:
    pip install selenium
"""

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from datetime import datetime, timedelta
import time
import re
import os

# ============= CONFIGURATION =============
PARENTZONE_USERNAME = "your_email@example.com"  # CHANGE THIS
PARENTZONE_PASSWORD = "your_password"           # CHANGE THIS
MONTHS_TO_SCRAPE = 3                            # How many months ahead to scrape

# ============= APPOINTMENT CUSTOMIZATION =============
# Choose how appointments appear in your calendar:

# Option 1: Use actual booking times (DEFAULT)
# Events will show the full duration (e.g., 09:00-17:00)
USE_ACTUAL_TIMES = True

# Option 2: Create reminder appointments at a specific time
# Set USE_ACTUAL_TIMES = False and configure below:
REMINDER_TIME = "07:00"           # Time for reminder (24-hour format, e.g., "07:00", "18:30")
REMINDER_DURATION_MINUTES = 30    # How long the reminder lasts (e.g., 30 = 30 minute appointment)

# Summary format options:
# True = Include time in summary (e.g., "Felicity 9am-5pm")
# False = Name only (e.g., "Felicity Nursery")
INCLUDE_TIME_IN_SUMMARY = True
# ====================================================

# URLs
PARENTZONE_LOGIN_URL = "https://www.parentzone.me/login"
PARENTZONE_BOOKINGS_URL = "https://www.parentzone.me/bookings"


def setup_driver():
    """Configure Chrome driver with options for visibility"""
    chrome_options = Options()
    
    # Comment out the headless line if you want to SEE the browser working
    # chrome_options.add_argument("--headless")
    
    chrome_options.add_argument("--window-size=1920,1080")
    chrome_options.add_argument("--disable-blink-features=AutomationControlled")
    chrome_options.add_experimental_option("excludeSwitches", ["enable-automation"])
    chrome_options.add_experimental_option('useAutomationExtension', False)
    
    driver = webdriver.Chrome(options=chrome_options)
    return driver


def login_to_parentzone(driver, username, password):
    """Log into ParentZone"""
    print("🔐 Logging into ParentZone...")
    driver.get(PARENTZONE_LOGIN_URL)
    
    wait = WebDriverWait(driver, 15)
    
    try:
        email_field = wait.until(EC.presence_of_element_located((By.NAME, "email")))
        email_field.clear()
        email_field.send_keys(username)
        
        password_field = driver.find_element(By.NAME, "password")
        password_field.clear()
        password_field.send_keys(password)
        
        login_button = driver.find_element(By.XPATH, "//button[@type='submit']")
        login_button.click()
        
        time.sleep(5)
        print("✅ Login successful!")
        return True
        
    except Exception as e:
        print(f"❌ Login failed: {e}")
        return False


def extract_bookings_from_page(driver):
    """Extract bookings from the current calendar view"""
    
    wait = WebDriverWait(driver, 15)
    
    try:
        wait.until(EC.presence_of_element_located(
            (By.CSS_SELECTOR, "h6.MuiTypography-root.MuiTypography-h6.MuiTypography-noWrap.css-1dw86cl-titleWithButtons")
        ))
    except:
        print("⚠️ Calendar not loaded properly, waiting longer...")
        time.sleep(5)
    
    bookings = []
    
    try:
        month_year_element = driver.find_element(
            By.CSS_SELECTOR, 
            "h6.MuiTypography-root.MuiTypography-h6.MuiTypography-noWrap.css-1dw86cl-titleWithButtons"
        )
        full_text = month_year_element.text.strip()
        month_year_str = full_text.replace('Bookings - ', '').strip()
        
        parts = month_year_str.split(' ')
        if len(parts) != 2:
            print(f"⚠️ Could not parse month/year: {month_year_str}")
            return bookings
        
        month_abbrev = parts[0]
        year_str = parts[1]
        
        month_map = {
            'Jan': 1, 'Feb': 2, 'Mar': 3, 'Apr': 4, 'May': 5, 'Jun': 6,
            'Jul': 7, 'Aug': 8, 'Sep': 9, 'Oct': 10, 'Nov': 11, 'Dec': 12
        }
        
        current_month = month_map.get(month_abbrev)
        current_year = int(year_str)
        
        if not current_month:
            print(f"⚠️ Invalid month abbreviation: {month_abbrev}")
            return bookings
        
        print(f"📅 Processing {month_abbrev} {current_year}...")
        
        day_elements = driver.find_elements(
            By.CSS_SELECTOR, 
            "div.css-1btmizi-day.css-1eqmmqv-dayDesktop.css-1ke78x2-dayBorder"
        )
        
        print(f"  Found {len(day_elements)} day elements")
        
        for day_el in day_elements:
            try:
                date_element = day_el.find_element(
                    By.CSS_SELECTOR,
                    "p.MuiTypography-root.MuiTypography-body2.css-68o8xu"
                )
                date_text = date_element.text.strip()
                
                if not date_text:
                    continue
                
                date_parts = date_text.split(' ')
                day_number = int(date_parts[0])
                
                day_month = current_month
                day_year = current_year
                
                if len(date_parts) > 1:
                    day_month_abbrev = date_parts[1]
                    day_month = month_map.get(day_month_abbrev, current_month)
                    
                    if day_month_abbrev == 'Dec' and month_abbrev == 'Jan':
                        day_year = current_year - 1
                    elif day_month_abbrev == 'Jan' and month_abbrev == 'Dec':
                        day_year = current_year + 1
                
                booking_containers = day_el.find_elements(
                    By.CSS_SELECTOR,
                    "div.css-jvibwz-buttonContainer"
                )
                
                if booking_containers:
                    print(f"  Day {date_text}: Found {len(booking_containers)} booking(s)")
                
                day_bookings = []
                
                for container in booking_containers:
                    try:
                        child_name_el = container.find_element(
                            By.CSS_SELECTOR,
                            "span.css-cypr81-childName"
                        )
                        child_name = child_name_el.text.strip() if child_name_el else "Felicity Nursery"
                        
                        session_time_el = container.find_element(
                            By.CSS_SELECTOR,
                            "span.css-11fzqss-sessionTime"
                        )
                        session_time = session_time_el.text.strip()
                        
                        if not session_time:
                            continue
                        
                        time_match = re.search(r'(\d{1,2}:\d{2})\s*-\s*(\d{1,2}:\d{2})', session_time)
                        if not time_match:
                            print(f"    ⚠️ Could not parse time: {session_time}")
                            continue
                        
                        start_time = time_match.group(1)
                        end_time = time_match.group(2)
                        
                        date_str = f"{day_year}-{day_month:02d}-{day_number:02d}"
                        start_datetime = datetime.strptime(f"{date_str} {start_time}", "%Y-%m-%d %H:%M")
                        end_datetime = datetime.strptime(f"{date_str} {end_time}", "%Y-%m-%d %H:%M")
                        
                        day_bookings.append({
                            'child_name': child_name,
                            'start': start_datetime,
                            'end': end_datetime
                        })
                        
                    except Exception as e:
                        print(f"    ⚠️ Failed to parse booking: {e}")
                        continue
                
                if day_bookings:
                    day_bookings.sort(key=lambda x: x['start'])
                    
                    i = 0
                    while i < len(day_bookings):
                        current = day_bookings[i]
                        combined_start = current['start']
                        combined_end = current['end']
                        child_name = current['child_name']
                        
                        j = i + 1
                        while j < len(day_bookings):
                            next_booking = day_bookings[j]
                            if (next_booking['child_name'] == child_name and 
                                combined_end == next_booking['start']):
                                combined_end = next_booking['end']
                                j += 1
                            else:
                                break
                        
                        if USE_ACTUAL_TIMES:
                            event_start = combined_start
                            event_end = combined_end
                        else:
                            reminder_hour, reminder_minute = map(int, REMINDER_TIME.split(':'))
                            event_start = combined_start.replace(hour=reminder_hour, minute=reminder_minute)
                            event_end = event_start + timedelta(minutes=REMINDER_DURATION_MINUTES)
                        
                        if INCLUDE_TIME_IN_SUMMARY:
                            start_hour = combined_start.hour
                            end_hour = combined_end.hour
                            
                            if start_hour == 0:
                                start_12h = "12am"
                            elif start_hour < 12:
                                start_12h = f"{start_hour}am"
                            elif start_hour == 12:
                                start_12h = "12pm"
                            else:
                                start_12h = f"{start_hour - 12}pm"
                            
                            if end_hour == 0:
                                end_12h = "12am"
                            elif end_hour < 12:
                                end_12h = f"{end_hour}am"
                            elif end_hour == 12:
                                end_12h = "12pm"
                            else:
                                end_12h = f"{end_hour - 12}pm"
                            
                            time_str = f"{start_12h}-{end_12h}"
                            summary = f"{child_name} {time_str}"
                        else:
                            summary = child_name
                        
                        bookings.append({
                            'summary': summary,
                            'start': event_start,
                            'end': event_end,
                            'description': f'ParentZone booking: {child_name} ({combined_start.strftime("%H:%M")}-{combined_end.strftime("%H:%M")})',
                            'original_start': combined_start.strftime("%H:%M"),
                            'original_end': combined_end.strftime("%H:%M")
                        })
                        
                        print(f"    ✓ {summary}")
                        i = j
                    
            except Exception as e:
                continue
        
        print(f"  Total: {len(bookings)} booking(s) this month\n")
        
    except Exception as e:
        print(f"❌ Error extracting bookings: {e}")
    
    return bookings


def click_next_month(driver):
    """Click the next month button and wait for page to update"""
    try:
        month_year_element = driver.find_element(
            By.CSS_SELECTOR, 
            "h6.MuiTypography-root.MuiTypography-h6.MuiTypography-noWrap.css-1dw86cl-titleWithButtons"
        )
        old_month_year = month_year_element.text.replace('Bookings - ', '').strip()
        
        next_button = None
        
        try:
            next_button = driver.find_element(By.CSS_SELECTOR, 'button[data-test-id="next_btn"]')
            print(f"  Found next month button via data-test-id")
        except:
            try:
                next_button = driver.find_element(
                    By.CSS_SELECTOR,
                    "button.MuiButtonBase-root.MuiIconButton-root.MuiIconButton-sizeSmall.css-1j7qk7u"
                )
                print(f"  Found next month button via CSS classes")
            except:
                try:
                    chevron_icon = driver.find_element(By.CSS_SELECTOR, 'svg[data-testid="ChevronRightIcon"]')
                    next_button = chevron_icon.find_element(By.XPATH, "..")
                    print(f"  Found next month button via ChevronRightIcon")
                except:
                    pass
        
        if not next_button:
            print("⚠️ Could not find next month button")
            return False
        
        next_button.click()
        
        wait = WebDriverWait(driver, 5)
        
        def month_changed(driver):
            try:
                element = driver.find_element(
                    By.CSS_SELECTOR,
                    "h6.MuiTypography-root.MuiTypography-h6.MuiTypography-noWrap.css-1dw86cl-titleWithButtons"
                )
                new_month_year = element.text.replace('Bookings - ', '').strip()
                return new_month_year != old_month_year
            except:
                return False
        
        wait.until(month_changed)
        time.sleep(1)
        
        print(f"  ➡️ Advanced to next month")
        return True
        
    except Exception as e:
        print(f"⚠️ Could not advance to next month: {e}")
        return False


def generate_ical_per_month(all_bookings, output_folder="."):
    """Generate separate iCal files for each month"""
    
    bookings_by_month = {}
    
    for booking in all_bookings:
        month_key = booking['start'].strftime('%Y-%m')
        month_name = booking['start'].strftime('%Y_%B')
        
        if month_key not in bookings_by_month:
            bookings_by_month[month_key] = {
                'name': month_name,
                'bookings': []
            }
        
        bookings_by_month[month_key]['bookings'].append(booking)
    
    created_files = []
    
    for month_key in sorted(bookings_by_month.keys()):
        month_data = bookings_by_month[month_key]
        month_bookings = month_data['bookings']
        month_name = month_data['name']
        
        seen = set()
        unique_bookings = []
        
        for booking in month_bookings:
            original_start = booking.get('original_start', booking['start'].strftime('%H:%M'))
            original_end = booking.get('original_end', booking['end'].strftime('%H:%M'))
            date_str = booking['start'].strftime('%Y-%m-%d')
            
            booking_key = (
                booking['summary'],
                date_str,
                original_start,
                original_end
            )
            
            if booking_key not in seen:
                seen.add(booking_key)
                unique_bookings.append(booking)
        
        duplicates_removed = len(month_bookings) - len(unique_bookings)
        
        filename = f"parentzone_bookings_{month_name}.ics"
        filepath = os.path.join(output_folder, filename)
        
        ical_content = [
            "BEGIN:VCALENDAR",
            "VERSION:2.0",
            "PRODID:-//ParentZone Scraper//EN",
            "CALSCALE:GREGORIAN",
            "METHOD:PUBLISH",
            f"X-WR-CALNAME:ParentZone {month_name.replace('_', ' ')}",
            "X-WR-TIMEZONE:Europe/London",
        ]
        
        for idx, booking in enumerate(unique_bookings):
            original_start = booking.get('original_start', booking['start'].strftime('%H:%M'))
            original_end = booking.get('original_end', booking['end'].strftime('%H:%M'))
            date_str = booking['start'].strftime('%Y%m%d')
            uid = f"parentzone-{date_str}-{original_start.replace(':', '')}-{idx}@parentzone.me"
            
            dtstart = booking['start'].strftime('%Y%m%dT%H%M%S')
            dtend = booking['end'].strftime('%Y%m%dT%H%M%S')
            dtstamp = datetime.now().strftime('%Y%m%dT%H%M%SZ')
            
            ical_content.extend([
                "BEGIN:VEVENT",
                f"UID:{uid}",
                f"DTSTAMP:{dtstamp}",
                f"DTSTART:{dtstart}",
                f"DTEND:{dtend}",
                f"SUMMARY:{booking['summary']}",
                f"DESCRIPTION:{booking['description']}",
                "STATUS:CONFIRMED",
                "END:VEVENT"
            ])
        
        ical_content.append("END:VCALENDAR")
        
        with open(filepath, 'w', encoding='utf-8') as f:
            f.write('\n'.join(ical_content))
        
        created_files.append(filename)
        
        print(f"✅ Created: {filename}")
        print(f"   📊 {len(unique_bookings)} unique event(s)", end="")
        if duplicates_removed > 0:
            print(f" ({duplicates_removed} duplicate(s) removed)")
        else:
            print()
    
    return created_files


def main():
    """Main execution"""
    print("=" * 60)
    print("ParentZone Calendar Scraper")
    print("=" * 60)
    print()
    
    if PARENTZONE_USERNAME == "your_email@example.com":
        print("❌ ERROR: Please edit the script and add your credentials!")
        print("   Update PARENTZONE_USERNAME and PARENTZONE_PASSWORD")
        input("\nPress Enter to exit...")
        return
    
    driver = None
    all_bookings = []
    scraped_months = []
    
    try:
        print("🌐 Starting Chrome browser...")
        driver = setup_driver()
        
        if not login_to_parentzone(driver, PARENTZONE_USERNAME, PARENTZONE_PASSWORD):
            print("❌ Login failed. Please check your credentials.")
            input("\nPress Enter to exit...")
            return
        
        print("📆 Navigating to bookings page...")
        driver.get(PARENTZONE_BOOKINGS_URL)
        time.sleep(5)
        
        print(f"\n🔍 Scraping up to {MONTHS_TO_SCRAPE} month(s)...\n")
        
        for month_idx in range(MONTHS_TO_SCRAPE):
            try:
                month_year_element = driver.find_element(
                    By.CSS_SELECTOR,
                    "h6.MuiTypography-root.MuiTypography-h6.MuiTypography-noWrap.css-1dw86cl-titleWithButtons"
                )
                current_month_year = month_year_element.text.replace('Bookings - ', '').strip()
                
                if current_month_year in scraped_months:
                    print(f"⚠️ Already scraped {current_month_year}, stopping to avoid duplicates.")
                    break
                
                scraped_months.append(current_month_year)
                
            except:
                print("⚠️ Could not read current month, stopping.")
                break
            
            bookings = extract_bookings_from_page(driver)
            all_bookings.extend(bookings)
            
            if month_idx < MONTHS_TO_SCRAPE - 1:
                if not click_next_month(driver):
                    print("⚠️ Could not navigate further, stopping here.")
                    break
        
        print("\n" + "=" * 60)
        if all_bookings:
            print("📝 Generating iCal files...\n")
            created_files = generate_ical_per_month(all_bookings)
            
            print("\n" + "=" * 60)
            print("✨ Success! Next steps:")
            print(f"\n📁 Created {len(created_files)} file(s):")
            for filename in created_files:
                print(f"   - {filename}")
            print(f"\n📂 Location: {os.path.abspath('.')}")
            print("\n🔄 To import to Google Calendar:")
            print("1. Go to https://calendar.google.com")
            print("2. Click ⚙️ Settings → Settings")
            print("3. Click 'Import & export' in left sidebar")
            print("4. Click 'Select file from your computer'")
            print("5. Import each .ics file (or select all at once)")
            print("6. Choose which calendar to import into")
            print("7. Click 'Import'")
            print("\n💡 Tip: You can import all files at once by selecting them together!")
        else:
            print("⚠️ No bookings found. This could mean:")
            print("   - No bookings exist for these months")
            print("   - CSS selectors need updating (ParentZone changed their UI)")
            print("   - Page didn't load properly")
            print("\n💡 Try running again with the browser visible:")
            print("   Comment out the --headless line in setup_driver()")
        
    except Exception as e:
        print(f"\n❌ Unexpected error: {e}")
        print("\nTroubleshooting tips:")
        print("1. Make sure Chrome browser is installed")
        print("2. Install chromedriver matching your Chrome version")
        print("3. Check if ParentZone website structure changed")
        print("4. Try running with browser visible (comment out --headless)")
        
    finally:
        if driver:
            print("\n🔒 Closing browser...")
            driver.quit()
        
        print("\n" + "=" * 60)
        input("Press Enter to exit...")


if __name__ == "__main__":
    main()
