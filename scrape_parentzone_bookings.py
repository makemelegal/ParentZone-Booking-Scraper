"""
ParentZone Calendar Scraper for Windows
Generates an iCal (.ics) file that can be imported into Google Calendar

Requirements:
- Python 3.7+
- Chrome browser installed
- selenium package
- chromedriver matching your Chrome version

Install dependencies:
    pip install selenium
"""

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from datetime import datetime
import time
import re
import os

# ============= CONFIGURATION =============
PARENTZONE_USERNAME = "INPUT EMAIL HERE"  # CHANGE THIS
PARENTZONE_PASSWORD = "INPUT PASSWORD HERE"           # CHANGE THIS
MONTHS_TO_SCRAPE = 3                            # How many months ahead to scrape
OUTPUT_FILE = "parentzone_bookings.ics"         # Output filename

# URLs
PARENTZONE_LOGIN_URL = "https://www.parentzone.me/login"
PARENTZONE_BOOKINGS_URL = "https://www.parentzone.me/bookings"

# Month mapping
MONTH_MAP = {
    'January': '01', 'February': '02', 'March': '03', 'April': '04',
    'May': '05', 'June': '06', 'July': '07', 'August': '08',
    'September': '09', 'October': '10', 'November': '11', 'December': '12'
}
# =========================================


def setup_driver():
    """Configure Chrome driver with options for visibility (so you can see what's happening)"""
    chrome_options = Options()
    
    # Comment out the headless line if you want to SEE the browser working
    # chrome_options.add_argument("--headless")
    
    chrome_options.add_argument("--window-size=1920,1080")
    chrome_options.add_argument("--disable-blink-features=AutomationControlled")
    chrome_options.add_experimental_option("excludeSwitches", ["enable-automation"])
    chrome_options.add_experimental_option('useAutomationExtension', False)
    
    # Chrome will automatically find chromedriver if it's in PATH
    # If not, you'll need to specify the path to chromedriver.exe
    driver = webdriver.Chrome(options=chrome_options)
    return driver


def login_to_parentzone(driver, username, password):
    """Log into ParentZone"""
    print("🔐 Logging into ParentZone...")
    driver.get(PARENTZONE_LOGIN_URL)
    
    # Wait for login form to load
    wait = WebDriverWait(driver, 15)
    
    try:
        # Find email field (may need adjustment based on actual HTML)
        email_field = wait.until(EC.presence_of_element_located((By.NAME, "email")))
        email_field.clear()
        email_field.send_keys(username)
        
        # Find password field
        password_field = driver.find_element(By.NAME, "password")
        password_field.clear()
        password_field.send_keys(password)
        
        # Find and click login button
        login_button = driver.find_element(By.XPATH, "//button[@type='submit']")
        login_button.click()
        
        # Wait for navigation (adjust time if needed)
        time.sleep(5)
        print("✅ Login successful!")
        return True
        
    except Exception as e:
        print(f"❌ Login failed: {e}")
        print("💡 Tip: Check if email/password field names have changed")
        return False


def extract_bookings_from_page(driver):
    """Extract bookings from the current calendar view using the actual ParentZone structure"""
    
    # Wait for calendar to fully load
    wait = WebDriverWait(driver, 15)
    
    try:
        # Wait for the title with month/year
        wait.until(EC.presence_of_element_located(
            (By.CSS_SELECTOR, "h6.MuiTypography-root.MuiTypography-h6.MuiTypography-noWrap.css-1dw86cl-titleWithButtons")
        ))
    except:
        print("⚠️ Calendar not loaded properly, waiting longer...")
        time.sleep(5)
    
    bookings = []
    
    try:
        # Get current month/year from page header (e.g., "Bookings - Jan 2025")
        month_year_element = driver.find_element(
            By.CSS_SELECTOR, 
            "h6.MuiTypography-root.MuiTypography-h6.MuiTypography-noWrap.css-1dw86cl-titleWithButtons"
        )
        full_text = month_year_element.text.strip()
        # Remove "Bookings - " prefix
        month_year_str = full_text.replace('Bookings - ', '').strip()
        
        # Parse "Jan 2025" format
        parts = month_year_str.split(' ')
        if len(parts) != 2:
            print(f"⚠️ Could not parse month/year: {month_year_str}")
            return bookings
        
        month_abbrev = parts[0]  # e.g., "Jan"
        year_str = parts[1]       # e.g., "2025"
        
        # Map month abbreviations to numbers
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
        
        # Get all day containers using the actual ParentZone CSS classes
        day_elements = driver.find_elements(
            By.CSS_SELECTOR, 
            "div.css-1btmizi-day.css-1eqmmqv-dayDesktop.css-1ke78x2-dayBorder"
        )
        
        print(f"  Found {len(day_elements)} day elements")
        
        for day_el in day_elements:
            try:
                # Get the date from the day element (e.g., "1 Jan" or just "1")
                date_element = day_el.find_element(
                    By.CSS_SELECTOR,
                    "p.MuiTypography-root.MuiTypography-body2.css-68o8xu"
                )
                date_text = date_element.text.strip()
                
                if not date_text:
                    continue
                
                # Parse the date - could be "1 Jan" or just "1"
                date_parts = date_text.split(' ')
                day_number = int(date_parts[0])
                
                # Handle month overlap (e.g., "31 Dec" showing in Jan view)
                day_month = current_month
                day_year = current_year
                
                if len(date_parts) > 1:
                    day_month_abbrev = date_parts[1]
                    day_month = month_map.get(day_month_abbrev, current_month)
                    
                    # Handle year boundaries
                    if day_month_abbrev == 'Dec' and month_abbrev == 'Jan':
                        day_year = current_year - 1
                    elif day_month_abbrev == 'Jan' and month_abbrev == 'Dec':
                        day_year = current_year + 1
                
                # Find all booking containers for this day
                booking_containers = day_el.find_elements(
                    By.CSS_SELECTOR,
                    "div.css-jvibwz-buttonContainer"
                )
                
                if booking_containers:
                    print(f"  Day {date_text}: Found {len(booking_containers)} booking(s)")
                
                for container in booking_containers:
                    try:
                        # Extract child name
                        child_name_el = container.find_element(
                            By.CSS_SELECTOR,
                            "span.css-cypr81-childName"
                        )
                        child_name = child_name_el.text.strip() if child_name_el else "Felicity Nursery"
                        
                        # Extract session time (e.g., "09:00 - 17:00")
                        session_time_el = container.find_element(
                            By.CSS_SELECTOR,
                            "span.css-11fzqss-sessionTime"
                        )
                        session_time = session_time_el.text.strip()
                        
                        if not session_time:
                            continue
                        
                        # Parse time range
                        time_match = re.search(r'(\d{1,2}:\d{2})\s*-\s*(\d{1,2}:\d{2})', session_time)
                        if not time_match:
                            print(f"    ⚠️ Could not parse time: {session_time}")
                            continue
                        
                        start_time = time_match.group(1)
                        end_time = time_match.group(2)
                        
                        # Build datetime objects
                        date_str = f"{day_year}-{day_month:02d}-{day_number:02d}"
                        start_datetime = datetime.strptime(f"{date_str} {start_time}", "%Y-%m-%d %H:%M")
                        end_datetime = datetime.strptime(f"{date_str} {end_time}", "%Y-%m-%d %H:%M")
                        
                        bookings.append({
                            'summary': child_name,
                            'start': start_datetime,
                            'end': end_datetime,
                            'description': f'ParentZone booking: {child_name}'
                        })
                        
                        print(f"    ✓ {child_name}: {start_time}-{end_time}")
                        
                    except Exception as e:
                        print(f"    ⚠️ Failed to parse booking: {e}")
                        continue
                    
            except Exception as e:
                # Day with no bookings - skip silently
                continue
        
        print(f"  Total: {len(bookings)} booking(s) this month\n")
        
    except Exception as e:
        print(f"❌ Error extracting bookings: {e}")
    
    return bookings


def click_next_month(driver):
    """Click the next month button and wait for page to update"""
    try:
        # Get current month/year BEFORE clicking
        month_year_element = driver.find_element(
            By.CSS_SELECTOR, 
            "h6.MuiTypography-root.MuiTypography-h6.MuiTypography-noWrap.css-1dw86cl-titleWithButtons"
        )
        old_month_year = month_year_element.text.replace('Bookings - ', '').strip()
        
        # Find the next month button using data-test-id (most reliable)
        next_button = None
        
        try:
            # Primary method: Use the data-test-id attribute
            next_button = driver.find_element(By.CSS_SELECTOR, 'button[data-test-id="next_btn"]')
            print(f"  Found next month button via data-test-id")
        except:
            # Fallback: Use the full class combination
            try:
                next_button = driver.find_element(
                    By.CSS_SELECTOR,
                    "button.MuiButtonBase-root.MuiIconButton-root.MuiIconButton-sizeSmall.css-1j7qk7u"
                )
                print(f"  Found next month button via CSS classes")
            except:
                # Last resort: Find button containing ChevronRightIcon
                try:
                    chevron_icon = driver.find_element(By.CSS_SELECTOR, 'svg[data-testid="ChevronRightIcon"]')
                    next_button = chevron_icon.find_element(By.XPATH, "..")
                    print(f"  Found next month button via ChevronRightIcon")
                except:
                    pass
        
        if not next_button:
            print("⚠️ Could not find next month button")
            return False
        
        # Click the button
        next_button.click()
        
        # Wait for the month/year to actually change (max 5 seconds)
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
        
        # Small additional pause to let the day elements render
        time.sleep(1)
        
        print(f"  ➡️ Advanced to next month")
        return True
        
    except Exception as e:
        print(f"⚠️ Could not advance to next month: {e}")
        return False


def generate_ical_per_month(all_bookings, output_folder="."):
    """Generate separate iCal files for each month"""
    
    # Group bookings by month
    bookings_by_month = {}
    
    for booking in all_bookings:
        # Get year and month from start datetime
        month_key = booking['start'].strftime('%Y-%m')  # e.g., "2025-01"
        month_name = booking['start'].strftime('%Y_%B')  # e.g., "2025_January"
        
        if month_key not in bookings_by_month:
            bookings_by_month[month_key] = {
                'name': month_name,
                'bookings': []
            }
        
        bookings_by_month[month_key]['bookings'].append(booking)
    
    # Generate a separate iCal file for each month
    created_files = []
    
    for month_key in sorted(bookings_by_month.keys()):
        month_data = bookings_by_month[month_key]
        month_bookings = month_data['bookings']
        month_name = month_data['name']
        
        # Remove duplicates within this month
        seen = set()
        unique_bookings = []
        
        for booking in month_bookings:
            booking_key = (
                booking['summary'],
                booking['start'].isoformat(),
                booking['end'].isoformat()
            )
            
            if booking_key not in seen:
                seen.add(booking_key)
                unique_bookings.append(booking)
        
        duplicates_removed = len(month_bookings) - len(unique_bookings)
        
        # Create filename
        filename = f"parentzone_bookings_{month_name}.ics"
        filepath = os.path.join(output_folder, filename)
        
        # Start iCal file
        ical_content = [
            "BEGIN:VCALENDAR",
            "VERSION:2.0",
            "PRODID:-//ParentZone Scraper//EN",
            "CALSCALE:GREGORIAN",
            "METHOD:PUBLISH",
            f"X-WR-CALNAME:ParentZone {month_name.replace('_', ' ')}",
            "X-WR-TIMEZONE:Europe/London",
        ]
        
        # Add each unique booking as an event
        for idx, booking in enumerate(unique_bookings):
            # Generate unique ID
            uid = f"parentzone-{booking['start'].strftime('%Y%m%d%H%M%S')}-{idx}@parentzone.me"
            
            # Format datetime for iCal
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
        
        # Write to file
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
    
    # Validate configuration
    if PARENTZONE_USERNAME == "your_email@example.com":
        print("❌ ERROR: Please edit the script and add your credentials!")
        print("   Update PARENTZONE_USERNAME and PARENTZONE_PASSWORD")
        input("\nPress Enter to exit...")
        return
    
    driver = None
    all_bookings = []
    scraped_months = []  # Track which months we've scraped to avoid duplicates
    
    try:
        # Setup Chrome driver
        print("🌐 Starting Chrome browser...")
        driver = setup_driver()
        
        # Login
        if not login_to_parentzone(driver, PARENTZONE_USERNAME, PARENTZONE_PASSWORD):
            print("❌ Login failed. Please check your credentials.")
            input("\nPress Enter to exit...")
            return
        
        # Navigate to bookings
        print("📆 Navigating to bookings page...")
        driver.get(PARENTZONE_BOOKINGS_URL)
        time.sleep(5)
        
        # Scrape multiple months
        print(f"\n🔍 Scraping up to {MONTHS_TO_SCRAPE} month(s)...\n")
        
        for month_idx in range(MONTHS_TO_SCRAPE):
            # Get current month name to track what we've scraped
            try:
                month_year_element = driver.find_element(
                    By.CSS_SELECTOR,
                    "h6.MuiTypography-root.MuiTypography-h6.MuiTypography-noWrap.css-1dw86cl-titleWithButtons"
                )
                current_month_year = month_year_element.text.replace('Bookings - ', '').strip()
                
                # Check if we've already scraped this month (means we've looped back)
                if current_month_year in scraped_months:
                    print(f"⚠️ Already scraped {current_month_year}, stopping to avoid duplicates.")
                    break
                
                scraped_months.append(current_month_year)
                
            except:
                print("⚠️ Could not read current month, stopping.")
                break
            
            # Scrape current month
            bookings = extract_bookings_from_page(driver)
            all_bookings.extend(bookings)
            
            # Move to next month (except on last iteration)
            if month_idx < MONTHS_TO_SCRAPE - 1:
                if not click_next_month(driver):
                    print("⚠️ Could not navigate further, stopping here.")
                    break
        
        # Generate iCal files (one per month)
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