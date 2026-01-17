# ParentZone-Booking-Scraper
a simple python code to scrape bookings from ParentZone and produce ical files for each month

This was created using a mixture of Google Dev AI and Claude - i kno w very little about coding and scraping but it works!

Pre-requisities:
Python (with environmental varables)
Web-driver Manager (can install from cmd pip command once python installed using this prompt: pip install webdriver-manager)
Selenium as Web Driver (pip install selenium)

Take the .py file.
Add your login details where it is commented to do so and Save
Run the script whenever you like and it will scrape the current month and as many as you have stated in the commented section (default to 3). It will pull multiple entries for a day if there are any and remove duplicates where the next month has the previous month dates in due to the display.
