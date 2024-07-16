from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from PIL import Image
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import pytesseract
import os

# API key to access the Google Sheet
scope = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]
# Use raw string for file path or ensure the path is correct
creds_path = r"/Users/manukyan/Downloads/polymer-info-automation-5e7438fb95d5 (1).json"
if not os.path.exists(creds_path):
    raise FileNotFoundError(f"The credentials file was not found at the path: {creds_path}")

creds = ServiceAccountCredentials.from_json_keyfile_name(creds_path, scope)
client = gspread.authorize(creds)

# Access the Google Sheet
sheet_url = "https://docs.google.com/spreadsheets/d/1wYjfBPM1oQ2RrOyZst8i9LyWlGMglCGTK02PNRXBjZg/edit?gid=2006529038"
sheet = client.open_by_url(sheet_url).sheet1  # Assuming you want the first sheet

# Define the PolyInfo URL
polyinfo_url = "https://polymer.nims.go.jp/PoLyInfo/search"  # Adjust to the correct PolyInfo URL

# Polymer Info Column 
column_i_values = sheet.col_values(9) 

try:
    # Function to get information from PolyInfo
    def get_info(pid):
        driver = webdriver.Chrome()  # Ensure ChromeDriver is in your PATH
        driver.get(polyinfo_url)

        # Enter the PID
        search_box = driver.find_element(By.ID, 'column_b')  # Adjust ID as needed
        search_box.send_keys(pid)

        # Click on the "POLYMER SEARCH" button
        polymer_search_button = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.LINK_TEXT, "POLYMER SEARCH"))
        )
        polymer_search_button.click()

        # Click on the first link
        first_link = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, "//a[contains(@class, 'polymer_name_link')]"))
        )
        first_link.click()

        # Click on the "Polymerization paths & Candidate monomers" link
        polymerization_link = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.LINK_TEXT, "Candidate monomers"))
        )
        polymerization_link.click()

        # Click on the samples link 
        samples_link = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.LINK_TEXT, "Samples"))
        )
        samples_link.click()

        # Scroll to the text "Polymerization informations:"
        polymerization_info = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, "//*[contains(text(), 'Polymerization informations:')]"))
        )
        driver.execute_script("arguments[0].scrollIntoView(true);", polymerization_info)

        # Take a screenshot of the specific section
        screenshot_section(driver, polymerization_info, 'polymerization_info.png')

        # Close the driver
        driver.quit()

    # Function to take a screenshot of a specific section
    def screenshot_section(driver, element, filename):
        # Get the location and size of the element
        location = element.location_once_scrolled_into_view
        size = element.size
        driver.save_screenshot("full_screenshot.png")

        # Open the full screenshot and crop it to the element's size
        image = Image.open("full_screenshot.png")
        left = location['x']
        top = location['y']
        right = location['x'] + size['width']
        bottom = location['y'] + size['height']
        image = image.crop((left, top, right, bottom))
        image.save(filename)

    def ocr_image_to_text(image_path):
        image = Image.open(image_path)
        text = pytesseract.image_to_string(image)
        return text.strip()

    def update_sheet(row_index, text):
        sheet.update_cell(row_index, 9, text)
        # After updating the sheet
        if os.path.exists("full_screenshot.png"):
            os.remove("full_screenshot.png")

    # Example usage within your loop
    for row_index, value in enumerate(column_i_values, start=1):
        if not value:  # If the slot is empty
            pid = sheet.cell(row_index, 2).value
            if pid:
                get_info(pid)
        else:
            continue

except Exception as e:
    print("CAPTCHA detected")
    driver = webdriver.Chrome()
    # Pause the script to allow manual CAPTCHA solving
    while True:
        user_input = input("Press Enter after solving the CAPTCHA and loading the page, or type 'skip' to skip this PID: ")
        if user_input.lower() == "skip":
            break
        try:
            # Check if CAPTCHA is still present
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, 'captcha_element')))
            print("CAPTCHA still present. Please solve it.")
        except:
            print("CAPTCHA solved or not present. Resuming.")
            break

print("Script completed")