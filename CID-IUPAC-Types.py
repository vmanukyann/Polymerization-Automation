from oauth2client.service_account import ServiceAccountCredentials
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.by import By
from selenium import webdriver
from PIL import Image, ImageEnhance
import pytesseract
import pyautogui
import gspread
import time
import os
import re

# API key to access the Google Sheet
scope = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]
creds_path = r'C:\Users\nsluser\Desktop\code\polyinfo-automation-for-sjee-ba26337d9e5d.json'

if not os.path.exists(creds_path):
    raise FileNotFoundError(f"The credentials file was not found at the path: {creds_path}")

creds = ServiceAccountCredentials.from_json_keyfile_name(creds_path, scope)
client = gspread.authorize(creds)

# Access the Google Sheet
sheet_url = "https://docs.google.com/spreadsheets/d/1c75WRiYRwgtIH34SJrRfr4YrzMWk8P9MvtZDF7hpeD0/edit?gid=2006529038#gid=2006529038"
try:
    sheet = client.open_by_url(sheet_url).sheet1
except gspread.exceptions.SpreadsheetNotFound:
    raise Exception(f"Could not find the spreadsheet with the URL: {sheet_url}")

# The PolyInfo URL
polyinfo_url = "https://polymer.nims.go.jp/PoLyInfo/search"

def login():
    driver = webdriver.Chrome()  # Ensure the driver path is correctly set if needed
    driver.get(polyinfo_url)
    print("Opened PolyInfo URL")

    # Wait for the user to manually complete the login
    print("Please complete the login manually. Waiting for 30 seconds...")
    time.sleep(30)  # Adjust sleep time as necessary
    print("Login wait period over")

    return driver

def get_to_info(driver, pid):

    print("Attempting to locate the first input field...")
    search_boxes = WebDriverWait(driver, 10).until(
        EC.presence_of_all_elements_located((By.TAG_NAME, 'input'))
    )
    search_box = search_boxes[0]

    print(f"Entering PID: {pid}")
    driver.execute_script("arguments[0].value = arguments[1];", search_box, pid)

    polymer_search_button = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.LINK_TEXT, "POLYMER SEARCH"))
    )
    polymer_search_button.click()

    time.sleep(5)

    first_link = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.XPATH, "//a[contains(@class, 'polymer_name_link')]"))
    )
    first_link.click()

    time.sleep(5)

    print("Clicking on Polymerization paths & Candidate monomers link")
    polymerization_link = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.ID, "polymerization_paths"))
    )
    driver.execute_script("arguments[0].click();", polymerization_link)
    print("Polymerization paths & Candidate monomers link clicked.")

def scroll_and_screenshot(driver, path):
    print("Screenshot time")
    driver.execute_script("window.scrollBy(0, 230);")
    time.sleep(5)
    
    # Capture the screenshot
    screenshot = pyautogui.screenshot()
    screenshot.save(path)
    print(f"Screenshot taken and saved as {path}")

def preprocess_image(image_path):
    """Preprocess the image for better OCR results."""
    image = Image.open(image_path)
    # Convert to grayscale
    image = image.convert('L')
    # Increase contrast
    image = ImageEnhance.Contrast(image).enhance(2)
    # Apply thresholding
    image = image.point(lambda x: 0 if x < 140 else 255, '1')
    return image

def ocr_image(image_path):
    pytesseract.pytesseract.tesseract_cmd = r"C:\Users\nsluser\AppData\Local\Programs\Tesseract-OCR\tesseract.exe"
    print(f"Performing OCR on image: {image_path}")
    image = preprocess_image(image_path)
    text = pytesseract.image_to_string(image)
    return text

def get_proper_info(image_path):
    text = ocr_image(image_path)
    print("Extracted text from image:")
    print(text)  # Debugging output

    # Regex to capture CIDs, IUPAC names, and types
    cids = re.findall(r'CID:\s*(\d+)', text)
    iupac_names = re.findall(r'IUPAC name:\s*([^\n\r]+)', text)
    types = re.findall(r'Type:\s*([^\n\r]+)', text)

    # Debugging output
    print(f"Extracted CIDs: {cids}")
    print(f"Extracted IUPAC Names: {iupac_names}")
    print(f"Extracted Types: {types}")

    # Clean up the data
    iupac_names = [name.strip() for name in iupac_names]
    types = [type_.strip() for type_ in types]

    print(f"Cleaned IUPAC Names: {iupac_names}")
    print(f"Cleaned Types: {types}")

    return cids, iupac_names, types

def update_sheet(row_index, cids, iupac_names, types):
    print(f"Updating sheet for row {row_index} with CIDs: {cids} and IUPACs: {iupac_names}")

    num_cids = len(cids)
    num_iupac_names = len(iupac_names)
    num_types = len(types)

    # Update the original row with the first CID and associated data
    if num_cids > 0:
        sheet.update_cell(row_index, 5, cids[0])
        if num_iupac_names > 0:
            sheet.update_cell(row_index, 7, iupac_names[0])
        if num_types > 0:
            sheet.update_cell(row_index, 4, types[0])

    # Initialize the starting index for new rows
    new_row_index = row_index

    # Process remaining info in pairs
    for i in range(1, num_cids):
        new_row_index += 1
        sheet.insert_row([''] * sheet.col_count, new_row_index)
        
        # Update current row with the current CID and IUPAC
        sheet.update_cell(new_row_index, 5, cids[i])
        if i < num_iupac_names:
            sheet.update_cell(new_row_index, 7, iupac_names[i])
        if i < num_types:
            sheet.update_cell(new_row_index, 4, types[i])

    print(f"Update complete for row {new_row_index}")

def process(driver, row_index, pid):
    get_to_info(driver, pid)
    screenshot_path = f"screenshot_{row_index}.png"
    scroll_and_screenshot(driver, screenshot_path)
    
    # Extract CIDs, IUPAC names, and types from the image
    cids, iupac_names, types = get_proper_info(screenshot_path)
    
    # Update the sheet with the extracted data
    update_sheet(row_index, cids, iupac_names, types)


    if os.path.exists(screenshot_path):
        print(f"Deleting {screenshot_path}")
        os.remove(screenshot_path)

# Directly call the process function for the first row
def main():
    driver = login()
    print("Starting the process for the first row")
    
    # Hardcoded row index and PID
    row_index = 1
    pid = sheet.cell(row_index, 2).value
    
    if pid:
        process(driver, row_index, pid)
    #except:
        # CAPTCHA SOLVER
            
    driver.quit()
    print("Driver quit")

# Run the script
main()
