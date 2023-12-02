import importlib
import subprocess
import sys

# Check if Selenium is installed
try:
    importlib.import_module('selenium')
except ImportError:
    print("Selenium is not installed. Installing it now...")

    # Run a script to install Selenium using subprocess
    try:
        subprocess.run([sys.executable, '-m', 'pip', 'install', 'selenium'], check=True)
    except subprocess.CalledProcessError as e:
        print(f"Error installing Selenium: {e}")
        sys.exit(1)

from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.options import Options
def typeAndClick(element, textToType):
    element.send_keys(textToType)
    element.send_keys(Keys.ENTER)

def successPopup():
    import tkinter as tk
    from tkinter import messagebox

    # Create the main window
    root = tk.Tk()
    root.withdraw()  # Hide the main window

    # Create a popup window
    result = messagebox.showinfo("Auto Laborer", "Associate(s) Labor Tracked")
    if result == 'ok':
        root.destroy()  # Close the main window
        sys.exit()      # Exit the program

def LT(badgeIDs,laborPath):
    website_url = "https://fcmenu-iad-regionalized.corp.amazon.com/HDC3/laborTrackingKiosk"
    options = Options()
    options.add_argument("--headless")
    driver = webdriver.Chrome()
    driver.get(website_url)

    # Find an input field and type characters
    input_element = driver.find_element('xpath', '//*[@id="calmCode"]')
    typeAndClick(input_element,laborPath)

    elements = driver.find_elements('xpath', '//*[@id="badgeScanGuidance"]/h1')
    if elements:
        loginBadge = '12730876'
        input_element = driver.find_element('xpath', '//*[@id="badgeBarcodeId"]')
        typeAndClick(input_element,loginBadge)

    # input_element = driver.find_elements('//*[@id="qlInput"]')
    driver.get(website_url)
    input_element = driver.find_element('xpath', '//*[@id="calmCode"]')
    typeAndClick(input_element,laborPath)
    input_element = driver.find_element('xpath', '//*[@id="trackingBadgeId"]')
    typeAndClick(input_element,badgeIDs)
    successPopup()

def readExcel():
    try:
        importlib.import_module('pandas')
    except ImportError:
        print("Pandas is not installed. Installing it now...")

        # Run a script to install Pandas using subprocess
        try:
            subprocess.run([sys.executable, '-m', 'pip', 'install', 'pandas'], check=True)
            import pandas as pd
        except subprocess.CalledProcessError as e:
            print(f"Error installing Pandas: {e}")
            sys.exit(1)
    import pandas as pd
    # Specify the path to your Excel file
    excel_file_path = "C:/Users/nuneadon/Downloads/StaffingBoard_P3.xlsm"

    # Specify the sheet name (if there are multiple sheets)
    sheet_name = 'DOCK'

    # Specify the range of columns you want to read (M2 through M7)
    column_range = 'M2:M7'

    # Read the Excel file into a DataFrame
    df = pd.read_excel(excel_file_path, sheet_name=sheet_name, usecols=column_range)

    # Access the specified column
    column_data = df['M2']

    # Print or use the data as needed
    print(column_data)

# LT('','TOTOL')
readExcel()
