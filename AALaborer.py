import importlib
import subprocess
import sys
#check if pandas is is installed
try:
    importlib.import_module('pandas')
except ImportError:
    print("Pandas is not installed. Installing it now...")
    # Run a script to install Pandas using subprocess
    try:
        subprocess.run([sys.executable, '-m', 'pip', 'install', 'pandas'], check=True)
    except subprocess.CalledProcessError as e:
        print(f"Error installing Pandas: {e}")
        sys.exit(1)
import pandas as pd
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
excel_file_path = "C:/Users/adn51/Downloads/StaffingBoard_P3.xlsm"

def HELPER_typeAndClick(element, textToType):
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

def LT(badgeIDs, laborPath):
    website_url = "https://fcmenu-iad-regionalized.corp.amazon.com/HDC3/laborTrackingKiosk"
    options = Options()
    options.add_argument("--headless")
    driver = webdriver.Chrome()
    driver.get(website_url)

    # Find an input field and type characters
    input_element = driver.find_element('xpath', '//*[@id="calmCode"]')
    HELPER_typeAndClick(input_element,laborPath)

    elements = driver.find_elements('xpath', '//*[@id="badgeScanGuidance"]/h1')
    if elements:
        loginBadge = '12730876'
        input_element = driver.find_element('xpath', '//*[@id="badgeBarcodeId"]')
        HELPER_typeAndClick(input_element,loginBadge)

    # input_element = driver.find_elements('//*[@id="qlInput"]')
    driver.get(website_url)
    input_element = driver.find_element('xpath', '//*[@id="calmCode"]')
    HELPER_typeAndClick(input_element,laborPath)
    input_element = driver.find_element('xpath', '//*[@id="trackingBadgeId"]')
    HELPER_typeAndClick(input_element,badgeIDs)
    successPopup()

def getBadges():
    palletize_AAs = palletizeDock()
    print(palletize_AAs)

def palletizeDock():
    dataFrameVariable = pd.read_excel(io=excel_file_path, sheet_name='DOCK', usecols="M:M", skiprows=0, nrows=28)
    badgeIDs = HELPER_convertToSingleLineStr(dataFrameVariable.values)
    return badgeIDs

def HELPER_convertToSingleLineStr(input):
    badgeList = str()
    for badge in input:
        n = str(badge)[1:-1]
        n = n.replace(".","")
        if n != 'nan':
            badgeList += n + " "
    return badgeList
LT(palletizeDock(),'TOTOL')
# getBadges()
# print(sys.path)

