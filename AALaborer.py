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
excel_file_path = "C:/Users/adn51/Downloads/StaffingBoard_P1.xlsm"

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
    
def named_range(range_start, range_end, excel_path=excel_file_path):
    srange = pd.ExcelFile(excel_path).book.defined_names[range_start].value # returns 'SHEET!$col$row:$col$row'
    erange = pd.ExcelFile(excel_path).book.defined_names[range_end].value
    sheet_name = str(srange.split('!')[0])
    column = str(srange.split('$')[1])
    start_row = int(srange.split('$')[2].split(':')[0])
    end_row = int(erange.split('$')[2].split(':')[0])
    return {'sheet':sheet_name, 'column':column, 'srow':start_row, 'erow':end_row}

def palletize_Dock():
    range = 'badges_OB_Line_Load'
    sheet = named_range(range)['sheet']
    column = named_range(range)['column']
    srow = named_range(range)['srow']
    erow = named_range(range)['erow']
    dataFrameVariable = pd.read_excel(io=excel_file_path, header=None, sheet_name=sheet, usecols=column, skiprows=srow-1, nrows=erow)
    badgeIDs = HELPER_convertToSingleLineStr(dataFrameVariable.values)
    return badgeIDs

def UIS_Dock():
    range = 'badge_UIS_DOCK'
    sheet = named_range(range)['sheet']
    column = named_range(range)['column']
    srow = named_range(range)['srow']
    erow = named_range(range)['erow']
    dataFrameVariable = pd.read_excel(io=excel_file_path, header=None, sheet_name=sheet, usecols=column, skiprows=srow-1, nrows=erow)
    badgeIDs = HELPER_convertToSingleLineStr(dataFrameVariable.values)
    return badgeIDs

def directed_counts():
    range_start = 'directedCounts_start'
    range_end = 'directedCounts_end'
    sheet = named_range(range_start, range_end)['sheet']
    column = named_range(range_start, range_end)['column']
    srow = named_range(range_start, range_end)['srow']
    erow = named_range(range_start, range_end)['erow']
    dataFrameVariable = pd.read_excel(io=excel_file_path, header=None, sheet_name=sheet, usecols=column, skiprows=srow-1, nrows=erow-srow+1)
    badgeIDs = HELPER_convertToSingleLineStr(dataFrameVariable.values)
    return badgeIDs

def quality_audits():
    range = 'quality_audits'
    sheet = named_range(range)['sheet']
    column = named_range(range)['column']
    srow = named_range(range)['srow']
    erow = named_range(range)['erow']
    dataFrameVariable = pd.read_excel(io=excel_file_path, header=None, sheet_name=sheet, usecols=column, skiprows=srow-1, nrows=erow)
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


# LT(palletizeDock(),'TOTOL')
directed_counts()

