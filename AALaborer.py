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

nrow = { #sheet name > path > role > number of rows down
    'PS_ICQA_Learning': {
        'directed counts': 10,
        'pallet racking' : 10,
        'simple bin counts' : 10,
        'andons' : 10,
        'IBJP': 5,
        'Stow Kiosk': 5,
        'Damages': 5,
        'IOL': 5,
        'Pallets': 5,
        'OBJP': 5,
        'runner': 5,
        'UIS': 5,
        'CPT Chase': 5,
        'Pick Skips': 5,
        'general' : 13,
        'pit': 13
    },
    'PIT': {
        'stow': 99,
        'pick': 99,
        'OB forklift': 6,
        'IB forklift': 6,
        'IB tugger': 6,
        'OB tugger': 6,
        'tugger bus route': 6,
        'center rider': 4,
        'turret': 5,
        'clamp': 5
    },
    'DOCK':{
        'UIS': 6,
        'Tote Palletize': 3,
        'Tote Palletize WS': 2,
        'Dock Auditor': 3,
        'Palletize': 30,
        'Palletize WS': 5,
        'OB Line Load': 15,
        'OB Line Load WS': 4,
        'QA Scanner': 8,
        'Grey Totes WS': 1
    }
}

laborPath = { #range-name, path, process, labor-Code
    # PS_ICQA_Learning
    'counts' : ['directedCounts_start','directed counts','ICQAPS'],
    'pallet rack audits': ['palletRacking_start' 'pallet racking','ICQAQA'],
    'SBC': ['SimpleBinCount_start','simple bin counts', 'ICQAQA'],
    'andons':  ['andon_start','andons','ICQAQA'],
    'IBJP' : ['IBJackpot_start','IBJP','IBPS'],
    'stowK' : ['stowKiosk_start','Stow Kiosk','IBPS'],
    'dmg': ['damage_start','Damages','IBPS'],
    'iol': ['IOL_start','IOL','IBPS'],
    'palletPS': ['palletPS_start', 'Pallets','IBPS'],
    'OBJP': ['OBJackpot_start','OBJP','PSTOPS'],
    'OBrunner':['OBRunner_start','runner','PSTOPS'],
    'UISPS': ['UISPS_start','UIS','PSTOPS'],
    'CPT': ['CPTChase_start','CPT Chase','PSTOPS'],
    'skips': ['pickSkips_start','Pick Skips','PSTOPS'],
    'LearningGeneral': ['learningGeneral_start','general','FCSCH'],
    'LearningPIT': ['learningPIT_start','pit','PITCLASS'],
    # PIT
    'stow': ['stow_start','PIT','stow',''], # stowers don't get labor tracked
    'OBForklift': ['OBForklift_start','PIT','OB forklift',''],                                  # needs labor code
    'IBForklift': ['IBForklift_start','PIT','IB forklift',''],                                  # needs labor code
    'IBTug': ['IBTugger_start','PIT','IB tugger',''],                                           # needs labor code
    'OBTug': ['OBTugger_start','PIT','OB tugger',''],                                           # needs labor code
    'BusTug': ['tuggerBusRoute_start','PIT','tugger bus route',''],                             # needs labor code
    'CR': ['centerRider_start','PIT','center rider',''],                                        # needs labor code
    'Turret': ['turret_start','PIT','turret',''], #pallet stower, not labor tracked
    'clamp': ['clampTruck_start','PIT','clamp',''],                                             # needs labor code
    'pick': ['pick_start','PIT','pick',''], # pickers don't get labor tracked
    # DOCK
    'UIS': ['UISDock_start', 'DOCK','UIS',''],                                                  # needs labor code
    'totePalletize': ['totePalletize_start','DOCK','Tote Palletize',''],                        # needs labor code
    'totePalletizeWS': ['totePalletizeWS_start','DOCK','Tote Palletize WS',''],                 # needs labor code
    'dockAudit': ['dockAuditor_start','DOCK','Dock Auditor',''],                                # needs labor code
    'palletize': ['palletize_start','DOCK','Palletize',''],                                     # needs labor code
    'palletizeWS': ['palletizeWS_start','DOCK','Palletize WS',''],                              # needs labor code
    'OBLineload': ['OBLineLoad_start','DOCK','OB Line Load',''],                                # needs labor code
    'OBLineLoadWS': ['OBLineLoadWS_start','DOCK','OB Line Load WS',''],                         # needs labor code
    'QAScanner': ['QAScanner_start','DOCK','QA Scanner',''],                                    # needs labor code
    'grayTotesWS': ['grayTotes_start', 'DOCK','grayTotes_start', '']                            # needs labor code
}

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

    
def named_range(range_start, laborPath_label, excel_path=excel_file_path):
    srange = pd.ExcelFile(excel_path).book.defined_names[range_start].value # returns 'SHEET!$col$row:$col$row'
    sheet_name = str(srange.split('!')[0])
    column = str(srange.split('$')[1])
    start_row = int(srange.split('$')[2].split(':')[0])
    return {'sheet':sheet_name, 'column':column, 'srow':start_row, 'erow':nrow[sheet_name][laborPath_label]}
    
def badges(range_name, pathlabel):
    info = named_range(range_name, pathlabel)
    dataFrameVariable = pd.read_excel(io=excel_file_path, header=None, sheet_name=info['sheet'], usecols=info['column'], skiprows=info['srow']-1, nrows=info['erow'])
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


def OBLineLoad():
    return badges(laborPath['OBLineload'][0], laborPath['OBLineload'][2]), laborPath['OBLineload'][3]

def UIS_Dock():
    return badges(laborPath['UIS'][0], laborPath['UIS'][3])

def directed_counts():
    return badges(laborPath['counts'][0], laborPath['counts'][1]), laborPath['counts'][2]

def palletRackingAudits():
    return badges(laborPath['pallet rack audits'][0], laborPath['pallet rack audits'][3])

def processLaborTracking(function):
    badges, CALM = function()
    # LT(badges,CALM)

processLaborTracking(directed_counts)

