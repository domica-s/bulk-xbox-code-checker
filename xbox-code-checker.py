# imports for selenium
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.common.keys import Keys
# import for JSON (to obtain login details)
import json
# import for reading excel files (to obtain code)
import xlsxwriter
import pandas as pd
# import for other functions
from datetime import datetime
import time
import random
import pyautogui

# setup selenium webdriver
# add selenium driver path
driver_path = "\\Program Files\\chromedriver\\chromedriver.exe"
# custom selenium options
chr_options = Options()
chr_options.add_experimental_option("detach", True)
chr_options.add_experimental_option('excludeSwitches', ['enable-logging'])
chr_driver = webdriver.Chrome(driver_path, options=chr_options)

# getting dataset from excel file
xbox_codes = pd.read_excel('active-codes.xlsx', usecols="A").values.tolist()

# read JSON file to get config details
f = open('config.json')
data_dict = json.load(f)
xbox_email = data_dict["xbox_email"]
xbox_password = data_dict["xbox_password"]
xbox_ss_path = data_dict["xbox_ss_path"]
xbox_err_path = data_dict["xbox_err_path"]

# get the current time (time when "now" is first executed)
now = datetime.now()
date_time = now.strftime("%d-%m-%Y %H:%M:%S")
date_time_file = now.strftime("%Y-%m-%d_%H_%M_%S")

# give delay when entering keys to prevent robot detection
def sendKeyDelay(target_id, target_text):
    for char in target_text:
        target_id.send_keys(char)
        time.sleep(random.randrange(0, 1))

# get screenshot of successful (valid) codes
def getSS(game, code):
    if (game == "xbox"):
        desired_PATH = xbox_ss_path

    #saving the screenshot to the proper directory
    myScreenshot = pyautogui.screenshot()
    myScreenshot.save(desired_PATH + code + ".jpg")

# XBOX methods
xbox_error = []
xbox_updated = []

def loginXbox():
    # load url
    xbox_link = "https://redeem.microsoft.com/?wa=wsignin1.0"
    chr_driver.get(xbox_link)
    time.sleep(random.randrange(2, 3)) # let page load fully

    # fill in username (email)
    email_box = chr_driver.find_element_by_id("i0116")
    sendKeyDelay(email_box, xbox_email)

    # click next button
    time.sleep(random.randrange(1, 2))
    next_btn = chr_driver.find_element_by_id("idSIButton9")
    next_btn.click()

    # send password with delay
    time.sleep(random.randrange(3, 5))
    pass_box = chr_driver.find_element_by_id("i0118")
    sendKeyDelay(pass_box, xbox_password)

    # click sign in button
    time.sleep(random.randrange(1, 2))
    sign_in_btn = chr_driver.find_element_by_id("idSIButton9")
    sign_in_btn.click()

def checkXbox(code):
    # now that redeem page is active
    # access 'iframe' for the code box
    time.sleep(random.randrange(7, 10))
    chr_driver.switch_to.frame('wb_auto_blend_container')

    # target boxes for the codes
    code_box = chr_driver.find_element_by_xpath("//*[@id='tokenString']")
    submit_btn = chr_driver.find_element_by_id("nextButton")

    # send the keys to the code_box
    sendKeyDelay(code_box, code)

    # if key is invalid, report, else (if valid) return denomination
    # if key is invalid, the next button will be disabled
    try:
        # if functions runs until here, it means code is valid
        # click next button in order to get denomination
        time.sleep(random.randrange(2, 3))  
        submit_btn.click()
        # if error -> Exception has occurred: ElementClickInterceptedException
        # find element that states the amount, and get the text
        time.sleep(random.randrange(3, 5))
        denom_message = chr_driver.find_element_by_xpath("//*[@id='pageContent']/div/div[2]/h2").text
        # sample text is /10.00 USD Microsoft gift card/
        # get the string before /USD/
        idx_msoft = denom_message.index("Microsoft")
        final_denom = denom_message[:idx_msoft]
        print("Success: " + code, final_denom)
        # get screenshot of this screen
        getSS("xbox", code)
    except:
        # if code is invalid, will go here
        # report time checked
        print("Error: " + code, date_time)
        # exit the code checking for this code
        xbox_error.append(code)

    time.sleep(random.randrange(2, 3))    
    chr_driver.refresh()

def updateXbox():
    if xbox_error:
        for x in xbox_codes:
            str_x = ''.join(x)
            if str_x not in xbox_error:
                xbox_updated.append(str_x)
                continue

        filename = xbox_err_path + date_time_file + ".txt"
        with open(filename, 'w') as file:
            file.write("ERROR CODES: \n")
            for y in xbox_error:
                file.write(y + "\n")

def runXbox():
    if xbox_codes:
        loginXbox()
        for x in xbox_codes:
            str_x = ''.join(x)
            checkXbox(str_x)
            #refresh page
            chr_driver.refresh()
        chr_driver.close() 
        updateXbox()

def updateExcel():
    #update .xlxs file
    workbook = xlsxwriter.Workbook('active-codes.xlsx')
    worksheet = workbook.add_worksheet()

    #prepare header
    worksheet.write(0, 0, "XBOX") 

    for i in range(len(xbox_updated)):
        worksheet.write(i + 1, 0, xbox_updated[i])

    workbook.close()

# 'MAIN'
runXbox()
updateExcel()