# -*- coding: utf-8 -*-
"""
Created on Fri Jan  7 14:22:09 2023

@author: fedig
"""

from selenium import webdriver

from selenium.webdriver.common.by import By

import time
import pandas as pd
pd.options.mode.chained_assignment = None  # default='warn'





def getStatusIndex(excelFile):
    members = pd.read_excel(excelFile)
    compt=0
    for i in members.columns:
        compt+=1
    
    return members.columns[compt-1]

def addStatusColumn(excelFile):
    members = pd.read_excel(excelFile)
    for i in members.columns:
        if i.upper() =="STATUS":
            return 0
    members["Status"]= "inactive"
    members.to_excel(excelFile)

    

addStatusColumn("VolunteerList.xlsx")

driver = webdriver.Chrome()
driver.implicitly_wait(2)
driver.get("https://www.ieee.org/mv")
email = driver.find_element(by=By.ID, value="username")
pwd = driver.find_element(by=By.ID, value="password")
signin_button = driver.find_element(by=By.ID, value="modalWindowRegisterSignInBtn")

email.send_keys("")
pwd.send_keys("")

signin_button.click()



#Excel file name, default =  "VolunteerListUpdatedStatus1.xlsx"
members = pd.read_excel("VolunteerListUpdatedStatus1.xlsx")





for index, row in members.iterrows():
    member_id_email = driver.find_element(by=By.ID, value="number-or-email")
    email_address = row.loc["Email Address  "]
    #print(f"Row {index + 2}: {email_address}")

    error_msg = driver.find_element(by=By.ID, value="error-div")
    membership_status = driver.find_element(by=By.XPATH, value="//div[3]/p")
    #if(members.at[index, "Status"] =="inactive" or members.at[index, "Status"] =="Error" or members.at[index, "Status"] =="INACTIVE"):
    if (members.at[index, "Status"] == "Error"):
        member_id_email.send_keys(email_address)
        submit_button = driver.find_element(by=By.ID, value="submit-membership-validator-form")
        submit_button.click()
        time.sleep(2)
        if(membership_status.text == "Membership validation status"):
            members.at[index, "Status"] = "ACTIVE"
            print("ACTIVE")
        elif (error_msg.text == "Member not found or membership status is not active."):
            members.at[index, "Status"] = "INACTIVE"
            print("INACTIVE")
        elif (error_msg.text =="To discourage use of automated scripts, IEEE Membership Validator limits number of validations. You can check additional members in one minute."):
            print("Validation Limit")
            break
        else:
            members.at[index, "Status"] = "Error"
            print("Error")
        member_id_email.clear()
        members.to_excel("VolunteerListUpdatedStatus1.xlsx")

    else:
        continue




driver.quit()