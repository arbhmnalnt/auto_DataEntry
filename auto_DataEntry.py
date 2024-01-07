import time
import pandas as pd
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
# ============================= functions part 
def login():
    driver.find_element(By.ID, 'txtloginName').send_keys('csb28@10542')
    driver.find_element(By.ID, 'txtpassword').send_keys('csb28@10542')
    driver.find_element(By.ID, "btnsave").click()

# ============================= functions part 
# Open the web application
chrome_driver_path='./chromedriver.exe'

driver = webdriver.Chrome(service=Service(chrome_driver_path))

driver.get('http://41.33.97.114/serviceprovider/CentersDefault.aspx')

# Read Excel data
excel_data = pd.read_excel('./excell/letters.xlsx')

print(excel_data.head())

# login if necessary
login()

# Iterate through each row
for index, row in excel_data.iloc[30:].iterrows():
    driver.implicitly_wait(2)
    # get data to vars from excell
    id          = str(row['id'])
    naid        = str(row['naId'])
    doctor      = str(row['doctor'])
    digns       = str(row['digns'])
    desc        = str(row['desc'])
    cost        = str(row['cost'])    


    # click search 
    driver.find_element(By.ID, 'ContentPlaceHolder1_txtSearch').clear()
    driver.find_element(By.ID, 'ContentPlaceHolder1_txtSearch').send_keys(str(naid))

    # check if service is exist in first letters then skip

    # click create new letter
    driver.find_element(By.ID, "ContentPlaceHolder1_btnsearch").click()
    # wait if user may stop the programm for any cause 
    print(f"naid => {naid}")
    # user_input = input("Press 'Enter' to continue or 'q' to skip to following: ")
    # if user_input.lower() == 'q':
    #     print("skipping the script as per user input.")
    #     continue  # Exit the loop and end the script
    
    # start make new letter -> 
    driver.find_element(By.ID, 'ContentPlaceHolder1_LinkButton1').click()
    # ayada 
    driver.find_element(By.ID, 'ContentPlaceHolder1_ddlclinic_chzn').click()
    driver.find_element(By.ID, 'ContentPlaceHolder1_ddlclinic_chzn_o_1').click()
    # province
    driver.find_element(By.ID, 'ContentPlaceHolder1_drpGov_chzn').click()
    driver.find_element(By.ID, 'ContentPlaceHolder1_drpGov_chzn_o_21').click()

    # address
    driver.find_element(By.ID, 'ContentPlaceHolder1_txtAddress').send_keys("اسوان")

    # doctor 
    driver.find_element(By.ID, 'ContentPlaceHolder1_txtDocName').send_keys(str(doctor))
    driver.find_element(By.ID, 'ContentPlaceHolder1_txtNote').send_keys(f"التشخيص : {digns} / {desc}")
    
    driver.find_element(By.ID, 'ContentPlaceHolder1_txtFees').clear()
    driver.find_element(By.ID, 'ContentPlaceHolder1_txtPhone').send_keys("0")

    # cost
    driver.find_element(By.ID, 'ContentPlaceHolder1_txtFees').clear()
    driver.find_element(By.ID, 'ContentPlaceHolder1_txtFees').send_keys(str(cost))

    user_input = input("Press 'Enter' to continue or 'q' to skip to following: ")
    if user_input.lower() == 'q':
        print("skipping the script as per user input.")
        driver.implicitly_wait(2)
        driver.get("http://41.33.97.114/serviceprovider/CentersDefault.aspx")
# Close the browser
driver.quit()


