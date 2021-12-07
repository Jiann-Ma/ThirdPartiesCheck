# -*- coding: utf-8 -*-
"""
Created on Tue Dec  7 09:09:08 2021

@author: ESB20914
"""

# selenium libraries
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait

import pandas as pd
import time, random

import time
import os


    
def checkScreenshot(driver, row):
    print(f"[INFO] 開始檢查第 {row} 列的representatives")
    picture_time = time.strftime("%Y-%m-%d-%H_%M_%S", time.localtime(time.time()))
    print(picture_time)
    try:
        picture_url=driver.get_screenshot_as_file('Y:\\治理規劃部\\MJT\\RPA\\ThirdPartiesCheck\\ThirdPartiesScreenshot'+ picture_time +'.png')
        print("%s：[INFO] 拍完該筆的照片了！" % picture_url)
    except BaseException as msg:
        print(msg)  

def set_environment_chrome():
    chrome_options = webdriver.ChromeOptions()
    chrome_options.add_argument("--incognito")
    chrome_options.add_argument("--ignore-certificate-errors")
    chrome_options.add_argument("--start-maximized")
    driver_chrome = webdriver.Chrome(executable_path=r'C:\chromedriver.exe',options=chrome_options)
    return driver_chrome

def action(driver, row, forVerify):        
    # 輸入要檢查的項目
    needsCheck = WebDriverWait(driver, random.randint(2, 4)).until(
        EC.visibility_of_element_located((By.ID, 'txtKW')))
    
    needsCheck.send_keys(forVerify)
    
    time.sleep(random.randint(2, 4))
    
    driver.find_element(By.XPATH, '//*[@id="btnSimpleQry"]').click()
    time.sleep(2)
    checkScreenshot(driver, row)
    driver.refresh()
    driver.find_element(By.ID, 'txtKW').clear()
        
def main(input_path):
    """
    至 https://law.judicial.gov.tw/FJUD/default.aspx 執行檢索
    input_path: str, 輸入資料源的路徑 ( 含副檔名 )
    """
    # 讀取資料輸入源
    df_input = pd.read_csv('representatives.csv', encoding= 'utf-8-sig',
                             converters={'representatives':str})
    print(f'[INFO] 已成功讀取資料，筆數: {len(df_input)}，內容如以下：')
    print(df_input['representatives']) 
    

    driver_chrome = set_environment_chrome()     
    driver = driver_chrome
    # 寫入所需資料
    for row in df_input.index:
        URL = 'https://law.judicial.gov.tw/FJUD/default.aspx'
        driver.get(URL)
        time.sleep(random.randint(2, 4))
        forVerify = df_input.loc[row, 'representatives']
        print(f'[INFO] 現在做第{row}筆')
        action(driver, row, forVerify)
        
    print(f"[INFO] 總共{len(df_input)}筆，全部完成了！")
        

if __name__ == '__main__':
    input_path = 'representatives.csv'
    main(input_path)