# -*- coding: utf-8 -*-
"""
Created on Mon Dec 13 17:08:48 2021

@author: ESB20914
"""

# selenium libraries
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
# from selenium.webdriver.common.keys import Keys

import pandas as pd
import time, os


def set_environment_chrome():
    chrome_options = webdriver.ChromeOptions()
    chrome_options.add_argument("--incognito")
    chrome_options.add_argument("--ignore-certificate-errors")
    chrome_options.add_argument("--start-maximized")
    driver_chrome = webdriver.Chrome(executable_path='../chromedriver.exe',options=chrome_options)
    return driver_chrome


driver = set_environment_chrome() 
URL = 'https://findbiz.nat.gov.tw/fts/query/QueryBar/queryInit.do'
driver.get(URL)
time.sleep(2)
      
        # 點擊「分公司」
needsCheck = WebDriverWait(driver, 2).until(EC.visibility_of_element_located((By.XPATH, '(//input[@name="qryType"])[2]')))
needsCheck.click()
        
        # 輸入要檢查的項目
driver.find_element(By.XPATH, '//*[@id="qryCond"]').send_keys("玉山商業銀行股份有限公司")
    
time.sleep(2)
        #按下送出查詢
driver.find_element(By.XPATH, '//*[@id="qryBtn"]').click()
time.sleep(2)
        
        # 點擊搜尋出來的第一筆資料
driver.find_element(By.XPATH, '//*[@id="vParagraph"]/div[1]/div[1]/a').click()
time.sleep(2)
        

# frames = driver.find_elements(By.TAG_NAME, "iframe")

frames = driver.find_element(By.ID, '//*[@id="tabAdvanced"]')

link = "b66e947f49b6811a8bf438040d1582232d3232d7"
final_website_link = f"https://ipcsa.nat.gov.tw/miv/shareHolder/index?token={link}"

driver.get(link)

control_frame = None
for index, frame in enumerate(frames):
    print(frames)
# for index, frame in enumerate(frames):
#     if frame.get_attribute("title") == "reCAPTCHA":
#         control_frame = frame
#     if not (control_frame):
#         print("[ERROR] 找不到Recaptcha，請重新再試！")
#         # 點選法人董監網路試用版 (要切換frame)
#  frame = driver.find_element(By.ID, '//*[@id="mivFrme"]') 

# print(frames)
# driver.switch_to.frame(frame)
        
# driver.find_element(By.ID, '//*[@id="tabAdvanced"]').click()
# time.sleep(1)