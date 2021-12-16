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

def build_directory(driver, forVerify):
    # 獲取當前目錄
    print(os.getcwd())
    try:
        File_Path = os.getcwd() + '\\' + forVerify + '\\'
        if not os.path.exists(File_Path):
            os.makedirs(File_Path)
            print('[INFO] 目錄新建成功：%s' % File_Path)
        else:
            print('[INFO] 目錄已存在！')
    except BaseException as msg:
        print('[INFO] 目錄新建失敗：%s' % msg)
    except BaseException as msg:
        print(msg)
    time.sleep(2)

df_input = pd.read_csv('..//companies.csv', encoding= 'utf-8-sig',
                             converters={'companies':str})
print(f'[INFO] 已成功讀取資料，筆數: {len(df_input)}，內容如以下：')
print(df_input['companies'])
for row in df_input.index:
    forVerify = df_input.loc[row, 'companies']
    forVerifyNo = f'{row}' + ". " + forVerify
        

        

driver = set_environment_chrome() 
URL = 'https://findbiz.nat.gov.tw/fts/query/QueryBar/queryInit.do'
driver.get(URL)
time.sleep(2)
      
        # 點擊「分公司」
needsCheck = WebDriverWait(driver, 2).until(EC.visibility_of_element_located((By.XPATH, '(//input[@name="qryType"])[2]')))
needsCheck.click()
        
        # 輸入要檢查的項目
# driver.find_element(By.XPATH, '//*[@id="qryCond"]').send_keys(forVerify)
driver.find_element(By.XPATH, '//*[@id="qryCond"]').send_keys("IHS Markit")
    
time.sleep(2)
        #按下送出查詢
driver.find_element(By.XPATH, '//*[@id="qryBtn"]').click()
time.sleep(2)
        
        # 點擊搜尋出來的第一筆資料
        
try:
    driver.find_element(By.XPATH, '//*[@id="vParagraph"]/div[1]/div[1]/a').click()
except:
    # 建立目錄，將照片存在該位置
    directory_time = time.strftime("%Y-%m-%d", time.localtime(time.time()))
    # 獲取當前目錄
    print(os.getcwd())
            
    try:
        File_Path = os.getcwd() + '\\' + directory_time + '\\'
        if not os.path.exists(File_Path):
            os.makedirs(File_Path)
            print('[INFO] 目錄新建成功：%s' % File_Path)
        else:
            print('[INFO] 目錄已存在！')
    except BaseException as msg:
        print('[INFO] 目錄新建失敗：%s' % msg)
            
    driver.save_screenshot(File_Path + '\\' + f'{forVerifyNo}.png')
    print("[INFO] 無查詢結果，拍照留存！") 
time.sleep(2)
        




# iframe = driver.find_element(By.XPATH, '//*[@id="tabAdvanced"]') 
# iframe.click()
# time.sleep(2)

# frame = driver.find_element(By.XPATH, '//*[@id="mivFrme"]') 
# driver.switch_to.frame(frame)

# button = driver.find_element(By.XPATH, '//*[@id="tab_table"]/div[2]/button[2]')
# button.click()

# time.sleep(2)

# nodes = driver.find_elements(By.XPATH, '//*[@id="chart"]')
# # print("No of frames present in the web page are: ", len(nodes))

# time.sleep(2)

# # https://ithelp.ithome.com.tw/articles/10202725
# for i in range(len(nodes)):
#     print(nodes[i].text)

# with open("test.txt","a+") as f:
#     f.write(nodes[i].text)
