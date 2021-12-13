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
# from selenium.webdriver.common.keys import Keys

import pandas as pd
import time, os


#Reference: https://www.cnblogs.com/hong-fithing/p/9656221.html
def checkScreenshot(driver, i):
    picture_time = time.strftime("%Y-%m-%d-%H_%M_%S", time.localtime(time.time()))
    directory_time = time.strftime("%Y-%m-%d", time.localtime(time.time()))
    print(picture_time)
    print(directory_time)
    
    #獲取當前目錄
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
    
    try:
        driver.save_screenshot('.\\' + directory_time + '\\' + picture_time + '.png')
        print(f'[INFO] 拍完第{i}筆的照片了！')
    except BaseException as msg:
        print(msg)
    time.sleep(2)
    

def set_environment_chrome():
    chrome_options = webdriver.ChromeOptions()
    chrome_options.add_argument("--incognito")
    chrome_options.add_argument("--ignore-certificate-errors")
    chrome_options.add_argument("--start-maximized")
    driver_chrome = webdriver.Chrome(executable_path='../chromedriver.exe',options=chrome_options)
    return driver_chrome

def change_iframe(driver):
    frame = driver.find_element(By.ID, '//*[@id="mivFrme"]') 
    driver.switch_to.frame(frame)
    
    driver.find_element(By.CLASS_NAME, "recaptcha-checkbox-border").click()

        
def main(input_path):
    """
    至 https://findbiz.nat.gov.tw/fts/query/QueryBar/queryInit.do 執行檢索
    input_path: str, 輸入資料源的相對路徑 ( 含副檔名 )
    """
    # Reference: https://blog.impochun.com/excel-big5-utf8-issue/
    df_input = pd.read_csv('..//companies.csv', encoding= 'utf-8-sig',
                             converters={'companies':str})
    print(f'[INFO] 已成功讀取資料，筆數: {len(df_input)}，內容如以下：')
    print(df_input['companies']) 
    

    driver = set_environment_chrome()     

    # 寫入所需資料
    i = 1
    for row in df_input.index:
        
        URL = 'https://findbiz.nat.gov.tw/fts/query/QueryBar/queryInit.do'
        driver.get(URL)
        time.sleep(2)
        forVerify = df_input.loc[row, 'companies']
        
        # 點擊「分公司」
        needsCheck = WebDriverWait(driver, 2).until(
            EC.visibility_of_element_located((By.XPATH, '(//input[@name="qryType"])[2]')))
        needsCheck.click()
        
        # 輸入要檢查的項目
        driver.find_element(By.XPATH, '//*[@id="qryCond"]').send_keys(forVerify)
    
        time.sleep(2)
        #按下送出查詢
        driver.find_element(By.XPATH, '//*[@id="qryBtn"]').click()
        time.sleep(2)
        
        # 點擊搜尋出來的第一筆資料
        driver.find_element(By.XPATH, '//*[@id="vParagraph"]/div[1]/div[1]/a').click()
        time.sleep(2)
        
        # 點選法人董監網路試用版 (要切換frame)
        frame = driver.find_element(By.ID, '//*[@id="mivFrme"]') 
        driver.switch_to.frame(frame)
        
        driver.find_element(By.ID, '//*[@id="tabAdvanced"]').click()
        time.sleep(1)
        
        # 點選展開全部
        driver.find_element(By.XPATH, '//button[text()="展開全部"]').click()
        
        
        #截圖
        # print(f'[INFO] 現在拍第{i}筆的照片！')
        checkScreenshot(driver, i)
        time.sleep(2)
        #點下「判決書查詢」，回到查詢頁面
        driver.find_element(By.XPATH, '//a[@id="tabAdvanced"]').click()
        #清除輸入的字(cookies)
        driver.delete_all_cookies()

        print(f'[INFO] 現在做完第{i}筆')
        i = i + 1

    print(f"[INFO] 總共{len(df_input)}筆，全部完成了！")
    driver.close()

if __name__ == '__main__':
    input_path = '..//companies.csv'
    main(input_path)
