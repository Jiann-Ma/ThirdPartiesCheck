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
def checkScreenshot(driver, forVerifyNo, row):
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
    print(f'[INFO] 拍完第{row}筆的照片了！')

    time.sleep(2)
    

def set_environment_chrome():
    chrome_options = webdriver.ChromeOptions()
    chrome_options.add_argument("--incognito")
    chrome_options.add_argument("--ignore-certificate-errors")
    chrome_options.add_argument("--start-maximized")
    driver_chrome = webdriver.Chrome(executable_path='../chromedriver.exe',options=chrome_options)
    return driver_chrome

        
def main(input_path):
    """
    至 https://law.judicial.gov.tw/FJUD/default.aspx 執行檢索
    input_path: str, 輸入資料源的相對路徑 ( 含副檔名 )
    """
    # Reference: https://blog.impochun.com/excel-big5-utf8-issue/
    df_input = pd.read_csv('..//companies.csv', encoding= 'utf-8-sig',
                             converters={'companies':str})
    print(f'[INFO] 已成功讀取資料，筆數: {len(df_input)}，內容如以下：')
    print(df_input['companies']) 
    

    driver = set_environment_chrome()     

    # 寫入所需資料
    # i = 1
    for row in df_input.index:
        
        URL = 'https://law.judicial.gov.tw/FJUD/default.aspx'
        driver.get(URL)
        time.sleep(2)
        forVerify = df_input.loc[row, 'companies']
        
        # 加上編號
        forVerifyNo = f'{row}' + ". " + forVerify
        
        # 輸入要檢查的項目
        needsCheck = WebDriverWait(driver, 2).until(
            EC.visibility_of_element_located((By.ID, 'txtKW')))
    
        needsCheck.send_keys(forVerify)
    
        time.sleep(2)
        
        #按下送出查詢
        driver.find_element(By.XPATH, '//*[@id="btnSimpleQry"]').click()
        time.sleep(2)
        
        #截圖
        checkScreenshot(driver, forVerifyNo, row)
        time.sleep(2)
        
        #點下「判決書查詢」，回到查詢頁面
        driver.find_element(By.XPATH, '//a[@href="/FJUD/default.aspx"]').click()
        
        #清除輸入的字(cookies)
        driver.delete_all_cookies()

        print(f'[INFO] 現在做完第{row}筆')
        # i = i + 1

    print(f"[INFO] 總共{len(df_input)}筆，全部完成了！")
    driver.close()

if __name__ == '__main__':
    input_path = '..//companies.csv'
    main(input_path)
