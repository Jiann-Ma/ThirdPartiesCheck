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
def build_directory(driver, forVerify):
    #獲取當前目錄
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
    

def set_environment_chrome():
    chrome_options = webdriver.ChromeOptions()
    chrome_options.add_argument("--incognito")
    chrome_options.add_argument("--ignore-certificate-errors")
    chrome_options.add_argument("--start-maximized")
    driver_chrome = webdriver.Chrome(executable_path='../chromedriver.exe',options=chrome_options)
    return driver_chrome


        
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
        iframe = driver.find_element(By.XPATH, '//*[@id="tabAdvanced"]') 
        iframe.click()
        driver.switch_to.frame(iframe)
        time.sleep(2)
        
        # 點選展開全部
        button = driver.find_element(By.XPATH, '//*[@id="tab_table"]/div[2]/button[2]')
        button.click()
        
        nodes = driver.find_elements(By.XPATH, '//*[@id="chart"]')
        
        time.sleep(2)
        # https://ithelp.ithome.com.tw/articles/10202725
        for i in range(len(nodes)):
            print(nodes[i].text)

        with open("test.txt","a+") as f:
            f.write(nodes[i].text)
        time.sleep(2)

        print(f'[INFO] 現在做完第{i}筆')
        i = i + 1

    print(f"[INFO] 總共{len(df_input)}筆，全部完成了！")
    driver.close()

if __name__ == '__main__':
    input_path = '..//companies.csv'
    main(input_path)
