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
import time, os


#Reference: https://www.cnblogs.com/hong-fithing/p/9656221.html
def build_directory():
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
    except BaseException as msg:
        print(msg)

def set_environment_chrome():
    chrome_options = webdriver.ChromeOptions()
    chrome_options.add_argument("--incognito")
    chrome_options.add_argument("--ignore-certificate-errors")
    chrome_options.add_argument("--start-maximized")
    driver_chrome = webdriver.Chrome(executable_path='../chromedriver.exe',options=chrome_options)
    return driver_chrome


        
def main(input_path, File_Path):
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
    
    build_directory()

    # 寫入所需資料
    for row in df_input.index:
        
        URL = 'https://findbiz.nat.gov.tw/fts/query/QueryBar/queryInit.do'
        driver.get(URL)
        time.sleep(2)
        forVerify = df_input.loc[row, 'companies']
        # 加上編號
        forVerifyNo = f'{row}' + ". " + forVerify
        # 點擊「分公司」
        needsCheck = WebDriverWait(driver, 2).until(
            EC.visibility_of_element_located((By.XPATH, '(//input[@name="qryType"])[2]')))
        needsCheck.click()
        
        # 輸入要檢查的項目
        driver.find_element(By.XPATH, '//*[@id="qryCond"]').send_keys(forVerify)
    
        time.sleep(2)
        # 按下送出查詢
        driver.find_element(By.XPATH, '//*[@id="qryBtn"]').click()
        time.sleep(2)
        
        # 點擊搜尋出來的第一筆資料，如果沒有資料就拍照留存。
        try:
            driver.find_element(By.XPATH, '//*[@id="vParagraph"]/div[1]/div[1]/a').click()
        except:
            driver.save_screenshot(File_Path + '\\' + f'{forVerifyNo}.png')
            print("[INFO] 無查詢結果，拍照留存！") 
        time.sleep(2)
        
        # 點選法人董監網路試用版 (要切換frame)
        try:
            iframe = driver.find_element(By.XPATH, '//*[@id="tabAdvanced"]') 
            iframe.click()
        except:
           driver.switchTo().alert().accept()
           iframe = driver.find_element(By.XPATH, '//*[@id="tabAdvanced"]') 
           iframe.click()
        time.sleep(2)
        
        # 點選展開全部
        iframe2 = driver.find_element(By.XPATH, '//*[@id="mivFrme"]') 
        driver.switch_to.frame(iframe2)
        button = driver.find_element(By.XPATH, '//*[@id="tab_table"]/div[2]/button[2]')
        button.click()
        time.sleep(2)
        
        # 找到所有有關係的公司
        nodes = driver.find_elements(By.XPATH, '//*[@id="chart"]')
        time.sleep(2)
        # https://ithelp.ithome.com.tw/articles/10202725
        for i in range(len(nodes)):
            # 如果沒資料，則txt檔案中顯示「無資料」
            if nodes[i].text is None:
                print("無資料")
                with open(File_Path + '\\' + f"{forVerifyNo}.txt","w") as f:
                    f.write("無資料")
            else:
                print(nodes[i].text)
                with open(File_Path + '\\' + f"{forVerifyNo}.txt","w") as f:
                    f.write(nodes[i].text)
        print(f'[INFO] 現在做完第{row}筆')
        time.sleep(1)
        
    print(f"[INFO] 總共{len(df_input)}筆，全部完成了！")
    driver.close()

if __name__ == '__main__':
    input_path = '..//companies.csv'
    directory_time = time.strftime("%Y-%m-%d", time.localtime(time.time()))
    File_Path = os.getcwd() + '\\' + directory_time + '\\'
    main(input_path, File_Path)
