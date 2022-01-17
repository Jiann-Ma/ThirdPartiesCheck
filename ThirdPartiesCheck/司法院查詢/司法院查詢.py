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
import time, os, re
import win32gui
import win32con
import logging

#Reference: https://eip.esunbank.com.tw/sites/PT/DF_RPA/_layouts/15/start.aspx#/Lists/List/Flat.aspx?RootFolder=%2fsites%2fPT%2fDF%5fRPA%2fLists%2fList%2f%5bChromedriver%5d%20%e8%87%aa%e5%8b%95%e6%8a%93%e5%8f%96%e5%b0%8d%e6%87%89%20Chrome%20%e7%89%88%e6%9c%ac%e7%9a%84%e7%a8%8b%e5%bc%8f&FolderCTID=0x012002003EAAE742E9793F48917DC83158D94F98
def get_chrome_driver():
    pattern = r'(\d+)\.(\d+)\.(\d+)'
    cmd = r'reg query "HKEY_CURRENT_USER\Software\Google\Chrome\BLBeacon" /v version'
    stdout = os.popen(cmd).read()
    version = re.search(pattern, stdout)
    if not version:
        raise ValueError(f'Could not get version for Chrome with this command: {cmd}')
    current_version = version.group(1)
    chromedriver_path = f"\\\\eip.esunbank.com.tw@SSL\\DavWWWRoot\\sites\\C010\\DocLib1\\kentsai\\chromedriver\\{current_version}\\chromedriver.exe"
    if not os.access(chromedriver_path, os.W_OK):
        logging.info('看不到暫存資料夾, 打開看看')
        dir_path = os.path.dirname(os.path.dirname(chromedriver_path))
        os.startfile(dir_path)
        for i in range(60):
            hwnd = win32gui.FindWindow(None, 'chromedriver')
            if hwnd != 0:
                win32gui.PostMessage(hwnd,win32con.WM_OPEN,0,0)
                logging.info('關閉資料夾')
                break
            logging.info('沒看到資料夾被打開')
            time.sleep(1)
    
    if os.path.exists(chromedriver_path):
        return chromedriver_path
    else:
        logging.info(f'請與RPA小組聯繫, 缺chromedriver v{current_version}, 改讀同目錄chromedriver.exe')
        chromedriver_path = os.path.abspath('chromedriver.exe')
        if not os.path.exists(chromedriver_path):
            logging.info('也不存在exe, 請下載解壓縮, 將chromedriver.exe複製到RPA當前目錄, 再次執行程式')
            os.popen('start chrome https://chromedriver.chromium.org/downloads')
        return chromedriver_path



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
    driver_chrome = webdriver.Chrome(get_chrome_driver(),options=chrome_options)
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

    print(f"[INFO] 總共{len(df_input)}筆，全部完成了！")
    driver.close()

if __name__ == '__main__':
    input_path = '..//companies.csv'
    main(input_path)
