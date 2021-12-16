# -*- coding: utf-8 -*-
"""
Created on Mon Dec 13 17:08:48 2021

@author: ESB20914
"""

# selenium libraries
# from selenium import webdriver
# from selenium.webdriver.common.by import By
# from selenium.webdriver.support import expected_conditions as EC
# from selenium.webdriver.support.ui import WebDriverWait
# from selenium.webdriver.common.keys import Keys

import pandas as pd
import time, os


def build_directory(forVerify):
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
    time.sleep(1)
        

def main(input_path):
    df_input = pd.read_csv('..//companies.csv', encoding= 'utf-8-sig',
                             converters={'companies':str})
    print(f'[INFO] 已成功讀取資料，筆數: {len(df_input)}，內容如以下：')
    # print(df_input['companies'])
    # print(type(df_input))
    # print(df_input)
    
    
    
    for row in df_input.index:
        forVerify = df_input.loc[row, 'companies']
        forVerify = f'{row}' + ". " + forVerify

        build_directory(forVerify)
    
if __name__ == '__main__':
    input_path = '..//companies.csv'
    main(input_path)