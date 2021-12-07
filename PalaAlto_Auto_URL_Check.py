"""
Created on Thu Oct 28 11:19:01 2021

@author: ESB20914
"""
# selenium libraries
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.keys import Keys
from msedge.selenium_tools import Edge, EdgeOptions

import pandas as pd
import time, random

# system libraries
import os
import urllib

# recaptcha libraries
import pydub
import speech_recognition as sr


def delay2to4():
    time.sleep(random.randint(2, 4))
    
def goodURLscreenshot(driver, row):
    print(f"[INFO] 開始檢查第 {row} 列的URL/Filename")
    driver.save_screenshot(f'Y:\\治理規劃部\\MJT\\RPA\\goodURLscreenshot\\goodURLscreenshot{row}.png')
    print("[INFO] 拍完該筆的照片了！") 

def set_environment_chrome():
    chrome_options = webdriver.ChromeOptions()
    # chrome_options.add_argument("--incognito")
    chrome_options.add_argument("--ignore-certificate-errors")
    chrome_options.add_argument("--start-maximized")
    # chrome_options.add_experimental_option('excludeSwitches', ['enable-automation']) #處理瀏覽器頂部顯示 ” Chrome正在由自動軟體測試控制”之問題
    driver_chrome = webdriver.Chrome(executable_path=r'C:\chromedriver.exe',options=chrome_options)
    return driver_chrome

def set_environment_edge():
    edge_options = EdgeOptions()
    edge_options.add_argument("--incognito")
    edge_options.add_argument("--start-maximized")
    driver_edge = Edge(executable_path=r'C:\msedgedriver.exe',options=edge_options)
    return driver_edge

def checkBox_and_audioButton_statusCheck(driver):
    # 點擊recaocha的checkBox
    # Auto locate recaptcha frames
    frames = driver.find_elements(By.TAG_NAME, "iframe")
    recaptcha_control_frame = None
    recaptcha_challenge_frame = None
    for index, frame in enumerate(frames):
        if frame.get_attribute("title") == "reCAPTCHA":
            recaptcha_control_frame = frame
        if frame.get_attribute("title") == "reCAPTCHA 驗證測試":
            recaptcha_challenge_frame = frame
    if not (recaptcha_control_frame and recaptcha_challenge_frame):
        print("[ERROR] 找不到Recaptcha，請重新再試！")
        
    # switch to recaptcha frame
    frames = driver.find_elements(By.TAG_NAME, "iframe")
    driver.switch_to.frame(recaptcha_control_frame)
    
    # 不管怎樣都一定要 click on checkbox to activate recaptcha
    driver.find_element(By.CLASS_NAME, "recaptcha-checkbox-border").click()
    
    # 第一個情境是出現alert
    statusCode = None
    try:
        WebDriverWait(driver, 3).until(EC.alert_is_present())
        alert = driver.switch_to.alert
        alert.accept()
        print("alert accepted")
        driver.switch_to.default_content()
        statusCode = "alert"
        return statusCode  #return後這function就結束了。
    except TimeoutException as e3:
        driver.switch_to.default_content()
        print("no alert")
        print(e3)
    
    # 第二個情境是只要打勾就好
    hasAudioButton = False
    # 切到recaptcha_challenge_frame
    frames = driver.find_elements(By.TAG_NAME, "iframe")
    driver.switch_to.frame(recaptcha_challenge_frame)
    try:
        wait = WebDriverWait(driver, 3)
        wait.until(EC.presence_of_element_located((By.XPATH, "//*[@id='recaptcha-audio-button']")))
        has_audioButton = True
    except TimeoutException as timeout:
        statusCode = "passDirectly"
        print(timeout)
        return statusCode
        
    # 第三個情境是被block，第四個情境是還要處理audio button (都有audio button)
    if has_audioButton == True:
        # get the mp3 audio file
        try:
            if driver.find_element(By.ID, "audio-source").get_attribute("src") == True:
                statusCode = "audioSrcFound"
            return statusCode
        except Exception:
            statusCode = "beBlocked"
            return statusCode

def action(driver, row, URLforVerify):        
    # 輸入要檢查的URL
    URLNeedsCheck = WebDriverWait(driver, random.randint(2, 4)).until(
        EC.visibility_of_element_located((By.ID, 'id_url')))
    URLNeedsCheck.send_keys(URLforVerify)
    
    delay2to4()

    while(True):
        statusCode = checkBox_and_audioButton_statusCheck(driver)
        # 出現情境1. alert時，動作是refresh畫面。
        if statusCode == "alert":
            driver.refresh()
            continue
        # 出現情境2. passDirectly時，動作是直接按下Search並送出。
        elif statusCode == "passDirectly":
            break
        # 出現情境3. beBlocked時，動作是換使用Edge做。
        elif statusCode == "beBlocked":
            raise Exception("changeToEdge")
        # 出現情境4. audioSrcFound時，動作點選recaptcha-audio-button。
        elif statusCode == "audioSrcFound":
            continue
            
    # wait
    print(f"[INFO] 開始等待4~9秒，等第{row}筆的recaptcha-audio-button出現")
    wait = WebDriverWait(driver, random.randint(4, 9))

    try:
        # 試著等等看recaptcha-audio-button會不會出來
        wait.until(EC.presence_of_element_located((By.XPATH, "//*[@id='recaptcha-audio-button']")))
        recaptchaAudioButton = driver.find_element(By.XPATH, "//*[@id='recaptcha-audio-button']")
        if recaptchaAudioButton.is_displayed() == True and recaptchaAudioButton.is_enabled() == True:
            print("因為 recaptchaAudioButton 有 display 與 enable，所以接下來只要打勾後就可以直接submit了。")
            driver.find_element(By.ID, "recaptcha-audio-button").click()
            wait = WebDriverWait(driver, random.randint(5, 8))
            # 只要打勾後就可以直接submit了
            driver.find_element(By.XPATH, "//div[@role='presentation']").click()
            # submit
            wait = WebDriverWait(driver, random.randint(6, 7))
            driver.find_element(By.XPATH, "//*[@id='threat-vault-app']/div[2]/div/form/div[1]/div[2]/button").click()            
        else:
            print("[WARNING] 儘管只要打勾並按下submit，但還是出錯了ˊ...所不管了，等待2~4秒後就先Submit，其他再說。")
            wait = WebDriverWait(driver, random.randint(2, 4))
            driver.find_element(By.XPATH, "//*[@id='threat-vault-app']/div[2]/div/form/div[1]/div[2]/button").click()
            # 其實重整就好
    except Exception as e1:
        try:
            # get the mp3 audio file
            src = driver.find_element(By.ID, "audio-source").get_attribute("src")
            print(f"[INFO] 第{row}筆的 Audio Source: {src}")
            
            path_to_mp3 = os.path.normpath(os.path.join(os.getcwd(), "sample.mp3"))
            path_to_wav = os.path.normpath(os.path.join(os.getcwd(), "sample.wav"))
            
            # download the mp3 audio file from the source
            urllib.request.urlretrieve(src, path_to_mp3)    
                
            # load downloaded mp3 audio file as .wav
            sound = pydub.AudioSegment.from_mp3(path_to_mp3)
            sound.export(path_to_wav, format="wav")
            sample_audio = sr.AudioFile(path_to_wav)
        
            # translate audio to text with google voice recognition
            r = sr.Recognizer()
            with sample_audio as source:
                audio = r.record(source)
            key = r.recognize_google(audio)
            print(f"[INFO] 第{row}筆的Recaptcha Passcode: {key}")
                
            # key in results and submit
            driver.find_element(By.ID, "audio-response").send_keys(key)
            driver.find_element(By.ID, "audio-response").send_keys(Keys.ENTER)
            driver.switch_to.default_content()
            
            # submit
            driver.find_element(By.XPATH, "//*[@id='threat-vault-app']/div[2]/div/form/div[1]/div[2]/button").click()    
        
            #如果第一個Category不是顯示"Malware"或"Phishing"，就截圖後繼續執行下一筆
            if driver.find_element(By.XPATH, "//*[@id='threat-vault-app']/div[2]/div/form/div[2]/ul/li[2]").text != "Category: Copyright Infringement" or "Category: Cryptocurrency" or "Category: Extremism" or "Category: Gambling" or "Category: Hacking" or "Category: Malware" or "Category: Phishing":
                driver.execute_script("window.scrollTo(0, 400)")
                goodURLscreenshot(driver, row)
        except Exception as e2:
            print(f"[WARNING] 在執行第{row}筆時，例外事件發生了，所以第{row}筆沒有成功！例外事件細節資訊：\n")
            print(e2)
        print(f"[WARNING] 在執行第{row}筆時，例外事件發生了，所以第{row}筆沒有成功！例外事件細節資訊：\n")
        print(e1)
        
def main(input_path):
    """
    至 Pala Alto 完成URL安全性驗證
    input_path: str, 輸入資料源的路徑 ( 含副檔名 )
    """
    # 讀取資料輸入源
    df_input = pd.read_csv('海外分行上網行為日報-2021-10-27.csv', encoding= 'unicode_escape',
                             converters={'URL/Filename':str})
    print(f'[INFO] 已成功讀取資料，筆數: {len(df_input)}，內容如以下：')
    print(df_input['URL/Filename']) 
    

    driver_chrome = set_environment_chrome()     
    driver_edge = set_environment_edge()
    
    driver = driver_chrome
    # 寫入所需資料
    for row in df_input.index:
        while(True):
            # URL這邊以後還是要拆個function比較好
            URL = 'https://urlfiltering.paloaltonetworks.com/'
            driver.get(URL)
            # wait
            delay2to4()
            
            checkBox_and_audioButton_statusCheck(driver)
            try:
                URLforVerify = df_input.loc[row, 'URL/Filename']
                print(f'[INFO] 現在做第{row}筆')
        
                action(driver, row, URLforVerify)
                # (for edge完執行成功，切回瀏覽器Chorme)
                driver = driver_chrome
                break
            except Exception:
                driver = driver_edge
        delay2to4() 
    print(f"[INFO] 總共{len(df_input)}筆，全部完成了！")
        

if __name__ == '__main__':
    input_path = '海外分行上網行為日報-2021-10-27.csv'
    main(input_path)