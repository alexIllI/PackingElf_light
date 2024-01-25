from operation.account_manage import EncryptedAccountManager

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait 
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service

import os
import sys
from subprocess import CREATE_NO_WINDOW
from configparser import ConfigParser
from enum import Enum

def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)

#====================== Enum ===============================
class ReturnType(Enum):
    MULTIPLE_TAB = "MULTIPLE_TAB"
    POPUP_UNSOLVED = "POPUP_UNSOLVED"
    ALREADY_FINISH = "ALREADY_FINISH"
    ORDER_NOT_FOUND = "ORDER_NOT_FOUND"
    ORDER_NOT_FOUND_ERROR = "ORDER_NOT_FOUND_ERROR"
    CHECKBOX_NOT_FOUND = "CHECKBOX_NOT_FOUND"
    CLICKING_CHECKBOX_ERROR = "CLICKING_CHECKBOX_ERROR"
    CLICKING_PRINT_ORDER_ERROR = "CLICKING_PRINT_ORDER_ERROR"
    ORDER_CANCELED = "ORDER_CANCELED"
    STORE_CLOSED = "STORE_CLOSED"
    SWITCH_TAB_ERROR = "SWITCH_TAB_ERROR"
    LOAD_HTML_BODY_ERROR = "LOAD_HTML_BODY_ERROR"
    EXCUTE_PRINT_ERROR = "EXCUTE_PRINT_ERROR"
    CLOSED_TAB_ERROR = "CLOSED_TAB_ERROR"
    SUCCESS = "SUCCESS"

#====================== Config ===============================
config = ConfigParser()
config.read(resource_path("config.ini"))

#====================== Path ===============================
OUTTER_PATH = os.path.abspath(os.getcwd())
URL = "https://www.myacg.com.tw/login.php?done=http%3A%2F%2Fwww.myacg.com.tw%2Findex.php"

#======================== Chrome Crawler Setting ===============================
options = webdriver.ChromeOptions()
options.add_experimental_option('excludeSwitches', ['enable-logging'])
options.add_experimental_option("excludeSwitches", ["enable-automation"])
options.add_experimental_option('useAutomationExtension', False)
options.add_experimental_option("prefs", {"profile.password_manager_enabled": False, "credentials_enable_service": False})
options.add_experimental_option("detach", True)
options.add_argument('--enable-print-browser')
options.add_argument('--kiosk-printing')
service = Service()
service.creation_flags = CREATE_NO_WINDOW


class MyAcg():
    def __init__(self):
        
        #============== Variables ==============
        self.last = ""
        
        #============== Account ==============
        try:
            info = EncryptedAccountManager()
            info.load_and_decrypt()
            
            # if there are multiple accounts, change
            account_info = info.get_account_by_name("meridian")
            account = account_info["account"]
            password = account_info["password"]
        except:
            print("decrypt info ERROR!!")
            return 
        
        #============== Login ==============
        try:
            self.driver = webdriver.Chrome(service=service,options=options)
            self.driver.get(URL)
        except:
            print("error occured when creating webdriver")
            return 
        
        #login account
        try:
            account_element = WebDriverWait(self.driver, config["WebOperation"]["waittime"]).until(
                EC.presence_of_element_located((By.NAME, "account")))
            account_element.clear()
            account_element.send_keys(account)
        except:
            print("can't find 'login account' element, or connection timed out")
            return 
        
        #login password
        try:
            password_element = WebDriverWait(self.driver, config["WebOperation"]["waittime"]).until(
                EC.presence_of_element_located((By.NAME, "password")))
            password_element.clear()
            password_element.send_keys(password)
        except:
            print("can't find 'login password' element, or connection timed out")
            return 
        
        #login button
        try:
            Login_btn = WebDriverWait(self.driver, config["WebOperation"]["waittime"]).until(
                EC.presence_of_element_located((By.XPATH, '//*[@id="form1"]/div/div/div[2]/div[5]/div[1]/a')))
            Login_btn.click()
        except:
            print("can't find 'login button' element, or connection timed out")
            return 

        #find 我的賣場 element and click
        try:
            locate_store = (By.XPATH, '//*[@id="topbar"]/div/ul/li[1]/a')
            Store = WebDriverWait(self.driver, config["WebOperation"]["longerwaittime"]).until(
                EC.presence_of_element_located(locate_store),
                "Can't find my store button")
        except:
            print("我的賣場按鈕連線超時")
            return False

        Store.click()

    #find search bar and search
    def printer(self, order):
        
        if len(self.driver.window_handles) > 1:
            return ReturnType.MULTIPLE_TAB
        
        #check if last one is closed, the popup window had been handled
        try:
            search_bar = self.driver.find_element(By.NAME, 'o_num') #search bar element
            search_bar.clear()
            search_bar.send_keys(order)
            search = self.driver.find_element(By.XPATH, '//*[@id="search_goods"]/div[4]/ul/li[2]/a') #search button element
            search.click()
        except:
            return ReturnType.POPUP_UNSOLVED
        
        #check if the order exist
        try:
            no_order_wait = self.driver.find_element(By.XPATH, '//*[@id="wrap"]/div[2]/div/div[2]/div/span[2]/a')
            if no_order_wait:
                return ReturnType.ORDER_NOT_FOUND
        except:
            pass
        
        #check if it's canceled
        try:
            self.driver.find_element(By.XPATH, '//*[@id="wrap"]/div[2]/div[2]/div[1]/table/tbody/tr[1]/td[1]/div[1]/div/span[2]')
            return ReturnType.ORDER_CANCELED
        except:
            pass
        
        #check if the order already finished
        try:
            #locate print order button as indicator of whether it's finished
            self.driver.find_element(By.XPATH, '//*[@id="wrap"]/div[2]/div[2]/div[1]/table/tbody/tr/td[7]/a[4]')
        except:
            return ReturnType.ALREADY_FINISH
        
        #check if there is closed tag
        try:
            self.driver.find_element(By.XPATH, '//*[@id="wrap"]/div[2]/div[2]/div[1]/table/tbody/tr[1]/td[7]/span')
            return ReturnType.STORE_CLOSED
        except:
            pass
        
        #======================================= TEST RETURN ====================================================
        # try:
        #     no_order = self.driver.find_element(By.XPATH, '//*[@id="wrap"]/div[2]/div/div[2]/div/span[1]')
        #     no_order_text = no_order.text
        #     if no_order_text == "您沒有訂單，趕快到買動漫逛逛吧！":
        #         return ReturnType.ORDER_CANCELED
        # except:
        #     pass
        
        # return ReturnType.SUCCESS
        #=================================================================================================
        
        #等待直到check box出現並勾選
        try:
            checkbox = WebDriverWait(self.driver, config["WebOperation"]["waittime"]).until(
                EC.presence_of_element_located((By.ID, "oid_check_" + order[3:])))
        except:
            return ReturnType.CHECKBOX_NOT_FOUND

        # use Javascript to click checkbox
        try:
            self.driver.execute_script("arguments[0].click();", checkbox)  
        except:
            return ReturnType.CLICKING_CHECKBOX_ERROR
        
        #click print order
        try:
            print_order = self.driver.find_element(By.ID, 'PrintBatch')
            print_order.click()
        except:
            return ReturnType.CLICKING_PRINT_ORDER_ERROR

        #測試是否有開啟新分頁
        try:
            self.driver.switch_to.window(self.driver.window_handles[1])
        except:
            #測試是否可以handle pop up
            try:
                alert = self.driver.switch_to.alert
                print(f"popup alert showing message: {alert.text}")
                alert.accept()
                # self.driver.execute_script("switchTo().alert().dismiss();")
                return ReturnType.STORE_CLOSED
            except:
                pass
            return ReturnType.SWITCH_TAB_ERROR
        
        #列印出貨單(找出出貨單元素)
        try:
            wait = WebDriverWait(self.driver, config["WebOperation"]["longerwaittime"])
            wait.until(EC.presence_of_element_located((By.TAG_NAME, "body")))
        except:
            return ReturnType.LOAD_HTML_BODY_ERROR
        
        #excute printing
        try:
            self.driver.execute_script('window.print();')
            print("成功列印")
        except:
            return ReturnType.EXCUTE_PRINT_ERROR
        
        #close opend tab
        try:
            self.driver.close()
            self.driver.switch_to.window(self.driver.window_handles[0])
        except:
            return ReturnType.CLOSED_TAB_ERROR
        
        return ReturnType.SUCCESS
        
    def shut_down(self):
        self.driver.quit()
        print("close webdriver")