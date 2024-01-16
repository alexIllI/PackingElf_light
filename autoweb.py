import time
import os
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait 
from selenium.webdriver.support import expected_conditions as EC
import account_manage as Acc
from selenium.webdriver.chrome.service import Service
from subprocess import CREATE_NO_WINDOW

OUTTER_PATH = os.path.abspath(os.getcwd())
URL = "https://www.myacg.com.tw/login.php?done=http%3A%2F%2Fwww.myacg.com.tw%2Findex.php"

#Chrome Crawler Setting
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
        
        #============== Account ==============
        info = Acc.EncryptedAccountManager()
        info.load_and_decrypt()
        
        # if there are multiple accounts, change
        account_info = info.get_account_by_name("meridian")
        account = account_info["account"]
        password = account_info["password"]
        
        #============== Login ==============
        self.driver = webdriver.Chrome(service=service,options=options)
        self.driver.get(URL)
        account_element = self.driver.find_element(by = By.NAME, value="account")
        password_element = self.driver.find_element(by = By.NAME, value="password")
        account_element.clear()
        password_element.clear()
        account_element.send_keys(account)
        password_element.send_keys(password)

        Login_btn = self.driver.find_element(by = By.XPATH, value = '//*[@id="form1"]/div/div/div[2]/div[5]/div[1]/a')
        Login_btn.click()
        time.sleep(5)

        #============== Variables ==============
        self.last = ""

        #find 我的賣場 element and click

        try:
            locate_store = (By.XPATH, '//*[@id="topbar"]/div/ul/li[1]/a')
            Store = WebDriverWait(self.driver, 10).until(
                EC.presence_of_element_located(locate_store),
                "Can't find my store button")
        except:
            print("我的賣場按鈕連線超時")
            return False

        Store.click()

    #find search bar and search
    def printer(self, order):
        with open('成功列印的出貨單.txt', 'r') as file:
            for line in file:
                line_text = line.strip()
                if line_text == order:
                    print("這單可能重複了喔~你再想想")
                    return False

        try:
            search_bar = self.driver.find_element(By.NAME, 'o_num')
            search_bar.clear()
            search_bar.send_keys(order)
            
            search = self.driver.find_element(By.XPATH, '//*[@id="search_goods"]/div[4]/ul/li[2]/a')
            search.click()
        except:
            print("你可能沒有按'我知道了',打開買動漫,按下去")
            return False
        
        #等待直到check box出現並勾選
        # self.doc = Ex.Excel()
        try:
            checkbox = WebDriverWait(self.driver, 8).until(
                EC.presence_of_element_located((By.ID, "oid_check_" + order[3:])))
        except:
            print("找不到checkbox,可能沒有這一單,自己開買動漫看一下")
            return False

        # use Javascript to click checkbox
        self.driver.execute_script("arguments[0].click();", checkbox)
        
        print_order = self.driver.find_element(By.ID, 'PrintBatch')
        print_order.click()

        #測試是否有開啟新分頁
        try:
            self.driver.switch_to.window(self.driver.window_handles[1])
        except:
            try:
                cancel = self.driver.find_element(By.XPATH, '//*[@id="wrap"]/div[2]/div[2]/div[1]/table/tbody/tr[1]/td[1]/div[1]/div/span[2]')
                element_text = cancel.text
                if element_text == "取消原因":
                    print("無法切換視窗(可能是取消，去看買動漫)")
                else:
                    print("無法切換視窗(可能是關轉，去看買動漫，如果有我知道了按下去，不然我會當掉)")
            except:
                print("無法切換視窗(可能是關轉，去看買動漫，如果有我知道了按下去，不然我會當掉)")

            return False
        
        #列印出貨單(找出出貨單元素)
        try:
            wait = WebDriverWait(self.driver, 10)
            wait.until(EC.presence_of_element_located((By.TAG_NAME, "body")))
        except:
            print("列印發生錯誤!")
            return False
        
        #列印
        self.driver.execute_script('window.print();')
        
        print("成功列印")
        
        self.driver.close()
        self.driver.switch_to.window(self.driver.window_handles[0])

        self.last = order
        return True

    def save(self, order):
        try:
            with open('成功列印的出貨單.txt', 'a') as file:
                file.write(order + '\n')
                print(f"成功寫入貨單: {order}")
        except:
            print("寫入貨單時發生錯誤，檢查一下筆記本裡有沒有存到這筆貨單編號")
        return
        
    def shut_down(self):
        self.driver.quit()
        print("close webdriver")
        
if __name__ == "__main__":
    web = MyAcg()
    while True:
        user_input = input("\n輸入PG0後的貨單號碼,例如:1900723,輸入'stop'來中止程式,再印一次的話要先刪掉舊的: ")
        
        if user_input.lower() == 'stop':
            web.shut_down()
            break
        
        if len(user_input) == 7 and user_input.isdigit():
            if web.printer("PG0" + user_input):
                web.save("PG0" + user_input)
        else:
            print("請輸入七位數字")

        print(f"剛剛輸入的貨單為: PG0{user_input}")