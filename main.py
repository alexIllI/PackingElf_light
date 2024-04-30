from operation.myacg_manager import MyAcg, ReturnType, AccountReturnType
from operation.database_operation import DataBase, DBreturnType

import customtkinter as ctk
import tkinter
from tkinter import ttk
from tkinter import messagebox
from PIL import Image
import pyglet
import sys
import os
import datetime

from configparser import ConfigParser

def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)

class App(ctk.CTk):
    def __init__(self, *args, **kwargs):
        try:
            pyglet.font.add_file(resource_path("font\\Iansui-Regular.ttf"))
        except:
            print("add font error")
        
        #============================ VARIABLES ==================================
        self.current_account = "子午計畫"
        
        #====================== Operation =========================
        self.myacg_manager = MyAcg()
        self.database = DataBase('save')  
        while True:
            check_result = self.database.check_previous_records()
            if check_result == DBreturnType.SUCCESS:
                break
            elif check_result == DBreturnType.CHECK_PREVIOUS_RECORD_ERROR:
                print("excel file is opend, should be closed while checking previous records")
                messagebox.showwarning("匯出貨單錯誤", "發生意料之外的狀況, 請關閉應用程式並記錄Error log")
            elif check_result == DBreturnType.EXPORT_UNRECORDED_ERROR:
                print("excel unrecord data error")
                messagebox.showwarning("匯出貨單錯誤", "發生錯誤，檢測到有未紀錄的貨單，但無法匯出")
                break
            elif check_result != None:
                print("excel file is opend, should be closed while checking previous records")
                messagebox.showwarning("匯出貨單錯誤", f"Excel在開啟時無法匯出, 請關閉紀錄貨單的Excel: {check_result}.xlsx 後, 再按下確定")
        
        #====================== Config ===============================
        config = ConfigParser()
        config.read(resource_path("config.ini"))
        self.theme_color_dark = config["ThemeColor_dark"]["theme_color_dark"]
        self.theme_color = config["ThemeColor_dark"]["theme_color"]
        self.dark0_color = config["ThemeColor_dark"]["dark0"]
        self.dark1_color = config["ThemeColor_dark"]["dark1"]
        self.dark2_color = config["ThemeColor_dark"]["dark2"]
        self.dark3_color = config["ThemeColor_dark"]["dark3"]
        self.dark4_color = config["ThemeColor_dark"]["dark4"]
        self.dark5_color = config["ThemeColor_dark"]["dark5"]
        self.cancel_color = config["ThemeColor_dark"]["cancel"]
        self.close_color = config["ThemeColor_dark"]["close"]
        
        #========================= Root Frame ============================
        ctk.CTk.__init__(self, *args, **kwargs)
        self.geometry("1000x600")
        ctk.set_appearance_mode("dark")
        self.resizable(0,0)
        self.title("包貨小精靈")
        self.iconbitmap(resource_path("images\icon.ico"))
        
        # Create the root frame
        self.root_container = ctk.CTkFrame(self)
        self.root_container.pack(side="top", fill="both", expand=True)
        self.root_container.grid_rowconfigure(0, weight=1)
        self.root_container.grid_columnconfigure(0, weight=1)
        self.root_container.grid_columnconfigure(1, weight=4)
        
        #========================== Side Bar ===========================
        self.sidebar_frame = ctk.CTkFrame(master=self.root_container, fg_color=self.dark2_color, corner_radius=0)
        self.sidebar_frame.grid(row=0, column=0, sticky='nsew')
        
        logo_img_data = Image.open(resource_path("images\icon_meridian_white.png"))
        logo_img = ctk.CTkImage(dark_image=logo_img_data, light_image=logo_img_data, size=(135, 136.3))
        ctk.CTkLabel(master=self.sidebar_frame, text="", image=logo_img).pack(pady=(40, 0), anchor="center")
        
        package_img_data = Image.open(resource_path("images\printer.png"))
        package_img = ctk.CTkImage(dark_image=package_img_data, light_image=package_img_data)

        self.print_prder_btn = ctk.CTkButton(master=self.sidebar_frame, width=200, image=package_img, text="列印出貨單", fg_color=self.dark1_color, font=("Iansui", 18), 
                hover_color=self.dark3_color, anchor="center", command=lambda: self.show_frame("frame_PrintOrder"))
        self.print_prder_btn.pack(anchor="center", ipady=5, pady=(135, 0))
        
        #==============================================
        self.current_selected_btn = self.print_prder_btn
        
        list_img_data = Image.open(resource_path("images\list_icon.png"))
        list_img = ctk.CTkImage(dark_image=list_img_data, light_image=list_img_data)
        self.saveSetting_btn = ctk.CTkButton(master=self.sidebar_frame, width=200, image=list_img, text="儲存管理", fg_color="transparent", font=("Iansui", 18), 
                hover_color=self.dark3_color, anchor="center", command=lambda: self.show_frame("frame_SaveSetting"))
        self.saveSetting_btn.pack(anchor="center", ipady=5, pady=(16, 0))

        settings_img_data = Image.open(resource_path("images\settings_icon.png"))
        settings_img = ctk.CTkImage(dark_image=settings_img_data, light_image=settings_img_data)
        ctk.CTkButton(master=self.sidebar_frame, width=200, image=settings_img, text="設定", fg_color="transparent", font=("Iansui", 18), 
                hover_color=self.dark3_color, anchor="center").pack(anchor="center", ipady=5, pady=(16, 0),)
        
        settings_img_data = Image.open(resource_path("images\person_icon.png"))
        settings_img = ctk.CTkImage(dark_image=settings_img_data, light_image=settings_img_data)
        self.account_name_btn = ctk.CTkButton(master=self.sidebar_frame, width=200, image=settings_img, text="子午計畫", fg_color="transparent", font=("Iansui", 18), 
                hover_color=self.dark3_color, anchor="center", command=lambda: self.show_frame("frame_Account"))
        self.account_name_btn.pack(anchor="s", ipady=5, pady=(80, 0),)
        
        #========================== Other Pages ===========================
        self.frames = {}
        for F in (frame_Account, frame_PrintOrder, frame_SaveSetting):
            page_name = F.__name__
            frame = F(parent_frame=self.root_container, parent=self)
            self.frames[page_name] = frame
            frame.grid(row=0, column=1, sticky="nsew")

        self.show_frame("frame_PrintOrder")

    def show_frame(self, page_name):
        '''Show a frame for the given page name'''
        frame = self.frames[page_name]
        frame.tkraise()
        
        self.current_selected_btn.configure(fg_color="transparent")
        
        if page_name == "frame_PrintOrder":
            self.current_selected_btn = self.print_prder_btn
            self.print_prder_btn.configure(fg_color=self.dark1_color)
            
        elif page_name == "frame_Account":
            self.current_selected_btn = self.account_name_btn
            self.account_name_btn.configure(fg_color=self.dark1_color)
            
        elif page_name == "frame_SaveSetting":
            self.current_selected_btn = self.saveSetting_btn
            self.saveSetting_btn.configure(fg_color=self.dark1_color)
        
    def on_closing(self):
        if messagebox.askokcancel("退出包貨小精靈", "確定要退出? (會自動匯出本次所有貨單)"):
            self.myacg_manager.shut_down()
            while True:
                close_result = self.database.close_database()
                if close_result == DBreturnType.CLOSE_AND_SAVE_ERROR:
                    messagebox.showwarning("匯出貨單時發生錯誤", "匯出貨單時發生錯誤, 包貨紀錄將不會匯出至excel!\n(下次啟動應用程式時將會嘗試匯出至excel)")
                    break
                elif close_result == DBreturnType.SUCCESS:
                    break
                elif close_result != None:
                    messagebox.showwarning("匯出貨單時發生錯誤", f"匯出貨單時發生錯誤, 請先將儲存貨單的excel: {close_result} 關閉!")
            
            self.destroy()
            print("close app")
        
class frame_Account(ctk.CTkFrame):
    def __init__(self, parent_frame, parent):
        ctk.CTkFrame.__init__(self, parent_frame, fg_color=parent.dark0_color, corner_radius=0)
        self.account_manager = parent.myacg_manager
        self.current_account = parent.current_account
        self.account_list = self.account_manager.get_all_account_names()
        self.account_name_btn = parent.account_name_btn
        #=============================== TITLE ======================================

        title_frame = ctk.CTkFrame(master=self, fg_color="transparent")
        title_frame.pack(anchor="n", fill="x",  padx=27, pady=(29, 0))
        ctk.CTkLabel(master=title_frame, text="切換登入帳號", font=("Iansui", 32), text_color=parent.theme_color).pack(anchor="nw", side="left") 
        
        #=============================== CHOOSE ACCOUNT ======================================

        choose_account_container = ctk.CTkFrame(master=self, height=37.5, fg_color="transparent")
        choose_account_container.pack(fill="x", pady=(35, 0), padx=30)

        ctk.CTkLabel(master=choose_account_container, text="選擇帳號: ", text_color="#fff", font=("Iansui", 24)).pack(side="left", padx=(13, 0), pady=5)
        self.choose_account_combobox = ctk.CTkComboBox(master=choose_account_container, state="readonly", width=300, height = 40, font=("Iansui", 20), values=self.account_list, button_color=parent.theme_color, border_color=parent.theme_color, 
                    border_width=2, button_hover_color=parent.theme_color_dark, dropdown_hover_color=parent.theme_color_dark, dropdown_fg_color=parent.theme_color, dropdown_text_color=parent.dark0_color, command=self.change_current_account)
        self.choose_account_combobox.pack(side="left", padx=(13, 0), pady=15)
        self.choose_account_combobox.set(self.current_account)
            
        choose_account_btn_container = ctk.CTkFrame(master=self, height=37.5, fg_color="transparent")
        choose_account_btn_container.pack(fill="x", pady=(10, 0), padx=30)
        choose_account_btn_container.grid_rowconfigure(0, weight=1)
        choose_account_btn_container.grid_columnconfigure(0, weight=1)
        choose_account_btn_container.grid_columnconfigure(1, weight=1)
        ctk.CTkButton(master=choose_account_btn_container, width=100, height = 40, text="登入", font=("Iansui", 20), text_color=parent.dark0_color, fg_color=parent.theme_color, hover_color=parent.theme_color_dark, command=self.switch_account).grid(row=0, column=0, pady=(10,0), padx=10, sticky="e")
        ctk.CTkButton(master=choose_account_btn_container, width=100, height = 40, text="刪除", font=("Iansui", 20), text_color=parent.dark0_color, fg_color=parent.theme_color, hover_color=parent.theme_color_dark, command=self.delete_account).grid(row=0, column=1, pady=(10,0), padx=10, sticky="w")
        
        #=============================== CREATE ACCOUNT ======================================
        
        create_name_container = ctk.CTkFrame(master=self, height=37.5, fg_color="transparent")
        create_name_container.pack(fill="x", pady=(35, 0), padx=30)
        
        ctk.CTkLabel(master=create_name_container, text="名稱:      ", text_color="#fff", font=("Iansui", 24)).pack(side="left", padx=(13, 0), pady=5)
        self.create_name_entry = ctk.CTkEntry(master=create_name_container, width=300, height = 40, font=("Iansui", 20), placeholder_text="請輸入作為紀錄的名稱", border_color=parent.theme_color, border_width=2)
        self.create_name_entry.pack(side="left", padx=(13, 0), pady=5)
        
        create_account_container = ctk.CTkFrame(master=self, height=37.5, fg_color="transparent")
        create_account_container.pack(fill="x", pady=(10, 0), padx=30)
        
        ctk.CTkLabel(master=create_account_container, text="帳號:      ", text_color="#fff", font=("Iansui", 24)).pack(side="left", padx=(13, 0), pady=5)
        self.create_account_entry = ctk.CTkEntry(master=create_account_container, width=300, height = 40, font=("Iansui", 20), placeholder_text="請輸入買動漫帳號", border_color=parent.theme_color, border_width=2)
        self.create_account_entry.pack(side="left", padx=(13, 0), pady=5)
        
        create_password_container = ctk.CTkFrame(master=self, height=37.5, fg_color="transparent")
        create_password_container.pack(fill="x", pady=(10, 0), padx=30)
         
        ctk.CTkLabel(master=create_password_container, text="密碼:      ", text_color="#fff", font=("Iansui", 24)).pack(side="left", padx=(13, 0), pady=5)
        self.create_password_entry = ctk.CTkEntry(master=create_password_container, width=300, height = 40, font=("Iansui", 20), placeholder_text="請輸入密碼", border_color=parent.theme_color, border_width=2, show="*")
        self.create_password_entry.pack(side="left", padx=(13, 0), pady=5)
        
        create_password_check_container = ctk.CTkFrame(master=self, height=37.5, fg_color="transparent")
        create_password_check_container.pack(fill="x", pady=(10, 0), padx=30)
        
        ctk.CTkLabel(master=create_password_check_container, text="確認密碼: ", text_color="#fff", font=("Iansui", 24)).pack(side="left", padx=(13, 0), pady=5)
        self.check_password_entry = ctk.CTkEntry(master=create_password_check_container, width=300, height = 40, font=("Iansui", 20), placeholder_text="請確認密碼", border_color=parent.theme_color, border_width=2, show="*")
        self.check_password_entry.pack(side="left", padx=(13, 0), pady=5)
        self.check_password_entry.bind('<Return>', lambda event: self.create_new_account())
        
        create_button_container = ctk.CTkFrame(master=self, height=37.5, fg_color="transparent")
        create_button_container.pack(fill="x", pady=(10, 0), padx=30)
        create_button_container.grid_rowconfigure(0, weight=1)
        create_button_container.grid_columnconfigure(0, weight=1)
        create_button_container.grid_columnconfigure(1, weight=1)
        ctk.CTkButton(master=create_button_container, width=120, height = 40, text="創建新帳號", font=("Iansui", 20), text_color=parent.dark0_color, fg_color=parent.theme_color, hover_color=parent.theme_color_dark, command=self.create_new_account).grid(row=0, column=0, pady=(10,0), padx=10, sticky="e")
        ctk.CTkButton(master=create_button_container, width=120, height = 40, text="修改帳號", font=("Iansui", 20), text_color=parent.dark0_color, fg_color=parent.theme_color, hover_color=parent.theme_color_dark, command=self.modify_account_by_name).grid(row=0, column=1, pady=(10,0), padx=10, sticky="w")
    
    def change_current_account(self, cur_value):
        self.current_account = cur_value
            
    def create_new_account(self):
        try:
            self.name = self.create_name_entry.get()
        except:
            messagebox.showwarning("創建新帳號錯誤", "名稱欄位不得為空!")
            print("create new account with empty name")
            return       
        try:
            self.account = self.create_account_entry.get()
        except:
            messagebox.showwarning("創建新帳號錯誤", "帳號欄位不得為空!")
            print("create new account with empty name")
            return       
        try:
            self.password = self.create_password_entry.get()
        except:
            messagebox.showwarning("創建新帳號錯誤", "密碼欄位不得為空!")
            print("create new account with empty password")
            return        
        try:
            self.check_password = self.check_password_entry.get()
        except:
            messagebox.showwarning("創建新帳號錯誤", "確認密碼欄位不得為空!")
            print("create new account with empty check_password")
            return
        
        if self.password != self.check_password:
            messagebox.showwarning("創建新帳號錯誤", "確認密碼與密碼不相同!")
            self.check_password_entry.delete(0, 'end')
            print("wrong check password")
            return
        
        create_account_result = self.account_manager.create_account(self.name, self.account, self.password)
        
        if create_account_result == AccountReturnType.USERNAME_REPEAT:
            messagebox.showwarning("創建新帳號錯誤", "使用者名稱重複!")
            print("username repeated")
            return
        elif create_account_result == AccountReturnType.LOAD_AND_DECRYPT_ERROR or create_account_result == AccountReturnType.ADD_ACCOUNT_ERROR:
            messagebox.showwarning("創建新帳號錯誤", "後端存取帳號異常，請稍後再試!")
            print("something went wrong in account maneger")
            return
        elif create_account_result == AccountReturnType.SUCCESS:
            messagebox.showinfo("創建新帳號", f"成功創建新帳號: {self.name}!")
            self.account_list = self.account_manager.get_all_account_names()
            self.choose_account_combobox.configure(values=self.account_list)
            print("create new account success")
            self.create_name_entry.delete(0, 'end')
            self.create_account_entry.delete(0, 'end')
            self.create_password_entry.delete(0, 'end')
            self.check_password_entry.delete(0, 'end')
            
    def modify_account_by_name(self):
        try:
            self.name = self.create_name_entry.get()
        except:
            messagebox.showwarning("修改帳號錯誤", "必須輸入已存在帳號名稱!")
            print("create new account with empty name")
            return
        if self.name == "子午計畫":
            messagebox.showwarning("修改帳號錯誤", "子午計畫帳號是不能被修改的!")
            print("modify meridian is not allowed")
            return
        try:
            self.account = self.create_account_entry.get()
        except:
            messagebox.showwarning("修改帳號錯誤", "帳號欄位不得為空!")
            print("create new account with empty name")
            return       
        try:
            self.password = self.create_password_entry.get()
        except:
            messagebox.showwarning("修改帳號錯誤", "密碼欄位不得為空!")
            print("create new account with empty password")
            return        
        try:
            self.check_password = self.check_password_entry.get()
        except:
            messagebox.showwarning("修改帳號錯誤", "確認密碼欄位不得為空!")
            print("create new account with empty check_password")
            return
        
        if self.password != self.check_password:
            messagebox.showwarning("修改帳號錯誤", "確認密碼與密碼不相同!")
            self.check_password_entry.delete(0, 'end')
            print("wrong check password")
            return
        
        if messagebox.askyesno("修改帳號提示", f"確定要修改帳號: {self.name} ?"):
        
            modify_account_result = self.account_manager.modify_account(self.name, self.account, self.password)
            
            if modify_account_result == AccountReturnType.USERNAME_NOT_FOUND:
                messagebox.showwarning("修改帳號錯誤", "輸入的使用者名稱不存在!")
                print("username not found")
                return
            elif modify_account_result == AccountReturnType.LOAD_AND_DECRYPT_ERROR or modify_account_result == AccountReturnType.MODIFY_ACCOUNT_ERROR:
                messagebox.showwarning("創建新帳號錯誤", "後端存取帳號異常，請稍後再試!")
                print("something went wrong in account maneger")
                return
            elif modify_account_result == AccountReturnType.SUCCESS:
                print("modify account success")
                self.create_name_entry.delete(0, 'end')
                self.create_account_entry.delete(0, 'end')
                self.create_password_entry.delete(0, 'end')
                self.check_password_entry.delete(0, 'end')

    def delete_account(self):
        try:
            self.name = self.choose_account_combobox.get()
        except:
            messagebox.showwarning("帳號刪除選擇錯誤", "需要從下拉選單選擇一帳號刪除!")
            print("not choosing delete account name")
            return
        
        if self.name == "子午計畫":
            messagebox.showwarning("帳號刪除選擇錯誤", "子午計畫帳號是不能被刪除的!")
            print("delete meridian is not allowed")
            return
               
        if messagebox.askyesno("刪除帳號提示", f"確定要刪除: {self.name} ?"):
            delete_account_result = self.account_manager.delete_account(self.name)
            
            if delete_account_result == AccountReturnType.USERNAME_NOT_FOUND:
                messagebox.showwarning("刪除帳號錯誤", "找不到該使用者名稱的帳號!")
                print("username not found")
                return
            elif delete_account_result == AccountReturnType.LOAD_AND_DECRYPT_ERROR or delete_account_result == AccountReturnType.DELETE_ACCOUNT_ERROR:
                messagebox.showwarning("刪除帳號錯誤", "後端存取帳號異常，請稍後再試!")
                print("something went wrong in account maneger")
                return
            elif delete_account_result == AccountReturnType.SUCCESS:
                messagebox.showinfo("刪除帳號提示", f"成功刪除帳號: {self.name}!")
                self.account_list = self.account_manager.get_all_account_names()
                self.choose_account_combobox.configure(values=self.account_list)
                if self.current_account == self.name:
                    self.choose_account_combobox.set("")
                print("delete account success")
    
    def switch_account(self):        
        if messagebox.askyesno("選擇登入帳號提示", f"確定要改登入帳號: {self.current_account} ?"):
            switch_account_result = self.account_manager.switch_account(self.current_account)
            
            if switch_account_result == AccountReturnType.USERNAME_NOT_FOUND:
                messagebox.showwarning("選擇登入帳號錯誤", "找不到該帳號!")
                print("username not found")
                return
            elif switch_account_result == AccountReturnType.LOAD_AND_DECRYPT_ERROR:
                messagebox.showwarning("登入其他帳號錯誤", "後端存取帳號異常，請稍後再試!")
                print("username not found")
                return
            elif switch_account_result == AccountReturnType.GET_ACCOUNT_INFO_ERROR:
                messagebox.showwarning("登入其他帳號錯誤", f"在存取名稱: {self.current_account} 的帳號密碼時發生錯誤!")
                print("get account info error")
                return
            elif switch_account_result == AccountReturnType.GET_LOGIN_PAGE_ERROR:
                messagebox.showwarning("網頁自動化錯誤", "無法進入登入畫面, 請嘗試手動登出後再進入到登入畫面!")
                print("get login page error")
                return
            elif switch_account_result == AccountReturnType.LOGIN_ACCOUNT_ERROR or switch_account_result == AccountReturnType.LOGIN_PASSWORD_ERROR:
                messagebox.showwarning("網頁自動化錯誤", "找不到帳號密碼輸入框, 請嘗試手動輸入!")
                print("login entry element error")
                return
            elif switch_account_result == AccountReturnType.LOGIN_BTN_ERROR:
                messagebox.showwarning("網頁自動化錯誤", "找不到登入按鈕, 請手動點擊!")
                print("login btn error")
                return
            elif switch_account_result == AccountReturnType.WRONG_ACCOUNT_INFO:
                messagebox.showwarning("登入其他帳號錯誤", "輸入的帳號或密碼錯誤, 請嘗試修改名稱的該帳號密碼!")
                print("wrong account info")
                return
            elif switch_account_result == AccountReturnType.MY_STORE_NOT_FOUND:
                messagebox.showwarning("網頁自動化錯誤", "找不到'我的賣場'按鈕, 請手動點擊!")
                print("my store btn not found")
                return
            elif switch_account_result == AccountReturnType.SUCCESS:
                messagebox.showinfo("重新登入帳號提示", f"成功登入帳號: {self.current_account}!")
                self.account_name_btn.configure(text = self.current_account)
                print(f"login to {self.current_account} success")
                return

class frame_SaveSetting(ctk.CTkFrame):
    def __init__(self, parent_frame, parent):
        ctk.CTkFrame.__init__(self, parent_frame, fg_color=parent.dark0_color, corner_radius=0)
        
        excel_extention = '.xlsx'
        self.excel_list = [filename.rstrip(excel_extention) for filename in os.listdir("save") if filename.endswith(excel_extention)]
        self.saving_value = ctk.StringVar(value = "default")
        #=============================== TITLE ======================================

        title_frame = ctk.CTkFrame(master=self, fg_color="transparent")
        title_frame.pack(anchor="n", fill="x",  padx=27, pady=(29, 0))
        ctk.CTkLabel(master=title_frame, text="儲存管理", font=("Iansui", 32), text_color=parent.theme_color).pack(anchor="nw", side="left")
        
        #=============================== DEFAULT SAVE PATH ======================================
        default_excel_container = ctk.CTkFrame(master=self, height=37.5, fg_color="transparent")
        default_excel_container.pack(fill="x", pady=(50, 0), padx=30)

        ctk.CTkLabel(master=default_excel_container, text=f"預設儲存的Excel: 000", text_color="#fff", font=("Iansui", 26)).pack(fill="x", pady=5)
        
        #=============================== CREATE NEW EXCEL ======================================
        create_excel_container = ctk.CTkFrame(master=self, height=37.5, fg_color="transparent")
        create_excel_container.pack(fill="x", pady=(50, 0), padx=30)

        self.newname_checkbox = ctk.CTkCheckBox(create_excel_container, text=" 儲存至新建Excel: ", text_color="#fff", font=("Iansui", 24), command=lambda:self.checkbox_event(self.saving_value.get()), variable=self.saving_value, onvalue="newName_on", offvalue="newName_off")
        self.newname_checkbox.pack(side="left")
        self.create_excel_entry = ctk.CTkEntry(master=create_excel_container, width=300, height = 40, font=("Iansui", 20), placeholder_text="請輸入檔案名稱", border_color=parent.theme_color, border_width=2)
        self.create_excel_entry.pack(side="left", padx=(13, 0), pady=5)
        
        #=============================== APPEND EXCEL ======================================
        append_excel_container = ctk.CTkFrame(master=self, height=37.5, fg_color="transparent")
        append_excel_container.pack(fill="x", pady=(15, 0), padx=30)
        
        self.append_checkbox = ctk.CTkCheckBox(append_excel_container, text=" 附加至已存在的Excel: ", text_color="#fff", font=("Iansui", 24), command=lambda:self.checkbox_event(self.saving_value.get()), variable=self.saving_value, onvalue="append_on", offvalue="append_off")
        self.append_checkbox.pack(side="left")
        
        self.append_excel_combobox = ctk.CTkComboBox(master=append_excel_container, state="readonly", width=320, height = 40, font=("Iansui", 20), values=self.excel_list, button_color=parent.theme_color, border_color=parent.theme_color, 
                    border_width=2, button_hover_color=parent.theme_color_dark, dropdown_hover_color=parent.theme_color_dark, dropdown_fg_color=parent.theme_color, dropdown_text_color=parent.dark0_color)
        self.append_excel_combobox.pack(side="left", padx=(13, 0), pady=15)
        if len(self.excel_list) > 0:
            self.append_excel_combobox.set((self.excel_list)[0])
        else:
            self.append_excel_combobox.set("沒有已存在的Excel")
        
        #=============================== REPLACE EXCEL ======================================
        replace_excel_container = ctk.CTkFrame(master=self, height=37.5, fg_color="transparent")
        replace_excel_container.pack(fill="x", pady=(15, 0), padx=30)
        
        self.replace_checkbox = ctk.CTkCheckBox(replace_excel_container, text=" 覆蓋已存在的Excel: ", text_color="#fff", font=("Iansui", 24), command=lambda:self.checkbox_event(self.saving_value.get()), variable=self.saving_value, onvalue="replace_on", offvalue="replace_off")
        self.replace_checkbox.pack(side="left")
        
        self.replace_excel_combobox = ctk.CTkComboBox(master=replace_excel_container, state="readonly", width=320, height = 40, font=("Iansui", 20), values=self.excel_list, button_color=parent.theme_color, border_color=parent.theme_color, 
                    border_width=2, button_hover_color=parent.theme_color_dark, dropdown_hover_color=parent.theme_color_dark, dropdown_fg_color=parent.theme_color, dropdown_text_color=parent.dark0_color)
        self.replace_excel_combobox.pack(side="left", padx=(13, 0), pady=15)
        if len(self.excel_list) > 0:
            self.replace_excel_combobox.set((self.excel_list)[0])
        else:
            self.replace_excel_combobox.set("沒有已存在的Excel")
            
        #=============================== BUTTON SECTION ======================================
        save_excel_btn_container = ctk.CTkFrame(master=self, height=37.5, fg_color="transparent")
        save_excel_btn_container.pack(fill="x", pady=(35, 0), padx=30)
        save_excel_btn_container.grid_rowconfigure(0, weight=1)
        save_excel_btn_container.grid_columnconfigure(0, weight=1)
        save_excel_btn_container.grid_columnconfigure(1, weight=1)
        self.confirm_btn = ctk.CTkButton(master=save_excel_btn_container, width=75, height = 40, text="確認", font=("Iansui", 18), text_color=parent.dark0_color, 
                                          fg_color=parent.theme_color, hover_color=parent.theme_color_dark, command=self.confirm_setting)
        self.confirm_btn.grid(row=0, column=0, pady=(10,0), padx=10, sticky="e")
        self.reset_btn = ctk.CTkButton(master=save_excel_btn_container, width=75, height = 40, text="恢復預設", font=("Iansui", 18), text_color=parent.dark0_color, 
                                          fg_color=parent.theme_color, hover_color=parent.theme_color_dark, command=self.reset_setting)
        self.reset_btn.grid(row=0, column=1, pady=(10,0), padx=10, sticky="w")
        
    # turn other checkbox into DISABLE    
    def checkbox_event(self, checkbox_value):
        print(f"checkbox toggled: {checkbox_value}")
        if checkbox_value in ["newName_on", "append_on", "replace_on"]:
            for C in (self.newname_checkbox, self.append_checkbox, self.replace_checkbox, self.create_excel_entry, self.append_excel_combobox, self.replace_excel_combobox):
                if checkbox_value == "newName_on" and (C == self.newname_checkbox or C == self.create_excel_entry):
                    continue
                elif checkbox_value == "append_on" and (C == self.append_checkbox or C == self.append_excel_combobox):
                    continue
                elif checkbox_value == "replace_on" and (C == self.replace_checkbox or C == self.replace_excel_combobox):
                    continue
                C.configure(state=tkinter.DISABLED)
        else:
            for C in (self.newname_checkbox, self.append_checkbox, self.replace_checkbox, self.create_excel_entry, self.append_excel_combobox, self.replace_excel_combobox):
                C.configure(state=tkinter.NORMAL)
            self.saving_value.set(value="default")
    
    def confirm_setting(self):
        command = self.saving_value.get()
        if command == "default":
            print("save setting: default")
        elif command == "newName_on":
            new_excel_name = self.create_excel_entry.get()
            if not new_excel_name:
                messagebox.showwarning("名稱錯誤", "請輸入欲創建的Excel名稱後, 再按下確認")
                print("empty new excel name entry")
                return
            
            print(f"save setting: save to new excel '{new_excel_name}'")
        elif command == "append_on":
            append_excel_name = self.append_excel_combobox.get()
            print(f"saving setting: append to excel '{append_excel_name}'")
        elif command == "replace_on":
            replace_excel_name = self.replace_excel_combobox.get()
            print(f"saving setting: replace excel '{replace_excel_name}'")
    
    def reset_setting(self):
        for C in (self.newname_checkbox, self.append_checkbox, self.replace_checkbox):
            C.deselect()
        for C in (self.newname_checkbox, self.append_checkbox, self.replace_checkbox, self.create_excel_entry, self.append_excel_combobox, self.replace_excel_combobox):
            C.configure(state=tkinter.NORMAL)
            
        self.create_excel_entry.delete(0, 'end')
        if len(self.excel_list) > 0:
            self.replace_excel_combobox.set((self.excel_list)[0])
            self.append_excel_combobox.set((self.excel_list)[0])
        else:
            self.replace_excel_combobox.set("沒有已存在的Excel")
            self.append_excel_combobox.set("沒有已存在的Excel")
        self.saving_value.set(value="default")
        print("save setting: default")
                   
class frame_PrintOrder(ctk.CTkFrame):
    def __init__(self, parent_frame, parent):
        ctk.CTkFrame.__init__(self, parent_frame, fg_color=parent.dark0_color, corner_radius=0)    
        #=============================== VARIABLES ======================================
        self.myacg_manager = parent.myacg_manager
        self.database = parent.database
        
        self.cancel_color = parent.cancel_color
        self.close_color = parent.close_color
        self.last_order = ""
        self.table_list = self.database.get_output_excel_options()

        #=============================== TITLE ======================================

        title_frame = ctk.CTkFrame(master=self, fg_color="transparent")
        title_frame.pack(anchor="n", fill="x",  padx=27, pady=(29, 0))
        ctk.CTkLabel(master=title_frame, text="列印出貨單", font=("Iansui", 32), text_color=parent.theme_color).pack(anchor="nw", side="left")

        #=============================== SELECT TABLE ======================================

        store_table_container = ctk.CTkFrame(master=self, fg_color="transparent")
        store_table_container.pack(fill="x", pady=(10, 0), padx=30)

        ctk.CTkLabel(master=store_table_container, text="選擇紀錄位置: ", text_color="#fff", font=("Iansui", 24)).pack(side="left", padx=(13, 0), pady=5)
        self.store_table_combobox = ctk.CTkComboBox(master=store_table_container, state="readonly", width=320, height = 40, font=("Iansui", 20), values=self.table_list, button_color=parent.theme_color, border_color=parent.theme_color, 
                    border_width=2, button_hover_color=parent.theme_color_dark, dropdown_hover_color=parent.theme_color_dark, 
                    dropdown_fg_color=parent.theme_color, dropdown_text_color=parent.dark0_color, command= lambda value: self.update_table(self.view_status_enrty.get(), value))
        self.store_table_combobox.pack(side="left", padx=(13, 0), pady=15)
        self.store_table_combobox.set(self.table_list[0])

        #=============================== INVOICE SCANNER ======================================

        invoice_container = ctk.CTkFrame(master=self, fg_color="transparent")
        invoice_container.pack(fill="x", pady=(10, 0), padx=30)

        ctk.CTkLabel(master=invoice_container, text="發票編號: ", text_color="#fff", font=("Iansui", 24)).pack(side="left", padx=(13, 0), pady=5)
        
        self.vcmd = (self.register(self.on_validate), '%P')
        self.invoice_entry = ctk.CTkEntry(master=invoice_container, validate="key", validatecommand=self.vcmd, width=225, height = 40, font=("Iansui", 20), border_color=parent.theme_color, border_width=2)
        self.invoice_entry.pack(side="left", padx=(60, 0), pady=5)
        
        # hotkey binding
        self.invoice_entry.bind('<Return>', lambda event: self.switch_to_order_entry()) 

        #=============================== PRINTER ORDER ======================================

        prnit_order_container = ctk.CTkFrame(master=self, height=37.5, fg_color="transparent")
        prnit_order_container.pack(fill="x", pady=(0, 0), padx=30)

        ctk.CTkLabel(master=prnit_order_container, text="PG", text_color="#fff", font=("Iansui", 24)).pack(side="left", padx=(13, 0), pady=5)
        self.order_combobox = ctk.CTkComboBox(master=prnit_order_container, state="readonly", width=105, height = 40, font=("Iansui", 20), values=["018", "019", "020", "021"], button_color=parent.theme_color, border_color=parent.theme_color, 
                    border_width=2, button_hover_color=parent.theme_color_dark, dropdown_hover_color=parent.theme_color_dark, dropdown_fg_color=parent.theme_color, dropdown_text_color=parent.dark0_color)
        self.order_combobox.pack(side="left", padx=(13, 0), pady=15)
        self.order_combobox.set("019")
        self.order_entry = ctk.CTkEntry(master=prnit_order_container, width=225, height = 40, font=("Iansui", 20), placeholder_text="請輸入貨單後五碼", border_color=parent.theme_color, border_width=2)
        self.order_entry.pack(side="left", padx=(23, 0), pady=5)
        
        # hotkey binding
        self.order_entry.bind('<Return>', lambda event: self.printToprinter())

        ctk.CTkButton(master=prnit_order_container, width=75, height = 40, text="列印", font=("Iansui", 18), text_color=parent.dark0_color, 
                                          fg_color=parent.theme_color, hover_color=parent.theme_color_dark, command=self.printToprinter).pack(anchor="ne", padx=(70, 0), pady=15, side="left")
        
        ctk.CTkButton(master=prnit_order_container, width=75, height = 40, text="重印上一單", font=("Iansui", 18), text_color=parent.dark0_color, 
                                          fg_color=parent.theme_color, hover_color=parent.theme_color_dark, command=self.printLastOrder).pack(anchor="ne", padx=(10, 40), pady=15, side="left")
        
        #=============================== ORDER COUNT ======================================

        counting_container = ctk.CTkFrame(master=self, height=60, fg_color="transparent")
        counting_container.pack(fill="x", pady=(30, 0), padx=40)

        total_count_metric = ctk.CTkFrame(master=counting_container, fg_color=parent.dark1_color,width=300, height=37.5)
        total_count_metric.pack(side="left")
        
        success_count_metric = ctk.CTkFrame(master=counting_container, fg_color=parent.dark1_color,width=300, height=37.5)
        success_count_metric.pack(side="right")

        self.label_total_order_number = ctk.CTkLabel(master=total_count_metric, text="目前貨單總數: 0", text_color="#fff", font=("Iansui", 22))
        self.label_total_order_number.pack(side="left", padx=20, pady=5)
        
        self.label_success_order_number = ctk.CTkLabel(master=success_count_metric, text="成功列印貨單總數: 0", text_color="#fff", font=("Iansui", 22))
        self.label_success_order_number.pack(side="left", padx=20, pady=5)
        
        #=============================== SEARCH BAR ======================================

        search_container = ctk.CTkFrame(master=self, height=37.5, fg_color="transparent")
        search_container.pack(fill="x", pady=(20, 0), padx=30)

        self.search_entry = ctk.CTkEntry(master=search_container, width=225, height = 30, font=("Iansui", 16), placeholder_text="搜尋貨單", border_color=parent.theme_color, border_width=2)
        self.search_entry.pack(side="left", padx=(13, 0), pady=5)
        self.search_entry.bind('<Return>', lambda event: self.search_order())
        
        ctk.CTkButton(master=search_container, width=75, height = 30, text="搜尋", font=("Iansui", 16), text_color=parent.dark0_color, fg_color=parent.theme_color, hover_color=parent.theme_color_dark, command=self.search_order).pack(anchor="ne", padx=(13, 0), pady=5, side="left")
        ctk.CTkButton(master=search_container, width=75, height = 30, text="刪除", font=("Iansui", 16), text_color=parent.dark0_color, fg_color=parent.theme_color, hover_color=parent.theme_color_dark, command=self.btn_delete_items).pack(anchor="ne", padx=(13, 0), pady=5, side="left")
        
        self.show_recorded_checkbox = ctk.CTkCheckBox(search_container, text="顯示過去紀錄", text_color="#fff", font=("Iansui", 16),
                                                      command=lambda: self.update_table(self.view_status_enrty.get(), self.store_table_combobox.get()))
        self.show_recorded_checkbox.pack(side="right",padx=15, pady=5)
        
        self.view_status_enrty = ctk.CTkComboBox(master=search_container, state="readonly", width=105, height = 30, font=("Iansui", 16), values=["顯示全部", "成功出貨", "關轉", "取消"], button_color=parent.theme_color, border_color=parent.theme_color, 
                    border_width=2, button_hover_color=parent.theme_color_dark, dropdown_hover_color=parent.theme_color_dark, dropdown_fg_color=parent.theme_color, dropdown_text_color=parent.dark0_color, command= lambda value: self.update_table(value, self.store_table_combobox.get()))
        self.view_status_enrty.pack(side="right", padx=(13, 0), pady=5)
        self.view_status_enrty.set("顯示全部")
        
        #============================ TABLE ============================             

        table_frame = ctk.CTkFrame(master=self, fg_color="transparent")
        table_frame.pack(expand=True, fill="both", padx=27, pady=10)
        
        tree_scroll = ctk.CTkScrollbar(table_frame, button_color=parent.theme_color, button_hover_color=parent.theme_color_dark)
        tree_scroll.pack(side="right", fill="y")
        
        printed_order_table_style = ttk.Style()
        printed_order_table_style.theme_use("clam")
        printed_order_table_style.configure("Treeview.Heading", font=("Iansui", 16))
        printed_order_table_style.configure("Treeview", rowheight = 50, font=("Iansui", 14), background="#fff")
        printed_order_table_style.layout("Treeview", [('Treeview.treearea', {'sticky': 'nswe'})])
        
        self.printed_order_table = ttk.Treeview(table_frame, columns = ('id', 'time', 'order', 'status', 'invoice', 'output_excel'), style="Treeview", show = 'headings', yscrollcommand=tree_scroll.set)
        self.printed_order_table.column("# 1", anchor="center",width=30)
        self.printed_order_table.column("# 2", anchor="center",width=145)
        self.printed_order_table.column("# 3", anchor="center")
        self.printed_order_table.column("# 4", anchor="center", width=120)
        self.printed_order_table.column("# 5", anchor="center")
        self.printed_order_table.column("# 6", anchor="center")
        self.printed_order_table.heading('id', text = 'id')
        self.printed_order_table.heading('time', text = '時間')
        self.printed_order_table.heading('order', text = '貨單編號')
        self.printed_order_table.heading('status', text = '狀態')
        self.printed_order_table.heading('invoice', text = '發票號碼')
        self.printed_order_table.heading('output_excel', text = '儲存位置')
        self.printed_order_table.pack(fill = 'both', expand = True)
        
        self.printed_order_table.tag_configure('cancel', background=parent.cancel_color)
        self.printed_order_table.tag_configure('close', background=parent.close_color)
        
        tree_scroll.configure(command=self.printed_order_table.yview)
        
        # events
        def item_select(_):
            print(self.printed_order_table.selection())
            for i in self.printed_order_table.selection():
                print(self.printed_order_table.item(i)['values'])

        def delete_items(_):
            print('delete pressed')
            if len(self.printed_order_table.selection()) > 1:
                messagebox.showwarning("刪除貨單警告", "一次只能刪除一筆貨單，請只選擇一筆刪除!")
                return
            
            if messagebox.askokcancel("刪除貨單", f"確定要刪除 {self.printed_order_table.item(self.printed_order_table.selection())['values'][2]} ? (刪除後不可復原)"):
                print(f"delete {self.printed_order_table.item(self.printed_order_table.selection())['values'][2]}")
                self.database.delete_data(self.printed_order_table.item(self.printed_order_table.selection())['values'][2])
                self.update_table(self.view_status_enrty.get(), self.store_table_combobox.get())
                
        def deselect_items(_):
            self.printed_order_table.selection_remove(*self.printed_order_table.selection())
            print("deselect all items")

        self.printed_order_table.bind('<<TreeviewSelect>>', item_select)
        self.printed_order_table.bind('<Delete>', delete_items)
        self.printed_order_table.bind('<Escape>', deselect_items)
    
    def update_table(self, status, table_name):
        self.printed_order_table.delete(*self.printed_order_table.get_children())
        record = False
        show_recorded = self.show_recorded_checkbox.get()
        if show_recorded:
            record = True
            
        if status == "顯示全部":
            datas = self.database.fetch_table_data("all", record, table_name)
        elif status == "成功出貨":
            datas = self.database.fetch_table_data('success', record, table_name)
        elif status == "關轉":
            datas = self.database.fetch_table_data('close', record, table_name)
        elif status == "取消":
            datas = self.database.fetch_table_data('cancel', record, table_name)
                
        for data in datas:
            if data[3] == 'success':
                self.printed_order_table.insert(parent = '', index = 0, values = (data[0], data[1], data[2], "成功", data[4], data[5]))
            if data[3] == 'close':
                self.printed_order_table.insert(parent = '', index = 0, values = (data[0], data[1], data[2], "關轉", data[4], data[5]), tags = ("close",))
            if data[3] == 'cancel':
                self.printed_order_table.insert(parent = '', index = 0, values = (data[0], data[1], data[2], "取消", data[4], data[5]), tags = ("cancel",))
        
        order_numbers = self.database.count_records(self.store_table_combobox.get(), record) 
        self.label_total_order_number.configure(text = f"目前貨單總數: {order_numbers[0]}")
        self.label_success_order_number.configure(text = f"成功列印貨單總數: {order_numbers[1]}")
                
    def search_order(self):
        if not self.search_entry.get():
            messagebox.showwarning("搜尋貨單結果", "老哥，先輸入再搜尋好嗎")
            print("empty search entry")
            return
        
        if len(self.search_entry.get()) != 10:
            messagebox.showwarning("搜尋貨單結果", "請輸入完整貨單!")
            print("search for wrong order length")
            self.order_entry.delete(0, 'end')
            return
        
        order_number = self.search_entry.get()
        self.search_entry.delete(0, 'end')
        result = self.database.search_order(order_number)
        
        if result == DBreturnType.ORDER_NOT_FOUND:
            messagebox.showwarning("搜尋貨單結果", "搜尋的貨單不存在!")
            print(f"search for {order_number} result doesn't exist")
        elif result == DBreturnType.SEARCH_ORDER_ERROR:
            messagebox.showwarning("刪除貨單錯誤", "發生意料外的狀況, 請紀錄並回報")
        else:
            try:
                self.printed_order_table.selection_remove(self.printed_order_table.selection())
                for child in self.printed_order_table.get_children():
                    if order_number in self.printed_order_table.item(child)['values'][2]:
                        print(f"find {self.printed_order_table.item(child)['values'][2]} within treeview, id: {result[0]}")
                        self.printed_order_table.selection_set(child)
            except:
                print(f"current table do not contain children: {order_number}")
                        
            if result[6] == "unrecorded":
                messagebox.showinfo("搜尋貨單結果", f"列印時間: {result[1].strftime('%m/%d %H:%M:%S')}\nID: {result[0]}\n狀態: {result[3]}\n發票編號: {result[4]}\n儲存位置: {result[5]}\n紀錄狀態: 尚未儲存至Excel")
                return
            else:
                messagebox.showinfo("搜尋貨單結果", f"列印時間: {result[1].strftime('%m/%d %H:%M:%S')}\nID: {result[0]}\n狀態: {result[3]}\n發票編號: {result[4]}\n儲存位置: {result[5]}\n紀錄狀態: 已儲存至Excel")
                return
        
    def btn_delete_items(self):
        #if search entry is empty
        if not self.search_entry.get():
            #if trere was selected row in treeview
            if self.printed_order_table.selection():
                print('delete selected row using btn')
                if len(self.printed_order_table.selection()) > 1:
                    messagebox.showwarning("刪除貨單警告", "一次只能刪除一筆貨單，請只選擇一筆刪除!")
                    return
                
                if messagebox.askokcancel("刪除貨單", f"確定要刪除 {self.printed_order_table.item(self.printed_order_table.selection())['values'][2]} ? (刪除後不可復原)"):
                    print(f"delete {self.printed_order_table.item(self.printed_order_table.selection())['values'][2]}")
                    self.database.delete_data(self.printed_order_table.item(self.printed_order_table.selection())['values'][2])
                    self.update_table(self.view_status_enrty.get(), self.store_table_combobox.get())
                    return
            else:
                messagebox.showwarning("刪除貨單結果", "老哥，先輸入再刪除好嗎")
                print("empty search entry")
                return
        
        if len(self.search_entry.get()) != 10:
            messagebox.showwarning("刪除貨單結果", "請輸入完整貨單!")
            print("search for wrong order length")
            self.order_entry.delete(0, 'end')
            return
        
        # delete order from search entry
        if messagebox.askokcancel("刪除貨單", f"確定要刪除 {self.search_entry.get()} ? (刪除後不可復原)"):
            print(f"delete {self.search_entry.get()} using btn")
            result = self.database.search_order(self.search_entry.get())
            if result == DBreturnType.ORDER_NOT_FOUND:
                messagebox.showwarning("刪除貨單結果", "欲刪除的貨單不存在!")
                print("the order trying to delete doesn't exist")
            elif result == DBreturnType.SEARCH_ORDER_ERROR:
                messagebox.showwarning("刪除貨單錯誤", "發生意料外的狀況, 請紀錄並回報")
            else:
                self.database.delete_data(self.search_entry.get())
                self.search_entry.delete(0, 'end')
                self.update_table(self.view_status_enrty.get(), self.store_table_combobox.get())
    
    # Validate if the input is english and number or not        
    def on_validate(self, P):
        """Validate new content if it's purely ASCII."""
        if P.isascii():
            return True
        else:
            messagebox.showwarning("輸入法錯誤", "喇咪說這樣不能印，請將輸入法切換成英文後重新掃描一次條碼")
            self.after(1, lambda: self.clear_entry(self.invoice_entry))  # Schedule the clear_entry to avoid issues with direct deletion within validation
            return False
        
    def clear_entry(self, target_entry):
        """Clears the content of the entry widget."""
        target_entry.delete(0, 'end')
            
    def switch_to_order_entry(self):
        invoice_text = self.invoice_entry.get()  # Get the scan result of the entry
        invoice = invoice_text[5:15]     # Extract the invoice number of the scan result
        self.invoice_entry.delete(0, 'end')  # Clear the current entry
        self.invoice_entry.insert(0, invoice)
        self.order_entry.focus_set()
                        
    def printLastOrder(self):
        def print_cancel_close(_status:str, order, invoice):
            self.database.insert_data(order, _status, invoice, self.store_table_combobox.get())
            self.update_table(self.view_status_enrty.get(), self.store_table_combobox.get())
        
        def print_success(order, invoice):
            self.database.insert_data(order, 'success', invoice, self.store_table_combobox.get())
            self.update_table(self.view_status_enrty.get(), self.store_table_combobox.get())
            
        if self.last_order == "":
            messagebox.showwarning("無法重印上一單", "找不到上一單的資訊!")
            print("no last order")
            return
        
        self.database.delete_data(self.last_order)        
        result = self.myacg_manager.printer(self.last_order)
        
        if result == ReturnType.MULTIPLE_TAB:
            messagebox.showwarning("網頁自動化錯誤", "請開啟瀏覽器將買動漫以外的分頁關閉!")
            print("meltiple tab detected")
            
        elif result == ReturnType.POPUP_UNSOLVED:
            messagebox.showwarning("網頁自動化錯誤", "你可能沒有按'我知道了',打開買動漫,按下去")
            print("popup unsolved")
        
        elif result == ReturnType.ALREADY_FINISH:
            messagebox.showwarning("列印出貨單錯誤", "該訂單已出貨或已完成!")
            print("already finished")
            
        elif result == ReturnType.ORDER_CANCELED:
            messagebox.showwarning("取消", "此單已被取消!請至買動漫確認)")
            print_cancel_close("cancel", self.last_order, self.last_invoice)
            print("order canceled")
        
        elif result == ReturnType.ORDER_NOT_FOUND:
            messagebox.showwarning("貨單後號碼錯誤", "沒有這一單!")
            print("order not found")
            
        elif result == ReturnType.ORDER_NOT_FOUND_ERROR:
            print("order not found raises error")
            
        elif result == ReturnType.CHECKBOX_NOT_FOUND:
            messagebox.showwarning("網頁自動化錯誤", "找不到checkbox,可能沒有這一單,自己開買動漫看一下")
            print("check box not found")
            
        elif result == ReturnType.CLICKING_CHECKBOX_ERROR:
            messagebox.showwarning("網頁自動化錯誤", "無法勾選checkbox,可能沒有這一單,自己開買動漫看一下")
            print("can't click checkbox")
            
        elif result == ReturnType.CLICKING_PRINT_ORDER_ERROR:
            messagebox.showwarning("網頁自動化錯誤", "無法點選列印出貨單,可能沒有這一單,自己開買動漫看一下")
            print("can't click print order")
            
        elif result == ReturnType.STORE_CLOSED:
            messagebox.showwarning("寄送商店關轉", "該筆貨單寄送商店關轉中!")
            print_cancel_close("close", self.last_order, self.last_invoice)
            print("store closed")
            
        elif result == ReturnType.SWITCH_TAB_ERROR:
            messagebox.showwarning("網頁自動化錯誤", "無法切換視窗(可能是關轉，去看買動漫，如果有我知道了按下去，不然我會當掉)")
            print("switch tab error")
            
        elif result == ReturnType.LOAD_HTML_BODY_ERROR:
            print("loading html body error")
            
        elif result == ReturnType.EXCUTE_PRINT_ERROR:
            messagebox.showwarning("列印失敗", "列印失敗，不會記錄貨單!請再列印一次")
            print("excute print error")
            
        elif result == ReturnType.CLOSED_TAB_ERROR:
            messagebox.showwarning("網頁自動化錯誤", "成功列印出貨單，已經記錄在文件中。但關閉分頁時發生異常，請手動關閉'出貨單'分頁!!!!")
            print_success(self.last_order, self.last_invoice)
            print("closed tab error")
            
        elif result == ReturnType.SUCCESS:
            print_success(self.last_order, self.last_invoice)
            print("Success!")
        
        else:
            print("Enum error!")
            return
        
    def printToprinter(self):
        validate_string = self.invoice_entry.get()
        print(validate_string)
        if not validate_string.isascii():
            self.invoice_entry.delete(0, 'end')
            self.invoice_entry.focus_set()
            messagebox.showerror("輸入法錯誤", "喇咪說這樣不能印，請將輸入法切換成英文後重新掃描一次條碼")
            return
        
        def print_cancel_close(_status:str, order, invoice):
            self.database.insert_data(order, _status, invoice, self.store_table_combobox.get())
            self.update_table(self.view_status_enrty.get(), self.store_table_combobox.get())
        
        def print_success(order, invoice):
            self.database.insert_data(order, 'success', invoice, self.store_table_combobox.get())
            self.update_table(self.view_status_enrty.get(), self.store_table_combobox.get())
        
        if not self.order_combobox.get():
            messagebox.showwarning("沒有選擇輸入前綴", "請選擇貨單PG後數字!")
            print("empty combobox")
            return
        
        if len(self.order_entry.get()) != 5:
            messagebox.showwarning("貨單後號碼錯誤", "請輸入貨單'後五碼'!")
            print("wrong order length")
            self.order_entry.delete(0, 'end')
            return
            
        if not self.order_entry.get().isdigit():
            messagebox.showwarning("貨單後號碼錯誤", "請只輸入數字!")
            print("wrong order number type")
            self.order_entry.delete(0, 'end')
            return
        
        if len(self.invoice_entry.get()) < 9:
            messagebox.showwarning("發票號碼錯誤", "喇咪說你很笨ㄟ 怎麼會忘記輸入發票")
            print("wrong invoice number")
            self.invoice_entry.focus_set()
            return
        
        current_order = f"PG{self.order_combobox.get()}{self.order_entry.get()}"
        current_invoice = self.invoice_entry.get()
        self.order_entry.delete(0, 'end')
        self.invoice_entry.delete(0, 'end')
        self.last_order = current_order
        self.last_invoice = current_invoice
        self.invoice_entry.focus_set()
        
        #check if it is repeated order
        check_repeat_result =  self.database.check_repeated(current_order)
        if check_repeat_result == DBreturnType.SEARCH_ORDER_ERROR:
            messagebox.showwarning("fatal error", "發生意料外的狀況, 請紀錄並回報")
            print("printToprinter -> check repeated error")
            return
        elif check_repeat_result[6] == "unrecorded":
            messagebox.showwarning("貨單後號碼錯誤", "出貨單重複!如果想重印這一單, 請先將下方同樣貨單編號的紀錄刪除")
            print(f"repeat unrecorded order: {current_order}")
            return
        elif check_repeat_result[6] == "recorded":
            result_delete_previous_record = messagebox.askretrycancel("貨單後號碼錯誤", f"出貨單重複!在 {check_repeat_result[1].strftime('%m-%d %H:%M:%S')} 時曾經列印過該出貨單, 並記錄在Excel: {check_repeat_result[5]} 中。若希望刪除資料庫中該紀錄(需要刪除才能列印, 但不會刪除在Excel的紀錄), 請選擇'重試', 若希望繼續, 請按'取消'")
            if result_delete_previous_record:
                self.database.delete_data(current_order)
            else:
                print(f"repeat recorded order: {current_order}")
                return
            
        #print to printer using autoweb
        result = self.myacg_manager.printer(current_order)
        
        if result == ReturnType.MULTIPLE_TAB:
            messagebox.showwarning("網頁自動化錯誤", "請開啟瀏覽器將買動漫以外的分頁關閉!")
            print("meltiple tab detected")
            
        elif result == ReturnType.POPUP_UNSOLVED:
            messagebox.showwarning("網頁自動化錯誤", "你可能沒有按'我知道了',打開買動漫,按下去")
            print("popup unsolved")
        
        elif result == ReturnType.ALREADY_FINISH:
            messagebox.showwarning("列印出貨單錯誤", "該訂單已出貨或已完成!")
            print("already finished")
            
        elif result == ReturnType.ORDER_CANCELED:
            messagebox.showwarning("取消", "此單已被取消!請至買動漫確認)")
            print_cancel_close("cancel", current_order, current_invoice)
            print("order canceled")
        
        elif result == ReturnType.ORDER_NOT_FOUND:
            messagebox.showwarning("貨單後號碼錯誤", "沒有這一單!")
            print("order not found")
            
        elif result == ReturnType.ORDER_NOT_FOUND_ERROR:
            print("order not found raises error")
            
        elif result == ReturnType.CHECKBOX_NOT_FOUND:
            messagebox.showwarning("網頁自動化錯誤", "找不到checkbox,可能沒有這一單,自己開買動漫看一下")
            print("check box not found")
            
        elif result == ReturnType.CLICKING_CHECKBOX_ERROR:
            messagebox.showwarning("網頁自動化錯誤", "無法勾選checkbox,可能沒有這一單,自己開買動漫看一下")
            print("can't click checkbox")
            
        elif result == ReturnType.CLICKING_PRINT_ORDER_ERROR:
            messagebox.showwarning("網頁自動化錯誤", "無法點選列印出貨單,可能沒有這一單,自己開買動漫看一下")
            print("can't click print order")
            
        elif result == ReturnType.STORE_CLOSED:
            messagebox.showwarning("寄送商店關轉", "該筆貨單寄送商店關轉中!")
            print_cancel_close("close", current_order, current_invoice)
            print("store closed")
            
        elif result == ReturnType.SWITCH_TAB_ERROR:
            messagebox.showwarning("網頁自動化錯誤", "無法切換視窗(可能是關轉，去看買動漫，如果有我知道了按下去，不然我會當掉)")
            print("switch tab error")
            
        elif result == ReturnType.LOAD_HTML_BODY_ERROR:
            print("loading html body error")
            
        elif result == ReturnType.EXCUTE_PRINT_ERROR:
            messagebox.showwarning("列印失敗", "列印失敗，不會記錄貨單!請再列印一次")
            print("excute print error")
            
        elif result == ReturnType.CLOSED_TAB_ERROR:
            messagebox.showwarning("網頁自動化錯誤", "成功列印出貨單，已經記錄在文件中。但關閉分頁時發生異常，請手動關閉'出貨單'分頁!!!!")
            print_success(current_order, current_invoice)
            print("closed tab error")
            
        elif result == ReturnType.SUCCESS:
            print_success(current_order, current_invoice)
            print("Success!")
        
        else:
            print("Enum error!")
            return
        
if __name__ == "__main__":
    app = App()
    app.protocol("WM_DELETE_WINDOW", app.on_closing)
    
    app.mainloop()