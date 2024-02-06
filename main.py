from operation.myacg_manager import MyAcg, ReturnType, AccountReturnType
from operation.database_operation import DataBase, DBreturnType

import customtkinter as ctk
from tkinter import ttk
from tkinter import messagebox
from PIL import Image
import pyglet
import sys
import os

from datetime import datetime, timedelta
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
        self.total_order = 0
        self.success_order = 0
        self.current_id = 0
        self.current_time_name = datetime.now().strftime('Printed_Order_%Y_%m_%d')
        self.current_account = "子午計畫"
        
        #====================== Operation =========================
        self.myacg_manager = MyAcg()
        self.database = DataBase(self.current_time_name, 'operation\\data.db', 'save')
        while True:
            check_result = self.database.check_previous_records(self.current_time_name, (datetime.now()-timedelta(days=1)).strftime('Printed_Order_%Y_%m_%d'))
            if check_result == DBreturnType.SUCCESS:
                break
            elif check_result == DBreturnType.PERMISSION_ERROR:
                print("excel file is opend, should be closed while checking previous records")
                messagebox.showwarning("匯出貨單錯誤", f"Excel在開啟時無法匯出, 請關閉紀錄貨單的Excel: {self.current_time_name}.xlsx 後, 再按下確定")
            elif check_result == DBreturnType.EXPORT_UNRECORDED_ERROR:
                print("excel unrecord data error")
                messagebox.showwarning("匯出貨單錯誤", "發生錯誤，檢測到有未紀錄的貨單，但無法匯出")
                break
        
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
        self.geometry("900x600")
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
        ctk.CTkButton(master=self.sidebar_frame, width=200, image=list_img, text="儲存管理", fg_color="transparent", font=("Iansui", 18), 
                hover_color=self.dark3_color, anchor="center").pack(anchor="center", ipady=5, pady=(16, 0))

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
        for F in (frame_Account, frame_PrintOrder):
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
        
    def on_closing(self):
        if messagebox.askokcancel("退出包貨小精靈", "確定要退出? (會自動匯出本次所有貨單)"):
            self.myacg_manager.shut_down()
            while True:
                close_result = self.database.close_database()
                if close_result == DBreturnType.CLOSE_AND_SAVE_ERROR:
                    messagebox.showwarning("匯出貨單時發生錯誤", "匯出貨單時發生錯誤, 包貨紀錄將不會匯出至excel!\n(下次啟動應用程式時將會嘗試匯出至excel)")
                    break
                elif close_result == DBreturnType.PERMISSION_ERROR:
                    messagebox.showwarning("匯出貨單時發生錯誤", "匯出貨單時發生錯誤, 請先將儲存貨單的excel關閉!")
                elif close_result == DBreturnType.SUCCESS:
                    break
            
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
        self.create_password_entry = ctk.CTkEntry(master=create_password_container, width=300, height = 40, font=("Iansui", 20), placeholder_text="請輸入密碼", border_color=parent.theme_color, border_width=2)
        self.create_password_entry.pack(side="left", padx=(13, 0), pady=5)
        
        create_password_check_container = ctk.CTkFrame(master=self, height=37.5, fg_color="transparent")
        create_password_check_container.pack(fill="x", pady=(10, 0), padx=30)
        
        ctk.CTkLabel(master=create_password_check_container, text="確認密碼: ", text_color="#fff", font=("Iansui", 24)).pack(side="left", padx=(13, 0), pady=5)
        self.check_password_entry = ctk.CTkEntry(master=create_password_check_container, width=300, height = 40, font=("Iansui", 20), placeholder_text="請確認密碼", border_color=parent.theme_color, border_width=2)
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
            
class frame_PrintOrder(ctk.CTkFrame):
    def __init__(self, parent_frame, parent):
        ctk.CTkFrame.__init__(self, parent_frame, fg_color=parent.dark0_color, corner_radius=0)    
        #=============================== VARIABLES ======================================
        self.myacg_manager = parent.myacg_manager
        self.database = parent.database
        
        self.total_order = parent.total_order
        self.success_order = parent.success_order
        self.current_id = parent.current_id
        self.cur_time_name = parent.current_time_name
        self.cancel_color = parent.cancel_color
        self.close_color = parent.close_color
        
        #=============================== TITLE ======================================

        title_frame = ctk.CTkFrame(master=self, fg_color="transparent")
        title_frame.pack(anchor="n", fill="x",  padx=27, pady=(29, 0))
        ctk.CTkLabel(master=title_frame, text="列印出貨單", font=("Iansui", 32), text_color=parent.theme_color).pack(anchor="nw", side="left")
        
        #=============================== STORAGE PATH ======================================

        storage_path_container = ctk.CTkFrame(master=self, height=37.5, fg_color="transparent")
        storage_path_container.pack(fill="x", pady=(35, 0), padx=30)

        ctk.CTkLabel(master=storage_path_container, text="儲存位置: ", text_color="#fff", font=("Iansui", 24)).pack(side="left", padx=(13, 0), pady=5)
        self.save_path_combobox = ctk.CTkComboBox(master=storage_path_container, state="readonly", width=200, height = 40, font=("Iansui", 20), values=["(現在沒功能)"], button_color=parent.theme_color, border_color=parent.theme_color, 
                    border_width=2, button_hover_color=parent.theme_color_dark, dropdown_hover_color=parent.theme_color_dark, dropdown_fg_color=parent.theme_color, dropdown_text_color=parent.dark0_color)
        self.save_path_combobox.pack(side="left", padx=(13, 0), pady=15)
        self.save_path_combobox.set("(現在沒功能)")

        #=============================== PRINTER ORDER ======================================

        prnit_order_container = ctk.CTkFrame(master=self, height=37.5, fg_color="transparent")
        prnit_order_container.pack(fill="x", pady=(10, 0), padx=30)

        ctk.CTkLabel(master=prnit_order_container, text="PG", text_color="#fff", font=("Iansui", 24)).pack(side="left", padx=(13, 0), pady=5)
        self.order_combobox = ctk.CTkComboBox(master=prnit_order_container, state="readonly", width=105, height = 40, font=("Iansui", 20), values=["018", "019", "020", "021"], button_color=parent.theme_color, border_color=parent.theme_color, 
                    border_width=2, button_hover_color=parent.theme_color_dark, dropdown_hover_color=parent.theme_color_dark, dropdown_fg_color=parent.theme_color, dropdown_text_color=parent.dark0_color)
        self.order_combobox.pack(side="left", padx=(13, 0), pady=15)
        self.order_combobox.set("019")
        self.order_entry = ctk.CTkEntry(master=prnit_order_container, width=225, height = 40, font=("Iansui", 20), placeholder_text="請輸入貨單後五碼", border_color=parent.theme_color, border_width=2)
        self.order_entry.pack(side="left", padx=(13, 0), pady=5)
        
        # hotkey binding
        self.order_entry.bind('<Return>', lambda event: self.printToprinter())

        ctk.CTkButton(master=prnit_order_container, width=75, height = 40, text="列印", font=("Iansui", 18), text_color=parent.dark0_color, 
                                          fg_color=parent.theme_color, hover_color=parent.theme_color_dark, command=self.printToprinter).pack(anchor="ne", padx=(40, 0), pady=15, side="left")
        
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
        self.view_status_enrty = ctk.CTkComboBox(master=search_container, state="readonly", width=105, height = 30, font=("Iansui", 16), values=["顯示全部", "成功出貨", "關轉", "取消"], button_color=parent.theme_color, border_color=parent.theme_color, 
                    border_width=2, button_hover_color=parent.theme_color_dark, dropdown_hover_color=parent.theme_color_dark, dropdown_fg_color=parent.theme_color, dropdown_text_color=parent.dark0_color, command=self.update)
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
        
        self.printed_order_table = ttk.Treeview(table_frame, columns = ('id', 'time', 'order', 'status', 'save_status'), style="Treeview", show = 'headings', yscrollcommand=tree_scroll.set)
        self.printed_order_table.column("# 1", anchor="center",width=35)
        self.printed_order_table.column("# 2", anchor="center",width=150)
        self.printed_order_table.column("# 3", anchor="center")
        self.printed_order_table.column("# 4", anchor="center", width=120)
        self.printed_order_table.column("# 5", anchor="center")
        self.printed_order_table.heading('id', text = 'id')
        self.printed_order_table.heading('time', text = '時間')
        self.printed_order_table.heading('order', text = '貨單編號')
        self.printed_order_table.heading('status', text = '狀態')
        self.printed_order_table.heading('save_status', text = '儲存位置')
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
                if self.printed_order_table.item(self.printed_order_table.selection())['values'][3] == "成功":
                    self.success_order -= 1
                self.total_order -= 1
                self.update(self.view_status_enrty.get())
                
        def deselect_items(_):
            self.printed_order_table.selection_remove(*self.printed_order_table.selection())
            print("deselect all items")

        self.printed_order_table.bind('<<TreeviewSelect>>', item_select)
        self.printed_order_table.bind('<Delete>', delete_items)
        self.printed_order_table.bind('<Escape>', deselect_items)
    
    def update(self, status):
        self.printed_order_table.delete(*self.printed_order_table.get_children())
        if status == "顯示全部":
            datas = self.database.fetch_all_unrecorded("all")
        elif status == "成功出貨":
            datas = self.database.fetch_all_unrecorded('success')
        elif status == "關轉":
            datas = self.database.fetch_all_unrecorded('close')
        elif status == "取消":
            datas = self.database.fetch_all_unrecorded('cancel')
        
        for data in datas:
            if data[3] == 'success':
                self.printed_order_table.insert(parent = '', index = 0, values = (data[0], data[1], data[2], "成功", data[4]))
            if data[3] == 'close':
                self.printed_order_table.insert(parent = '', index = 0, values = (data[0], data[1], data[2], "關轉", data[4]), tags = ("close",))
            if data[3] == 'cancel':
                self.printed_order_table.insert(parent = '', index = 0, values = (data[0], data[1], data[2], "取消", data[4]), tags = ("cancel",))
                
        self.label_total_order_number.configure(text = f"目前貨單總數: {self.total_order}")
        self.label_success_order_number.configure(text = f"成功列印貨單總數: {self.success_order}")
                
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
        else:
            if result[5] == "unrecorded":
                self.printed_order_table.selection_remove(self.printed_order_table.selection())
                for child in self.printed_order_table.get_children():
                    if order_number in self.printed_order_table.item(child)['values'][2]:
                        print(f"find {self.printed_order_table.item(child)['values'][2]} within treeview, id: {result[0]}")
                        self.printed_order_table.selection_set(child)
                        messagebox.showinfo("搜尋貨單結果", f"貨單編號: {order_number}\nID: {result[0]}\n狀態: {result[3]}\n儲存位置: {result[4]}")
                        return
            else:
                messagebox.showwarning("搜尋貨單結果", "資料庫中有紀錄這筆貨單，但不是在本次應用程式執行後紀錄，因此不會顯示在下方表格!")
                print(f"search for {order_number} already exist, but is recorded")
        
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
                    if self.printed_order_table.item(self.printed_order_table.selection())['values'][3] == "成功":
                        self.success_order -= 1
                    self.database.delete_data(self.printed_order_table.item(self.printed_order_table.selection())['values'][2])
                    self.total_order -= 1
                    self.update(self.view_status_enrty.get())
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
            else:
                self.database.delete_data(self.search_entry.get())
                self.search_entry.delete(0, 'end')
                if result[3] == 'success':
                    self.success_order -= 1
                self.total_order -= 1
                self.update(self.view_status_enrty.get())
        
    def printToprinter(self):
        def print_cancel_close(_status:str, order):
            self.total_order += 1
            self.current_id += 1
            self.label_total_order_number.configure(text = f"目前貨單總數: {self.total_order}")
            cur_time = datetime.now().strftime('%H:%M:%S')
            self.database.insert_data(self.current_id, cur_time, order, _status, self.cur_time_name)
            if _status == "close":
                self.printed_order_table.insert(parent = '', index = 0, values = (self.current_id, cur_time, order, "關轉", self.cur_time_name), tags = (_status,))
            else:
                self.printed_order_table.insert(parent = '', index = 0, values = (self.current_id, cur_time, order, "取消", self.cur_time_name), tags = (_status,))
            self.update(self.view_status_enrty.get())
        
        def print_success(order):
            self.total_order += 1
            self.success_order += 1
            self.current_id += 1
            self.label_total_order_number.configure(text = f"目前貨單總數: {self.total_order}")
            self.label_success_order_number.configure(text = f"成功列印貨單總數: {self.success_order}")
            cur_time = datetime.now().strftime('%H:%M:%S')
            self.database.insert_data(self.current_id, cur_time, order, 'success', self.cur_time_name)
            self.printed_order_table.insert(parent = '', index = 0, values = (self.current_id, cur_time, order, '成功', self.cur_time_name))
            self.update(self.view_status_enrty.get())
        
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
        
        current_order = f"PG{self.order_combobox.get()}{self.order_entry.get()}"
        self.order_entry.delete(0, 'end')
        
        #check if it is repeated order
        check_repeat_result =  self.database.search_order(current_order)
        if check_repeat_result == DBreturnType.ORDER_NOT_FOUND:
            pass
        elif check_repeat_result[5] == "unrecorded":
            messagebox.showwarning("貨單後號碼錯誤", "出貨單重複!如果想重印這一單, 請先將下方同樣貨單編號的紀錄刪除")
            print(f"repeat unrecorded order: {current_order}")
            return
        elif check_repeat_result[5] == "recorded":
            result_delete_previous_record = messagebox.askretrycancel("貨單後號碼錯誤", f"出貨單重複!在 {check_repeat_result[1]} 時曾經列印過該出貨單, 並記錄在Excel: {check_repeat_result[4]} 中。若希望刪除資料庫中該紀錄(需要刪除才能列印, 但不會刪除在Excel的紀錄), 請選擇'重試', 若希望繼續, 請按'取消'")
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
            print_cancel_close("cancel", current_order)
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
            print_cancel_close("close", current_order)
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
            print_success(current_order)
            print("closed tab error")
            
        elif result == ReturnType.SUCCESS:
            print_success(current_order)
            print("Success!")
        
        else:
            print("Enum error!")
            return
        
if __name__ == "__main__":
    app = App()
    app.protocol("WM_DELETE_WINDOW", app.on_closing)
    
    app.mainloop()