from operation.myacg_manager import MyAcg
from operation.database_operation import DataBase, DBreturnType

import customtkinter as ctk
from tkinter import messagebox
from PIL import Image
import pyglet
import sys
import os

from configparser import ConfigParser

from frames.setting_frame import frame_SaveSetting
from frames.account_frame import frame_Account
from frames.printorder_frame import frame_PrintOrder

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
        
if __name__ == "__main__":
    app = App()
    app.protocol("WM_DELETE_WINDOW", app.on_closing)
    
    app.mainloop()