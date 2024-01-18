from autoweb import MyAcg

from tkinter import messagebox
import customtkinter as ctk
from CTkTable import CTkTable
from PIL import Image, ImageTk

from configparser import ConfigParser

class App(ctk.CTk):
    def __init__(self):
        super().__init__()
        
        #====================== Chromedriver =========================
        # self.driver = MyAcg()
        # while True:
        #     user_input = input("\n輸入PG0後的貨單號碼,例如:1900723,輸入'stop'來中止程式,再印一次的話要先刪掉舊的: ")
        
        #====================== Config ===============================
        config = ConfigParser()
        config.read("config.ini")
        self.theme_color_dark = config["ThemeColor"]["theme_color_dark"]
        self.theme_color = config["ThemeColor"]["theme_color"]
        self.dark0_color = config["ThemeColor"]["dark0"]
        self.dark1_color = config["ThemeColor"]["dark1"]
        self.dark2_color = config["ThemeColor"]["dark2"]
        self.dark3_color = config["ThemeColor"]["dark3"]
        self.dark4_color = config["ThemeColor"]["dark4"]
        self.dark5_color = config["ThemeColor"]["dark5"]
        
        self.geometry("1200x800")
        self.resizable(0,0)
        self.title("包貨小精靈")
        self.iconbitmap("images\icon.ico")
        
        ctk.set_appearance_mode("dark")
        
        #========================== Side Bar ===========================
        self.sidebar = SideBar(self)
        self.Page_printorder = PrintOrder(self)
        
    def on_closing(self):
        if messagebox.askokcancel("退出包貨小精靈", "確定要退出? (會自動匯出未登記的貨單)"):
            # self.driver.shut_down()
            print("close app")
            self.destroy()
        
class SideBar(ctk.CTkFrame):
    def __init__(self, parent):
        super().__init__(parent)
        self.sidebar_frame = ctk.CTkFrame(master=parent, fg_color=parent.dark2_color,  width=300, height=800, corner_radius=0)
        self.sidebar_frame.pack_propagate(0)
        self.sidebar_frame.pack(fill="y", anchor="w", side="left")
        
        logo_img_data = Image.open("images\icon_meridian_white.png")
        logo_img = ctk.CTkImage(dark_image=logo_img_data, light_image=logo_img_data, size=(180, 186))
        ctk.CTkLabel(master=self.sidebar_frame, text="", image=logo_img).pack(pady=(60, 0), anchor="center")
        
        package_img_data = Image.open("images\printer.png")
        package_img = ctk.CTkImage(dark_image=package_img_data, light_image=package_img_data)

        ctk.CTkButton(master=self.sidebar_frame, width=250, image=package_img, text="列印出貨單", fg_color=parent.dark1_color, font=("Iansui", 24), 
                hover_color=parent.dark3_color, anchor="n").pack(anchor="center", ipady=5, pady=(180, 0))

        list_img_data = Image.open("images\list_icon.png")
        list_img = ctk.CTkImage(dark_image=list_img_data, light_image=list_img_data)
        ctk.CTkButton(master=self.sidebar_frame, width=250, image=list_img, text="儲存管理", fg_color="transparent", font=("Iansui", 24), 
                hover_color=parent.dark3_color, anchor="n").pack(anchor="center", ipady=5, pady=(16, 0))

        settings_img_data = Image.open("images\settings_icon.png")
        settings_img = ctk.CTkImage(dark_image=settings_img_data, light_image=settings_img_data)
        ctk.CTkButton(master=self.sidebar_frame, width=250, image=settings_img, text="設定", fg_color="transparent", font=("Iansui", 24), 
                hover_color=parent.dark3_color, anchor="n").pack(anchor="center", ipady=5, pady=(16, 0),)

class PrintOrder(ctk.CTkFrame):
    def __init__(self, parent):
        super().__init__(parent)
        
        main_view = ctk.CTkFrame(master=parent, fg_color=parent.dark0_color,  width=900, height=800, corner_radius=0)
        main_view.pack_propagate(0)
        main_view.pack(side="left")

        title_frame = ctk.CTkFrame(master=main_view, fg_color="transparent")
        title_frame.pack(anchor="n", fill="x",  padx=27, pady=(29, 0))
        ctk.CTkLabel(master=title_frame, text="列印出貨單", font=("Iansui", 36), text_color=parent.theme_color).pack(anchor="nw", side="left")
        
        #=============================== STORAGE PATH ======================================

        storage_path_container = ctk.CTkFrame(master=main_view, height=50, fg_color="transparent")
        storage_path_container.pack(fill="x", pady=(45, 0), padx=30)

        ctk.CTkLabel(master=storage_path_container, text="儲存位置: ", text_color="#fff", font=("Iansui", 24)).pack(side="left", padx=(13, 0), pady=5)
        ctk.CTkComboBox(master=storage_path_container, state="readonly", width=200, height = 40, font=("Iansui", 20), values=["貨單.xlsx"], button_color=parent.theme_color, border_color=parent.theme_color, 
                    border_width=2, button_hover_color=parent.theme_color_dark, dropdown_hover_color=parent.theme_color_dark, dropdown_fg_color=parent.theme_color, dropdown_text_color=parent.dark0_color).pack(side="left", padx=(13, 0), pady=15)

        #=============================== PRINTER ORDER ======================================

        prnit_order_container = ctk.CTkFrame(master=main_view, height=50, fg_color="transparent")
        prnit_order_container.pack(fill="x", pady=(10, 0), padx=30)

        ctk.CTkLabel(master=prnit_order_container, text="PG", text_color="#fff", font=("Iansui", 24)).pack(side="left", padx=(13, 0), pady=5)
        self.order_combobox = ctk.CTkComboBox(master=prnit_order_container, state="readonly", width=140, height = 40, font=("Iansui", 20), values=["018", "019", "020", "021"], button_color=parent.theme_color, border_color=parent.theme_color, 
                    border_width=2, button_hover_color=parent.theme_color_dark, dropdown_hover_color=parent.theme_color_dark, dropdown_fg_color=parent.theme_color, dropdown_text_color=parent.dark0_color)
        self.order_combobox.pack(side="left", padx=(13, 0), pady=15)
        self.order_entry = ctk.CTkEntry(master=prnit_order_container, width=300, height = 40, font=("Iansui", 20), placeholder_text="請輸入貨單後五碼", border_color=parent.theme_color, border_width=2)
        self.order_entry.pack(side="left", padx=(13, 0), pady=5)
        
        # hotkey binding
        self.order_entry.bind('<Return>', lambda event: self.printToprinter())
        # self.order_entry.bind('<Control-8>', lambda event: self.order_combobox.current(0))
        # self.order_entry.bind('<Control-9>', lambda event: self.order_combobox.current(1))

        ctk.CTkButton(master=prnit_order_container, width=100, height = 40, text="列印", font=("Iansui", 20), text_color=parent.dark0_color, 
                                          fg_color=parent.theme_color, hover_color=parent.theme_color_dark, command=self.printToprinter).pack(anchor="ne", padx=(40, 0), pady=15, side="left")
        
        #=============================== ORDER COUNT ======================================

        counting_container = ctk.CTkFrame(master=main_view, height=80, fg_color="transparent")
        counting_container.pack(fill="x", pady=(70, 0), padx=40)

        # total_order = 0
        # sucess_order = 0
        # total_order_string = ""
        # total_order_string_var = ctk.StringVar()
        # sucess_order_string = ctk.StringVar()
        # total_order_string_var.set("目前貨單總數: 0")
        # sucess_order_string.set("成功列印數量: 0")

        # def click():
        #     global total_order
        #     total_order = total_order + 1
        #     total_order_string = "儲存位置: " + str(total_order)
        #     total_order_string_var.set(total_order_string)

        total_count_metric = ctk.CTkFrame(master=counting_container, fg_color=parent.dark1_color,width=400, height=50)
        total_count_metric.pack(side="left")

        self.total_order_number = ctk.CTkLabel(master=total_count_metric, text="目前貨單總數: 0", text_color="#fff", font=("Iansui", 24))
        self.total_order_number.pack(side="left", padx=20, pady=5)
        
        #=============================== SEARCH BAR ======================================

        search_container = ctk.CTkFrame(master=main_view, height=50, fg_color="transparent")
        search_container.pack(fill="x", pady=(10, 0), padx=30)

        ctk.CTkEntry(master=search_container, width=300, height = 40, font=("Iansui", 20), placeholder_text="搜尋貨單", border_color=parent.theme_color, border_width=2).pack(side="left", padx=(13, 0), pady=5)

        ctk.CTkButton(master=search_container, width=100, height = 40, text="搜尋", font=("Iansui", 20), text_color=parent.dark0_color, fg_color=parent.theme_color, hover_color=parent.theme_color_dark, command=self.search_order).pack(anchor="ne", padx=(13, 0), pady=5, side="left")
        ctk.CTkButton(master=search_container, width=100, height = 40, text="刪除", font=("Iansui", 20), text_color=parent.dark0_color, fg_color=parent.theme_color, hover_color=parent.theme_color_dark).pack(anchor="ne", padx=(13, 0), pady=5, side="left")
        ctk.CTkComboBox(master=search_container, state="readonly", width=140, height = 40, font=("Iansui", 20), values=["顯示全部", "成功出貨", "關轉", "取消", "其他異常"], button_color=parent.theme_color, border_color=parent.theme_color, 
                    border_width=2, button_hover_color=parent.theme_color_dark, dropdown_hover_color=parent.theme_color_dark, dropdown_fg_color=parent.theme_color, dropdown_text_color=parent.dark0_color).pack(side="right", padx=(13, 0), pady=5)

        #============================ TABLE ============================

        table_data = [
            ["時間", "貨單編號", "狀態", "儲存狀態"],
            ["-","-","-","-"]
        ]

        table_frame = ctk.CTkScrollableFrame(master=main_view, fg_color="transparent")
        table_frame.pack(expand=True, fill="both", padx=27, pady=5)
        table = CTkTable(master=table_frame, height=36, font=("Iansui", 20), values=table_data, colors=[parent.dark3_color, parent.dark4_color], header_color=parent.theme_color, hover_color=parent.dark5_color)
        table.edit_row(0, text_color=parent.dark0_color, hover_color=parent.theme_color)
        table.pack(expand=True, fill="both", padx=10, pady=10)
        
    def printToprinter(self):
        if not self.order_combobox.get():
            messagebox.showwarning("沒有選擇輸入前綴", "請選擇貨單PG後數字!")
            print("empty combobox")
            return
        
        if len(self.order_entry.get()) == 5 and self.order_entry.get().isdigit():
            print(f"剛剛輸入的貨單為: PG0{self.order_combobox.get()}{self.order_entry.get()}")
            self.order_entry.delete(0, 'end')
        else:
            messagebox.showwarning("貨單後號碼錯誤", "請輸入正確貨單後五碼!")
            print("wrong order_entry input")
            self.order_entry.delete(0, 'end')
        
    def search_order(self):
        self.total_order_number.configure(text = "目前貨單總數11111")
        

if __name__ == "__main__":
    app = App()
    app.protocol("WM_DELETE_WINDOW", app.on_closing)
    app.mainloop()