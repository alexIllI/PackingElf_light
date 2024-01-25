from operation.autoweb import MyAcg, ReturnType
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
    def __init__(self):
        super().__init__()
        try:
            pyglet.font.add_file(resource_path("font\\Iansui-Regular.ttf"))
        except:
            print("add font error")
        self.total_order = 0
        self.success_order = 0
        self.current_id = 0
        self.current_time_name = datetime.now().strftime('Printed_Order_%Y_%m_%d')
        
        #====================== Operation =========================
        self.driver = MyAcg()
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
        
        self.geometry("900x600")
        # self.resizable(0,0)
        self.title("包貨小精靈")
        self.iconbitmap(resource_path("images\icon.ico"))
        
        ctk.set_appearance_mode("dark")
        
        #========================== Side Bar ===========================
        self.sidebar = SideBar(self)
        self.Page_printorder = PrintOrder(self)
        
    def on_closing(self):
        if messagebox.askokcancel("退出包貨小精靈", "確定要退出? (會自動匯出本次所有貨單)"):
            self.driver.shut_down()
            while True:
                close_result = self.database.close_database()
                if close_result == DBreturnType.CLOSE_AND_SAVE_ERROR:
                    messagebox.showwarning("匯出貨單時發生錯誤", "匯出貨單時發生錯誤, 包貨紀錄將不會匯出至excel!\n(下次啟動應用程式時將會嘗試匯出)")
                    break
                elif close_result == DBreturnType.PERMISSION_ERROR:
                    messagebox.showwarning("匯出貨單時發生錯誤", "匯出貨單時發生錯誤, 請先將儲存貨單的excel關閉!")
                elif close_result == DBreturnType.SUCCESS:
                    break
            
            self.destroy()
            print("close app")
        
class SideBar(ctk.CTkFrame):
    def __init__(self, parent):
        super().__init__(parent)
        self.sidebar_frame = ctk.CTkFrame(master=parent, fg_color=parent.dark2_color,  width=225, height=600, corner_radius=0)
        self.sidebar_frame.pack_propagate(0)
        self.sidebar_frame.pack(fill="y", anchor="w", side="left")
        
        logo_img_data = Image.open(resource_path("images\icon_meridian_white.png"))
        logo_img = ctk.CTkImage(dark_image=logo_img_data, light_image=logo_img_data, size=(135, 136.3))
        ctk.CTkLabel(master=self.sidebar_frame, text="", image=logo_img).pack(pady=(60, 0), anchor="center")
        
        package_img_data = Image.open(resource_path("images\printer.png"))
        package_img = ctk.CTkImage(dark_image=package_img_data, light_image=package_img_data)

        ctk.CTkButton(master=self.sidebar_frame, width=200, image=package_img, text="列印出貨單", fg_color=parent.dark1_color, font=("Iansui", 18), 
                hover_color=parent.dark3_color, anchor="center").pack(anchor="center", ipady=5, pady=(135, 0))

        list_img_data = Image.open(resource_path("images\list_icon.png"))
        list_img = ctk.CTkImage(dark_image=list_img_data, light_image=list_img_data)
        ctk.CTkButton(master=self.sidebar_frame, width=200, image=list_img, text="儲存管理", fg_color="transparent", font=("Iansui", 18), 
                hover_color=parent.dark3_color, anchor="center").pack(anchor="center", ipady=5, pady=(16, 0))

        settings_img_data = Image.open(resource_path("images\settings_icon.png"))
        settings_img = ctk.CTkImage(dark_image=settings_img_data, light_image=settings_img_data)
        ctk.CTkButton(master=self.sidebar_frame, width=200, image=settings_img, text="設定", fg_color="transparent", font=("Iansui", 18), 
                hover_color=parent.dark3_color, anchor="center").pack(anchor="center", ipady=5, pady=(16, 0),)

class PrintOrder(ctk.CTkFrame):
    def __init__(self, parent):
        super().__init__(parent)
        
        #=============================== VARIABLES ======================================
        self.total_order = parent.total_order
        self.success_order = parent.success_order
        self.current_id = parent.current_id
        self.cur_time_name = parent.current_time_name
        self.driver = parent.driver
        self.database = parent.database
        self.cancel_color = parent.cancel_color
        self.close_color = parent.close_color
        
        #=============================== SET UP ======================================
        main_view = ctk.CTkFrame(master=parent, fg_color=parent.dark0_color,  width=675, height=600, corner_radius=0)
        main_view.pack_propagate(0)
        main_view.pack(side="left")

        title_frame = ctk.CTkFrame(master=main_view, fg_color="transparent")
        title_frame.pack(anchor="n", fill="x",  padx=27, pady=(29, 0))
        ctk.CTkLabel(master=title_frame, text="列印出貨單", font=("Iansui", 32), text_color=parent.theme_color).pack(anchor="nw", side="left")
        
        #=============================== STORAGE PATH ======================================

        storage_path_container = ctk.CTkFrame(master=main_view, height=37.5, fg_color="transparent")
        storage_path_container.pack(fill="x", pady=(35, 0), padx=30)

        ctk.CTkLabel(master=storage_path_container, text="儲存位置: ", text_color="#fff", font=("Iansui", 24)).pack(side="left", padx=(13, 0), pady=5)
        self.save_path_combobox = ctk.CTkComboBox(master=storage_path_container, state="readonly", width=200, height = 40, font=("Iansui", 20), values=["(現在沒功能)"], button_color=parent.theme_color, border_color=parent.theme_color, 
                    border_width=2, button_hover_color=parent.theme_color_dark, dropdown_hover_color=parent.theme_color_dark, dropdown_fg_color=parent.theme_color, dropdown_text_color=parent.dark0_color)
        self.save_path_combobox.pack(side="left", padx=(13, 0), pady=15)
        self.save_path_combobox.set("(現在沒功能)")

        #=============================== PRINTER ORDER ======================================

        prnit_order_container = ctk.CTkFrame(master=main_view, height=37.5, fg_color="transparent")
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

        counting_container = ctk.CTkFrame(master=main_view, height=60, fg_color="transparent")
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

        search_container = ctk.CTkFrame(master=main_view, height=37.5, fg_color="transparent")
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

        table_frame = ctk.CTkFrame(master=main_view, fg_color="transparent")
        table_frame.pack(expand=True, fill="both", padx=27, pady=10)
        
        tree_scroll = ctk.CTkScrollbar(table_frame, button_color=parent.theme_color, button_hover_color=parent.theme_color_dark)
        tree_scroll.pack(side="right", fill="y")
        
        printed_order_table_style = ttk.Style()
        printed_order_table_style.theme_use("clam")
        printed_order_table_style.configure("Treeview.Heading", font=("Iansui", 20))
        printed_order_table_style.configure("Treeview", rowheight = 50, font=("Iansui", 16), background="#fff")
        printed_order_table_style.layout("Treeview", [('Treeview.treearea', {'sticky': 'nswe'})])
        
        self.printed_order_table = ttk.Treeview(table_frame, columns = ('id', 'time', 'order', 'status', 'save_status'), style="Treeview", show = 'headings', yscrollcommand=tree_scroll.set)
        self.printed_order_table.column("# 1", anchor="center",width=35)
        self.printed_order_table.column("# 2", anchor="center",width=180)
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
        
        #======================================================== TEST ========================================================
        # for i in range(10):
        #     self.total_order += 1
        #     self.success_order += 1
        #     self.current_id += 1
        #     self.database.insert_data(self.current_id, datetime.now().strftime('%H:%M:%S'), f"PG000{i}", 'success', self.cur_time_name)
        #     self.printed_order_table.insert(parent = '', index = 0, values = (self.current_id, datetime.now().strftime('%H:%M:%S'), f"PG000{i}", '成功', self.cur_time_name))
        
        # self.total_order += 1
        # self.current_id += 1
        # self.database.insert_data(self.current_id, datetime.now().strftime('%H:%M:%S'), f"PG000{11}", 'close', self.cur_time_name)
        # self.printed_order_table.insert(parent = '', index = 0, values = (self.current_id, datetime.now().strftime('%H:%M:%S'), f"PG000{11}", '關轉', self.cur_time_name), tags = ("close",))
        
        # self.total_order += 1
        # self.current_id += 1
        # self.database.insert_data(self.current_id, datetime.now().strftime('%H:%M:%S'), f"PG000{22}", 'close', self.cur_time_name)
        # self.printed_order_table.insert(parent = '', index = 0, values = (self.current_id, datetime.now().strftime('%H:%M:%S'), f"PG000{22}", '關轉', self.cur_time_name), tags = ("close",))
        #======================================================== TEST ========================================================
        
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
        result = self.driver.printer(current_order)
        
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