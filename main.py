from operation.autoweb import MyAcg, ReturnType
from operation.database_operation import DataBase

import customtkinter as ctk
import tkinter
from tkinter import ttk
from CTkTable import CTkTable
from tkinter import messagebox
from PIL import Image, ImageTk

from datetime import datetime, timedelta
from configparser import ConfigParser

class App(ctk.CTk):
    def __init__(self):
        super().__init__()
        
        #====================== Operation =========================
        self.driver = MyAcg()
        self.database = DataBase(datetime.now().strftime('Printed_Order_%Y_%m_%d'), 'operation\\data.db')
        self.database.check(datetime.now().strftime('Printed_Order_%Y_%m_%d'), (datetime.now()-timedelta(days=1)).strftime('Printed_Order_%Y_%m_%d'))
        
        self.total_order = 0
        self.success_order = 0
        self.current_id = 0
        
        #====================== Config ===============================
        config = ConfigParser()
        config.read("config.ini")
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
            self.driver.shut_down()
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
        
        #=============================== VARIABLES ======================================
        self.total_order = parent.total_order
        self.success_order = parent.success_order
        self.current_id = parent.current_id
        self.driver = parent.driver
        self.database = parent.database
        self.cancel_color = parent.cancel_color
        self.close_color = parent.close_color
        
        #=============================== SET UP ======================================
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

        ctk.CTkButton(master=prnit_order_container, width=100, height = 40, text="列印", font=("Iansui", 20), text_color=parent.dark0_color, 
                                          fg_color=parent.theme_color, hover_color=parent.theme_color_dark, command=self.printToprinter).pack(anchor="ne", padx=(40, 0), pady=15, side="left")
        
        #=============================== ORDER COUNT ======================================

        counting_container = ctk.CTkFrame(master=main_view, height=80, fg_color="transparent")
        counting_container.pack(fill="x", pady=(70, 0), padx=40)

        total_count_metric = ctk.CTkFrame(master=counting_container, fg_color=parent.dark1_color,width=400, height=50)
        total_count_metric.pack(side="left")
        
        success_count_metric = ctk.CTkFrame(master=counting_container, fg_color=parent.dark1_color,width=400, height=50)
        success_count_metric.pack(side="right")

        self.label_total_order_number = ctk.CTkLabel(master=total_count_metric, text="目前貨單總數: 0", text_color="#fff", font=("Iansui", 24))
        self.label_total_order_number.pack(side="left", padx=20, pady=5)
        
        self.label_success_order_number = ctk.CTkLabel(master=success_count_metric, text="成功列印貨單總數: 0", text_color="#fff", font=("Iansui", 24))
        self.label_success_order_number.pack(side="left", padx=20, pady=5)
        
        #=============================== SEARCH BAR ======================================

        search_container = ctk.CTkFrame(master=main_view, height=50, fg_color="transparent")
        search_container.pack(fill="x", pady=(10, 0), padx=30)

        self.search_entry = ctk.CTkEntry(master=search_container, width=300, height = 40, font=("Iansui", 20), placeholder_text="搜尋貨單", border_color=parent.theme_color, border_width=2)
        self.search_entry.pack(side="left", padx=(13, 0), pady=5)
        self.search_entry.bind('<Return>', lambda event: self.search_order())
        
        ctk.CTkButton(master=search_container, width=100, height = 40, text="搜尋", font=("Iansui", 20), text_color=parent.dark0_color, fg_color=parent.theme_color, hover_color=parent.theme_color_dark, command=self.search_order).pack(anchor="ne", padx=(13, 0), pady=5, side="left")
        ctk.CTkButton(master=search_container, width=100, height = 40, text="刪除", font=("Iansui", 20), text_color=parent.dark0_color, fg_color=parent.theme_color, hover_color=parent.theme_color_dark, command=self.btn_delete_items).pack(anchor="ne", padx=(13, 0), pady=5, side="left")
        self.view_status_enrty = ctk.CTkComboBox(master=search_container, state="readonly", width=140, height = 40, font=("Iansui", 20), values=["顯示全部", "成功出貨", "關轉", "取消"], button_color=parent.theme_color, border_color=parent.theme_color, 
                    border_width=2, button_hover_color=parent.theme_color_dark, dropdown_hover_color=parent.theme_color_dark, dropdown_fg_color=parent.theme_color, dropdown_text_color=parent.dark0_color, command=self.update)
        self.view_status_enrty.pack(side="right", padx=(13, 0), pady=5)

        #============================ TABLE ============================             

        table_frame = ctk.CTkFrame(master=main_view, fg_color="transparent")
        table_frame.pack(expand=True, fill="both", padx=27, pady=10)
        
        tree_scroll = ctk.CTkScrollbar(table_frame, button_color=parent.theme_color, button_hover_color=parent.theme_color_dark)
        tree_scroll.pack(side="right", fill="y")
        
        printed_order_table_style = ttk.Style()
        printed_order_table_style.theme_use("clam")
        printed_order_table_style.configure("Treeview.Heading", font=("Iansui", 25))
        printed_order_table_style.configure("Treeview", rowheight = 50, font=("Iansui", 20), background="#fff")
        printed_order_table_style.layout("Treeview", [('Treeview.treearea', {'sticky': 'nswe'})])
        
        self.printed_order_table = ttk.Treeview(table_frame, columns = ('id', 'time', 'order', 'status', 'save_status'), style="Treeview", show = 'headings', yscrollcommand=tree_scroll.set)
        self.printed_order_table.column("# 1", anchor="center",width=40)
        self.printed_order_table.column("# 2", anchor="center")
        self.printed_order_table.column("# 3", anchor="center")
        self.printed_order_table.column("# 4", anchor="center")
        self.printed_order_table.column("# 5", anchor="center")
        self.printed_order_table.heading('id', text = 'id')
        self.printed_order_table.heading('time', text = '時間')
        self.printed_order_table.heading('order', text = '貨單編號')
        self.printed_order_table.heading('status', text = '狀態')
        self.printed_order_table.heading('save_status', text = '儲存狀態')
        self.printed_order_table.pack(fill = 'both', expand = True)
        
        self.printed_order_table.tag_configure('cancel', background=parent.cancel_color)
        self.printed_order_table.tag_configure('close', background=parent.close_color)
        
        tree_scroll.configure(command=self.printed_order_table.yview)
        for i in range(10):
            self.total_order += 1
            self.success_order += 1
            self.current_id += 1
            self.database.insert_data(self.current_id, datetime.now().strftime('%H:%M:%S'), f"PG000{i}", 'success', "-")
            self.printed_order_table.insert(parent = '', index = 0, values = (self.current_id, datetime.now().strftime('%H:%M:%S'), f"PG000{i}", 'success', "-"))
        
        self.total_order += 1
        self.success_order += 1
        self.current_id += 1
        self.database.insert_data(self.current_id, datetime.now().strftime('%H:%M:%S'), f"PG000{11}", 'close', "-")
        self.printed_order_table.insert(parent = '', index = 0, values = (self.current_id, datetime.now().strftime('%H:%M:%S'), f"PG000{11}", 'close', "-"), tags = ("close",))
        
        self.total_order += 1
        self.success_order += 1
        self.current_id += 1
        self.database.insert_data(self.current_id, datetime.now().strftime('%H:%M:%S'), f"PG000{22}", 'close', "-")
        self.printed_order_table.insert(parent = '', index = 0, values = (self.current_id, datetime.now().strftime('%H:%M:%S'), f"PG000{22}", 'close', "-"), tags = ("close",))
        
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
                self.update()

        self.printed_order_table.bind('<<TreeviewSelect>>', item_select)
        self.printed_order_table.bind('<Delete>', delete_items)
    
    def update(self, status):
        self.printed_order_table.delete(*self.printed_order_table.get_children())
        if status == "顯示全部":
            datas = self.database.fetch_all_unrecorded("all")
        elif status == "成功出貨":
            datas = self.database.fetch_all_unrecorded("success")
        elif status == "關轉":
            datas = self.database.fetch_all_unrecorded("close")
        elif status == "取消":
            datas = self.database.fetch_all_unrecorded("cancel")
        
        for data in datas[1:]:
            if data[3] == 'success':
                self.printed_order_table.insert(parent = '', index = 0, values = (data[0], data[1], data[2], "成功", data[4]))
            if data[3] == 'close':
                self.printed_order_table.insert(parent = '', index = 0, values = (data[0], data[1], data[2], "關轉", data[4]), tags = (status,))
            if data[3] == 'cancel':
                self.printed_order_table.insert(parent = '', index = 0, values = (data[0], data[1], data[2], "取消", data[4]), tags = (status,))
                
    def search_order(self):
        if not self.search_entry.get():
            messagebox.showwarning("搜尋貨單結果", "老哥，先輸入再搜尋好嗎")
            print("empty combobox")
            return
        
        if len(self.search_entry.get()) != 10:
            messagebox.showwarning("搜尋貨單結果", "請輸入完整貨單!")
            print("wrong order length")
            self.order_entry.delete(0, 'end')
            return
        order_number = self.search_entry.get()
        self.search_entry.delete(0, 'end')
        
        # try:
        result_id = self.database.search_order(order_number)
        print(result_id)
        # except Exception as e:
        #     print(f"Error in main -> search_order: {e}")
        #     return
        
        if result_id:
            if result_id[1] == "unrecorded":
                self.printed_order_table.selection_remove(self.printed_order_table.selection())
                for child in self.printed_order_table.get_children():
                    if order_number in self.printed_order_table.item(child)['values'][2]:
                        print(f"find {self.printed_order_table.item(child)['values'][2]} within treeview, id: {result_id[0]}")
                        self.printed_order_table.selection_set(child)
                        messagebox.showinfo("搜尋貨單結果", f"貨單編號: {order_number}\nID: {result_id[0]}\n狀態: {self.printed_order_table.item(child)['values'][3]}\n儲存位置: {self.printed_order_table.item(child)['values'][4]}")
                        return
            else:
                messagebox.showwarning("搜尋貨單結果", "資料庫中有這筆貨單，但不是在本次應用程式執行後紀錄，因此無法顯示在下方表格!")
        else:
            messagebox.showwarning("搜尋貨單結果", "搜尋的貨單不存在!")
        
    def btn_delete_items(self):
        if len(self.printed_order_table.selection()) > 1:
                messagebox.showwarning("刪除貨單警告", "一次只能刪除一筆貨單，請只選擇一筆刪除!")
                return
        
        print(f"delete {self.printed_order_table.item(self.printed_order_table.selection())['values'][2]}")
        self.database.delete_data(self.printed_order_table.item(self.printed_order_table.selection())['values'][2])
        self.update(self.view_status_enrty.get())
        
    def printToprinter(self):
        def print_cancel_close(_status:str, order):
            self.total_order += 1
            self.current_id += 1
            self.label_total_order_number.configure(text = f"目前貨單總數: {self.total_order}")
            cur_time = datetime.now().strftime('%H:%M:%S')
            self.database.insert_data(self.current_id, cur_time, order, _status, "-")
            if _status == "close":
                self.printed_order_table.insert(parent = '', index = 0, values = (self.current_id, cur_time, order, "關轉", "-"), tags = (_status,))
            else:
                self.printed_order_table.insert(parent = '', index = 0, values = (self.current_id, cur_time, order, "取消", "-"), tags = (_status,))
        
        def print_success(order):
            self.total_order += 1
            self.success_order += 1
            self.current_id += 1
            self.label_total_order_number.configure(text = f"目前貨單總數: {self.total_order}")
            self.label_success_order_number.configure(text = f"成功列印貨單總數: {self.success_order}")
            cur_time = datetime.now().strftime('%H:%M:%S')
            self.database.insert_data(self.current_id, cur_time, order, 'success', "-")
            self.printed_order_table.insert(parent = '', index = 0, values = (self.current_id, cur_time, order, '成功', "-"))
        
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
        result = self.driver.printer(current_order)
          
        if result == ReturnType.REPEAT:
            messagebox.showwarning("貨單後號碼錯誤", "出貨單重複!如果想重印這一單，請先將下方同樣貨單編號的紀錄刪除")
            print("repeat")
            
        elif result == ReturnType.OPEN_FILE_ERROR:
            print("openfile error")
        
        elif result == ReturnType.MULTIPLE_TAB:
            messagebox.showwarning("網頁自動化錯誤", "請開啟瀏覽器將買動漫以外的分頁關閉!")
            print("meltiple tab detected")
            
        elif result == ReturnType.POPUP_UNSOLVED:
            messagebox.showwarning("網頁自動化錯誤", "你可能沒有按'我知道了',打開買動漫,按下去")
            print("popup unsolved")
        
        elif result == ReturnType.ALREADY_FINISH:
            messagebox.showwarning("列印出貨單錯誤", "該訂單已取貨!")
            print("already finished")
        
        elif result == ReturnType.ORDER_NOT_FOUND:
            messagebox.showwarning("貨單後號碼錯誤", "沒有這一單!")
            print("order not found")
            
        elif result == ReturnType.CHECKBOX_NOT_FOUND:
            messagebox.showwarning("網頁自動化錯誤", "找不到checkbox,可能沒有這一單,自己開買動漫看一下")
            print("check box not found")
            
        elif result == ReturnType.CLICKING_CHECKBOX_ERROR:
            print("can't click check box")
            
        elif result == ReturnType.ORDER_CANCELED:
            messagebox.showwarning("取消", "此單已被取消!請至買動漫確認)")
            print_cancel_close("cancel", current_order)
            print("order canceled")
            
        elif result == ReturnType.STORE_CLOSED:
            messagebox.showwarning("網頁自動化錯誤", "無法切換視窗(可能是關轉，去看買動漫，如果有我知道了按下去，不然我會當掉)")
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