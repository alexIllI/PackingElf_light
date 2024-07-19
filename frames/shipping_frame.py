from operation.myacg_manager import ReturnType
from operation.database_operation import DBreturnType

import customtkinter as ctk
from tkinter import ttk
from tkinter import messagebox

class frame_Shipping(ctk.CTkFrame):
    def __init__(self, parent_frame, parent):
        ctk.CTkFrame.__init__(self, parent_frame, fg_color=parent.dark0_color, corner_radius=0)    
        #=============================== VARIABLES ======================================
        self.myacg_manager = parent.myacg_manager
        self.database = parent.database
        
        self.cancel_color = parent.cancel_color
        self.close_color = parent.close_color
        self.table_list = self.database.get_output_excel_options()
        
        self.order_prefix = ["019", "020", "021", "022", "023"]
        self.default_order_prefix = "020"

        #=============================== TITLE ======================================

        title_frame = ctk.CTkFrame(master=self, fg_color="transparent")
        title_frame.pack(anchor="n", fill="x",  padx=27, pady=(29, 0))
        ctk.CTkLabel(master=title_frame, text="宅配紀錄", font=("Iansui", 32), text_color=parent.theme_color).pack(anchor="nw", side="left")

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

        #=============================== RECORD ORDER ======================================

        record_order_container = ctk.CTkFrame(master=self, height=37.5, fg_color="transparent")
        record_order_container.pack(fill="x", pady=(0, 0), padx=30)

        ctk.CTkLabel(master=record_order_container, text="PG", text_color="#fff", font=("Iansui", 24)).pack(side="left", padx=(13, 0), pady=5)
        self.order_combobox = ctk.CTkComboBox(master=record_order_container, state="readonly", width=105, height = 40, font=("Iansui", 20), values=self.order_prefix, button_color=parent.theme_color, border_color=parent.theme_color, 
                    border_width=2, button_hover_color=parent.theme_color_dark, dropdown_hover_color=parent.theme_color_dark, dropdown_fg_color=parent.theme_color, dropdown_text_color=parent.dark0_color)
        self.order_combobox.pack(side="left", padx=(13, 0), pady=15)
        self.order_combobox.set(self.default_order_prefix)
        self.order_entry = ctk.CTkEntry(master=record_order_container, width=225, height = 40, font=("Iansui", 20), placeholder_text="請輸入貨單後五碼", border_color=parent.theme_color, border_width=2)
        self.order_entry.pack(side="left", padx=(23, 0), pady=5)
        
        # hotkey binding
        self.order_entry.bind('<Return>', lambda event: self.StoreDeliveryOrder())

        ctk.CTkButton(master=record_order_container, width=75, height = 40, text="儲存", font=("Iansui", 18), text_color=parent.dark0_color, 
                                          fg_color=parent.theme_color, hover_color=parent.theme_color_dark, command=self.StoreDeliveryOrder).pack(anchor="ne", padx=(100, 0), pady=15, side="left")
        
        #=============================== ORDER COUNT ======================================

        counting_container = ctk.CTkFrame(master=self, height=60, fg_color="transparent")
        counting_container.pack(fill="x", pady=(30, 0), padx=40)

        total_count_metric = ctk.CTkFrame(master=counting_container, fg_color=parent.dark1_color,width=300, height=37.5)
        total_count_metric.pack(side="left")

        self.label_total_order_number = ctk.CTkLabel(master=total_count_metric, text="目前紀錄宅配貨單總數: 0", text_color="#fff", font=("Iansui", 22))
        self.label_total_order_number.pack(side="left", padx=20, pady=5)
        
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
        
        self.view_status_enrty = ctk.CTkComboBox(master=search_container, state="readonly", width=105, height = 30, font=("Iansui", 16), values=["顯示全部", "成功出貨", "關轉", "取消", "宅配"], button_color=parent.theme_color, border_color=parent.theme_color, 
                    border_width=2, button_hover_color=parent.theme_color_dark, dropdown_hover_color=parent.theme_color_dark, dropdown_fg_color=parent.theme_color, dropdown_text_color=parent.dark0_color, command= lambda value: self.update_table(value, self.store_table_combobox.get()))
        self.view_status_enrty.pack(side="right", padx=(13, 0), pady=5)
        self.view_status_enrty.set("顯示全部")
        
        #============================ TABLE ============================             

        table_frame = ctk.CTkFrame(master=self, fg_color="transparent")
        table_frame.pack(expand=True, fill="both", padx=27, pady=10)
        
        tree_scroll = ctk.CTkScrollbar(table_frame, button_color=parent.theme_color, button_hover_color=parent.theme_color_dark)
        tree_scroll.pack(side="right", fill="y")
        
        delivery_order_table_style = ttk.Style()
        delivery_order_table_style.theme_use("clam")
        delivery_order_table_style.configure("Treeview.Heading", font=("Iansui", 16))
        delivery_order_table_style.configure("Treeview", rowheight = 50, font=("Iansui", 14), background="#fff")
        delivery_order_table_style.layout("Treeview", [('Treeview.treearea', {'sticky': 'nswe'})])
        
        self.delivery_order_table = ttk.Treeview(table_frame, columns = ('id', 'time', 'order', 'status', 'invoice', 'output_excel'), style="Treeview", show = 'headings', yscrollcommand=tree_scroll.set)
        self.delivery_order_table.column("# 1", anchor="center",width=30)
        self.delivery_order_table.column("# 2", anchor="center",width=145)
        self.delivery_order_table.column("# 3", anchor="center")
        self.delivery_order_table.column("# 4", anchor="center", width=120)
        self.delivery_order_table.column("# 5", anchor="center")
        self.delivery_order_table.column("# 6", anchor="center")
        self.delivery_order_table.heading('id', text = 'id')
        self.delivery_order_table.heading('time', text = '時間')
        self.delivery_order_table.heading('order', text = '貨單編號')
        self.delivery_order_table.heading('status', text = '狀態')
        self.delivery_order_table.heading('invoice', text = '發票號碼')
        self.delivery_order_table.heading('output_excel', text = '儲存位置')
        self.delivery_order_table.pack(fill = 'both', expand = True)
        
        self.delivery_order_table.tag_configure('cancel', background=parent.cancel_color)
        self.delivery_order_table.tag_configure('close', background=parent.close_color)
        
        tree_scroll.configure(command=self.delivery_order_table.yview)
        
        # events
        def item_select(_):
            print(self.delivery_order_table.selection())
            for i in self.delivery_order_table.selection():
                print(self.delivery_order_table.item(i)['values'])

        def delete_items(_):
            print('delete pressed')
            if len(self.delivery_order_table.selection()) > 1:
                messagebox.showwarning("刪除貨單警告", "一次只能刪除一筆貨單，請只選擇一筆刪除!")
                return
            
            if messagebox.askokcancel("刪除貨單", f"確定要刪除 {self.delivery_order_table.item(self.delivery_order_table.selection())['values'][2]} ? (刪除後不可復原)"):
                print(f"delete {self.delivery_order_table.item(self.delivery_order_table.selection())['values'][2]}")
                self.database.delete_data(self.delivery_order_table.item(self.delivery_order_table.selection())['values'][2])
                self.update_table(self.view_status_enrty.get(), self.store_table_combobox.get())
                
        def deselect_items(_):
            self.delivery_order_table.selection_remove(*self.delivery_order_table.selection())
            print("deselect all items")

        self.delivery_order_table.bind('<<TreeviewSelect>>', item_select)
        self.delivery_order_table.bind('<Delete>', delete_items)
        self.delivery_order_table.bind('<Escape>', deselect_items)
        
    def on_validate(self, P):
        """Validate new content if it's purely ASCII."""
        if P.isascii():
            return True
        else:
            messagebox.showwarning("輸入法錯誤", "喇咪說這樣不能印，請將輸入法切換成英文後重新掃描一次條碼")
            self.after(1, lambda: self.clear_entry(self.invoice_entry))  # Schedule the clear_entry to avoid issues with direct deletion within validation
            return False
    
    def update_table(self, status, table_name):
        self.delivery_order_table.delete(*self.delivery_order_table.get_children())
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
        elif status == "宅配":
            datas = self.database.fetch_table_data('delivery', record, table_name)
                
        for data in datas:
            if data[3] == 'success':
                self.delivery_order_table.insert(parent = '', index = 0, values = (data[0], data[1], data[2], "成功", data[4], data[5]))
            elif data[3] == 'close':
                self.delivery_order_table.insert(parent = '', index = 0, values = (data[0], data[1], data[2], "關轉", data[4], data[5]), tags = ("close",))
            elif data[3] == 'cancel':
                self.delivery_order_table.insert(parent = '', index = 0, values = (data[0], data[1], data[2], "取消", data[4], data[5]), tags = ("cancel",))
            elif data[3] == 'delivery':
                self.delivery_order_table.insert(parent = '', index = 0, values = (data[0], data[1], data[2], "宅配", data[4], data[5]), tags = ("delivery",))
        
        order_numbers = self.database.count_records(self.store_table_combobox.get(), record) 
        self.label_total_order_number.configure(text = f"目前貨單總數: {order_numbers[0]}")
                
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
                self.delivery_order_table.selection_remove(self.delivery_order_table.selection())
                for child in self.delivery_order_table.get_children():
                    if order_number in self.delivery_order_table.item(child)['values'][2]:
                        print(f"find {self.delivery_order_table.item(child)['values'][2]} within treeview, id: {result[0]}")
                        self.delivery_order_table.selection_set(child)
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
            if self.delivery_order_table.selection():
                print('delete selected row using btn')
                if len(self.delivery_order_table.selection()) > 1:
                    messagebox.showwarning("刪除貨單警告", "一次只能刪除一筆貨單，請只選擇一筆刪除!")
                    return
                
                if messagebox.askokcancel("刪除貨單", f"確定要刪除 {self.delivery_order_table.item(self.delivery_order_table.selection())['values'][2]} ? (刪除後不可復原)"):
                    print(f"delete {self.delivery_order_table.item(self.delivery_order_table.selection())['values'][2]}")
                    self.database.delete_data(self.delivery_order_table.item(self.delivery_order_table.selection())['values'][2])
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
                
    def StoreDeliveryOrder(self):
        validate_string = self.invoice_entry.get()
        print(validate_string)
        if not validate_string.isascii():
            self.invoice_entry.delete(0, 'end')
            self.invoice_entry.focus_set()
            messagebox.showerror("輸入法錯誤", "喇咪說這樣不能印，請將輸入法切換成英文後重新掃描一次條碼")
            return
        
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
        self.invoice_entry.focus_set()
        
        self.database.insert_data(current_order, 'delivery', current_invoice, self.store_table_combobox.get(), "-", "-")
        self.update_table(self.view_status_enrty.get(), self.store_table_combobox.get())