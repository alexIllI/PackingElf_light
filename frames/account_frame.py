import customtkinter as ctk
from tkinter import messagebox
from operation.myacg_manager import AccountReturnType

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
