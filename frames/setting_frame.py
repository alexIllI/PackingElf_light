import customtkinter as ctk
import tkinter
import os
from tkinter import messagebox

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