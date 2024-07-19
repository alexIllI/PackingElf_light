import customtkinter as ctk
import tkinter
import os
from tkinter import messagebox

class frame_Shipping(ctk.CTkFrame):
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
        
        
    # turn other checkbox into DISABLE    
    def checkbox_event(self, checkbox_value):
        print(f"checkbox toggled: {checkbox_value}")
        