import customtkinter as ctk
from CTkTable import CTkTable
from PIL import Image, ImageTk

from configparser import ConfigParser

class App(ctk.CTk):
    def __init__(self):
        super().__init__()
        
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
        
        self.geometry("1200x800")
        self.resizable(0,0)
        self.title("包貨小精靈")
        self.iconbitmap("images\icon.ico")
        
        ctk.set_appearance_mode("dark")
        
        #========================== Side Bar =============================
        self.sidebar = SideBar(self)
        
        
class SideBar(ctk.CTkFrame):
    def __init__(self, parent):
        super().__init__(parent)
        self.sidebar_frame = ctk.CTkFrame(master=parent, fg_color=parent.dark2_color,  width=300, height=800, corner_radius=0)
        self.sidebar_frame.pack_propagate(0)
        self.sidebar_frame.pack(fill="y", anchor="w", side="left")
        
        logo_img_data = Image.open("images\icon_meridian_white.png")
        logo_img = ctk.CTkImage(dark_image=logo_img_data, light_image=logo_img_data, size=(180, 186))
        ctk.CTkLabel(master=self.sidebar_frame, text="", image=logo_img).pack(pady=(60, 0), anchor="center")
        
# class SideBar(ctk.CTkFrame):
#     def __init__(self, parent):
#         super().__init__(parent)
        
        
if __name__ == "__main__":
    app = App()
    app.mainloop()