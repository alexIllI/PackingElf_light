import tkinter as tk
from tkinter import ttk
from tkinter import filedialog, messagebox
import pandas as pd
from sqlalchemy import create_engine, inspect

def connect_db():
    try:
        # Create the database URL for SQLAlchemy
        # db_host = '192.168.1.11'
        db_host = 'localhost'
        database_url = f'mysql+pymysql://root:Meridian0723@{db_host}/myacg_data'
        # Create the engine
        engine = create_engine(database_url)
        return engine
    except Exception as e:
        print(f"Error connecting to MySQL Database: {e}")
        return None

def fetch_table_names():
    engine = connect_db()
    if engine is not None:
        try:
            inspector = inspect(engine)
            table_names = inspector.get_table_names()
            print(f"Fetched table names: {table_names}")
            return table_names
        except Exception as e:
            print(f"Error fetching table names: {e}")
            return []
    else:
        print("Failed to connect to database for fetching tables.")
        return []

def check_column_exists(table_name, column_name):
    engine = connect_db()
    if engine is not None:
        try:
            inspector = inspect(engine)
            columns = [col['name'] for col in inspector.get_columns(table_name)]
            return column_name in columns
        except Exception as e:
            print(f"Error checking column existence: {e}")
            return False
    else:
        print("Failed to connect to database for checking column.")
        return False

def export_table():
    table_name = combobox.get()
    folder_path = folder_path_entry.get()
    if not table_name:
        messagebox.showwarning("Warning", "請選擇匯出月份")
        return
    if not folder_path:
        messagebox.showwarning("Warning", "請選擇匯出目的資料夾")
        return

    if not check_column_exists(table_name, "invoice_time"):
        messagebox.showwarning("Warning", "請先補上發票時間後再匯出")
        return
    
    engine = connect_db()
    if engine is not None and folder_path:
        try:
            # Specify the columns to fetch
            query = f"SELECT time AS 時間, order_establish_date AS 訂單成立時間, invoice AS 發票號碼, invoice_time AS 發票開立時間, using_coupon AS 使用優惠券 FROM `{table_name}`"
            df = pd.read_sql(query, engine)

            # Export to Excel file, replacing if it already exists
            file_path = f"{folder_path}/{table_name}.xlsx"
            df.to_excel(file_path, index=False)
            print(f"Data exported to file: {file_path}")

        except Exception as e:
            print(f"Error exporting table {table_name}: {e}")
    else:
        if not engine:
            print("Failed to connect to database.")
        if not folder_path:
            print("Folder path not specified.")

# Setting up the GUI
root = tk.Tk()
root.title("Export Table to Excel")

# Dropdown box to select table name
label = ttk.Label(root, text="選擇輸出的月份:")
label.pack(pady=5)

table_names = fetch_table_names()
combobox = ttk.Combobox(root, values=table_names)
combobox.pack(pady=5)

# Export button
export_button = ttk.Button(root, text="Export to Excel", command=export_table)
export_button.pack(pady=20)

# Select folder button
def select_folder():
    folder_path = filedialog.askdirectory()
    if folder_path:
        folder_path_entry.delete(0, tk.END)
        folder_path_entry.insert(0, folder_path)
        print(f"Selected folder path: {folder_path}")

select_folder_button = ttk.Button(root, text="Select Folder", command=select_folder)
select_folder_button.pack(pady=5)

# Folder path entry
folder_path_entry = ttk.Entry(root)
folder_path_entry.pack(pady=5)

root.mainloop()
