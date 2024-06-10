import tkinter as tk
from tkinter import ttk
from tkinter import filedialog
import pandas as pd
import pymysql

def connect_db():
    try:
        connection = pymysql.connect(
            host='localhost',
            user='root',
            password='Meridian0723',
            database='myacg_data'
        )
        return connection
    except pymysql.MySQLError as e:
        print(f"Error connecting to MySQL Database: {e}")
        return None

def fetch_table_names():
    connection = connect_db()
    if connection is not None:
        try:
            cursor = connection.cursor()
            cursor.execute("SHOW TABLES")
            table_names = [table[0] for table in cursor.fetchall()]
            return table_names
        except pymysql.MySQLError as e:
            print(f"Error fetching table names: {e}")
            return []
        finally:
            connection.close()
    else:
        print("Failed to connect to database for fetching tables.")
        return []

def export_table():
    table_name = combobox.get()
    connection = connect_db()
    if connection is not None:
        try:
            # Specify the columns to fetch
            query = f"SELECT time AS 時間, order_number AS 貨單編號, invoice AS 發票號碼 FROM `{table_name}`"
            df = pd.read_sql(query, connection)

            # Export to Excel file, replacing if it already exists
            file_path = f"{table_name}.xlsx"
            df.to_excel(file_path, index=False)
            print(f"Data exported to file: {file_path}")

        except Exception as e:
            print(f"Error exporting table {table_name}: {e}")
        finally:
            connection.close()
    else:
        print("Failed to connect to database.")

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

select_folder_button = ttk.Button(root, text="Select Folder", command=select_folder)
select_folder_button.pack(pady=5)

# Folder path entry
folder_path_entry = ttk.Entry(root)
folder_path_entry.pack(pady=5)
root.mainloop()
