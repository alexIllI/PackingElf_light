import tkinter as tk
from tkinter import ttk, filedialog
import pandas as pd
import pymysql

def connect_db():
    try:
        connection = pymysql.connect(
            # host="192.168.1.114",
            host='localhost',
            user='root',
            password='Meridian0723',
            database='Myacg_data'
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

def get_table_columns(table_name):
    connection = connect_db()
    if connection is not None:
        try:
            cursor = connection.cursor()
            cursor.execute(f"SHOW COLUMNS FROM `{table_name}`")
            columns = [column[0] for column in cursor.fetchall()]
            return columns
        except pymysql.MySQLError as e:
            print(f"Error fetching columns for table {table_name}: {e}")
            return []
        finally:
            connection.close()
    else:
        print(f"Failed to connect to database to fetch columns for table {table_name}.")
        return []

def update_table_with_excel():
    table_name = combobox.get()
    excel_path = file_path_entry.get()
    if not table_name or not excel_path:
        print("Please select a table and an Excel file.")
        return
    
    try:
        df = pd.read_excel(excel_path)
        df = df.where(pd.notnull(df), None)  # Replace NaN with None
    except Exception as e:
        print(f"Error reading Excel file: {e}")
        return

    table_columns = get_table_columns(table_name)
    print(f"Columns in table `{table_name}`: {table_columns}")

    if '發票號碼' not in df.columns or '開立時間' not in df.columns:
        print("The Excel file must contain '發票號碼' and '開立時間' columns.")
        return

    if 'invoice' not in table_columns or 'invoice_time' not in table_columns:
        print("The database table must contain 'invoice' and 'invoice_time' columns.")
        return

    connection = connect_db()
    if connection is not None:
        try:
            cursor = connection.cursor()

            cursor.execute(f"SHOW COLUMNS FROM `{table_name}` LIKE 'invoice_time'")
            result = cursor.fetchone()
            if not result:
                cursor.execute(f"ALTER TABLE `{table_name}` ADD COLUMN invoice_time DATETIME")
                connection.commit()
            
            for index, row in df.iterrows():
                invoice_number = row['發票號碼']
                invoice_time = row['開立時間']

                update_query = f"UPDATE `{table_name}` SET invoice_time = %s WHERE invoice = %s"
                cursor.execute(update_query, (invoice_time, invoice_number))
                connection.commit()
            print(f"Table `{table_name}` updated successfully.")
        
        except pymysql.MySQLError as e:
            print(f"Error updating table {table_name}: {e}")
        finally:
            connection.close()
    else:
        print("Failed to connect to database.")

# Setting up the GUI
root = tk.Tk()
root.title("上傳發票開立時間")

# Dropdown box to select table name
label = ttk.Label(root, text="選擇要更新的月份:")
label.pack(pady=5)

table_names = fetch_table_names()
combobox = ttk.Combobox(root, values=table_names)
combobox.pack(pady=5)

# Button to select the Excel file
def select_file():
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
    if file_path:
        file_path_entry.delete(0, tk.END)
        file_path_entry.insert(0, file_path)

select_file_button = ttk.Button(root, text="選擇 Excel", command=select_file)
select_file_button.pack(pady=5)

# File path entry
file_path_entry = ttk.Entry(root, width=50)
file_path_entry.pack(pady=5)

# Update button
update_button = ttk.Button(root, text="上傳", command=update_table_with_excel)
update_button.pack(pady=20)

root.mainloop()
