import os
import sys
import sqlite3
import pandas as pd
from enum import Enum

def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)

#====================== Enum ===============================
class DBreturnType(Enum):
    SUCCESS = "SUCCESS"
    PERMISSION_ERROR = "PERMISSION_ERROR"
    EXPORT_UNRECORDED_ERROR = "EXPORT_UNRECORDED_ERROR"
    CHECK_YESTERDAY_ERROR = "CHECK_YESTERDAY_ERROR"
    CHECK_TODAY_ERROR = "CHECK_TODAY_ERROR"
    ORDER_NOT_FOUND = "ORDER_NOT_FOUND"
    CLOSE_AND_SAVE_ERROR = "CLOSE_AND_SAVE_ERROR"

class DataBase():
    def __init__(self, table_name:str, database_path:str ,save_path:str):
        #================= Date for DBname ========================
        self.table_name = table_name
        self.save_path = resource_path(save_path)
        
        self.connection = sqlite3.connect(resource_path(database_path))
        table_create_query = f"CREATE TABLE IF NOT EXISTS {self.table_name} (table_id INT, time TEXT, order_number TEXT, status TEXT, save_status TEXT, record TEXT)"
        self.connection.execute(table_create_query)
        self.cursor = self.connection.cursor()
        
    def export_unrecorded_to_excel(self, current_table_name):
        try:
            # Fetch unrecorded orders from the table for specific columns
            select_query = f"SELECT time, order_number, status FROM {current_table_name} WHERE record = ?"
            unrecorded_data = self.cursor.execute(select_query, ('unrecorded',)).fetchall()

            if unrecorded_data:
                new_data_df = pd.DataFrame(unrecorded_data, columns=['時間', '貨單號碼', '狀態'])
                
                 # Check if the file already exists
                file_path = os.path.join(self.save_path, f"{current_table_name}.xlsx")
                if os.path.exists(file_path):
                    # If the file exists, append the data to the first sheet
                    try:
                        with pd.ExcelWriter(file_path, mode='a', if_sheet_exists='overlay') as writer:
                            new_data_df.to_excel(writer, sheet_name='Sheet1', index=False, header=False, startrow=writer.sheets['Sheet1'].max_row)
                        print(f"export data to existed file: {current_table_name} success")
                    except PermissionError:
                        print(f"Error: Could not save Excel file. Please close the file and try again.")
                        return DBreturnType.PERMISSION_ERROR
                else:
                    # If the file does not exist, create a new file and write the data to it
                    new_data_df.to_excel(file_path, index=False)
                    print(f"export data to and create new file: {current_table_name} success")
                    
            return DBreturnType.SUCCESS
                    
        except Exception as e:
            print(f"Error in db_operation -> export_unrecorded_to_excel: {e}")
            return DBreturnType.EXPORT_UNRECORDED_ERROR

    def check_previous_records(self, today: str, yesterday: str):
        #first check if 'yesterday' table exist
        self.cursor.execute(f"SELECT name FROM sqlite_master WHERE type='table' AND name='{yesterday}'")
        result = self.cursor.fetchone()
        if result:
            try:
                # Check if there are any unrecorded orders for yesterday
                select_query_yesterday = f"SELECT COUNT(*) FROM {yesterday} WHERE record = ?"
                unrecorded_count_yesterday = self.cursor.execute(select_query_yesterday, ('unrecorded',)).fetchone()[0]
                print(f"find {unrecorded_count_yesterday} unrecorded data in table '{yesterday}'")

                if unrecorded_count_yesterday > 0:
                    result = self.export_unrecorded_to_excel(yesterday)
                    if result == DBreturnType.PERMISSION_ERROR:
                        print(f"'check_precious_record' using 'export_unrecorded_to_excel' in table '{yesterday}' get permission error")
                        return DBreturnType.PERMISSION_ERROR
                    elif result == DBreturnType.EXPORT_UNRECORDED_ERROR:
                        print(f"'check_precious_record' using 'export_unrecorded_to_excel' in table '{yesterday}' get export unrecorded error")
                        return DBreturnType.EXPORT_UNRECORDED_ERROR

                    update_query_yesterday = f"UPDATE {yesterday} SET record = ? WHERE record = ?"
                    self.cursor.execute(update_query_yesterday, ('recorded', 'unrecorded'))
                    self.connection.commit()
                    print(f"successfully update all unrecorded data in table: {yesterday}")
                    
            except Exception as e:
                print(f"Error in db_operation -> check_previous_records (yesterday): {e}")
                return DBreturnType.CHECK_YESTERDAY_ERROR

        try:
            select_query_today = f"SELECT COUNT(*) FROM {today} WHERE record = ?"
            unrecorded_count_today = self.cursor.execute(select_query_today, ('unrecorded',)).fetchone()[0]
            print(f"find {unrecorded_count_today} unrecorded data in table '{today}'")

            if unrecorded_count_today > 0:
                result = self.export_unrecorded_to_excel(today)
                if result == DBreturnType.PERMISSION_ERROR:
                    print(f"'check_precious_record' using 'export_unrecorded_to_excel' in table '{today}' get permission error")
                    return DBreturnType.PERMISSION_ERROR
                elif result == DBreturnType.EXPORT_UNRECORDED_ERROR:
                    print(f"'check_precious_record' using 'export_unrecorded_to_excel' in table '{today}' get export unrecorded error")
                    return DBreturnType.EXPORT_UNRECORDED_ERROR

                update_query_today = f"UPDATE {today} SET record = ? WHERE record = ?"
                self.cursor.execute(update_query_today, ('recorded', 'unrecorded'))
                self.connection.commit()
                print(f"successfully update all unrecorded data in table: {today}")
            
        except Exception as e:
            print(f"Error in db_operation -> check_previous_records (today): {e}")
            return DBreturnType.CHECK_TODAY_ERROR
        
        return DBreturnType.SUCCESS
        
    def insert_data(self, table_id:int, time:str, order_number:str, status:str, save_status:str):
        try:
            data_insert_query = f"INSERT INTO {self.table_name} (table_id, time, order_number, status, save_status, record) VALUES (?, ?, ?, ?, ?, ?)"
            data_insert_tuple = (table_id, time, order_number, status, save_status, "unrecorded")
            self.cursor.execute(data_insert_query, data_insert_tuple)
            self.connection.commit()
            print("Data inserted to database successfully!")
            return True
        except:
            print("'insert_data' in db_operation raises error")
            return False
        
    def search_order(self, order):
        try:
            select_order_data = f"SELECT * FROM {self.table_name} WHERE order_number = ?"
            self.cursor.execute(select_order_data, (order,))
            result = self.cursor.fetchone()
            print(f"search for order: {order} result: {result}")
            if result:
                return result
            else:
                return DBreturnType.ORDER_NOT_FOUND  # Return None if the order is not found
        except Exception as e:
            print(f"Error in db_operation -> search_order: {e}")
            return DBreturnType.ORDER_NOT_FOUND
    
    def fetch_all_unrecorded(self, status):
        if status == "all":
            select_unrecorded_data = f"SELECT * FROM {self.table_name} WHERE record = 'unrecorded'"
            self.cursor.execute(select_unrecorded_data)
            return self.cursor.fetchall()
        else: 
            select_unrecorded_data = f"SELECT * FROM {self.table_name} WHERE record = 'unrecorded' AND status = '{status}'"
            self.cursor.execute(select_unrecorded_data)
            return self.cursor.fetchall()
    
    def delete_data(self, order:str):
        try:
            delete_query = f"DELETE FROM {self.table_name} WHERE order_number = ?"
            self.cursor.execute(delete_query, (order,))
            self.connection.commit()
            print("Data deleted in db_operation successfully!")
            return True
        except:
            print("delete data in db_operation error")
            return False
    
    def close_database(self):
        try:
            # Export unrecorded orders to Excel
            if not self.export_unrecorded_to_excel(self.table_name):
                print("'close_database' calls 'export_unrecorded_to_excel' get close and save error")
                return DBreturnType.CLOSE_AND_SAVE_ERROR
            
            # Update all unrecorded orders to recorded
            update_query = f"UPDATE {self.table_name} SET record = ? WHERE record = ?"
            self.cursor.execute(update_query, ('recorded', 'unrecorded'))
            self.connection.commit()
            print(f"'close_database' successfully update all unrecorded and export to excel {self.table_name}")
        except Exception as e:
            print(f"Error in db_operation -> close_database: {e}")
            return DBreturnType.CLOSE_AND_SAVE_ERROR
        
        self.connection.close()
        print("Database connection closed successfully")
        return DBreturnType.SUCCESS
