import sqlite3
import pandas as pd
from openpyxl import load_workbook
from enum import Enum

#====================== Enum ===============================
class DBreturnType(Enum):
    SUCCESS = "SUCCESS"
    PERMISSION_ERROR = "PERMISSION_ERROR"
    CHECK_YESTERDAY_ERROR = "CHECK_YESTERDAY_ERROR"
    CHECK_TODAY_ERROR = "CHECK_TODAY_ERROR"
    ORDER_NOT_FOUND = "ORDER_NOT_FOUND"
    CLOSE_AND_SAVE_ERROR = "CLOSE_AND_SAVE_ERROR"

class DataBase():
    def __init__(self, db_name:str, database_path:str ,save_path:str):
        #================= Date for DBname ========================
        self.db_name = db_name
        self.save_path = save_path
        
        self.connection = sqlite3.connect(database_path)
        table_create_query = f"CREATE TABLE IF NOT EXISTS {self.db_name} (table_id INT, time TEXT, order_number TEXT, status TEXT, save_status TEXT, record TEXT)"
        self.connection.execute(table_create_query)
        self.cursor = self.connection.cursor()
        
    def export_unrecorded_to_excel(self, table_name):
        try:
            # Fetch unrecorded orders from the table for specific columns
            select_query = f"SELECT time, order_number, status FROM {table_name} WHERE record = ?"
            unrecorded_data = self.cursor.execute(select_query, ('unrecorded',)).fetchall()

            if unrecorded_data:
                # Convert the data to a DataFrame
                new_data_df = pd.DataFrame(unrecorded_data, columns=['time', 'order_number', 'status'])
                excel_file_path = f"{self.save_path}\\{table_name}.xlsx"
                try:
                    existing_data_df = pd.read_excel(excel_file_path)
                except FileNotFoundError:
                    existing_data_df = pd.DataFrame()

                # Concatenate existing data with new data
                combined_df = pd.concat([existing_data_df, new_data_df], ignore_index=True)

                # Export combined DataFrame to Excel with openpyxl engine
                with pd.ExcelWriter(excel_file_path, engine='openpyxl') as writer:
                    try:
                        combined_df.to_excel(writer, index=False)
                        return True
                    except PermissionError:
                        # Handle the case where the file is open
                        print(f"Error: Could not save Excel file. Please close the file and try again.")
                        return False
            return True
                    
        except Exception as e:
            print(f"Error in db_operation -> export_unrecorded_to_excel: {e}")

    def check_previous_records(self, today:str, yesterday:str):
        try:
            # Export unrecorded orders to Excel for yesterday
            if not self.export_unrecorded_to_excel(yesterday):
                return DBreturnType.PERMISSION_ERROR
            
            # Update all unrecorded orders to recorded for yesterday
            update_query = f"UPDATE {yesterday} SET record = ? WHERE record = ?"
            self.cursor.execute(update_query, ('recorded', 'unrecorded'))
            self.connection.commit()
        except Exception as e:
            print(f"Error in db_operation -> check_previous_records (yesterday): {e}")
            return DBreturnType.CHECK_YESTERDAY_ERROR

        try:
            # Export unrecorded orders to Excel for today
            if not self.export_unrecorded_to_excel(today):
                return DBreturnType.PERMISSION_ERROR
            
            # Update all unrecorded orders to recorded for today
            update_query = f"UPDATE {today} SET record = ? WHERE record = ?"
            self.cursor.execute(update_query, ('recorded', 'unrecorded'))
            self.connection.commit()
        except Exception as e:
            print(f"Error in db_operation -> check_previous_records (today): {e}")
            return DBreturnType.CHECK_TODAY_ERROR
        
    def insert_data(self, table_id:int, time:str, order_number:str, status:str, save_status:str):
        try:
            data_insert_query = f"INSERT INTO {self.db_name} (table_id, time, order_number, status, save_status, record) VALUES (?, ?, ?, ?, ?, ?)"
            data_insert_tuple = (table_id, time, order_number, status, save_status, "unrecorded")
            self.cursor.execute(data_insert_query, data_insert_tuple)
            self.connection.commit()
            print("Data inserted successfully!")
            return True
        except:
            print("insert data error")
            return False
        
    def search_order(self, order):
        try:
            select_order_data = f"SELECT * FROM {self.db_name} WHERE order_number = ?"
            self.cursor.execute(select_order_data, (order,))
            result = self.cursor.fetchone()
            print(result)
            if result:
                return result
            else:
                return DBreturnType.ORDER_NOT_FOUND  # Return None if the order is not found
        except Exception as e:
            print(f"Error in db_operation -> search_order: {e}")
            return DBreturnType.ORDER_NOT_FOUND
    
    def fetch_all_unrecorded(self, status):
        if status == "all":
            select_unrecorded_data = f"SELECT * FROM {self.db_name} WHERE record = 'unrecorded'"
            self.cursor.execute(select_unrecorded_data)
            return self.cursor.fetchall()
        else: 
            select_unrecorded_data = f"SELECT * FROM {self.db_name} WHERE record = 'unrecorded' AND status = '{status}'"
            self.cursor.execute(select_unrecorded_data)
            return self.cursor.fetchall()
    
    def delete_data(self, order:str):
        try:
            delete_query = f"DELETE FROM {self.db_name} WHERE order_number = ?"
            self.cursor.execute(delete_query, (order,))
            self.connection.commit()
            print("Data deleted successfully!")
            return True
        except:
            print("delete data error")
            return False
    
    def close_database(self):
        try:
            # Export unrecorded orders to Excel
            if not self.export_unrecorded_to_excel(f"{self.db_name}.xlsx"):
                return DBreturnType.CLOSE_AND_SAVE_ERROR
            
            # Update all unrecorded orders to recorded
            update_query = f"UPDATE {self.db_name} SET record = ? WHERE record = ?"
            self.cursor.execute(update_query, ('recorded', 'unrecorded'))
            self.connection.commit()
        except Exception as e:
            print(f"Error in db_operation -> close_database: {e}")
            return DBreturnType.CLOSE_AND_SAVE_ERROR
        
        self.connection.close()
        print("Database connection closed successfully")
        return DBreturnType.SUCCESS
