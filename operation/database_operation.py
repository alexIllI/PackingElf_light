import os
import sys
import datetime
import mysql.connector
from mysql.connector import Error
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
    SEARCH_ORDER_ERROR = "SEARCH_ORDER_ERROR"
    CLOSE_AND_SAVE_ERROR = "CLOSE_AND_SAVE_ERROR"

class DataBase():
    def __init__(self, table_name:str, save_path:str):
        self.table_name = table_name
        self.save_path = save_path
        self.database_name = 'MyACG_data'
        
        # MySQL connection setup
        self.connection = mysql.connector.connect(
            host='localhost',
            database = self.database_name,
            user='root',
            password='Meridian0723'
        )
        self.cursor = self.connection.cursor()
        table_create_query = f"""CREATE TABLE IF NOT EXISTS {self.table_name} (
            table_id INT AUTO_INCREMENT PRIMARY KEY,
            time TEXT,
            order_number TEXT,
            status TEXT,
            invoice TEXT,
            save_status TEXT,
            record TEXT
        )"""
        self.cursor.execute(table_create_query)
        self.connection.commit()
        
    def export_unrecorded_to_excel(self, current_table_name):
        try:
            select_query = f"SELECT time, order_number, status, invoice FROM {current_table_name} WHERE record = %s"
            self.cursor.execute(select_query, ('unrecorded',))
            unrecorded_data = self.cursor.fetchall()

            if unrecorded_data:
                new_data_df = pd.DataFrame(unrecorded_data, columns=['Time', 'Order Number', 'Status', 'Invoice Number'])
                
                file_path = os.path.join(self.save_path, f"{current_table_name}.xlsx")
                if os.path.exists(file_path):
                    try:
                        with pd.ExcelWriter(file_path, mode='a', if_sheet_exists='overlay') as writer:
                            new_data_df.to_excel(writer, sheet_name='Sheet1', index=False, header=False, startrow=writer.sheets['Sheet1'].max_row)
                        print(f"export data to existed file: {current_table_name} success")
                    except PermissionError:
                        print("Error: Could not save Excel file. Please close the file and try again.")
                        return DBreturnType.PERMISSION_ERROR
                else:
                    new_data_df.to_excel(file_path, index=False)
                    print(f"export data to and create new file: {current_table_name} success")
                    
            return DBreturnType.SUCCESS
                    
        except Error as e:
            print(f"Error in db_operation -> export_unrecorded_to_excel: {e}")
            return DBreturnType.EXPORT_UNRECORDED_ERROR

    def check_previous_records(self, today: str, yesterday: str):
        try:
            # Check if the 'yesterday' table exists and has unrecorded entries
            self.cursor.execute("SELECT COUNT(*) FROM information_schema.tables WHERE table_schema = %s AND table_name = %s", (self.database_name, yesterday))
            if self.cursor.fetchone()[0] == 1:
                check_record = f"SELECT COUNT(*) FROM {yesterday} WHERE record = %s"
                self.cursor.execute(check_record, ('unrecorded',))
                unrecorded_count_yesterday = self.cursor.fetchone()[0]
                print(f"find {unrecorded_count_yesterday} unrecorded data in table '{yesterday}'")

                if unrecorded_count_yesterday > 0:
                    result = self.export_unrecorded_to_excel(yesterday)
                    if result == DBreturnType.PERMISSION_ERROR:
                        print(f"'check_previous_records' using 'export_unrecorded_to_excel' in table '{yesterday}' get permission error")
                        return DBreturnType.PERMISSION_ERROR
                    elif result == DBreturnType.EXPORT_UNRECORDED_ERROR:
                        print(f"'check_previous_records' using 'export_unrecorded_to_excel' in table '{yesterday}' get export unrecorded error")
                        return DBreturnType.EXPORT_UNRECORDED_ERROR

                    update_yesterday_unrecorded = f"UPDATE {yesterday} SET record = %s WHERE record = %s"
                    self.cursor.execute(update_yesterday_unrecorded, ('recorded', 'unrecorded'))
                    self.connection.commit()
                    print(f"successfully updated all unrecorded data in table: {yesterday}")
            else:
                print(f"No table named '{yesterday}' found in the database.")

            # Repeat the check for 'today' table
            self.cursor.execute("SELECT COUNT(*) FROM information_schema.tables WHERE table_schema = %s AND table_name = %s", (self.database_name, today))
            if self.cursor.fetchone()[0] == 1:
                self.cursor.execute(f"SELECT COUNT(*) FROM {today} WHERE record = %s", ('unrecorded',))
                unrecorded_count_today = self.cursor.fetchone()[0]
                print(f"find {unrecorded_count_today} unrecorded data in table '{today}'")

                if unrecorded_count_today > 0:
                    result = self.export_unrecorded_to_excel(today)
                    if result == DBreturnType.PERMISSION_ERROR:
                        print(f"'check_previous_records' using 'export_unrecorded_to_excel' in table '{today}' get permission error")
                        return DBreturnType.PERMISSION_ERROR
                    elif result == DBreturnType.EXPORT_UNRECORDED_ERROR:
                        print(f"'check_previous_records' using 'export_unrecorded_to_excel' in table '{today}' get export unrecorded error")
                        return DBreturnType.EXPORT_UNRECORDED_ERROR

                    update_today_unrecorded = f"UPDATE {today} SET record = %s WHERE record = %s"
                    self.cursor.execute(update_today_unrecorded, ('recorded', 'unrecorded'))
                    self.connection.commit()
                    print(f"successfully updated all unrecorded data in table: {today}")
            else:
                print(f"No table named '{today}' found in the database.")

            return DBreturnType.SUCCESS

        except Error as e:
            print(f"Error in db_operation -> check_previous_records: {e}")
            return DBreturnType.CHECK_YESTERDAY_ERROR  # Or other appropriate error

    def insert_data(self, time:str, order_number:str, status:str, invoice:str, save_status:str):
        try:
            data_insert_query = f"INSERT INTO {self.table_name} (time, order_number, status, invoice, save_status, record) VALUES (%s, %s, %s, %s, %s, %s)"
            data_insert_tuple = (time, order_number, status, invoice, save_status, "unrecorded")
            self.cursor.execute(data_insert_query, data_insert_tuple)
            self.connection.commit()
            print("Data inserted to database successfully!")
            return True
        except Exception as e:
            print(f"'insert_data' in db_operation raises error: {e}")
            return False
        
    def search_order(self, order):
        try:
            select_order_data = f"SELECT * FROM {self.table_name} WHERE order_number = %s"
            self.cursor.execute(select_order_data, (order,))
            result = self.cursor.fetchone()
            print(f"search for order: {order} result: {result}")
            if result:
                return result
            else:
                return DBreturnType.ORDER_NOT_FOUND  # Return None if the order is not found
        except Exception as e:
            print(f"Error in db_operation -> search_order: {e}")
            return DBreturnType.SEARCH_ORDER_ERROR
        
    def fetch_table_data(self, status, record, table_name):
        if status == "all":
            select_unrecorded_data = f"SELECT * FROM {table_name} WHERE record = %s"
            self.cursor.execute(select_unrecorded_data, (record,))
        else:
            select_unrecorded_data = f"SELECT * FROM {table_name} WHERE record = %s AND status = %s"
            self.cursor.execute(select_unrecorded_data, (record, status))
        return self.cursor.fetchall()

    def delete_data(self, order:str):
        try:
            delete_query = f"DELETE FROM {self.table_name} WHERE order_number = %s"
            self.cursor.execute(delete_query, (order,))
            self.connection.commit()
            print("Data deleted in db_operation successfully!")
            return True
        except Exception as e:
            print(f"delete data in db_operation error: {e}")
            return False

    def count_records(self, table_name, show_recorded):
        record = 'unrecorded'
        try:
            if show_recorded:
                record = 'recorded'
            # Query to count all records in the table
            query_all = f"SELECT COUNT(*) FROM {table_name} WHERE record = %s"
            self.cursor.execute(query_all, (record,))
            total_count = self.cursor.fetchone()[0]  # Fetch the result of the COUNT(*)
            
            # Query to count records where status is 'success'
            query_success = f"SELECT COUNT(*) FROM {table_name} WHERE record = %s AND status = %s"
            self.cursor.execute(query_success, (record, "success"))
            success_count = self.cursor.fetchone()[0]
            
            return (total_count, success_count)
        except mysql.connector.Error as err:
            print(f"Error in counting records in {table_name}: {err}")
            return (-1, -1)
    
    def get_recent_tables(self):    # currently only 5 tables 
        try:
            query = """
            SELECT table_name 
            FROM information_schema.TABLES 
            WHERE table_schema = %s
            ORDER BY create_time DESC
            LIMIT 5
            """
            self.cursor.execute(query, (self.database_name,))  # Assuming self.database_name holds the name of your database
            result = self.cursor.fetchall()
            # This will return a list of tuples, so let's extract the table names
            return [table[0] for table in result]
        except Exception as e:
            print(f"Error fetching recent tables: {e}")
            return []
            
    def close_database(self):
        try:
            # Export unrecorded orders to Excel
            export_result = self.export_unrecorded_to_excel(self.table_name)
            if export_result != DBreturnType.SUCCESS:
                print(f"'close_database' calls 'export_unrecorded_to_excel' get error: {export_result}")
                return export_result

            # Update all unrecorded orders to recorded
            update_query = f"UPDATE {self.table_name} SET record = %s WHERE record = %s"
            self.cursor.execute(update_query, ('recorded', 'unrecorded'))
            self.connection.commit()
            print(f"'close_database' successfully updated all unrecorded and exported to excel {self.table_name}")
        except Exception as e:
            print(f"Error in db_operation -> close_database: {e}")
            return DBreturnType.CLOSE_AND_SAVE_ERROR
            
        self.connection.close()
        print("Database connection closed successfully")
        return DBreturnType.SUCCESS
