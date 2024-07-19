import os
import sys
import datetime
import mysql.connector
from mysql.connector import Error
import pandas as pd
from enum import Enum
from dotenv import load_dotenv

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
    EXPORT_UNRECORDED_ERROR = "EXPORT_UNRECORDED_ERROR"
    CHECK_PREVIOUS_RECORD_ERROR = "CHECK_PREVIOUS_RECORD_ERROR"
    CHECK_TODAY_ERROR = "CHECK_TODAY_ERROR"
    ORDER_NOT_FOUND = "ORDER_NOT_FOUND"
    SEARCH_ORDER_ERROR = "SEARCH_ORDER_ERROR"
    CLOSE_AND_SAVE_ERROR = "CLOSE_AND_SAVE_ERROR"

class DataBase():
    def __init__(self, save_path:str):
        load_dotenv()
        self.save_path = save_path
        self.database_name = os.getenv('DB_NAME')
        
        # Dynamically name the table based on the current year and month
        current_time = datetime.datetime.now()
        self.current_table_name = current_time.strftime("MyACG_data_%Y_%m")
        last_month_time = current_time - datetime.timedelta(days=current_time.day)
        self.last_month_table = last_month_time.strftime("MyACG_data_%Y_%m")
        
        self.output_excel_name = current_time.strftime("Printed_Order_%Y_%m_%d")
        
        # MySQL connection setup
        self.connection = mysql.connector.connect(
            host = "192.168.1.114",
            user = os.getenv('DB_USER'),
            password = os.getenv('DB_PASSWORD'),
            database=self.database_name
        )
        self.cursor = self.connection.cursor()

        table_create_query = f"""CREATE TABLE IF NOT EXISTS {self.current_table_name} (
            table_id INT AUTO_INCREMENT PRIMARY KEY,
            time TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            order_number TEXT,
            status TEXT,
            invoice TEXT,
            output_excel TEXT,
            record TEXT,
            using_coupon TEXT,
            order_establish_date TEXT
        )"""
        self.cursor.execute(table_create_query)
        self.connection.commit()
        
    def export_unrecorded_to_excel(self, table_name):
        try:
            # Select including 'output_excel' column to get the specific Excel file name
            select_query = f"SELECT time, order_number, status, invoice, output_excel, using_coupon, order_establish_date FROM {table_name} WHERE record = %s"
            print(f"exporting unrecorded data to Excel: {table_name}")
            self.cursor.execute(select_query, ('unrecorded',))
            unrecorded_data = self.cursor.fetchall()

            if unrecorded_data:
                for data_row in unrecorded_data:
                    # Unpack each row into respective columns including 'output_excel'
                    time, order_number, status, invoice, output_excel, using_coupon, order_establish_date = data_row
                    new_data_df = pd.DataFrame([[time, order_number, status, invoice, using_coupon, order_establish_date]],
                                               columns=['時間', '貨單編號', '狀態', '發票號碼', '使用優惠券', '訂單建立日期'])

                    # Create or append to the Excel file specified in 'output_excel'
                    file_path = os.path.join(self.save_path, f"{output_excel}.xlsx")
                    if os.path.exists(file_path):
                        try:
                            with pd.ExcelWriter(file_path, mode='a', if_sheet_exists='overlay') as writer:
                                new_data_df.to_excel(writer, sheet_name='Sheet1', index=False, header=False, startrow=writer.sheets['Sheet1'].max_row)
                            print(f"Data exported to existing file: {output_excel} successfully.")
                        except PermissionError:
                            print("Error: Could not save Excel file. Please close the file and try again.")
                            return output_excel
                    else:
                        new_data_df.to_excel(file_path, index=False)
                        print(f"Data exported to new file: {output_excel} successfully.")
                    
            return DBreturnType.SUCCESS
                    
        except Error as e:
            print(f"Error in db_operation -> export_unrecorded_to_excel: {e}")
            return DBreturnType.EXPORT_UNRECORDED_ERROR

    def check_previous_records(self):
        for table in [self.last_month_table, self.current_table_name]:
            try:
                self.cursor.execute("SELECT COUNT(*) FROM information_schema.tables WHERE table_schema = %s AND table_name = %s", (self.database_name, table))
                if self.cursor.fetchone()[0] == 1:
                    check_record_query = f"SELECT COUNT(*) FROM {table} WHERE record = %s"
                    self.cursor.execute(check_record_query, ('unrecorded',))
                    unrecorded_count = self.cursor.fetchone()[0]
                    print(f"find {unrecorded_count} unrecorded data in table '{table}'")

                    if unrecorded_count > 0:
                        result = self.export_unrecorded_to_excel(table)
                        if result == DBreturnType.SUCCESS:
                            pass
                        elif result == DBreturnType.EXPORT_UNRECORDED_ERROR:
                            print(f"'check_previous_records' using 'export_unrecorded_to_excel' in table '{table}' get export unrecorded error")
                            return DBreturnType.EXPORT_UNRECORDED_ERROR
                        else:
                            print(f"'check_previous_records' using 'export_unrecorded_to_excel' in table '{table}' get permission error")
                            return result

                    update_record_query = f"UPDATE {table} SET record = %s WHERE record = %s"
                    self.cursor.execute(update_record_query, ('recorded', 'unrecorded'))
                    self.connection.commit()
                    print(f"successfully updated all unrecorded data in table: {table}")
                else:
                    print(f"No table named '{table}' found in the database.")

            except Error as e:
                print(f"Error in db_operation -> check_previous_records: {e}")
                return DBreturnType.CHECK_PREVIOUS_RECORD_ERROR
            
        return DBreturnType.SUCCESS

    def insert_data(self, order_number:str, status:str, invoice:str, output_excel:str, using_coupon:str, order_establish_date:str):
        try:
            data_insert_query = f"INSERT INTO {'myacg_data_' + output_excel[14:21]} (order_number, status, invoice, output_excel, record, using_coupon, order_establish_date) VALUES (%s, %s, %s, %s, %s, %s, %s)"
            data_insert_tuple = (order_number, status, invoice, output_excel, "unrecorded", using_coupon, order_establish_date)
            self.cursor.execute(data_insert_query, data_insert_tuple)
            self.connection.commit()
            print("Data inserted to database successfully!")
            return True
        except Exception as e:
            print(f"'insert_data' in db_operation raises error: {e}")
            return False
        
    def search_order(self, order):
        try:
            select_order_data = f"SELECT * FROM {self.current_table_name} WHERE order_number = %s"
            self.cursor.execute(select_order_data, (order,))
            result = self.cursor.fetchone()
            print(f"search for order in current month: {order} result: {result}")
            if result:
                return result
            
            check_table_exists = f"SELECT COUNT(*) FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_SCHEMA = %s AND TABLE_NAME = %s"
            self.cursor.execute(check_table_exists, (self.database_name, self.last_month_table))
            if self.cursor.fetchone()[0] == 1:
                select_order_data = f"SELECT * FROM {self.last_month_table} WHERE order_number = %s"
                self.cursor.execute(select_order_data, (order,))
                result = self.cursor.fetchone()
                print(f"Search for order in last month: {order} result: {result}")
            if result:
                return result
            else:
                return DBreturnType.ORDER_NOT_FOUND  # Return None if the order is not found
        except Exception as e:
            print(f"Error in db_operation -> search_order: {e}")
            return DBreturnType.SEARCH_ORDER_ERROR
        
    def check_repeated(self, order):
        result = self.search_order(order)
        return result
           
    def fetch_table_data(self, status, record, output_excel):
        '''
        status: all, success, close, cancel
        record: unrecorded = False, recorded = True
        output_excel: output Excel name'''
        table_name = "myacg_data_" + output_excel[14:21]
        try:
            if status == "all":
                if record:
                    select_unrecorded_data = f"SELECT table_id, time, order_number, status, invoice, output_excel, record FROM {table_name} WHERE output_excel = %s"
                    params = (output_excel,)
                else:
                    select_unrecorded_data = f"SELECT table_id, time, order_number, status, invoice, output_excel, record FROM {table_name} WHERE record = %s AND output_excel = %s"
                    params = ("unrecorded", output_excel)
            else:
                if record:
                    select_unrecorded_data = f"SELECT table_id, time, order_number, status, invoice, output_excel, record FROM {table_name} WHERE status = %s AND output_excel = %s"
                    params = (status, output_excel)
                else:
                    select_unrecorded_data = f"SELECT table_id, time, order_number, status, invoice, output_excel, record FROM {table_name} WHERE record = %s AND status = %s AND output_excel = %s"
                    params = ("unrecorded", status, output_excel)

            # Execute query
            self.cursor.execute(select_unrecorded_data, params)
            results = self.cursor.fetchall()

            # Format time in Python
            formatted_results = []
            for result in results:
                formatted_result = list(result)
                formatted_result[1] = result[1].strftime('%H:%M:%S')  # Reformat the time field
                formatted_results.append(formatted_result)

            return formatted_results
        except mysql.connector.Error as err:
            print(f"Error in db_operation -> fetch_table_data: {err}")
            return []

    def delete_data(self, order:str):
        try:
            delete_query = f"DELETE FROM {self.current_table_name} WHERE order_number = %s"
            self.cursor.execute(delete_query, (order,))
            self.connection.commit()
            print("Data deleted in db_operation successfully!")
            return True
        except Exception as e:
            print(f"delete data in db_operation error: {e}")
            return False

    def count_records(self, output_excel, show_recorded):
        try:
            if show_recorded:
            # Query to count all records in the table
                query_all = f"SELECT COUNT(*) FROM {self.current_table_name} WHERE output_excel = %s"
                self.cursor.execute(query_all, (output_excel,))
                total_count = self.cursor.fetchone()[0]  # Fetch the result of the COUNT(*)
                
                # Query to count records where status is 'success'
                query_success = f"SELECT COUNT(*) FROM {self.current_table_name} WHERE status = %s AND output_excel = %s"
                self.cursor.execute(query_success, ("success", output_excel))
                success_count = self.cursor.fetchone()[0]
            else:
                query_all = f"SELECT COUNT(*) FROM {self.current_table_name} WHERE record = %s AND output_excel = %s"
                self.cursor.execute(query_all, ("unrecorded", output_excel,))
                total_count = self.cursor.fetchone()[0]  # Fetch the result of the COUNT(*)
                
                # Query to count records where status is 'success'
                query_success = f"SELECT COUNT(*) FROM {self.current_table_name} WHERE record = %s AND status = %s AND output_excel = %s"
                self.cursor.execute(query_success, ("unrecorded", "success", output_excel))
                success_count = self.cursor.fetchone()[0]
                
            return (total_count, success_count)
        
        except mysql.connector.Error as err:
            print(f"Error in counting records in {self.current_table_name}: {err}")
            return (-1, -1)
    
    def table_exists(self, table_name):
        try:
            self.cursor.execute("SELECT COUNT(*) FROM information_schema.tables WHERE table_schema = %s AND table_name = %s", (self.database_name, table_name))
            return self.cursor.fetchone()[0] == 1
        except mysql.connector.Error as e:
            print(f"Error checking if table {table_name} exists: {e}")
            return False
    
    def get_output_excel_options(self):
        output_excel_options = []
        for table_name in [self.current_table_name, self.last_month_table]:
            if self.table_exists(table_name):
                try:
                    query = f"SELECT DISTINCT output_excel FROM `{table_name}` WHERE output_excel IS NOT NULL"
                    self.cursor.execute(query)
                    result = self.cursor.fetchall()
                    output_excel_options.extend([item[0] for item in result if item[0]])
                except mysql.connector.Error as e:
                    print(f"Error fetching output_excel options from {table_name}: {e}")
            else:
                print(f"Table {table_name} does not exist.")
                
        if self.output_excel_name not in output_excel_options:
            output_excel_options.insert(0, self.output_excel_name)
        else:
            output_excel_options.remove(self.output_excel_name)
            output_excel_options.insert(0, self.output_excel_name)
            
        return output_excel_options

            
    def close_database(self):
        try:
            # Export unrecorded orders to Excel
            export_result = self.export_unrecorded_to_excel(self.current_table_name)
            if export_result != DBreturnType.SUCCESS:
                print(f"'close_database' calls 'export_unrecorded_to_excel' for current month get error: {export_result}")
                return export_result
            
            check_table_exists = f"SELECT COUNT(*) FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_SCHEMA = %s AND TABLE_NAME = %s"
            self.cursor.execute(check_table_exists, (self.database_name, self.last_month_table))
            if self.cursor.fetchone()[0] == 1:
                export_result = self.export_unrecorded_to_excel(self.last_month_table)
                if export_result != DBreturnType.SUCCESS:
                    print(f"'close_database' calls 'export_unrecorded_to_excel' for last month get error: {export_result}")
                    return export_result

            # Update all unrecorded orders to recorded
            update_query = f"UPDATE {self.current_table_name} SET record = %s WHERE record = %s"
            self.cursor.execute(update_query, ('recorded', 'unrecorded'))
            self.connection.commit()
            print("'close_database' successfully updated all unrecorded and exported to excel")
        except Exception as e:
            print(f"Error in db_operation -> close_database: {e}")
            return DBreturnType.CLOSE_AND_SAVE_ERROR
            
        self.connection.close()
        print("Database connection closed successfully")
        return DBreturnType.SUCCESS
