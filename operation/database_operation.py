import sqlite3

class DataBase():
    def __init__(self, db_name:str, save_path:str):
        #================= Date for DBname ========================
        
        self.db_name = db_name
        self.connection = sqlite3.connect('operation\data.db')
        table_create_query = f"CREATE TABLE IF NOT EXISTS {self.db_name} (table_id INT, time TEXT, order_number TEXT, status TEXT, save_status TEXT, record TEXT)"
        self.connection.execute(table_create_query)
        self.cursor = self.connection.cursor()
        self.save_path = save_path
        
    def check(self, today:str, yesterday:str):
        #check if there is any unrecorded order, if there were, than export all row with unrecorded order to excel
        try:
            # Query the database for unrecorded orders
            # select_unrecorded_data = f"select * from {yesterday} where record = 'unrecorded'"

            # df = pd.read_sql_query(select_unrecorded_data, self.connection)
            # # Export the results to an Excel file
            # df.to_excel(self.save_path, index=False)
            
            update_query = f"UPDATE {yesterday} SET record = 'recorded' WHERE record = 'unrecorded'"
            self.cursor.execute(update_query)
            self.connection.commit()
        except Exception as e:
            print(f"Error in db_operation -> check (yesterday): {e}")
            return
        
        try:
            update_query = f"UPDATE {today} SET record = 'recorded' WHERE record = 'unrecorded'"
            self.cursor.execute(update_query)
            self.connection.commit()
        except Exception as e:
            print(f"Error in db_operation -> check (today): {e}")
            return
    
        
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
                return (result[0], result[5])
            else:
                return None  # Return None if the order is not found
        except Exception as e:
            print(f"Error in db_operation -> search_order: {e}")
            return None
    
    def fetch_all_unrecorded(self, status):
        if status == "all":
            try:
                select_unrecorded_data = f"SELECT * FROM {self.db_name} WHERE record = 'unrecorded'"
                self.cursor.execute(select_unrecorded_data)
                return self.cursor.fetchall()
            except:
                return []
            
        try:
            select_unrecorded_data = f"SELECT * FROM {self.db_name} WHERE record = 'unrecorded' AND status = {status}"
            self.cursor.execute(select_unrecorded_data)
            return self.cursor.fetchall()
        except:
            return []
        
    def fetch_last_row(self):
        try:
            fetch_query = f"SELECT * FROM {self.db_name} WHERE ROWID = (SELECT MAX(ROWID) FROM {self.db_name});"
            self.cursor.execute(fetch_query)
            return self.cursor.fetchone()
        except:
            return
    
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
    
    # def save_to_xlsx(self):
    #     try:
    #         df = pd.read_sql_query(f"SELECT * from {self.db_name}", self.connection)
    #         df.to_excel(self.save_path, index=False)
    #         print("Data saved to Excel file successfully!")
    #         return True
    #     except:
    #         print("save to excel error")
    #         return False
    
    def closeDB(self):
        
        #save to excel and close db
        # try:
        #     df = pd.read_sql_query(f"SELECT * from {self.db_name}", self.connection)
        #     df.to_excel(self.save_path, index=False)
        #     print("Data saved to Excel file successfully!")
        # except:
        #     print("save to excel error")
        #     return False
        
        self.connection.close()
        print("Database connection closed.")
