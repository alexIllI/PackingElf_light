import sqlite3

class DataBase():
    def __init__(self, db_name:str):
        self.connection = sqlite3.connect('operation\data.db')
        table_create_query = f"CREATE TABLE IF NOT EXISTS {db_name} (time TEXT, order_number TEXT, status TEXT, save_status TEXT)"
        self.connection.execute(table_create_query)
        self.cursor = self.connection.cursor()
        self.db_name = db_name
        
    def insert_data(self, time, order, status, save_status):
        data_insert_query = f"INSERT INTO {self.db_name} (time, order_number, status, save_status) VALUES (?, ?, ?, ?)"
        data_insert_tuple = (time, order, status, save_status)
        self.cursor.execute(data_insert_query, data_insert_tuple)
        self.connection.commit()
        print("insert data")
    
    def fetch_data():
        pass
        
    def delete_data(self, order):
        pass
    
    def save_to_xlsx(self):
        pass
    
    def closeDB(self):
        self.connection.close()
        
db = DataBase("abc")
db.insert_data("123","abc","success", "貨單")