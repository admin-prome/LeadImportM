import os
import traceback
import pyodbc
from dotenv import load_dotenv

class DatabaseConnection:
    def __init__(self):
        load_dotenv()
        self.sql_user = os.getenv("USER_SQL_SERVER")
        self.sql_pass = os.getenv("PASSWORD_SQL_SERVER")
        self.sql_db = os.getenv("SQL_DB")
        self.sql_server = os.getenv("SQL_SERVER_HOST")
             
    def connect(self):
        try:
            conexion = pyodbc.connect(f'DRIVER={{ODBC Driver 17 for SQL Server}};SERVER={self.sql_server};DATABASE={self.sql_db};UID={self.sql_user};PWD={self.sql_pass}')
            # print(f'\nConexion exitosa a la Base de datos: {self.sql_server}')
            return conexion
        except Exception as e:
            print(f'Verifique si tiene conectada la VPN')
            print(f'Error al intentar conectarse a la base de "{self.sql_server}"')
            traceback.print_exc()