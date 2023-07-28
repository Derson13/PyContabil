import pandas as pd

#Database
import pyodbc 
from requests.utils import requote_uri
from sqlalchemy import create_engine

class Database:
    def __init__(self):
        self.strConn        = 'Driver={SQL Server};Server=38.242.238.194;Database=dbContabil;Trusted_Connection=yes;'

    staticmethod    
    def sqlSet(self, df, table, is_increment):
        quoted = requote_uri(self.strConn)
        engine = create_engine(f'mssql+pyodbc:///?odbc_connect={quoted}')
        if is_increment:
            df.to_sql(table, con=engine, if_exists='append', index=False)                   
        else:
            df.to_sql(table, con=engine, if_exists='replace', index=False)