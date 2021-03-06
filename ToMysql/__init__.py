# -*-  coding: utf-8  -*-
from mysql.connector.cursor import MySQLCursor

__author__ = '10409003'
import mysql.connector
from mysql.connector import errorcode

class ToMysql():
    host = '192.168.112.164'
    # host = 'localhost'
    #host = 'www.a-ber.com.tw'
    database = 'db_virtualbusiness'
    user = 'root'
    password = 'admin123'
    #host = '192.168.112.175'
    #database = 'db_virutalbusiness'
    #user = 'root'
    # password = 'mysql'

    def __init__(self):
        pass

    def connect(self):
        try:
            self.db = mysql.connector.connect(user=self.user, password=self.password,
                              host=self.host,database=self.database)
            self.db.commit()
        except mysql.connector.Error as err:
            if err.errno == errorcode.ER_ACCESS_DENIED_ERROR:
                print("Something is wrong with your user name or password")
            elif err.errno == errorcode.ER_BAD_DB_ERROR:
                print("Database does not exist")
            else:
                print(err)
        self.cursor= MySQLCursor(self.db)

    def setDatabase(self,dbname):
        self.database=dbname
        pass

    def setTable(self,tbname):
        self.tablename=tbname
        pass

    def setuser(self,username,pw):
        self.user=username
        self.password=pw
        pass

    def dbClose(self):
        self.cursor.close()
