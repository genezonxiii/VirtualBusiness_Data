# -*-  coding: utf-8  -*-
__author__ = '10409003'
import sys
# import MySQLdb
from aes_data import aes_data
from ToMysql import ToMysql
class Query_customer():
    def __init__(self):
        pass

    def GetDataContent(self,GroupID):
        ClientData = []
        try:
            mysqlconnect = ToMysql()
            mysqlconnect.connect()
            # mysql = MySQLdb.connect(host=setting.host,user=setting.user, passwd=setting.passwd, db=setting.db)
            mysql_cursor = mysqlconnect.cursor
            GroupId=GroupID
            mysql_cursor.callproc("sp_selectall_customer", [GroupId,])
            for result in mysql_cursor.stored_results():
                Data=result.fetchall()
            aes=aes_data()
            for results in Data:

                if results[2] != None:
                    results2 = aes.AESdecrypt("p@ssw0rd", results[2], True)
                else:
                    results2 = ""

                if results[3] != None:
                    results3 = aes.AESdecrypt("p@ssw0rd", results[3], True)
                else:
                    results3 = ""

                if results[4] != None:
                    results4 = aes.AESdecrypt("p@ssw0rd", results[4], True)
                else:
                    results4 = ""

                if results[5] != None:
                    results5 = aes.AESdecrypt("p@ssw0rd", results[5], True)
                else:
                    results5 = ""

                r = {"customer_id": results[0],
                     "group_id": results[1],
                     "name": results2,
                     "address": results3,
                     "email": results[6],
                     "post": results[7],
                     "class": results[8],
                     "memo": results[9],
                     "phone": results4,
                     "mobile": results5
                     }

                # print results[4]
                ClientData.append(r)
            return ClientData
        except Exception as e:
            return "MySQL Error %d:  %s" % (e.args[0], e.args[1])
