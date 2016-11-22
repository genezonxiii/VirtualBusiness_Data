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
            #mysql = MySQLdb.connect(host=setting.host,user=setting.user, passwd=setting.passwd, db=setting.db)
            mysql_cursor = mysqlconnect.cursor
            GroupId=GroupID
            mysql_cursor.callproc("sp_selectall_customer", [GroupId,])
            for result in mysql_cursor.stored_results():
                Data=result.fetchall()
            aes=aes_data()
            for results in Data:
                if results[2]!=None and results[4]==None and results[3]!=None and results[5]!=None:
                    r={"customer_id" : results[0],"group_id" : results[1],"name" :aes.AESdecrypt("p@ssw0rd",results[2], True),
                       "email" : results[6],"post" : results[7],
                       "class" :results[8],"memo" : results[9],"address" :aes.AESdecrypt("p@ssw0rd",results[3], True),
                       "mobile" :aes.AESdecrypt("p@ssw0rd",results[5], True)}
                elif results[2]!=None and results[4]==None and results[3]==None and results[5]!=None:
                    r={"customer_id" : results[0],"group_id" : results[1],"name" :aes.AESdecrypt("p@ssw0rd",results[2], True),
                       "email" : results[6],"post" : results[7],
                       "class" :results[8],"memo" : results[9],
                       "mobile" :aes.AESdecrypt("p@ssw0rd",results[5], True)}
                elif results[2]!=None and results[4]!=None and results[3]==None and results[5]!=None:
                    r={"customer_id" : results[0],"group_id" : results[1],"name" :aes.AESdecrypt("p@ssw0rd",results[2], True),
                       "email" : results[6],"post" : results[7],
                       "class" :results[8],"memo" : results[9],
                       "phone": aes.AESdecrypt("p@ssw0rd", results[4], True),"mobile" :aes.AESdecrypt("p@ssw0rd",results[5], True)}
                elif results[2]!=None and results[4]!=None and results[3]!=None and results[5]!=None:
                    r={"customer_id" : results[0],"group_id" : results[1],"name" :aes.AESdecrypt("p@ssw0rd",results[2], True),
                       "address" :aes.AESdecrypt("p@ssw0rd",results[3], True),"email" : results[6],"post" : results[7],
                       "class" :results[8],"memo" : results[9],
                       "phone": aes.AESdecrypt("p@ssw0rd", results[4], True),"mobile" :aes.AESdecrypt("p@ssw0rd",results[5], True)}
                elif results[2]==None and results[4]==None and results[3]==None and results[5]==None:
                    r={"customer_id" : results[0],"group_id" : results[1],"email" : results[6],"post" : results[7],
                       "class" :results[8],"memo" : results[9]}
                elif results[2]!=None and results[4]==None and results[3]==None and results[5]==None:
                    r={"customer_id" : results[0],"group_id" : results[1],"name" :aes.AESdecrypt("p@ssw0rd",results[2], True),
                       "email" : results[6],"post" : results[7],"class" :results[8],"memo" : results[9]}
                elif results[2]!=None and results[4]==None and results[3]!=None and results[5]==None:
                    r={"customer_id" : results[0],"group_id" : results[1],"name" :aes.AESdecrypt("p@ssw0rd",results[2], True),
                       "email" : results[6],"post" : results[7],
                       "class" :results[8],"memo" : results[9],
                       "address" :aes.AESdecrypt("p@ssw0rd",results[3], True)}
                

                # print results[4]
                ClientData.append(r)
            return ClientData
        except MySQLdb.Error, e:
            return "MySQL Error %d:  %s" % (e.args[0], e.args[1])
