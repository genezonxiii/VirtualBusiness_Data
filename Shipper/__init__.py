# -*-  coding: utf-8  -*-
__author__ = '10409003'
import sys
# import MySQLdb
from aes_data import aes_data
from ToMysql import ToMysql
from datetime import datetime

class ShipperData():
    def __init__(self):
        pass

    def GetDataContent(self,GroupID,StartDate,EndDate):
        ClientData = []
        try:
            mysqlconnect = ToMysql()
            mysqlconnect.connect()
            #mysql = MySQLdb.connect(host=setting.host,user=setting.user, passwd=setting.passwd, db=setting.db)
            mysql_cursor = mysqlconnect.cursor
            Data=(GroupID,StartDate,EndDate)
            mysql_cursor.callproc("sp_shipper",  Data)
            for result in mysql_cursor.stored_results():
                Data=result.fetchall()
            aes = aes_data()
            # print Data
            for results in Data:
                if results[20]!=None and results[21]!=None and results[22]== None and results[23]!= None:
                    r = {"sale_id": results[0], "seq_no": results[1],"group_id": results[2],"order_no" : results[3],"user_id" : results[4],
                     "product_id" : results[5],"product_name" : results[6],"c_product_id": results[7], "customer_id": results[8],
                     "quantity": results[9], "price": results[10],"invoice": results[11], "invoice_date": results[12],
                     "trans_list_date": results[13], "dis_date": results[14],"memo": results[15], "sale_date": results[16],
                     "order_source": results[17], "return_date": results[18],"isreturn": results[19],
                     "name": aes.AESdecrypt("p@ssw0rd", results[20], True),"address": aes.AESdecrypt("p@ssw0rd", results[21], True),
                     "mobile": aes.AESdecrypt("p@ssw0rd", results[23], True),
                     "post": results[24], "class": results[25],
                     }
                elif results[20]!=None and results[21]==None and results[22]== None and results[23]!= None:
                    r = {"sale_id": results[0], "seq_no": results[1],"group_id": results[2],"order_no" : results[3],"user_id" : results[4],
                         "product_id" : results[5],"product_name" : results[6],"c_product_id": results[7], "customer_id": results[8],
                         "quantity": results[9], "price": results[10],"invoice": results[11], "invoice_date": results[12],
                         "trans_list_date": results[13], "dis_date": results[14],"memo": results[15], "sale_date": results[16],
                         "order_source": results[17], "return_date": results[18],"isreturn": results[19],
                         "name": aes.AESdecrypt("p@ssw0rd", results[20], True),"mobile": aes.AESdecrypt("p@ssw0rd", results[23], True),
                         "post": results[24], "class": results[25],
                         }
                elif results[20]!=None and results[21]==None and results[22]==None and results[23]==None:
                    r = {"sale_id": results[0], "seq_no": results[1],"group_id": results[2],"order_no" : results[3],"user_id" : results[4],
                         "product_id" : results[5],"product_name" : results[6],"c_product_id": results[7], "customer_id": results[8],
                         "quantity": results[9], "price": results[10],"invoice": results[11], "invoice_date": results[12],
                         "trans_list_date": results[13], "dis_date": results[14],"memo": results[15], "sale_date": results[16],
                         "order_source": results[17], "return_date": results[18],"isreturn": results[19],
                         "name": aes.AESdecrypt("p@ssw0rd", results[20], True),"post": results[24], "class": results[25],
                         }
                elif results[20]!=None and results[21]==None and results[22]!=None and results[23]!=None:
                    r = {"sale_id": results[0], "seq_no": results[1],"group_id": results[2],"order_no" : results[3],"user_id" : results[4],
                         "product_id" : results[5],"product_name" : results[6],"c_product_id": results[7], "customer_id": results[8],
                         "quantity": results[9], "price": results[10],"invoice": results[11], "invoice_date": results[12],
                         "trans_list_date": results[13], "dis_date": results[14],"memo": results[15], "sale_date": results[16],
                         "order_source": results[17], "return_date": results[18],"isreturn": results[19],
                         "name": aes.AESdecrypt("p@ssw0rd", results[20], True),
                         "phone": aes.AESdecrypt("p@ssw0rd", results[22], True),"mobile": aes.AESdecrypt("p@ssw0rd", results[23], True),
                         "post": results[24], "class": results[25],
                         }
                elif results[20]!=None and results[21]!=None and results[22]!=None and results[23]!=None:
                    r = {"sale_id": results[0], "seq_no": results[1],"group_id": results[2],"order_no" : results[3],"user_id" : results[4],
                         "product_id" : results[5],"product_name" : results[6],"c_product_id": results[7], "customer_id": results[8],
                         "quantity": results[9], "price": results[10],"invoice": results[11], "invoice_date": results[12],
                         "trans_list_date": results[13], "dis_date": results[14],"memo": results[15], "sale_date": results[16],
                         "order_source": results[17], "return_date": results[18],"isreturn": results[19],
                         "name": aes.AESdecrypt("p@ssw0rd", results[20], True),"address": aes.AESdecrypt("p@ssw0rd", results[21], True),
                         "phone": aes.AESdecrypt("p@ssw0rd", results[22], True),"mobile": aes.AESdecrypt("p@ssw0rd", results[23], True),
                         "post": results[24], "class": results[25],
                         }
                ClientData.append(r)
            return ClientData
        except Exception as e :
            print e.message


class VBsale_Analytics():
    conn = None
    def __init__(self):
        pass

    # 取得 db 的連線
    #def getConnection(self):
    #        try:
    #            if (self.conn == None):
    #                config = Config()
    #                return mysql.connector.connect(user=config.dbUser, password=config.dbPwd,host=config.dbServer, database=config.dbName)
    #            else:
    #                return self.conn
    #        except mysql.connector.Error:
    #            print "Connection DB Error"
    #            raise
    #        except Exception as e:
    #            print e.message
    #            raise

    # 取得訂購人消費 Top 10 資料
    def get_buyer_top10(self,group_id,start_Date,end_date):
        args =(group_id,datetime.strptime(start_Date,'%Y-%m-%d'),datetime.strptime(end_date,'%Y-%m-%d'))
        data = self.getData('sp_buyer_top10',args)
        dncrypt = aes_data()
        result=[]
        for row in data:
            r={"customer_id":row[0],"quantity":float(row[1]),"name":dncrypt.AESdecrypt('p@ssw0rd',row[2],True)}
            result.append(r)
        return result

    # 取得單一通路訂購人消費排名 資料
    def get_buyer_channel(self, group_id, start_Date, end_date,order_source):
        args = (group_id, datetime.strptime(start_Date, '%Y-%m-%d'), datetime.strptime(end_date, '%Y-%m-%d'),order_source)
        data = self.getData('sp_buyer_channel', args)
        dncrypt = aes_data()
        result = []
        for row in data:
            r = {"customer_id": row[0], "quantity": float(row[1]), "name": dncrypt.AESdecrypt('p@ssw0rd', row[2], True)}
            result.append(r)
        return result

    # 從 db 取得 stored procedure 結果
    def getData(self, procedureName,parameter):
            try:
                mysqlconnect = ToMysql()
                mysqlconnect.connect()
                mysql_cursor = mysqlconnect.cursor
                #conn = self.getConnection()
                #cursor = conn.cursor()
                #cursor.callproc(procedureName,parameter)
                mysql_cursor.callproc(procedureName,parameter)
                data_row = []
                for row in mysql_cursor.stored_results():
                    data_row=row.fetchall()
                mysql_cursor.close()
                return data_row
            except Exception as e:
                print e.message
                raise


