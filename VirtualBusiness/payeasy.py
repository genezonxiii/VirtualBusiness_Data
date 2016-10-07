# -*-  coding: utf-8  -*-
__author__ = '10409003'
import xlrd
import uuid
import csv
from datetime import datetime
from aes_data import aes_data
from ToMongodb import ToMongodb
from ToMysql import ToMysql
import os
import sys, argparse, csv


class payeasy_Data():
    Data=None
    def __init__(self):
        pass
    def payeasy_Data(self,supplier,GroupID,path,UserID):
        #mysql connector object
        mysqlconnect=ToMysql()
        mysqlconnect.connect()
        print path
        mongoOrder=ToMongodb()
        mongoOrder.setCollection('co_order')
        mongoOrder.connect()
        mongodbClient=ToMongodb()
        mongodbClient.setCollection('co_client')
        mongodbClient.connect()

        Ordernum=""
        Clientnum=""


        # open csv file
        with open(path, 'rb') as f:
            reader = csv.reader(f, delimiter=',')
            ignore=0

            for row in reader:
                if ignore:
                    aes = aes_data()
                    OrderNo = row[6]
                    print OrderNo
                    PartName = row[8].decode("big5")
                    PartNo = row[9]
                    PartPrice = row[12]
                    PartQuility = row[11]
                    PartTotalPrice = row[12]
                    strShipmentDate = row[1]
                    temp_time = map(int,((strShipmentDate.split(' ')[0]).split('/')))#2015/5/23 07:25
                    _ShipmentDate = datetime(*(temp_time[0:3])).strftime('%Y-%m-%d')
                    ShipmentDate = datetime.strptime(_ShipmentDate, '%Y-%m-%d')
                    Name = row[14]
                    ClientName = aes.AESencrypt("p@ssw0rd", Name, True)
                    Phone = row[19]
                    ClientPhone = aes.AESencrypt("p@ssw0rd", Phone, True)
                    ClientZipCode = row[15]
                    Add= row[17]
                    ClientAdd = aes.AESencrypt("p@ssw0rd", Add, True)
                    firm = GroupID
                    supplier = supplier
                    UserID = UserID

                    # SupplySQL = (str(uuid.uuid4()), GroupID, supplier, "", "", "", "", "", "", "", "", "", "", "")
                    ProductSQL = (
                        str(uuid.uuid4()), GroupID, PartNo, PartName, supplier, None, None, 0, PartPrice, 0, None,
                        None,None,None)
                    CustomereSQL = (
                        str(uuid.uuid4()), GroupID, ClientName, ClientAdd,None, ClientPhone, None, ClientZipCode,
                        None,
                        None)
                    mysqlconnect.connect()
                    # mysqlconnect.cursor.callproc('p_tb_supply', SupplySQL)
                    mysqlconnect.cursor.callproc('p_tb_product', ProductSQL)

                    CustomereSQLsel = """select group_id,name from tb_customer where group_id='%s'""" % (GroupID)

                    CustomereSQLupd = """update tb_customer
                                set address= %s ,
                                    phone= %s,
                                    mobile = %s,
                                    email = %s,
                                    post= %s,
                                    class = %s,
                                    memo= %s

                                where group_id=%s and name=%s;"""
                    CustomereSQLins = """ insert into tb_customer
                                                    (customer_id,group_id ,name,address,phone,mobile,email,post,class,memo)
                                                    values(%s,%s,%s,%s,%s,%s,%s,%s,%s,%s);"""
                    mysqlconnect.cursor.execute(CustomereSQLsel)
                    result = mysqlconnect.cursor.fetchall()
                    print 'Name'

                    print result
                    Name_compare = []
                    if result != []:
                        for x in result:
                            Name_compare.append(aes.AESdecrypt("p@ssw0rd", x[1], True))
                        if Name in Name_compare:
                            print "update"
                            mysqlconnect.cursor.execute(CustomereSQLupd, (
                                ClientAdd,None, ClientPhone, None, ClientZipCode,None, None, GroupID, x[1]))
                        else:
                            print "select insert"
                            mysqlconnect.cursor.execute(CustomereSQLins, CustomereSQL)
                    else:
                        mysqlconnect.cursor.execute(CustomereSQLins, CustomereSQL)
                    mysqlconnect.cursor.execute(CustomereSQLsel)
                    result = mysqlconnect.cursor.fetchall()
                    customer_id_temp = []
                    SalestrSQLsel = "SELECT customer_id from tb_customer where name =%s;"
                    for x in result:
                        if aes.AESdecrypt('p@ssw0rd', x[1], True) == Name:
                            mysqlconnect.cursor.execute(SalestrSQLsel, (str(x[1]),))
                            customer_id_temp.append(mysqlconnect.cursor.fetchall()[0])
                    print customer_id_temp
                    for y in customer_id_temp:
                        for x in result:
                            if aes.AESdecrypt('p@ssw0rd', x[1], True) == Name:
                                SaleSQL = (
                                    GroupID, OrderNo, UserID, PartName, PartNo, y[0],
                                    x[1], PartQuility, PartTotalPrice, None, ShipmentDate, ShipmentDate, ShipmentDate,
                                    None,ShipmentDate,supplier)
                                mysqlconnect.cursor.callproc('p_tb_sale', SaleSQL)
                                # SalestrSQL = "INSERT INTO tb_sale (sale_id,seq_no,group_id,user_id,customer_id,name,c_product_id,quantity,price,sale_date,product_name)" \
                                #              "VALUES(%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)"
                                # mysqlconnect.cursor.execute(SalestrSQL, (
                                #     str(uuid.uuid4()), OrderNo, GroupID, '52345', y[0], x[1], PartNo, PartQuility,
                                #     PartTotalPrice, ShipmentDate,PartName))
                    mysqlconnect.db.commit()
                    # mysqlconnect.cursor.callproc('p_tb_customer', CustomereSQL)

                    if (Ordernum == OrderNo[0:13]):
                        print 'update'
                        self.updataOrder(mongoOrder, ShipmentDate,OrderNo, PartName, PartNo, PartPrice, \
                                         PartQuility, PartTotalPrice, \
                                         firm, supplier)

                    else:
                        print 'insert'
                        Ordernum = OrderNo
                        self.insertOrder(mongoOrder, ShipmentDate,OrderNo, PartName, PartNo, PartPrice, \
                                         PartQuility, PartTotalPrice, \
                                         firm, supplier)
                    if (Clientnum == ClientName):
                        print 'update'
                        self.updataClient(mongodbClient, OrderNo, ClientName, ClientPhone, ClientZipCode, ClientAdd, \
                                          firm, supplier)
                    else:
                        print 'insert'
                        Clientnum = ClientName
                        self.insertClient(mongodbClient, OrderNo, ClientName, ClientPhone, ClientZipCode, ClientAdd, \
                                          firm, supplier)
                    mysqlconnect.dbClose()
                    mongoOrder.dbClose()
                    mongodbClient.dbClose()
                else:
                    ignore = 1
        return 'success'






                #put the data into the corresponding variable



    # mongoDB storage   第�??��??�是丟�??��?mongoOrder or mongoClient
    def insertOrder(self,mongoOrder,_ShipmentDate,_OrderNo,_PartName,_PartNo,_PartPrice,\
                                _PartQuility,_PartTotalPrice,\
                                _firm,_supplier):
        businessorder_doc={'ShipmentDate':_ShipmentDate,'OrderNo':_OrderNo,\
                           'PartName':[_PartName],'PartNo':[_PartNo],'Price':[_PartPrice],\
                           'PartQuility':[_PartQuility],'PartTotalPrice':[_PartTotalPrice],'firm':[_firm],'supplier':[_supplier]
                            }

        mongoOrder.cursor.insert(businessorder_doc)
    def updataOrder(self,mongoOrder,_ShipmentDate,_OrderNo,_PartName,_PartNo,_PartPrice,\
                                _PartQuility,_PartTotalPrice,\
                                _firm,_supplier):
        mongoOrder.cursor.update({'ShipmentDate':_ShipmentDate, "OrderNo" : _OrderNo,'firm':_firm,'supplier':_supplier},\
                           {'$push':{'PartName':_PartName,'PartNo':_PartNo ,\
                           'Price':_PartPrice,'PartQuility':_PartQuility,'PartTotalPrice':_PartTotalPrice}}
                                                )
    def insertClient(self,mongodbClient,_OrderNo,_ClientName,_ClientPhone,_ClientZipCode,_ClientAdd,\
                                _firm,_supplier):

        businessorder_doc={'OrderNo':_OrderNo,'ClientName':[_ClientName],'ClientPhone':[_ClientPhone],'ClientZipCode':[_ClientZipCode],'ClientAdd':[_ClientAdd],\
                           'firm':[_firm],'supplier':[_supplier]}

        mongodbClient.cursor.insert(businessorder_doc)
    def updataClient(self,mongodbClient,_OrderNo,_ClientName,_ClientPhone,_ClientZipCode,_ClientAdd,\
                                _firm,_supplier):
        mongodbClient.cursor.update({ 'ClientName':_ClientName,'ClientPhone':[_ClientPhone],'ClientZipCode':_ClientZipCode,'ClientAdd':_ClientAdd,\
                            'firm':_firm,'supplier':_supplier}\
                           ,{'$push':{"OrderNo" : _OrderNo}})






