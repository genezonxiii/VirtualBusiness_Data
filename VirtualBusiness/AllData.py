# -*-  coding: utf-8  -*-
__author__ = '10409003'
import xlrd
import uuid
from datetime import datetime
from aes_data import aes_data
from ToMongodb import ToMongodb
from ToMysql import ToMysql
import logging
import time

class ASAP_Data():
    Data=None
    def __init__(self):
        pass
    def ASAP_Data(self,supplier,GroupID,path):
        logging.basicConfig(filename='pyupload.log', level=logging.DEBUG, format='%(asctime)s %(message)s', datefmt='%Y/%m/%d %I:%M:%S %p')
        logging.Formatter.converter = time.gmtime
        logging.info('===ASAP_Data. filename=AllData.py===')
        logging.debug('supplier:' + supplier)
        logging.debug('GroupID:' + GroupID)
        logging.debug('path:' + path)
        logging.debug('UserID:' + UserID)
        
        #mysql connector object
        mysqlconnect=ToMysql()
        mysqlconnect.connect()

        mongoOrder=ToMongodb()
        mongoOrder.setCollection('co_ordertest')
        mongoOrder.connect()
        mongodbClient=ToMongodb()
        mongodbClient.setCollection('co_co_clienttest')
        mongodbClient.connect()

        Ordernum=""
        Clientnum=""

        data=xlrd.open_workbook(path)
        table=data.sheets()[0]
        num_cols=table.ncols
        #put the data into the corresponding variable

        for row_index in range(1,table.nrows):
            for col_index in range(0,num_cols):
                aes=aes_data()
                strTurntDate    =xlrd.xldate_as_tuple(table.cell_value(row_index,0),data.datemode)
                TurntDate       =datetime(*strTurntDate[0:6]).strftime('%Y-%m-%d')
                OrderNo         =table.cell(row_index,1).value[0:13]
                PartNum         =str(table.cell(row_index,2).value)
                PartMaterial    =str(table.cell(row_index,3).value)
                PartName        =table.cell(row_index,4).value
                PartColor       =table.cell(row_index,5).value
                PartSize        =table.cell(row_index,6).value
                PartQuility     =table.cell(row_index,7).value
                Name            =table.cell(row_index, 8).value
                ClientName      = aes.AESencrypt("password",Name.encode('utf-8'), True)
                Phone           = table.cell(row_index, 9).value
                ClientPhone     = aes.AESencrypt("password", Phone.encode('utf-8'), True)
                Tel             = table.cell(row_index, 10).value
                ClientTel       = aes.AESencrypt("password", Tel.encode('utf-8'), True)
                Add             = table.cell(row_index, 11).value
                ClientAdd       = aes.AESencrypt("password", Add.encode('utf-8'), True)
                PartNo          =table.cell(row_index,12).value
                firm            =GroupID
                supplier        =supplier
                GroupID         =GroupID
            SupplySQL = (str(uuid.uuid4()),GroupID, supplier,"","","","","","","","","","","")
            ProductSQL = (str(uuid.uuid4()), GroupID, PartNo, PartName,supplier, "",PartSize,0,0,0,'NULL','NULL','NULL','NULL')
            CustomereSQL = (str(uuid.uuid4()), GroupID, ClientName, ClientAdd, ClientTel, ClientPhone,'NULL','NULL','NULL','NULL')

            mysqlconnect.cursor.callproc('p_tb_supply', SupplySQL)
            mysqlconnect.cursor.callproc('p_tb_product', ProductSQL)

            CustomereSQLsel = """select group_id,name from tb_customer where group_id=%s""" % (GroupID)

            CustomereSQLupd = """update tb_customer
                set address= %s ,
                    phone= %s,
                    mobile = %s,
                    email = %s,
                    post= %s,
                    class = %s,
                    memo= %s

                where group_id=%s and name=%s;"""
            print SupplySQL[0]
            CustomereSQLins = """ insert into tb_customer
                                    (customer_id,group_id ,name,address,phone,mobile,email,post,class,memo)
                                    values(%s,%s,%s,%s,%s,%s,%s,%s,%s,%s);"""
            mysqlconnect.cursor.execute(CustomereSQLsel)
            result = mysqlconnect.cursor.fetchall()

            print result
            Name_compare=[]
            if result!=[]:
                for x in result:
                    Name_compare.append(aes.AESdecrypt("password",x[1], True))
                if Name.encode('utf-8') in Name_compare :
                    print "update"
                    mysqlconnect.cursor.execute(CustomereSQLupd,(ClientAdd, ClientTel, ClientPhone,'NULL','NULL','NULL','NULL',GroupID,x[1]))
                else:
                    print "select insert"
                    mysqlconnect.cursor.execute(CustomereSQLins, CustomereSQL)
            else:
                mysqlconnect.cursor.execute(CustomereSQLins, CustomereSQL)
            mysqlconnect.cursor.execute(CustomereSQLsel)
            result = mysqlconnect.cursor.fetchall()
            customer_id_temp=[]
            SalestrSQLsel="SELECT customer_id from tb_customer where name =%s;"
            x_name=""
            for x in result:
                if aes.AESdecrypt('password',x[1],True)==Name.encode('utf-8'):
                    mysqlconnect.cursor.execute(SalestrSQLsel,(str(x[1]),))
                    customer_id_temp.append(mysqlconnect.cursor.fetchall()[0])
            print customer_id_temp
            for y in customer_id_temp:
                for x in result:
                    if aes.AESdecrypt('password', x[1], True) == Name.encode('utf-8'):
                        SalestrSQL = "INSERT INTO tb_sale (sale_id,seq_no,group_id,user_id,customer_id,name,c_product_id,quantity,trans_list_date,sale_date,product_name)"\
                        "VALUES(%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)"
                        mysqlconnect.cursor.execute(SalestrSQL,(str(uuid.uuid4()), OrderNo, GroupID, '12345', y[0], x[1], PartNo, PartQuility, TurntDate,TurntDate,PartName))
            mysqlconnect.db.commit()
            # mysqlconnect.cursor.callproc('p_tb_customer', CustomereSQL)

            if (Ordernum==OrderNo[0:13]):
                print 'update'
                self.updataOrder(mongoOrder,TurntDate,OrderNo,PartNum,PartMaterial,PartName,PartColor,PartSize,PartQuility,PartNo,firm,supplier)

            else:
                print 'insert'
                Ordernum=OrderNo
                self.insertOrder(mongoOrder,TurntDate,OrderNo,PartNum,PartMaterial,PartName,PartColor,PartSize,PartQuility,PartNo,firm,supplier)
            if (Clientnum == ClientName):
                print 'update'
                self.updataClient(mongodbClient, OrderNo, ClientName, ClientPhone, ClientTel, ClientAdd, firm, supplier)
            else:
                print 'insert'
                Clientnum = ClientName
                self.insertClient(mongodbClient, OrderNo, ClientName, ClientPhone, ClientTel, ClientAdd, firm, supplier)

        mysqlconnect.dbClose()
        mongoOrder.dbClose()
        mongodbClient.dbClose()

    # mongoDB storage   ç¬¬ä??‹å??¸æ˜¯ä¸Ÿä??¢ç?mongoOrder or mongoClient
    def insertOrder(self,mongoOrder,_TurntDate,_OrderNo,_PartNum,_PartMaterial,_PartName,_PartColor,_PartSize,_PartQuility,_PartNo,_firm,_supplier):
        businessorder_doc={ 'TurntDate':_TurntDate,
                            'OrderNo':_OrderNo,
                            'PartNum':[_PartNum],
                            'PartMaterial':[_PartMaterial],\
                            'PartName':[_PartName],
                            'PartColor':[_PartColor],
                            'PartSize':[_PartSize],\
                            'PartQuility':[_PartQuility],
                            'PartNo':[_PartNo],
                            'firm':[_firm],
                            'supplier':[_supplier]
                            }

        mongoOrder.cursor.insert(businessorder_doc)
    def updataOrder(self,mongoOrder,_TurntDate,_OrderNo,_PartNum,_PartMaterial,_PartName,_PartColor,_PartSize,_PartQuility,_PartNo,_firm,_supplier):
        mongoOrder.cursor.update( { "OrderNo" : _OrderNo,
                                    'firm':_firm,'supplier':_supplier},\
                                    {'$push':{  'TurntDate':_TurntDate,
                                                'PartNum':_PartNum ,
                                                'PartMaterial':_PartMaterial,
                                                'PartName':_PartName,\
                                                'PartColor':_PartColor,
                                                'PartSize':_PartSize,
                                                'PartNo':_PartNo,
                                                'PartQuility':_PartQuility}
                                                }
                                                )
    def insertClient(self,mongodbClient,_OrderNo,_ClientName,_ClientPhone,_ClientTel,_ClientAdd,_firm,_supplier):

        businessorder_doc={ 'OrderNo':_OrderNo,
                            'ClientName':[_ClientName],
                            'ClientPhone':[_ClientPhone],
                            'ClientTel':[_ClientTel],
                            'ClientAdd':[_ClientAdd],
                            'firm':[_firm],
                            'supplier':[_supplier]}

        mongodbClient.cursor.insert(businessorder_doc)
    def updataClient(self,mongodbClient,_OrderNo,_ClientName,_ClientPhone,_ClientTel,_ClientAdd,_firm,_supplier):
        mongodbClient.cursor.update( {  'ClientName':_ClientName,
                                        'ClientPhone':_ClientPhone,
                                        'ClientTel':[_ClientTel],
                                        'ClientAdd':_ClientAdd,
                                        'firm':_firm,
                                        'supplier':_supplier}\
                                        ,{'$push':{"OrderNo" : _OrderNo}})

class GoHappy_Data():
    Data=None
    def __init__(self):
        pass
    def GoHappy_Data(self,supplier,GroupID,path):
        #mysql connector object
        mysqlconnect=ToMysql()
        mysqlconnect.connect()

        mongoOrder=ToMongodb()
        mongoOrder.setCollection('co_ordertest')
        mongoOrder.connect()
        mongodbClient=ToMongodb()
        mongodbClient.setCollection('co_co_clienttest')
        mongodbClient.connect()

        Ordernum=""
        Clientnum=""

        data=xlrd.open_workbook(path)
        table=data.sheets()[0]
        num_cols=table.ncols
        #put the data into the corresponding variable

        for row_index in range(1,table.nrows):
            for col_index in range(0,num_cols):
                aes=aes_data()
                firm = GroupID
                firmNo = table.cell(row_index, 1).value
                ShipmentDateNo = table.cell(row_index, 3).value
                ReconciliationDate = table.cell(row_index, 4).value
                OrderNo = str(table.cell(row_index, 6).value[0:13])
                OrderStatus = table.cell(row_index, 7).value
                strOrderDate= xlrd.xldate_as_tuple(table.cell_value(row_index, 8), data.datemode)
                OrderDate = datetime(*strOrderDate[0:6]).strftime('%Y-%m-%d')
                PartName = table.cell(row_index, 9).value
                PartType = str(table.cell(row_index, 10).value)
                PartNo = table.cell(row_index, 11).value
                FormatNo = table.cell(row_index, 12).value
                PartCost = table.cell(row_index, 13).value
                PartPrice = table.cell(row_index, 14).value
                PartQuility = table.cell(row_index, 15).value
                PartTotalPrice = table.cell(row_index, 16).value
                OrderName = table.cell(row_index, 17).value
                Name = table.cell(row_index, 18).value
                ClientName = aes.AESencrypt("password", Name.encode('utf-8'), True)
                ClientZipCode = table.cell(row_index, 19).value
                Add = table.cell(row_index, 20).value
                ClientAdd = aes.AESencrypt("password", Add.encode('utf-8'), True)
                Tel = table.cell(row_index, 21).value
                ClientTel = aes.AESencrypt("password", Tel.encode('utf-8'), True)
                Phone = table.cell(row_index, 22).value
                ClientPhone = aes.AESencrypt("password", Phone.encode('utf-8'), True)
                PartNote = table.cell(row_index, 23).value
                PartShopNo = table.cell(row_index, 24).value
                PartAction = table.cell(row_index, 25).value
                OrderType = table.cell(row_index, 26).value
                strShipmentdate = xlrd.xldate_as_tuple(table.cell_value(row_index, 27), data.datemode)
                ShipmentDate = datetime(*strShipmentdate[0:6]).strftime('%Y-%m-%d')
                PartNum = str(table.cell(row_index, 28).value)
                supplier = supplier
                GroupID = GroupID
            SupplySQL = (str(uuid.uuid4()),GroupID, supplier,"","","","","","","","","","","")
            ProductSQL = (str(uuid.uuid4()), GroupID, PartNo, PartName,supplier, PartType,"",PartCost,PartPrice,0,'NULL','NULL','NULL','NULL')
            CustomereSQL = (str(uuid.uuid4()), GroupID, ClientName, ClientAdd, ClientTel, ClientPhone,'NULL',ClientZipCode,'NULL','NULL')

            mysqlconnect.cursor.callproc('p_tb_supply', SupplySQL)
            mysqlconnect.cursor.callproc('p_tb_product', ProductSQL)

            CustomereSQLsel = """select group_id,name from tb_customer where group_id=%s""" % (GroupID)

            CustomereSQLupd = """update tb_customer
                set address= %s ,
                    phone= %s,
                    mobile = %s,
                    email = %s,
                    post= %s,
                    class = %s,
                    memo= %s

                where group_id=%s and name=%s;"""
            print SupplySQL[0]
            CustomereSQLins = """ insert into tb_customer
                                    (customer_id,group_id ,name,address,phone,mobile,email,post,class,memo)
                                    values(%s,%s,%s,%s,%s,%s,%s,%s,%s,%s);"""
            mysqlconnect.cursor.execute(CustomereSQLsel)
            result = mysqlconnect.cursor.fetchall()

            print result
            Name_compare=[]
            if result!=[]:
                for x in result:
                    Name_compare.append(aes.AESdecrypt("password",x[1], True))
                if Name.encode('utf-8') in Name_compare :
                    print "update"
                    mysqlconnect.cursor.execute(CustomereSQLupd,(ClientAdd, ClientTel, ClientPhone,'NULL','NULL','NULL','NULL',GroupID,x[1]))
                else:
                    print "select insert"
                    mysqlconnect.cursor.execute(CustomereSQLins, CustomereSQL)
            else:
                mysqlconnect.cursor.execute(CustomereSQLins, CustomereSQL)
            mysqlconnect.cursor.execute(CustomereSQLsel)
            result = mysqlconnect.cursor.fetchall()
            customer_id_temp=[]
            SalestrSQLsel="SELECT customer_id from tb_customer where name =%s;"
            for x in result:
                if aes.AESdecrypt('password',x[1],True)==Name.encode('utf-8'):
                    mysqlconnect.cursor.execute(SalestrSQLsel,(str(x[1]),))
                    customer_id_temp.append(mysqlconnect.cursor.fetchall()[0])
            print customer_id_temp
            # print x[1]
            for y in customer_id_temp:
                for x in result:
                    if aes.AESdecrypt('password', x[1], True) == Name.encode('utf-8'):
                        SalestrSQL = "INSERT INTO tb_sale (sale_id,seq_no,group_id,user_id,product_name,c_product_id,customer_id,name,quantity,price,trans_list_date,sale_date)"\
                                        "VALUES(%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)"
                        mysqlconnect.cursor.execute(SalestrSQL,(str(uuid.uuid4()), OrderNo, GroupID, '22345', PartName ,PartNo, y[0], x[1],PartQuility, PartTotalPrice, OrderDate, ShipmentDate))
            mysqlconnect.db.commit()


            if (Ordernum==OrderNo):
                print 'update'
                self.updataOrder(mongoOrder,firm,firmNo,ShipmentDateNo,ReconciliationDate,OrderNo,OrderStatus,OrderDate,PartName,PartType,PartNo,FormatNo,\
                                PartCost,PartPrice,PartQuility,PartTotalPrice,PartNote,PartShopNo,
                                PartAction,OrderType,ShipmentDate,PartNum,supplier)

            else:
                print 'insert'
                Ordernum=OrderNo
                self.insertOrder(mongoOrder,firm,firmNo,ShipmentDateNo,ReconciliationDate,OrderNo,OrderStatus,OrderDate,PartName,PartType,PartNo,FormatNo,\
                                PartCost,PartPrice,PartQuility,PartTotalPrice,PartNote,PartShopNo,
                                PartAction,OrderType,ShipmentDate,PartNum,supplier)
            if (Clientnum == ClientName):
                print 'update'
                self.updataClient(mongodbClient, firm, OrderNo, OrderName, ClientName, ClientZipCode, ClientAdd, ClientTel,ClientPhone, supplier)
            else:
                print 'insert'
                Clientnum = ClientName
                self.insertClient(mongodbClient, firm, OrderNo, ClientName, ClientZipCode, ClientAdd, ClientTel,ClientPhone, supplier)

        mysqlconnect.dbClose()
        mongoOrder.dbClose()
        mongodbClient.dbClose()

    # mongoDB storage   ç¬¬ä??‹å??¸æ˜¯ä¸Ÿä??¢ç?mongoOrder or mongoClient
    def insertOrder(self,mongoOrder,_firm,_firmNo,_ShipmentDateNo,_ReconciliationDate,_OrderNo,_OrderStatus,_OrderDate,_PartName,_PartType,_PartNo,_FormatNo,\
                                _PartCost,_PartPrice,_PartQuility,_PartTotalPrice,_PartNote,_PartShopNo,\
                                _PartAction,_OrderType,_ShipmentDate,_PartNum,_supplier):
        businessorder_doc={ 'firm':[_firm],'firmNo':[_firmNo],'ShipmentDateNo':[_ShipmentDateNo],'ReconciliationDate':[_ReconciliationDate],\
                           '_OrderNo':_OrderNo,'OrderStatus':[_OrderStatus],'OrderDate':_OrderDate,\
                           'PartName':[_PartName],'PartType':[_PartType],'PartNo':[_PartNo],'FormatNo':[_FormatNo],\
                           'PartCost':[_PartCost],'PartPrice':[_PartPrice],'PartQuility':[_PartQuility],'PartTotalPrice':[_PartTotalPrice],\
                           'PartNote':[_PartNote],'PartShopNo':[_PartShopNo],'PartAction':[_PartAction],'OrderType':[_OrderType],\
                           'ShipmentDate':_ShipmentDate,'PartNum':[_PartNum],'supplier':[_supplier]
                            }

        mongoOrder.cursor.insert(businessorder_doc)
    def updataOrder(self,mongoOrder,_firm,_firmNo,_ShipmentDateNo,_ReconciliationDate,_OrderNo,_OrderStatus,_OrderDate,_PartName,_PartType,_PartNo,_FormatNo,\
                                _PartCost,_PartPrice,_PartQuility,_PartTotalPrice,_PartNote,_PartShopNo,\
                                _PartAction,_OrderType,_ShipmentDate,_PartNum,_supplier):
        mongoOrder.cursor.update( { "OrderNo" : _OrderNo,'firm':_firm,'supplier':_supplier,'firmNo':[_firmNo]},\
                           {'$push':{'ShipmentDateNo':_ShipmentDateNo,'ReconciliationDate':_ReconciliationDate,\
                           'OrderStatus':_OrderStatus,'OrderDate':_OrderDate,\
                           'PartName':_PartName,'PartType':_PartType,'PartNo':_PartNo,'FormatNo':_FormatNo,\
                           'PartCost':_PartCost,'PartPrice':_PartPrice,'PartQuility':_PartQuility,'PartTotalPrice':_PartTotalPrice,\
                           'PartNote':_PartNote,'PartShopNo':_PartShopNo,'PartAction':_PartAction,'OrderType':_OrderType,\
                           'ShipmentDate':_ShipmentDate,'PartNum':_PartNum}}
                                                )
    def insertClient(self,mongodbClient,_firm,_OrderNo,_ClientName,_ClientZipCode,_ClientAdd,_ClientTel,_ClientPhone,_supplier):

        businessorder_doc={ 'firm':[_firm],'OrderNo':_OrderNo,'ClientName':[_ClientName],'ClientZipCode':[_ClientZipCode],'ClientAdd':[_ClientAdd],'ClientTel':[_ClientTel],'ClientPhone':[_ClientPhone],'supplier':[_supplier]}

        mongodbClient.cursor.insert(businessorder_doc)
    def updataClient(self,mongodbClient,_firm,_OrderNo,_ClientName,_ClientZipCode,_ClientAdd,_ClientTel,_ClientPhone,_supplier):
        mongodbClient.cursor.update(  { 'firm':_firm,'ClientName':_ClientName,'ClientZipCode':_ClientZipCode,'ClientTel':_ClientTel,'ClientPhone':_ClientPhone,'supplier':_supplier}\
                           ,{'$push':{"OrderNo" : _OrderNo}})

class IBON_Data():
    Data=None
    def __init__(self):
        pass
    def IBON_Data(self,supplier,GroupID,path):
        #mysql connector object
        mysqlconnect=ToMysql()
        mysqlconnect.connect()

        mongoOrder=ToMongodb()
        mongoOrder.setCollection('co_ordertest')
        mongoOrder.connect()
        mongodbClient=ToMongodb()
        mongodbClient.setCollection('co_co_clienttest')
        mongodbClient.connect()

        Ordernum=""
        Clientnum=""

        data=xlrd.open_workbook(path)
        table=data.sheets()[0]
        num_cols=table.ncols
        #put the data into the corresponding variable

        for row_index in range(1,table.nrows):
            for col_index in range(0,num_cols):
                aes=aes_data()
                strTurntDate = xlrd.xldate_as_tuple(table.cell_value(row_index, 3), data.datemode)
                TurntDate = datetime(*strTurntDate[0:6]).strftime('%Y-%m-%d')
                OrderNo = table.cell(row_index, 4).value[0:13]
                OrderName = table.cell(row_index, 6).value
                Name = table.cell(row_index, 7).value
                ClientName = aes.AESencrypt("password", Name.encode('utf-8'), True)
                Phone = table.cell(row_index, 8).value
                ClientPhone = aes.AESencrypt("password", Phone.encode('utf-8'), True)
                Add = table.cell(row_index, 9).value
                ClientAdd = aes.AESencrypt("password", Add.encode('utf-8'), True)
                Mail = table.cell(row_index, 10).value
                ClientMail = aes.AESencrypt("password", Mail.encode('utf-8'), True)
                ClientZipCode = table.cell(row_index, 11).value
                OrderType = table.cell(row_index, 13).value
                PartNo = str(table.cell(row_index, 15).value)
                PartName = table.cell(row_index, 17).value
                PartSpec = table.cell(row_index, 18).value
                PartQuility = table.cell(row_index, 19).value
                PartCost = table.cell(row_index, 20).value
                strShipmentDate = xlrd.xldate_as_tuple(table.cell_value(row_index, 25), data.datemode)
                ShipmentDate = datetime(*strShipmentDate[0:6]).strftime('%Y-%m-%d')
                firm = GroupID
                supplier = supplier
            SupplySQL = (str(uuid.uuid4()),GroupID, supplier,"","","","","","","","","","","")
            ProductSQL = (str(uuid.uuid4()), GroupID, PartNo, PartName,supplier,"",PartSpec,PartCost,0,0,'NULL','NULL','NULL','NULL')
            CustomereSQL = (str(uuid.uuid4()), GroupID, ClientName, ClientAdd, "", ClientPhone,'',ClientZipCode,'NULL','NULL')

            mysqlconnect.cursor.callproc('p_tb_supply', SupplySQL)
            mysqlconnect.cursor.callproc('p_tb_product', ProductSQL)

            CustomereSQLsel = """select group_id,name from tb_customer where group_id=%s""" % (GroupID)

            CustomereSQLupd = """update tb_customer
                set address= %s ,
                    phone= %s,
                    mobile = %s,
                    email = %s,
                    post= %s,
                    class = %s,
                    memo= %s

                where group_id=%s and name=%s;"""
            print SupplySQL[0]
            CustomereSQLins = """ insert into tb_customer
                                    (customer_id,group_id ,name,address,phone,mobile,email,post,class,memo)
                                    values(%s,%s,%s,%s,%s,%s,%s,%s,%s,%s);"""
            mysqlconnect.cursor.execute(CustomereSQLsel)
            result = mysqlconnect.cursor.fetchall()

            print result
            Name_compare=[]
            if result!=[]:
                for x in result:
                    Name_compare.append(aes.AESdecrypt("password",x[1], True))
                if Name.encode('utf-8') in Name_compare :
                    print "update"
                    mysqlconnect.cursor.execute(CustomereSQLupd,(ClientAdd,'NULL', ClientPhone,"",'NULL','NULL','NULL',GroupID,x[1]))
                else:
                    print "select insert"
                    mysqlconnect.cursor.execute(CustomereSQLins, CustomereSQL)
            else:
                mysqlconnect.cursor.execute(CustomereSQLins, CustomereSQL)
            mysqlconnect.cursor.execute(CustomereSQLsel)
            result = mysqlconnect.cursor.fetchall()
            customer_id_temp=[]
            SalestrSQLsel="SELECT customer_id from tb_customer where name =%s;"
            for x in result:
                if aes.AESdecrypt('password',x[1],True)==Name.encode('utf-8'):
                    mysqlconnect.cursor.execute(SalestrSQLsel,(str(x[1]),))
                    customer_id_temp.append(mysqlconnect.cursor.fetchall()[0])
            print customer_id_temp
            # print x[1]
            for y in customer_id_temp:
                for x in result:
                    if aes.AESdecrypt('password', x[1], True) == Name.encode('utf-8'):
                        SalestrSQL = "INSERT INTO tb_sale (sale_id,seq_no,group_id,user_id,product_name,c_product_id,customer_id,name,quantity,price,trans_list_date,sale_date)"\
                                        "VALUES(%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)"
                        mysqlconnect.cursor.execute(SalestrSQL,(str(uuid.uuid4()), OrderNo, GroupID, '32345',PartName ,PartNo, y[0], x[1],  PartQuility, PartCost, TurntDate, ShipmentDate))
            mysqlconnect.db.commit()

            if (Ordernum==OrderNo):
                print 'update'
                self.updataOrder(mongoOrder,TurntDate,OrderNo,OrderType,PartNo,PartName,PartSpec,\
                                PartQuility,PartCost,ShipmentDate,firm,supplier)

            else:
                print 'insert'
                Ordernum=OrderNo
                self.insertOrder(mongoOrder,TurntDate,OrderNo,OrderType,PartNo,PartName,PartSpec,\
                                PartQuility,PartCost,ShipmentDate,firm,supplier)
            if (Clientnum == ClientName):
                print 'update'
                self.updataClient(mongodbClient, OrderNo,OrderName,ClientName,ClientPhone,ClientAdd,ClientMail,\
                                ClientZipCode,firm,supplier)
            else:
                print 'insert'
                Clientnum = ClientName
                self.insertClient(mongodbClient, OrderNo,OrderName,ClientName,ClientPhone,ClientAdd,ClientMail,\
                                ClientZipCode,firm,supplier)

        mysqlconnect.dbClose()
        mongoOrder.dbClose()
        mongodbClient.dbClose()

    # mongoDB storage   ç¬¬ä??‹å??¸æ˜¯ä¸Ÿä??¢ç?mongoOrder or mongoClient
    def insertOrder(self,mongoOrder,_TurntDate,_OrderNo,_OrderType,_PartNo,_PartName,_PartSpec,\
                                _PartQuility,_PartCost,_ShipmentDate,_firm,_supplier):
        businessorder_doc={ 'TurntDate':_TurntDate,'OrderNo':_OrderNo,'OrderType':[_OrderType],'PartNo':[_PartNo],\
                           'PartName':[_PartName],'PartSpec':[_PartSpec],\
                           'PartQuility':[_PartQuility],'PartCost':[_PartCost],'ShipmentDate':_ShipmentDate,'firm':[_firm],'supplier':[_supplier]
                            }

        mongoOrder.cursor.insert(businessorder_doc)
    def updataOrder(self,mongoOrder,_TurntDate,_OrderNo,_OrderType,_PartNum,_PartName,_PartSpec,\
                                _PartQuility,_PartCost,_ShipmentDate,_firm,_supplier):
        mongoOrder.cursor.update( { "OrderNo" : _OrderNo,'firm':_firm,'supplier':_supplier},\
                           {'$push':{'TurntDate':_TurntDate,'OrderType':_OrderType,'PartNum':_PartNum ,'PartName':_PartName,\
                           'PartSpec':_PartSpec,'PartQuility':_PartQuility,'PartCost':_PartCost,'ShipmentDate':_ShipmentDate}}
                                                )
    def insertClient(self,mongodbClient,_OrderNo,_OrderName,_ClientName,_ClientPhone,_ClientAdd,_ClientMail,\
                                _ClientZipCode,_firm,_supplier):

        businessorder_doc={'OrderNo':_OrderNo,'OrderName':[_OrderName],'ClientName':[_ClientName],'ClientPhone':[_ClientPhone],'ClientAdd':[_ClientAdd],\
                           'ClientMail':[_ClientMail],'ClientZipCode':[_ClientZipCode],'firm':[_firm],'supplier':[_supplier]}

        mongodbClient.cursor.insert(businessorder_doc)
    def updataClient(self,mongodbClient,_OrderNo,_OrderName,_ClientName,_ClientPhone,_ClientAdd,_ClientMail,\
                                _ClientZipCode,_firm,_supplier):
        mongodbClient.cursor.update(  {  'OrderName':_OrderName,'ClientName':_ClientName,'ClientPhone':[_ClientPhone],'ClientAdd':_ClientAdd,\
                             'ClientMail':_ClientMail,'ClientZipCode':_ClientZipCode,'firm':_firm,'supplier':_supplier}\
                           ,{'$push':{"OrderNo" : _OrderNo}})

class IBONC_Data():
    Data=None
    def __init__(self):
        pass
    def IBONC_Data(self,supplier,GroupID,path):
        #mysql connector object
        mysqlconnect=ToMysql()
        mysqlconnect.connect()

        mongoOrder=ToMongodb()
        mongoOrder.setCollection('co_ordertest')
        mongoOrder.connect()
        mongodbClient=ToMongodb()
        mongodbClient.setCollection('co_co_clienttest')
        mongodbClient.connect()

        Ordernum=""
        Clientnum=""

        data=xlrd.open_workbook(path)
        table=data.sheets()[0]
        num_cols=table.ncols
        #put the data into the corresponding variable

        for row_index in range(1,table.nrows):
            for col_index in range(0,num_cols):
                aes=aes_data()
                strTurntDate = xlrd.xldate_as_tuple(table.cell_value(row_index, 5), data.datemode)
                TurntDate = datetime(*strTurntDate[0:6]).strftime('%Y-%m-%d')
                strShipmentDate = xlrd.xldate_as_tuple(table.cell_value(row_index, 7), data.datemode)
                ShipmentDate = datetime(*strShipmentDate[0:6]).strftime('%Y-%m-%d')
                OrderNo = table.cell(row_index, 8).value[0:13]
                OrderName = table.cell(row_index, 11).value
                Phone = table.cell(row_index, 12).value
                ClientPhone = aes.AESencrypt("password", Phone.encode('utf-8'), True)
                Name = table.cell(row_index, 13).value
                ClientName = aes.AESencrypt("password", Name.encode('utf-8'), True)
                PartNo = table.cell(row_index, 21).value
                PartName = table.cell(row_index, 22).value
                PartSpec = table.cell(row_index, 23).value
                PartQuility = table.cell(row_index, 25).value
                PartCost = table.cell(row_index, 26).value
                firm = GroupID
                supplier = supplier
            SupplySQL = (str(uuid.uuid4()),GroupID, supplier,"","","","","","","","","","","")
            print SupplySQL
            ProductSQL = (str(uuid.uuid4()), GroupID, PartNo, PartName,supplier,"",PartSpec,PartCost,0,0,'NULL','NULL','NULL','NULL')
            print ProductSQL
            print len(PartName)
            CustomereSQL = (str(uuid.uuid4()), GroupID, ClientName,"", "", ClientPhone,'',"",'NULL','NULL')

            mysqlconnect.cursor.callproc('p_tb_supply', SupplySQL)
            mysqlconnect.cursor.callproc('p_tb_product', ProductSQL)

            CustomereSQLsel = """select group_id,name from tb_customer where group_id=%s""" % (GroupID)

            CustomereSQLupd = """update tb_customer
                set address= %s ,
                    phone= %s,
                    mobile = %s,
                    email = %s,
                    post= %s,
                    class = %s,
                    memo= %s

                where group_id=%s and name=%s;"""
            print SupplySQL[0]
            CustomereSQLins = """ insert into tb_customer
                                    (customer_id,group_id ,name,address,phone,mobile,email,post,class,memo)
                                    values(%s,%s,%s,%s,%s,%s,%s,%s,%s,%s);"""
            mysqlconnect.cursor.execute(CustomereSQLsel)
            result = mysqlconnect.cursor.fetchall()

            print result
            Name_compare=[]
            if result!=[]:
                for x in result:
                    Name_compare.append(aes.AESdecrypt("password",x[1], True))
                if Name.encode('utf-8') in Name_compare :
                    print "update"
                    mysqlconnect.cursor.execute(CustomereSQLupd,("",'NULL', ClientPhone,"",'NULL','NULL','NULL',GroupID,x[1]))
                else:
                    print "select insert"
                    mysqlconnect.cursor.execute(CustomereSQLins, CustomereSQL)
            else:
                mysqlconnect.cursor.execute(CustomereSQLins, CustomereSQL)
            mysqlconnect.cursor.execute(CustomereSQLsel)
            result = mysqlconnect.cursor.fetchall()
            customer_id_temp=[]
            SalestrSQLsel="SELECT customer_id from tb_customer where name =%s;"
            for x in result:
                if aes.AESdecrypt('password',x[1],True)==Name.encode('utf-8'):
                    mysqlconnect.cursor.execute(SalestrSQLsel,(str(x[1]),))
                    customer_id_temp.append(mysqlconnect.cursor.fetchall()[0])
            print customer_id_temp
            # print x[1]
            for y in customer_id_temp:
                for x in result:
                    if aes.AESdecrypt('password', x[1], True) == Name.encode('utf-8'):
                        print 'GroupID'
                        print PartName
                        SalestrSQL = "INSERT INTO tb_sale (sale_id,seq_no,group_id,user_id,product_name,c_product_id,customer_id,name,quantity,price,trans_list_date,sale_date)"\
                                        "VALUES(%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)"
                        mysqlconnect.cursor.execute(SalestrSQL,(str(uuid.uuid4()), OrderNo, GroupID, '32345', PartName ,PartNo, y[0], x[1], PartQuility, PartCost, TurntDate, ShipmentDate))
            mysqlconnect.db.commit()

            if (Ordernum==OrderNo[0:13]):
                print 'update'
                self.updataOrder(mongoOrder,TurntDate,ShipmentDate,OrderNo,PartNo,PartName,PartSpec,\
                                PartQuility,PartCost,firm,supplier)

            else:
                print 'insert'
                Ordernum=OrderNo
                self.insertOrder(mongoOrder,TurntDate,ShipmentDate,OrderNo,PartNo,PartName,PartSpec,\
                                PartQuility,PartCost,firm,supplier)
            if (Clientnum == ClientName):
                print 'update'
                self.updataClient(mongodbClient, OrderNo,OrderName,ClientPhone,ClientName,\
                                firm,supplier)
            else:
                print 'insert'
                Clientnum = ClientName
                self.insertClient(mongodbClient, OrderNo,OrderName,ClientPhone,ClientName,\
                                firm,supplier)

        mysqlconnect.dbClose()
        mongoOrder.dbClose()
        mongodbClient.dbClose()

    # mongoDB storage   ç¬¬ä??‹å??¸æ˜¯ä¸Ÿä??¢ç?mongoOrder or mongoClient
    def insertOrder(self,mongoOrder,_TurntDate,_ShipmentDate,_OrderNo,_PartNo,_PartName,_PartSpec,\
                                _PartQuility,_PartCost,_firm,_supplier):
        businessorder_doc={'TurntDate':_TurntDate,'ShipmentDate':_ShipmentDate,'OrderNo':_OrderNo,'PartNo':[_PartNo],\
                           'PartName':[_PartName],'PartSpec':[_PartSpec],\
                           'PartQuility':[_PartQuility],'PartCost':[_PartCost],'firm':[_firm],'supplier':[_supplier]
                            }

        mongoOrder.cursor.insert(businessorder_doc)
    def updataOrder(self,mongoOrder,_TurntDate,_ShipmentDate,_OrderNo,_PartNo,_PartName,_PartSpec,\
                                _PartQuility,_PartCost,_firm,_supplier):
        mongoOrder.cursor.update( { "OrderNo" : _OrderNo,'firm':_firm,'supplier':_supplier},\
                           {'$push':{'TurntDate':_TurntDate,'ShipmentDate':_ShipmentDate,'PartNo':_PartNo ,'PartName':_PartName,\
                           'PartSpec':_PartSpec,'PartQuility':_PartQuility,'PartCost':_PartCost}}
                                                )
    def insertClient(self,mongodbClient,_OrderNo,_OrderName,_ClientPhone,_ClientName,\
                                _firm,_supplier):

        businessorder_doc={'OrderNo':_OrderNo,'OrderName':[_OrderName],'ClientPhone':[_ClientPhone],'ClientName':[_ClientName],\
                           'firm':[_firm],'supplier':[_supplier]}

        mongodbClient.cursor.insert(businessorder_doc)
    def updataClient(self,mongodbClient,_OrderNo,_OrderName,_ClientPhone,_ClientName,\
                                _firm,_supplier):
        mongodbClient.cursor.update(  { 'OrderName':[_OrderName],'ClientName':_ClientName,'ClientPhone':[_ClientPhone],\
                             'firm':_firm,'supplier':_supplier}\
                           ,{'$push':{"OrderNo" : _OrderNo}})

