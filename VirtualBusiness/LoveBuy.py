# -*-  coding: utf-8  -*-
__author__ = '10409003'
import csv
import uuid
from datetime import datetime
from aes_data import aes_data
from ToMongodb import ToMongodb
from ToMysql import ToMysql
import logging
import time

class LoveBuy_Data():
    Data=None
    def __init__(self):
        pass
    def LoveBuy_Data(self,supplier,GroupID,path,UserID):
        logging.basicConfig(filename='pyupload.log', level=logging.DEBUG, format='%(asctime)s %(message)s', datefmt='%Y/%m/%d %I:%M:%S %p')
        logging.Formatter.converter = time.gmtime
        logging.info('===LoveBuy_Data===')
        logging.debug('supplier:' + supplier)
        logging.debug('GroupID:' + GroupID)
        logging.debug('path:' + path)
        logging.debug('UserID:' + UserID)
        
        #mysql connector object
        mysqlconnect=ToMysql()
        mysqlconnect.connect()

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
            ignore = 0

            for row in reader:
                if ignore:
                    aes=aes_data()
                    OrderNo = row[3].split("'")[-1]
                    Name = str(row[9].split("'")[-1]).decode('big5')
                    ClientName = aes.AESencrypt("p@ssw0rd", Name.encode('utf-8'), True)
                    Phone = str(row[12].split("'")[-1]).decode('big5')
                    ClientPhone = aes.AESencrypt("p@ssw0rd", Phone, True)
                    Tel = str(row[11].split("'")[-1]).decode('big5')
                    ClientTel = aes.AESencrypt("p@ssw0rd", Tel, True)
                    Add = str(row[10].split("'")[-1]).decode('big5')
                    ClientAdd = aes.AESencrypt("p@ssw0rd", Add.encode('utf-8'), True)
                    PartName = (row[5].split("'")[-1]).decode('big5')
                    PartNo = row[23].split("'")[-1]
                    PartTotalPrice = int(str(str(row[6]).split('(')[0]).split("'")[-1])
                    PartQuility=int(str(str(row[6]).split('(')[-1]).split(')')[0])
                    strTurnDate = row[0].split("'")[-1]
                    temp_time = map(int, ((strTurnDate.split(' ')[0]).split('-')))  # 2015/5/23 07:25
                    _TurnDate = datetime(*(temp_time[0:3])).strftime('%Y-%m-%d %H:%M:%S.%f')
                    TurnDate = datetime.strptime(_TurnDate, '%Y-%m-%d %H:%M:%S.%f')
                    strShipmentDate = row[1].split("'")[-1]
                    temp_time = map(int, ((strShipmentDate.split(' ')[0]).split('-')))  # 2015/5/23 07:25
                    _ShipmentDate = datetime(*(temp_time[0:3])).strftime('%Y-%m-%d %H:%M:%S.%f')
                    ShipmentDate = datetime.strptime(_ShipmentDate, '%Y-%m-%d %H:%M:%S.%f')
                    firm = GroupID
                    supplier = supplier
                    UserID = UserID
                    # SupplySQL = (str(uuid.uuid4()),GroupID, supplier,"","","","","","","","","","","")
                    ProductSQL = (str(uuid.uuid4()), GroupID, PartNo, PartName,supplier, None,None,0,PartTotalPrice,0,None,None,None,None)
                    CustomereSQL = (str(uuid.uuid4()), GroupID, ClientName, ClientAdd, ClientTel, ClientPhone,None,None,None,None)

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

                    print result
                    Name_compare=[]
                    if result!=[]:
                        for x in result:
                            Name_compare.append(aes.AESdecrypt("p@ssw0rd",x[1], True))
                        if Name.encode('utf-8') in Name_compare :
                            print "update"
                            mysqlconnect.cursor.execute(CustomereSQLupd,(ClientAdd, ClientTel, ClientPhone,None,None,None,None,GroupID,x[1]))
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
                        if aes.AESdecrypt('p@ssw0rd',x[1],True)==Name.encode('utf-8'):
                            mysqlconnect.cursor.execute(SalestrSQLsel,(str(x[1]),))
                            customer_id_temp.append(mysqlconnect.cursor.fetchall()[0])
                    print customer_id_temp
                    for y in customer_id_temp:
                        for x in result:
                            if aes.AESdecrypt('p@ssw0rd', x[1], True) == Name.encode('utf-8'):
                                SaleSQL = (
                                    GroupID, OrderNo, UserID, PartName, PartNo, y[0],
                                    x[1], PartQuility, PartTotalPrice, None, None, TurnDate, None,
                                    None,ShipmentDate,supplier)
                                mysqlconnect.cursor.callproc('p_tb_sale', SaleSQL)
                                # SalestrSQL = "INSERT INTO tb_sale (sale_id,seq_no,group_id,user_id,customer_id,name,c_product_id,quantity,price,trans_list_date,sale_date,product_name)"\
                                # "VALUES(%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)"
                                # mysqlconnect.cursor.execute(SalestrSQL,(str(uuid.uuid4()), OrderNo, GroupID, '12345', y[0], x[1], PartNo, PartQuility,PartTotalPrice, TurnDate,ShipmentDate,PartName))
                    mysqlconnect.db.commit()
                    # mysqlconnect.cursor.callproc('p_tb_customer', CustomereSQL)

                    if (Ordernum==OrderNo[0:13]):
                        print 'update'
                        self.updataOrder(mongoOrder,OrderNo,PartName,PartNo,\
                                        PartQuility,PartTotalPrice,\
                                        firm,supplier)

                    else:
                        print 'insert'
                        Ordernum=OrderNo
                        self.insertOrder(mongoOrder,OrderNo,PartName,PartNo,\
                                        PartQuility,PartTotalPrice,\
                                        firm,supplier)
                    if (Clientnum == ClientName):
                        print 'update'
                        self.updataClient(mongodbClient, OrderNo,ClientName,\
                                        ClientAdd,ClientTel,ClientPhone,\
                                       firm,supplier)
                    else:
                        print 'insert'
                        Clientnum = ClientName
                        self.insertClient(mongodbClient, OrderNo,ClientName,\
                                        ClientAdd,ClientTel,ClientPhone,\
                                       firm,supplier)
                else:
                    ignore=1
        mysqlconnect.dbClose()
        mongoOrder.dbClose()
        mongodbClient.dbClose()
        logging.info('===LoveBuy_Data SUCCESS===')
        return 'success'

    # mongoDB storage   ç¬¬ä??‹å??¸æ˜¯ä¸Ÿä??¢ç?mongoOrder or mongoClient
    def insertOrder(self,mongoOrder,_OrderNo,_PartName,_PartNo,\
                                _PartQuility,_PartTotalPrice,\
                                _firm,_supplier):
        businessorder_doc={ 'OrderNo':_OrderNo,\
                           'PartName':[_PartName],'PartNo':[_PartNo],\
                           'PartQuility':[_PartQuility],'Price':[_PartTotalPrice],\
                           'firm':[_firm],'supplier':[_supplier]
                            }

        mongoOrder.cursor.insert(businessorder_doc)
    def updataOrder(self,mongoOrder,_OrderNo,_PartName,_PartNo,\
                                _PartQuility,_PartTotalPrice,\
                                _firm,_supplier):
        mongoOrder.cursor.update({ "OrderNo" : _OrderNo,'firm':_firm,'supplier':_supplier},\
                           {'$push':{'PartName':_PartName,'PartNo':_PartNo ,\
                           'PartQuility':_PartQuility,'Price':_PartTotalPrice}}
                                                )
    def insertClient(self,mongodbClient,_OrderNo,_ClientName,_ClientAdd,_ClientTel,\
                   _ClientPhone,_firm,_supplier):
        businessorder_doc={'OrderNo':_OrderNo,'ClientName':[_ClientName],'ClientTel':[_ClientTel],'ClientPhone':[_ClientPhone],'ClientAdd':[_ClientAdd],\
                           'firm':[_firm],'supplier':[_supplier]}

        mongodbClient.cursor.insert(businessorder_doc)
    def updataClient(self,mongodbClient,_OrderNo,_ClientName,_ClientAdd,_ClientTel,\
                   _ClientPhone,_firm,_supplier):
        mongodbClient.cursor.update({ 'ClientName':_ClientName,'ClientAdd':_ClientAdd,'ClientTel':_ClientTel,'ClientPhone':_ClientPhone,\
                            'firm':_firm,'supplier':_supplier}\
                           ,{'$push':{"OrderNo" : _OrderNo}})







