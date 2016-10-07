# -*-  coding: utf-8  -*-
__author__ = '10409003'
import xlrd
import uuid
import csv
from datetime import datetime
from aes_data import aes_data
from ToMongodb import ToMongodb
from ToMysql import ToMysql



class LineMart_Data():
    Data=None
    def __init__(self):
        pass
    def LineMart_Data(self,supplier,GroupID,path,UserID):
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
            ignore=0

            for row in reader:
                if ignore:
                    aes = aes_data()
                    OrderNo = row[1]
                    PartName = row[3].decode('big5')
                    PartNum = row[4]
                    PartPrice = row[7]
                    PartQuility = row[8]
                    PartTotalPrice = row[9]
                    strTurntDate = row[10]
                    temp_time = map(int,((strTurntDate.split(' ')[0]).split('/')))#2015/5/23 07:25
                    _TurnDate = datetime(*(temp_time[0:3])).strftime('%Y-%m-%d')
                    TurnDate = datetime.strptime(_TurnDate, '%Y-%m-%d')

                    Name = row[11].decode('big5')
                    ClientName = aes.AESencrypt("p@ssw0rd", Name.encode('utf-8'), True)
                    Phone = row[12]
                    ClientPhone = aes.AESencrypt("p@ssw0rd", Phone, True)

                    ClientZipCode = row[13]
                    Add1 = row[14]
                    ClientAdd1 = aes.AESencrypt("p@ssw0rd", Add1, True)
                    Add2 = row[15]
                    ClientAdd2 = aes.AESencrypt("p@ssw0rd", Add2, True)
                    Add3 = row[16].decode('big5')

                    ClientAdd3 = aes.AESencrypt("p@ssw0rd", Add3.encode('utf-8'), True)
                    ClientAdd = ClientAdd3
                    firm = GroupID
                    supplier = supplier
                    UserID = UserID

                    # SupplySQL = (str(uuid.uuid4()), GroupID, supplier, "", "", "", "", "", "", "", "", "", "", "")
                    # ProductSQL = (
                    #     str(uuid.uuid4()), GroupID, PartNum, PartName, supplier, "", "", 0, PartPrice, 0, None,
                    #     None,
                    #     None, None)
                    CustomereSQL = (
                        str(uuid.uuid4()), GroupID, ClientName, ClientAdd,None, ClientPhone, None, ClientZipCode,
                        None,
                        None)
                    mysqlconnect.connect()
                    # mysqlconnect.cursor.callproc('p_tb_supply', SupplySQL)
                    # mysqlconnect.cursor.callproc('p_tb_product', ProductSQL)

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
                    print Name.decode('big5')
                    print result
                    Name_compare = []
                    if result != []:
                        for x in result:
                            Name_compare.append(aes.AESdecrypt("p@ssw0rd", x[1], True))
                        if (Name.decode('big5')).encode('utf-8') in Name_compare:
                            print "update"
                            mysqlconnect.cursor.execute(CustomereSQLupd, (
                                ClientAdd, "", ClientPhone,None, ClientZipCode, None, None, GroupID, x[1]))
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
                        if aes.AESdecrypt('p@ssw0rd', x[1], True) == (Name.decode('big5')).encode('utf-8'):
                            mysqlconnect.cursor.execute(SalestrSQLsel, (str(x[1]),))
                            customer_id_temp.append(mysqlconnect.cursor.fetchall()[0])
                    print customer_id_temp
                    for y in customer_id_temp:
                        for x in result:
                            if aes.AESdecrypt('p@ssw0rd', x[1], True) == (Name.decode('big5')).encode('utf-8'):
                                SaleSQL = (
                                    GroupID, OrderNo, UserID, PartName,PartNum, y[0],
                                    x[1],PartQuility, PartPrice, None, TurnDate, TurnDate, TurnDate, None,
                                    TurnDate,supplier)
                                mysqlconnect.cursor.callproc('p_tb_sale', SaleSQL)
                                # SalestrSQL = "INSERT INTO tb_sale (sale_id,seq_no,group_id,user_id,customer_id,name,c_product_id,quantity,price,sale_date,product_name)" \
                                #              "VALUES(%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)"
                                # mysqlconnect.cursor.execute(SalestrSQL, (
                                #     str(uuid.uuid4()), OrderNo, GroupID, '52345', y[0], x[1], PartNum, PartQuility,
                                #     PartTotalPrice, TurntDate,PartName))
                    mysqlconnect.db.commit()
                    # mysqlconnect.cursor.callproc('p_tb_customer', CustomereSQL)

                    if (Ordernum == OrderNo[0:13]):
                        print 'update'
                        self.updataOrder(mongoOrder, TurnDate,OrderNo, PartName, PartNum, PartPrice, \
                                         PartQuility, PartTotalPrice, \
                                         firm, supplier)

                    else:
                        print 'insert'
                        Ordernum = OrderNo
                        self.insertOrder(mongoOrder, TurnDate,OrderNo, PartName, PartNum, PartPrice, \
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
    def insertOrder(self,mongoOrder,_TurnDate,_OrderNo,_PartName,_PartNum,_PartPrice,\
                                _PartQuility,_PartTotalPrice,\
                                _firm,_supplier):
        businessorder_doc={'TurnDate':_TurnDate,'OrderNo':_OrderNo,\
                           'PartName':[_PartName],'PartNum':[_PartNum],'Price':[_PartPrice],\
                           'PartQuility':[_PartQuility],'PartTotalPrice':[_PartTotalPrice],'firm':[_firm],'supplier':[_supplier]
                            }

        mongoOrder.cursor.insert(businessorder_doc)
    def updataOrder(self,mongoOrder,_TurnDate,_OrderNo,_PartName,_PartNum,_PartPrice,\
                                _PartQuility,_PartTotalPrice,\
                                _firm,_supplier):
        mongoOrder.cursor.update({ "OrderNo" : _OrderNo,'firm':_firm,'supplier':_supplier},\
                           {'$push':{'TurnDate':_TurnDate,'PartName':_PartName,'PartNum':_PartNum ,\
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






