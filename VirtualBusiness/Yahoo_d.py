# -*-  coding: utf-8  -*-
__author__ = '10409003'
import xlrd
import uuid
import datetime
from aes_data import aes_data
from ToMongodb import ToMongodb
from ToMysql import ToMysql
import logging
import time


class Yahood_Data():
    Data = None

    def __init__(self):
        pass

    def Yahood_Data(self, supplier, GroupID, path, UserID):
        logging.basicConfig(filename='pyupload.log', level=logging.DEBUG, format='%(asctime)s %(message)s',
                            datefmt='%Y/%m/%d %I:%M:%S %p')
        logging.Formatter.converter = time.gmtime
        logging.info('===Yahood_Data===')
        logging.debug('supplier:' + supplier)
        logging.debug('GroupID:' + GroupID)
        logging.debug('path:' + path)
        logging.debug('UserID:' + UserID)

        # mysql connector object
        mysqlconnect = ToMysql()
        mysqlconnect.connect()

        mongoOrder = ToMongodb()
        mongoOrder.setCollection('co_order')
        mongoOrder.connect()
        mongodbClient = ToMongodb()
        mongodbClient.setCollection('co_client')
        mongodbClient.connect()

        Ordernum = ""
        Clientnum = ""

        data = xlrd.open_workbook(path)
        table = data.sheets()[0]
        num_cols = table.ncols
        # put the data into the corresponding variable
        for row_index in range(2, table.nrows):
            for col_index in range(0, num_cols):
                aes = aes_data()
                OrderNo = str(table.cell(row_index, 1).value)

                # strTurnDate = str(table.cell(row_index, 2).value)
                # _strTurnDate = datetime.datetime.strptime(strTurnDate,'%Y/%m/%d %H:%M ')
                # TurnDate = datetime.datetime.strftime(_strTurnDate,'%Y-%m-%d %H:%M')
                TurnDate = datetime.datetime.strptime(str(table.cell_value(row_index, 2)), '%Y/%m/%d %H:%M')
                ShipmentDate = datetime.datetime.strptime(str(table.cell_value(row_index, 3)), '%Y/%m/%d %H:%M')

                Name = table.cell(row_index, 6).value
                ClientName = aes.AESencrypt("p@ssw0rd", Name.encode('utf-8'), True)
                Phone = table.cell(row_index, 9).value
                ClientPhone = aes.AESencrypt("p@ssw0rd", Phone, True)
                Tel = table.cell(row_index, 7).value
                ClientTel = aes.AESencrypt("p@ssw0rd", Tel, True)
                ClientZipCode = table.cell(row_index, 10).value
                Add = table.cell(row_index, 11).value
                ClientAdd = aes.AESencrypt("p@ssw0rd", Add.encode('utf-8'), True)
                PartNo = str(table.cell(row_index, 13).value).split('.')[0]
                PartName = table.cell(row_index, 14).value

                PartQuility = table.cell(row_index, 18).value
                PartCost = table.cell(row_index, 19).value
                PartTotalPrice = table.cell(row_index, 20).value
                # strShipmentDate = xlrd.xldate_as_tuple(table.cell_value(row_index, 10), data.datemode)
                # ShipmentDate = datetime(*strShipmentDate[0:6]).strftime('%Y-%m-%d')

                firm = GroupID
                supplier = supplier
                UserID = UserID
            # SupplySQL = (str(uuid.uuid4()),GroupID, supplier,"","","","","","","","","","","")
            ProductSQL = (
            str(uuid.uuid4()), GroupID, PartNo, PartName, supplier, None, None, PartCost, PartTotalPrice, 0, None, None,
            None, None)
            CustomereSQL = (
            str(uuid.uuid4()), GroupID, ClientName, ClientAdd, ClientTel, ClientPhone, None, ClientZipCode, None, None)

            # mysqlconnect.cursor.callproc('p_tb_supply', SupplySQL)
            mysqlconnect.cursor.callproc('p_tb_product', ProductSQL)

            CustomereSQLsel = """select group_id,name,address,mobile from tb_customer where group_id='%s'""" % (GroupID)

            CustomereSQLupd_1 = """update tb_customer
                set phone= %s,
                    mobile = %s,
                    email = %s,
                    post= %s,
                    class = %s,
                    memo= %s
                where group_id=%s and name=%s and address= %s;"""
            CustomereSQLupd_2 = """update tb_customer
                set address= %s ,
                    phone= %s,
                    email = %s,
                    post= %s,
                    class = %s,
                    memo= %s

                where group_id=%s and name=%s and mobile=%s;"""

            CustomereSQLins = """ insert into tb_customer
                                    (customer_id,group_id ,name,address,phone,mobile,email,post,class,memo)
                                    values(%s,%s,%s,%s,%s,%s,%s,%s,%s,%s);"""
            mysqlconnect.cursor.execute(CustomereSQLsel)
            Dataresult = mysqlconnect.cursor.fetchall()

            print Dataresult
            same_name = []
            if Dataresult != []:
                for x in Dataresult:
                    Name_compare = aes.AESdecrypt("p@ssw0rd", x[1], True)
                    if Name.encode('utf-8') == Name_compare:
                        print "the same name data list"
                        same_name.append(x)
                if same_name == []:
                    mysqlconnect.cursor.execute(CustomereSQLins, CustomereSQL)
                for r in same_name:
                    address_compare = aes.AESdecrypt("p@ssw0rd", r[2], True)
                    mobile_compare = aes.AESdecrypt("p@ssw0rd", r[3], True)
                    if Add.encode('utf-8') == address_compare:
                        mysqlconnect.cursor.execute(CustomereSQLupd_1, (
                        ClientTel, ClientPhone, None, None, None, None, GroupID, r[1], r[2]))
                    elif Phone.encode('utf-8') == mobile_compare:
                        mysqlconnect.cursor.execute(CustomereSQLupd_2,
                                                    (ClientAdd, ClientTel, None, None, None, None, GroupID, r[1], r[3]))
                    else:
                        print "select insert"
                        mysqlconnect.cursor.execute(CustomereSQLins, CustomereSQL)
            else:
                mysqlconnect.cursor.execute(CustomereSQLins, CustomereSQL)
            mysqlconnect.cursor.execute(CustomereSQLsel)
            Finalresult = mysqlconnect.cursor.fetchall()
            customer_id_temp = []
            SalestrSQLsel_1 = "SELECT customer_id from tb_customer where name =%s and address= %s;"
            SalestrSQLsel_2 = "SELECT customer_id from tb_customer where name =%s and mobile=%s;"
            same_name = []
            for x in Finalresult:
                Name_compare = aes.AESdecrypt("p@ssw0rd", x[1], True)
                if Name.encode('utf-8') == Name_compare:
                    print "the same name data list"
                    same_name.append(x)

            for r in same_name:
                address_compare = aes.AESdecrypt("p@ssw0rd", r[2], True)
                mobile_compare = aes.AESdecrypt("p@ssw0rd", r[3], True)
                if Add.encode('utf-8') == address_compare:
                    mysqlconnect.cursor.execute(SalestrSQLsel_1, (str(r[1]), str(r[2])))
                    customer_id_temp.append(mysqlconnect.cursor.fetchall()[0])
                elif Phone.encode('utf-8') == mobile_compare:
                    mysqlconnect.cursor.execute(SalestrSQLsel_2, (str(r[1]), str(r[3])))
                    customer_id_temp.append(mysqlconnect.cursor.fetchall()[0])

            print customer_id_temp
            mysqlconnect.db.commit()
            for y in customer_id_temp:
                for x in Finalresult:
                    if aes.AESdecrypt('p@ssw0rd', x[1], True) == Name.encode('utf-8'):
                        SaleSQL = (
                            GroupID, OrderNo, UserID, PartName, PartNo, y[0],
                            x[1], PartQuility, PartTotalPrice, None, None, TurnDate, None,
                            None, ShipmentDate, supplier)
                        mysqlconnect.cursor.callproc('p_tb_sale', SaleSQL)
                        # SalestrSQL = "INSERT INTO tb_sale (sale_id,seq_no,group_id,user_id,c_product_id,customer_id,name,quantity,price,trans_list_date,sale_date,product_name)"\
                        # "VALUES(%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)"
                        # mysqlconnect.cursor.execute(SalestrSQL,(str(uuid.uuid4()), OrderNo, GroupID,'73345',PartNo , y[0], x[1], PartQuility,PartTotalPrice, TurnDate,ShipmentDate,PartName))
            mysqlconnect.db.commit()
            # mysqlconnect.cursor.callproc('p_tb_customer', CustomereSQL)

            if (Ordernum == OrderNo[0:13]):
                print 'update'
                self.updataOrder(mongoOrder, TurnDate, ShipmentDate, OrderNo, PartName, \
                                 PartQuility, PartTotalPrice, \
                                 firm, supplier)

            else:
                print 'insert'
                Ordernum = OrderNo
                self.insertOrder(mongoOrder, TurnDate, ShipmentDate, OrderNo, PartName, \
                                 PartQuility, PartTotalPrice, \
                                 firm, supplier)
            if (Clientnum == ClientName):
                print 'update'
                self.updataClient(mongodbClient, OrderNo, ClientName, \
                                  ClientTel, ClientPhone, \
                                  firm, supplier)
            else:
                print 'insert'
                Clientnum = ClientName
                self.insertClient(mongodbClient, OrderNo, ClientName, \
                                  ClientTel, ClientPhone, \
                                  firm, supplier)

        mysqlconnect.dbClose()
        mongoOrder.dbClose()
        mongodbClient.dbClose()
        logging.info('===Yahood_Data SUCCESS===')
        return 'success'

    # mongoDB storage   ç¬¬ä??å??¸æ¯ä¸ä??¢ç?mongoOrder or mongoClient
    def insertOrder(self, mongoOrder, _TurnDate, _ShipmentDate, _OrderNo, _PartName, \
                    _PartQuility, _PartTotalPrice, \
                    _firm, _supplier):
        businessorder_doc = {'TurnDate': _TurnDate, 'ShipmentDate': _ShipmentDate, 'OrderNo': _OrderNo, \
                             'PartName': [_PartName], \
                             'PartQuility': [_PartQuility], 'Price': [_PartTotalPrice], \
                             'firm': [_firm], 'supplier': [_supplier]}

        mongoOrder.cursor.insert(businessorder_doc)

    def updataOrder(self, mongoOrder, _TurnDate, _ShipmentDate, _OrderNo, _PartName, \
                    _PartQuility, _PartTotalPrice, \
                    _firm, _supplier):
        mongoOrder.cursor.update({'TurnDate': _TurnDate, "OrderNo": _OrderNo, 'firm': _firm, 'supplier': _supplier}, \
                                 {'$push': {'ShipmentDate': _ShipmentDate, 'PartName': _PartName, \
                                            'PartQuility': _PartQuility, 'Price': _PartTotalPrice}})

    def insertClient(self, mongodbClient, _OrderNo, _ClientName, _ClientTel, \
                     _ClientPhone, _firm, _supplier):
        businessorder_doc = {'OrderNo': _OrderNo, 'ClientName': [_ClientName], 'ClientTel': [_ClientTel],
                             'ClientPhone': [_ClientPhone], \
                             'firm': [_firm], 'supplier': [_supplier]}

        mongodbClient.cursor.insert(businessorder_doc)

    def updataClient(self, mongodbClient, _OrderNo, _ClientName, _ClientTel, \
                     _ClientPhone, _firm, _supplier):
        mongodbClient.cursor.update({'ClientName': _ClientName, 'ClientTel': _ClientTel, 'ClientPhone': _ClientPhone, \
                                     'firm': _firm, 'supplier': _supplier} \
                                    , {'$push': {"OrderNo": _OrderNo}})