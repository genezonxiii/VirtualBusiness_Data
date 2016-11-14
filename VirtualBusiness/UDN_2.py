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

class UDN_Data2():
    Data=None
    def __init__(self):
        pass
    def UDN_Data2(self,supplier,GroupID,path,UserID):
        logging.basicConfig(filename='pyupload.log', level=logging.DEBUG, format='%(asctime)s %(message)s', datefmt='%Y/%m/%d %I:%M:%S %p')
        logging.Formatter.converter = time.gmtime
        logging.info('===UDN_Data2===')
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

        data=xlrd.open_workbook(path)
        table=data.sheets()[0]
        num_cols=table.ncols
        #put the data into the corresponding variable
        for row_index in range(3,table.nrows):
            for col_index in range(0,num_cols):
                aes=aes_data()
                OrderNo = str(table.cell(row_index, 7).value)


                #strTurnDate = str(table.cell(row_index, 2).value)
                #_strTurnDate = datetime.datetime.strptime(strTurnDate,'%Y/%m/%d %H:%M ')
                #TurnDate = datetime.datetime.strftime(_strTurnDate,'%Y-%m-%d %H:%M')
                TurnDate = datetime.datetime.strptime(str(table.cell_value(row_index, 9)),'%Y/%m/%d')
                ShipmentDate = datetime.datetime.strptime(str(table.cell_value(row_index, 1)),'%Y/%m/%d')

                Name = table.cell(row_index, 10).value
                ClientName = aes.AESencrypt("p@ssw0rd", Name.encode('utf-8'), True)

                PartNo = str(table.cell(row_index, 15).value).split('.')[0]
                PartName = table.cell(row_index,19).value

                PartQuility = table.cell(row_index, 21).value
                PartCost = table.cell(row_index, 23).value
                PartTotalPrice = table.cell(row_index, 25).value
                # strShipmentDate = xlrd.xldate_as_tuple(table.cell_value(row_index, 10), data.datemode)
                # ShipmentDate = datetime(*strShipmentDate[0:6]).strftime('%Y-%m-%d')

                firm = GroupID
                supplier = supplier
                UserID = UserID
            # SupplySQL = (str(uuid.uuid4()),GroupID, supplier,"","","","","","","","","","","")
            ProductSQL = (str(uuid.uuid4()), GroupID, PartNo, PartName,supplier, None,None,PartCost,PartTotalPrice,0,None,None,None,None)
            CustomereSQL = (str(uuid.uuid4()), GroupID, ClientName, None, None, None,None,None,None,None)

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
            Name_compare = []
            if result != []:
                for x in result:
                    Name_compare.append(aes.AESdecrypt("p@ssw0rd", x[1], True))
                if Name.encode('utf-8') in Name_compare:
                    print "update"
                    mysqlconnect.cursor.execute(CustomereSQLupd, (
                        None, None, None, None, None, None, None, GroupID, x[1]))
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
                if aes.AESdecrypt('p@ssw0rd', x[1], True) == Name.encode('utf-8'):
                    mysqlconnect.cursor.execute(SalestrSQLsel, (str(x[1]),))
                    customer_id_temp.append(mysqlconnect.cursor.fetchall()[0])
            print customer_id_temp
            for y in customer_id_temp:
                for x in result:
                    if aes.AESdecrypt('p@ssw0rd', x[1], True) == Name.encode('utf-8'):
                        SaleSQL = (
                            GroupID, OrderNo, UserID, PartName, PartNo, y[0],
                            x[1], PartQuility, PartTotalPrice, None, None, TurnDate, None,
                            None, None, supplier)
                        mysqlconnect.cursor.callproc('p_tb_sale', SaleSQL)

            mysqlconnect.db.commit()
            # mysqlconnect.cursor.callproc('p_tb_customer', CustomereSQL)

            if (Ordernum==OrderNo[0:13]):
                print 'update'
                self.updataOrder(mongoOrder,TurnDate,ShipmentDate,OrderNo,PartName,\
                                PartQuility,PartTotalPrice,\
                                firm,supplier)

            else:
                print 'insert'
                Ordernum=OrderNo
                self.insertOrder(mongoOrder,TurnDate,ShipmentDate,OrderNo,PartName,\
                                PartQuility,PartTotalPrice,\
                                firm,supplier)
            if (Clientnum == ClientName):
                print 'update'
                self.updataClient(mongodbClient, OrderNo,ClientName,\
                            firm,supplier)
            else:
                print 'insert'
                Clientnum = ClientName
                self.insertClient(mongodbClient, OrderNo,ClientName,\
                               firm,supplier)

        mysqlconnect.dbClose()
        mongoOrder.dbClose()
        mongodbClient.dbClose()
        logging.info('===UDN_Data2 SUCCESS===')
        return 'success'

    # mongoDB storage   ?ï¿½ï¿½??????è±¢ï¿½??ï¿½ï¿½???ï¿?mongoOrder or mongoClient
    def insertOrder(self,mongoOrder,_TurnDate,_ShipmentDate,_OrderNo,_PartName,\
                                _PartQuility,_PartTotalPrice,\
                                _firm,_supplier):
        businessorder_doc={'TurnDate':_TurnDate,'ShipmentDate':_ShipmentDate,'OrderNo':_OrderNo,\
                           'PartName':[_PartName],\
                           'PartQuility':[_PartQuility],'Price':[_PartTotalPrice],\
                           'firm':[_firm],'supplier':[_supplier]}

        mongoOrder.cursor.insert(businessorder_doc)
    def updataOrder(self,mongoOrder,_TurnDate,_ShipmentDate,_OrderNo,_PartName,\
                                _PartQuility,_PartTotalPrice,\
                                _firm,_supplier):
        mongoOrder.cursor.update( {'TurnDate':_TurnDate, "OrderNo" : _OrderNo,'firm':_firm,'supplier':_supplier},\
                           {'$push':{'ShipmentDate':_ShipmentDate,'PartName':_PartName,\
                           'PartQuility':_PartQuility,'Price':_PartTotalPrice}})
    def insertClient(self,mongodbClient,_OrderNo,_ClientName,\
                   _firm,_supplier):
        businessorder_doc={'OrderNo':_OrderNo,'ClientName':[_ClientName],\
                           'firm':[_firm],'supplier':[_supplier]}

        mongodbClient.cursor.insert(businessorder_doc)
    def updataClient(self,mongodbClient,_OrderNo,_ClientName,\
                   _firm,_supplier):
        mongodbClient.cursor.update({ 'ClientName':_ClientName,\
                            'firm':_firm,'supplier':_supplier}\
                           ,{'$push':{"OrderNo" : _OrderNo}})







