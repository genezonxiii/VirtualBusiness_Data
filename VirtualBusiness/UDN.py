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

class UDN_Data():
    Data=None
    def __init__(self):
        pass
    def UDN_Data(self,supplier,GroupID,path,UserID):
        logging.basicConfig(filename='pyupload.log', level=logging.DEBUG, format='%(asctime)s %(message)s', datefmt='%Y/%m/%d %I:%M:%S %p')
        logging.Formatter.converter = time.gmtime
        logging.info('===UDN_Data===')
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

        for row_index in range(2,table.nrows):
            for col_index in range(0,num_cols):
                aes=aes_data()
                print table.cell_value(row_index, 3)
                # strShipmentDate = xlrd.xldate_as_tuple(table.cell_value(row_index, 1), data.datemode)
                # ShipmentDate = datetime(*strShipmentDate[0:6]).strftime('%Y-%m-%d')
                # strTurnDate = xlrd.xldate_as_tuple(table.cell_value(row_index, 4), data.datemode)
                # TurnDate = datetime(*strTurnDate[0:6]).strftime('%Y-%m-%d')
                strTurnDate = str(table.cell(row_index, 1).value).replace("/", "-")
                TurnDate = datetime.strptime(strTurnDate, '%Y-%m-%d')
                strShipmentDate = str(table.cell(row_index, 12).value).replace("/", "-")
                ShipmentDate = datetime.strptime(strShipmentDate, '%Y-%m-%d')
                OrderNo = str(table.cell(row_index, 0).value)
                #Name = table.cell(row_index, 6).value
                #ClientName = aes.AESencrypt("p@ssw0rd", Name, True)
                #Phone = table.cell(row_index, 7).value
                #ClientPhone = aes.AESencrypt("p@ssw0rd", Phone, True)
                #Tel = table.cell(row_index, 8).value
                #ClientTel = aes.AESencrypt("p@ssw0rd", Tel, True)
                #ClientZipCode = table.cell(row_index, 9).value
                #Add = table.cell(row_index, 10).value
                #ClientAdd = aes.AESencrypt("p@ssw0rd", Add, True)
                PartNo = str(table.cell(row_index, 2).value)
                PartName = table.cell(row_index, 6).value
                PartQuility = table.cell(row_index, 7).value
                PartTotalPrice = table.cell(row_index, 10).value
                PartCost = table.cell(row_index, 9).value
                firm = GroupID
                supplier = supplier
                UserID = UserID
            # SupplySQL = (str(uuid.uuid4()),GroupID, supplier,"","","","","","","","","","","")
            ProductSQL = (str(uuid.uuid4()), GroupID, PartNo, PartName,supplier, "",'',PartCost,PartTotalPrice,0,None,None,None,None)
            #CustomereSQL = (str(uuid.uuid4()), GroupID, ClientName, ClientAdd, ClientTel, ClientPhone,None,ClientZipCode,None,None)
            SaleSQL = (GroupID, OrderNo, UserID, PartName, PartNo, None,None, PartQuility, PartTotalPrice,None, None, TurnDate, None,
                       None, ShipmentDate,supplier)

            # mysqlconnect.cursor.callproc('p_tb_supply', SupplySQL)
            mysqlconnect.cursor.callproc('p_tb_product', ProductSQL)
            mysqlconnect.cursor.callproc('p_tb_sale', SaleSQL)

            mysqlconnect.db.commit()
            # mysqlconnect.cursor.callproc('p_tb_customer', CustomereSQL)

            if (Ordernum==OrderNo[0:13]):
                print 'update'
                self.updataOrder(mongoOrder,ShipmentDate,OrderNo,PartNo,PartName,\
                                PartQuility,PartTotalPrice,PartCost,\
                                firm,supplier)

            else:
                print 'insert'
                Ordernum=OrderNo
                self.insertOrder(mongoOrder,ShipmentDate,OrderNo,PartNo,PartName,\
                                PartQuility,PartTotalPrice,PartCost,\
                                firm,supplier)
           

        mysqlconnect.dbClose()
        mongoOrder.dbClose()
        mongodbClient.dbClose()
        logging.info('===UDN_Data SUCCESS===')
        return 'success'


    def insertOrder(self,mongoOrder,_ShipmentDate,_OrderNo,_PartNo,_PartName,\
                                _PartQuility,_PartTotalPrice,_PartCost,\
                                _firm,_supplier):
        businessorder_doc={'ShipmentDate':_ShipmentDate,'OrderNo':_OrderNo,\
                           'PartNo':[_PartNo],'PartName':[_PartName],\
                           'PartQuility':[_PartQuility],'PartTotalPrice':[_PartTotalPrice],\
                           'Price': [_PartCost],'firm':[_firm],'supplier':[_supplier]}

        mongoOrder.cursor.insert(businessorder_doc)
    def updataOrder(self,mongoOrder,_ShipmentDate,_OrderNo,_PartNo,_PartName,\
                                _PartQuility,_PartTotalPrice,_PartCost,\
                                _firm,_supplier):
        mongoOrder.cursor.update( { "OrderNo" : _OrderNo,'firm':_firm,'supplier':_supplier},\
                           {'$push':{'ShipmentDate':_ShipmentDate,'PartNo':_PartNo ,'PartName':_PartName,\
                           'PartQuility':_PartQuility,'PartTotalPrice':_PartTotalPrice,'Price': _PartCost}})
  



