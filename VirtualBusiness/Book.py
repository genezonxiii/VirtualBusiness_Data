# -*-  coding: utf-8  -*-
__author__ = '10409003'
import xlrd
import uuid
from datetime import datetime
from aes_data import aes_data
from ToMongodb import ToMongodb
from ToMysql import ToMysql

class Book_Data():
    Data=None
    def __init__(self):
        pass
    def Book_Data(self,supplier,GroupID,path,UserID):
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
        print path

        data=xlrd.open_workbook(path)
        table=data.sheets()[0]
        num_cols=table.ncols
        #put the data into the corresponding variable

        for row_index in range(3,table.nrows):
            for col_index in range(0,num_cols):
                aes=aes_data()
                PartName = table.cell(row_index, 0).value
                PartNo = table.cell(row_index, 1).value
                PartQuility = table.cell(row_index, 2).value
                firm = GroupID
                supplier = supplier
                UserID=UserID
            # SupplySQL = (str(uuid.uuid4()),GroupID, supplier,"","","","","","","","","","","")
            ProductSQL = (str(uuid.uuid4()), GroupID, PartNo, PartName,supplier, "",'',0,0,0,None,None,None,None)
            # SalestrSQLinsert = "INSERT INTO tb_sale (sale_id,seq_no,group_id,user_id,customer_id,c_product_id,quantity,product_name)" \
            #              "VALUES(%s,%s,%s,%s,%s,%s,%s,%s)"
            # SalestrSQL = (str(uuid.uuid4()), PartNo, GroupID, '12345','NULL',PartNo,PartQuility,PartName)
            SaleSQL = (
                 GroupID, None,UserID, PartName, PartNo, '',
                '', PartQuility, 0, None, '2016-06-06', '2016-06-06', '2016-06-06',
                None, '2016-06-06',supplier)

            # mysqlconnect.cursor.callproc('p_tb_supply', SupplySQL)
            mysqlconnect.cursor.callproc('p_tb_product', ProductSQL)
            mysqlconnect.cursor.callproc('p_tb_sale', SaleSQL)
            # mysqlconnect.cursor.execute(SalestrSQLinsert,SalestrSQL)



            mysqlconnect.db.commit()
            # mysqlconnect.cursor.callproc('p_tb_customer', CustomereSQL)



            self.insertOrder(mongoOrder,PartName,PartNo,\
                            PartQuility,firm,supplier)

        mysqlconnect.dbClose()
        mongoOrder.dbClose()
        mongodbClient.dbClose()
        return 'success'

    # mongoDB storage   第一個參數是丟上面的mongoOrder or mongoClient
    def insertOrder(self,mongoOrder,_PartName,_PartNo,\
                                _PartQuility,\
                                _firm,_supplier):
        businessorder_doc={
                           'PartName':[_PartName],'PartNo':[_PartNo],\
                           'PartQuility':[_PartQuility],
                           'firm':[_firm],'supplier':[_supplier]
                            }

        mongoOrder.cursor.insert(businessorder_doc)







