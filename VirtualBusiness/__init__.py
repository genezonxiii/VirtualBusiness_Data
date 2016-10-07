# -*-  coding: utf-8  -*-
__author__ = '10409003'
from pymongo import MongoClient
import xlrd
from datetime import datetime
from aes_data import aes_data
import uuid
import ToMysql
import ToMongodb

class ASAP():
    Data=None
    def __init__(self):
        pass

    def ASAP_Data(self,supplier,GroupID,path,collection_order,collection_client):
        #mysql connector object
        mysqlconnect=ToMysql()
        mysqlCursor=mysqlconnect.connect()

        #mongodb connector object儲存  先生出一個object 連線後 用 object.cursor.接你處理ＣＲＵＤ
        mongoOrder=ToMongodb()
        mongoOrder.setcollection('co_ordertest')
        mongoOrder.connect()
        mongoOrder.cursor.find()

        mongodbClient=ToMongodb()
        mongodbClient.setcollection('co_co_clienttest')
        mongodbClient.connect()
        mongodbClient.cursor.updateOne()

        Ordernum=""
        Clientnum=""

        data=xlrd.open_workbook(path)
        table=data.sheets()[0]
        num_cols=table.ncols
        #put the data into the corresponding variable
        for row_index in range(1,table.nrows):
            for col_index in range(0,num_cols):
                aes=aes_data()
                strTurntDate    =str(table.cell(row_index,0).value).replace("/","-")
                TurntDate       =datetime.strptime(strTurntDate,'%Y-%m-%d')
                OrderNo         =table.cell(row_index,1).value[0:13]
                PartNum         =str(table.cell(row_index,2).value)
                PartMaterial    =str(table.cell(row_index,3).value)
                PartName        =table.cell(row_index,4).value
                PartColor       =table.cell(row_index,5).value
                PartSize        =table.cell(row_index,6).value
                PartQuility     =table.cell(row_index,7).value
                Name            =table.cell(row_index, 8).value
                ClientName      = aes.AESencrypt("password", Name.encode('utf-8'), True)
                Phone           = table.cell(row_index, 9).value
                ClientPhone     = aes.AESencrypt("password", Phone.encode('utf-8'), True)
                Tel             = table.cell(row_index, 10).value
                ClientTel       = aes.AESencrypt("password", Tel.encode('utf-8'), True)
                Add             = table.cell(row_index, 11).value
                ClientAdd       = aes.AESencrypt("password", Add.encode('utf-8'), True)
                PartNo          =str(table.cell(row_index,12).value)
                firm            =table.cell(row_index,13).value
                supplier        =supplier
                GroupID         =GroupID

            # if (Ordernum==OrderNo[0:13]):
            #     print 'update'
            #     self.updataOrder(mongoOrder,TurntDate,OrderNo,PartNum,PartMaterial,PartName,PartColor,PartSize,PartQuility,PartNo,firm,supplier)

                mysqlCursor.execute(self.ToSql(GroupID,OrderNo,PartNo,PartQuility,TurntDate,PartName,PartSize,supplier,ClientName,ClientAdd,ClientPhone,ClientTel,1))
                mysqlCursor.callproc(self.ToSql())

            # else:
            #     print 'insert'
            #     Ordernum=OrderNo
            #     self.insertOrder(mongoOrder,TurntDate,OrderNo,PartNum,PartMaterial,PartName,PartColor,PartSize,PartQuility,PartNo,firm,supplier)
            # if (Clientnum == ClientName):
            #     print 'update'
            #     self.updataClient(mongodbClient, OrderNo, ClientName, ClientPhone, ClientTel, ClientAdd, firm, supplier)
            # else:
            #     print 'insert'
            #     Clientnum = ClientName
            #     self.insertClient(mongodbClient, OrderNo, ClientName, ClientPhone, ClientTel, ClientAdd, firm, supplier)

        mysqlCursor.close()
        mongoOrder.close()
        mongodbClient.close()

    #mongoDB storage   第一個參數是丟上面的mongoOrder or mongoClient
    # def insertOrder(self,mongoOrder,_TurntDate,_OrderNo,_PartNum,_PartMaterial,_PartName,_PartColor,_PartSize,_PartQuility,_PartNo,_firm,_supplier):
    #     businessorder_doc={ 'TurntDate':_TurntDate,
    #                         'OrderNo':_OrderNo,
    #                         'PartNum':[_PartNum],
    #                         'PartMaterial':[_PartMaterial],\
    #                         'PartName':[_PartName],
    #                         'PartColor':[_PartColor],
    #                         'PartSize':[_PartSize],\
    #                         'PartQuility':[_PartQuility],
    #                         'PartNo':[_PartNo],
    #                         'firm':[_firm],
    #                         'supplier':[_supplier]
    #                         }
    #
    #     mongoOrder.cursor.insert(businessorder_doc)
    # def updataOrder(self,mongoOrder,_TurntDate,_OrderNo,_PartNum,_PartMaterial,_PartName,_PartColor,_PartSize,_PartQuility,_PartNo,_firm,_supplier):
    #     mongoOrder.cursor.update( { "OrderNo" : _OrderNo,
    #                                 'firm':_firm,'supplier':_supplier},\
    #                                 {'$push':{  'TurntDate':_TurntDate,
    #                                             'PartNum':_PartNum ,
    #                                             'PartMaterial':_PartMaterial,
    #                                             'PartName':_PartName,\
    #                                             'PartColor':_PartColor,
    #                                             'PartSize':_PartSize,
    #                                             'PartNo':_PartNo,
    #                                             'PartQuility':_PartQuility}
    #                                             }
    #                                             )
    # def insertClient(self,mongodbClient,_OrderNo,_ClientName,_ClientPhone,_ClientTel,_ClientAdd,_firm,_supplier):
    #
    #     businessorder_doc={ 'OrderNo':_OrderNo,
    #                         'ClientName':[_ClientName],
    #                         'ClientPhone':[_ClientPhone],
    #                         'ClientTel':[_ClientTel],
    #                         'ClientAdd':[_ClientAdd],
    #                         'firm':[_firm],
    #                         'supplier':[_supplier]}
    #
    #     mongodbClient.cursor.insert(businessorder_doc)
    # def updataClient(self,mongodbClient,_OrderNo,_ClientName,_ClientPhone,_ClientTel,_ClientAdd,_firm,_supplier):
    #     mongodbClient.cursor.update( {  'ClientName':_ClientName,
    #                                     'ClientPhone':_ClientPhone,
    #                                     'ClientTel':[_ClientTel],
    #                                     'ClientAdd':_ClientAdd,
    #                                     'firm':_firm,
    #                                     'supplier':_supplier}\
    #                                     ,{'$push':{"OrderNo" : _OrderNo}})


    #The judgement of mysql sql by sqlselect parameter
    def ToSql(self,GroupID,OrderNo,PartNo,PartQuility,TurntDate,PartName,PartSize,supplier,ClientName,ClientAdd,ClientPhone,ClientTel,sqlselect):
        SupplySQL=("INSERT INTO tb_sale (supply_id,group_id,supply_name)"
                  "VALUES(%s,%s,%s)", str(uuid.uuid4()),GroupID,supplier)
        SalestrSQL =("INSERT INTO tb_sale (sale_id,group_id,seq_no,c_product_id,customer_id,name,quantity,trans_list_date,sale_date)"
                  "VALUES(%s,%s,%s,%s,%s,%s,%s)", str(uuid.uuid4()),GroupID,OrderNo, PartNo, PartQuility,TurntDate,TurntDate)
        ProductSQL=("INSERT INTO tb_product (product_id,group_id,c_product_id,product_name,unit_id,supply_name)"
                  "VALUES(%s,%s,%s,%s,%s)", str(uuid.uuid4()),GroupID, PartNo, PartName,PartSize,supplier)
        CustomereSQL=("INSERT INTO tb_customer (product_id,group_id,name,address,phone,mobile)"
                  "VALUES(%s,%s,%s,%s,%s,%s)", str(uuid.uuid4()),GroupID, ClientName, ClientAdd,ClientTel,ClientPhone)

        if sqlselect==1:
            sql = SupplySQL
        elif sqlselect==2:
            sql = SalestrSQL
        elif sqlselect == 3:
            sql=ProductSQL
        elif sqlselect==4:
            sql = CustomereSQL
        else:
            print "Sqlselect error!"
        return sql