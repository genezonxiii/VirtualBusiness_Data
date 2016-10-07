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

class UMall_Data():
    Data=None
    def __init__(self):
        pass
    def UMall_Data(self,supplier,GroupID,path,UserID):
        logging.basicConfig(filename='pyupload.log', level=logging.DEBUG, format='%(asctime)s %(message)s', datefmt='%Y/%m/%d %I:%M:%S %p')
        logging.Formatter.converter = time.gmtime
        logging.info('===UMall_Data===')
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

        for row_index in range(1,table.nrows):
            for col_index in range(0,num_cols):
                aes=aes_data()
                OrderNo = table.cell(row_index, 1).value[0:13]
                #strShipmentDate = str(table.cell(row_index, 22).value).replace("/", "-")
                #ShipmentDate = datetime.strptime(strShipmentDate, '%Y-%m-%d')
                strShipmentDate = xlrd.xldate_as_tuple(table.cell_value(row_index, 22), data.datemode)
                _ShipDate = datetime(*strShipmentDate[0:6]).strftime('%Y-%m-%d')
                ShipmentDate = datetime.strptime(_ShipDate, '%Y-%m-%d')
                Name = table.cell(row_index, 15).value
                ClientName = aes.AESencrypt("p@ssw0rd", Name.encode('utf-8'), True)
                Tel = table.cell(row_index, 17).value
                ClientTel = aes.AESencrypt("p@ssw0rd", Tel, True)
                Phone = table.cell(row_index, 16).value
                ClientPhone = aes.AESencrypt("p@ssw0rd", Phone, True)
                Add = table.cell(row_index, 18).value
                ClientAdd = aes.AESencrypt("p@ssw0rd", Add.encode('utf-8'), True)
                PartNo = str(table.cell(row_index, 6).value)
                PartName = table.cell(row_index, 7).value
                PartQuility = table.cell(row_index, 12).value
                PartPrice = table.cell(row_index, 13).value
                PartCost = table.cell(row_index, 14).value
                InvoiceNo = table.cell(row_index, 25).value
                #strInvoiceDate = str(table.cell(row_index, 27).value).replace("/", "-")
                #InvoiceDate = datetime.strptime(strInvoiceDate, '%Y-%m-%d')
                strInvoiceDate = xlrd.xldate_as_tuple(table.cell_value(row_index, 27), data.datemode)
                _InvoiceDate = datetime(*strInvoiceDate[0:6]).strftime('%Y-%m-%d')
                InvoiceDate = datetime.strptime(_InvoiceDate, '%Y-%m-%d')

                firm = GroupID
                supplier = supplier
                UserID = UserID
            # SupplySQL = (str(uuid.uuid4()),GroupID, supplier,"","","","","","","","","","","")
            ProductSQL = (str(uuid.uuid4()), GroupID, PartNo, PartName,supplier, "",'',PartCost,PartPrice,0,None,None,None,None)
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
                             GroupID, OrderNo,UserID, PartName, PartNo, y[0], x[1],
                            PartQuility,
                            0, None, ShipmentDate, ShipmentDate, ShipmentDate,None, ShipmentDate,supplier)
                        mysqlconnect.cursor.callproc('p_tb_sale', SaleSQL)
                        # SalestrSQL = "INSERT INTO tb_sale (sale_id,seq_no,group_id,user_id,customer_id,name,c_product_id,quantity,price,invoice,invoice_date,sale_date,product_name)"\
                        # "VALUES(%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)"
                        # mysqlconnect.cursor.execute(SalestrSQL,(str(uuid.uuid4()), OrderNo, GroupID, '72345', y[0], x[1], PartNo, PartQuility,PartPrice, InvoiceNo,InvoiceDate,ShipmentDate,PartName))
            mysqlconnect.db.commit()
            # mysqlconnect.cursor.callproc('p_tb_customer', CustomereSQL)

            if (Ordernum==OrderNo[0:13]):
                print 'update'
                self.updataOrder(mongoOrder,OrderNo,ShipmentDate,PartNo,PartName,PartQuility,PartPrice,PartCost,InvoiceNo,InvoiceDate,firm,supplier)

            else:
                print 'insert'
                Ordernum=OrderNo
                self.insertOrder(mongoOrder,OrderNo,ShipmentDate,PartNo,PartName,PartQuility,PartPrice,PartCost,InvoiceNo,InvoiceDate,firm,supplier)
            if (Clientnum == ClientName):
                print 'update'
                self.updataClient(mongodbClient, OrderNo,ClientName,ClientAdd,ShipmentDate,firm,supplier)
            else:
                print 'insert'
                Clientnum = ClientName
                self.insertClient(mongodbClient, OrderNo,ClientName,ClientAdd,ShipmentDate,firm,supplier)

        mysqlconnect.dbClose()
        mongoOrder.dbClose()
        mongodbClient.dbClose()
        logging.info('===UMall_Data SUCCESS===')
        return 'success'

    # mongoDB storage   第�??��??�是丟�??��?mongoOrder or mongoClient
    def insertOrder(self,mongoOrder,_OrderNo,_ShipmentDate,_PartNo,_PartName,_PartQuility,_PartPrice,_PartCost,_InvoiceNo,_InvoiceDate,_firm,_supplier):
        businessorder_doc={ 'OrderNo':_OrderNo,'ShipmentDate':_ShipmentDate,'PartNo':[_PartNo],'PartName':[_PartName],\
                           'PartQuility':[_PartQuility],'Price':[_PartPrice],'PartCost':[_PartCost],'InvoiceNo':[_InvoiceNo],\
                           'InvoiceDate':_InvoiceDate,'firm':[_firm],'supplier':[_supplier]
                            }

        mongoOrder.cursor.insert(businessorder_doc)
    def updataOrder(self,mongoOrder,_OrderNo,_ShipmentDate,_PartNo,_PartName,_PartQuility,_PartPrice,_PartCost,_InvoiceNo,_InvoiceDate,_firm,_supplier):
        mongoOrder.cursor.update({ "OrderNo" : _OrderNo,'InvoiceNo':_InvoiceNo,\
                           'InvoiceDate':_InvoiceDate,'firm':_firm,'supplier':_supplier},{'$push':{'ShipmentDate':_ShipmentDate ,'PartNo':_PartNo,'PartName':_PartName,\
                           'PartQuility':_PartQuility,'Price':_PartPrice,'PartCost':[_PartCost],}}
                                                )
    def insertClient(self,mongodbClient,_OrderNo,_ClientName,_ClientAdd,_ShipmentDate,_firm,_supplier):

        businessorder_doc={ 'OrderNo':_OrderNo,'ClientName':[_ClientName],'ClientAdd':[_ClientAdd],'ShipmentDate':_ShipmentDate,'firm':[_firm],'supplier':[_supplier]}

        mongodbClient.cursor.insert(businessorder_doc)
    def updataClient(self,mongodbClient,_OrderNo,_ClientName,_ClientAdd,_ShipmentDate,_firm,_supplier):
        mongodbClient.cursor.update({ 'ClientName':_ClientName,'ClientAdd':_ClientAdd,'firm':_firm,'supplier':_supplier}\
                           ,{'$push':{"OrderNo" : _OrderNo,'ShipmentDate':_ShipmentDate }})







