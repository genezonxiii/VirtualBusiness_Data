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

class Momo_Data():
    Data=None
    def __init__(self):
        pass
    def Momo_Data(self,supplier,GroupID,path,UserID):
        logging.basicConfig(filename='pyupload.log', level=logging.DEBUG, format='%(asctime)s %(message)s', datefmt='%Y/%m/%d %I:%M:%S %p')
        logging.Formatter.converter = time.gmtime
        logging.info('===Momo_Data===')
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
                #strTurnDate = str(table.cell(row_index, 8).value).replace("/", "-")
                #TurnDate = datetime.strptime(strTurnDate, '%Y-%m-%d')
                #strOrderDate= xlrd.xldate_as_tuple(table.cell_value(row_index, 6), data.datemode)
                #_OrderDate = datetime(*strOrderDate[0:6]).strftime('%Y-%m-%d')
                #TurnDate = datetime.strptime(_OrderDate, '%Y-%m-%d')
                #strShipmentDate = str(table.cell(row_index, 9).value).replace("/", "-")
                #ShipmentDate = datetime.strptime(strShipmentDate, '%Y-%m-%d')
                #strTurnDate = xlrd.xldate_as_tuple(table.cell_value(row_index, 7), data.datemode)
                #_TurnDate = datetime(*strTurnDate[0:6]).strftime('%Y-%m-%d')
                #ShipmentDate = datetime.strptime(_TurnDate, '%Y-%m-%d')
                TurnDate = datetime.datetime.strptime(str(table.cell_value(row_index, 6)),'%Y/%m/%d %H:%M')
                ShipmentDate = datetime.datetime.strptime(str(table.cell_value(row_index, 7)),'%Y/%m/%d')
                InvoiceDate = datetime.datetime.strptime(str(table.cell_value(row_index, 20)),'%Y/%m/%d')

                Name = table.cell(row_index, 8).value
                ClientName = aes.AESencrypt("p@ssw0rd", Name.encode('utf-8'), True)
                #Tel = table.cell(row_index, 11).value
                #ClientTel = aes.AESencrypt("p@ssw0rd", Tel, True)
                #Phone = table.cell(row_index, 12).value
                #ClientPhone = aes.AESencrypt("p@ssw0rd", Phone, True)
                #Add = table.cell(row_index, 13).value
                #ClientAdd = aes.AESencrypt("p@ssw0rd", Add.encode('utf-8'), True)
                PartNo = str(table.cell(row_index, 10).value)
                PartName = table.cell(row_index, 11).value
                PartQuility = table.cell(row_index, 14).value
                PartPrice = table.cell(row_index, 15).value
                InvoiceNo = table.cell(row_index, 19).value
                #strInvoiceDate = str(table.cell(row_index, 24).value).replace("/", "-")
                #InvoiceDate = datetime.strptime(strInvoiceDate, '%Y-%m-%d')
                #sstrTurnDate = xlrd.xldate_as_tuple(table.cell_value(row_index, 20), data.datemode)
                #_sTurnDate = datetime(*sstrTurnDate[0:6]).strftime('%Y-%m-%d')
                #InvoiceDate = datetime.strptime(_sTurnDate, '%Y-%m-%d')
                

                firm = GroupID
                supplier = supplier
                UserID = UserID
            # SupplySQL = (str(uuid.uuid4()),GroupID, supplier,"","","","","","","","","","","")
            ProductSQL = (str(uuid.uuid4()), GroupID, PartNo, PartName,supplier,None,None,0,PartPrice,0,None,None,None,None)
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
            Name_compare=[]
            if result!=[]:
                for x in result:
                    Name_compare.append(aes.AESdecrypt("p@ssw0rd",x[1], True))
                if Name.encode('utf-8') in Name_compare :
                    print "update"
                    mysqlconnect.cursor.execute(CustomereSQLupd,(None, None, None,None,None,None,None,GroupID,x[1]))
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
                            x[1], PartQuility, PartPrice,InvoiceNo, InvoiceDate, TurnDate, None,
                            None, ShipmentDate,supplier)
                        mysqlconnect.cursor.callproc('p_tb_sale', SaleSQL)
                        # SalestrSQL = "INSERT INTO tb_sale (sale_id,seq_no,group_id,user_id,product_name,c_product_id,customer_id,name,quantity,price,invoice,invoice_date,trans_list_date,sale_date)"\
                        # "VALUES(%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)"
                        # mysqlconnect.cursor.execute(SalestrSQL,(str(uuid.uuid4()), OrderNo, GroupID, '72345',PartName , PartNo, y[0], x[1], PartQuility,PartPrice, InvoiceNo,InvoiceDate,TurnDate,ShipmentDate))
            mysqlconnect.db.commit()
            # mysqlconnect.cursor.callproc('p_tb_customer', CustomereSQL)

            if (Ordernum==OrderNo[0:13]):
                print 'update'
                self.updataOrder(mongoOrder,OrderNo,ShipmentDate,PartNo,PartName,PartQuility,PartPrice,InvoiceNo,InvoiceDate,firm,supplier)

            else:
                print 'insert'
                Ordernum=OrderNo
                self.insertOrder(mongoOrder,OrderNo,ShipmentDate,PartNo,PartName,PartQuility,PartPrice,InvoiceNo,InvoiceDate,firm,supplier)
            if (Clientnum == ClientName):
                print 'update'
                self.updataClient(mongodbClient, OrderNo,ClientName,ShipmentDate,firm,supplier)
            else:
                print 'insert'
                Clientnum = ClientName
                self.insertClient(mongodbClient, OrderNo,ClientName,ShipmentDate,firm,supplier)

        mysqlconnect.dbClose()
        mongoOrder.dbClose()
        mongodbClient.dbClose()
        logging.info('===Momo_Data SUCCESS===')
        return 'success'

    # mongoDB storage   第�??��??�是丟�??��?mongoOrder or mongoClient
    def insertOrder(self,mongoOrder,_OrderNo,_ShipmentDate,_PartNo,_PartName,_PartQuility,_PartPrice,_InvoiceNo,_InvoiceDate,_firm,_supplier):
        businessorder_doc={ 'OrderNo':_OrderNo,'ShipmentDate':_ShipmentDate,'PartNo':[_PartNo],'PartName':[_PartName],\
                           'PartQuility':[_PartQuility],'Price':[_PartPrice],'InvoiceNo':[_InvoiceNo],\
                           'InvoiceDate':_InvoiceDate,'firm':[_firm],'supplier':[_supplier]
                            }

        mongoOrder.cursor.insert(businessorder_doc)
    def updataOrder(self,mongoOrder,_OrderNo,_ShipmentDate,_PartNo,_PartName,_PartQuility,_PartPrice,_InvoiceNo,_InvoiceDate,_firm,_supplier):
        mongoOrder.cursor.update({ "OrderNo" : _OrderNo,'InvoiceNo':_InvoiceNo,\
                           'InvoiceDate':_InvoiceDate,'firm':_firm,'supplier':_supplier},{'$push':{'ShipmentDate':_ShipmentDate ,'PartNo':_PartNo,'PartName':_PartName,\
                           'PartQuility':_PartQuility,'Price':_PartPrice}}
                                                )
    def insertClient(self,mongodbClient,_OrderNo,_ClientName,_ShipmentDate,_firm,_supplier):

        businessorder_doc={ 'OrderNo':_OrderNo,'ClientName':[_ClientName],'ShipmentDate':_ShipmentDate,'firm':[_firm],'supplier':[_supplier]}

        mongodbClient.cursor.insert(businessorder_doc)
    def updataClient(self,mongodbClient,_OrderNo,_ClientName,_ShipmentDate,_firm,_supplier):
        mongodbClient.cursor.update({ 'ClientName':_ClientName,'firm':_firm,'supplier':_supplier}\
                           ,{'$push':{"OrderNo" : _OrderNo,'ShipmentDate':_ShipmentDate }})







