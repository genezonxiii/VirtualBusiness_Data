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

class ASAP_Data():
    Data=None
    def __init__(self):
        pass
    def ASAP_Data(self,supplier,GroupID,path,UserID):
        logging.basicConfig(filename='pyupload.log', level=logging.DEBUG, format='%(asctime)s %(message)s', datefmt='%Y/%m/%d %I:%M:%S %p')
        logging.Formatter.converter = time.gmtime
        logging.info('===ASAP_Data===')
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
            aes=aes_data()
            
            logging.info('row_index:' + str(row_index))
            logging.info('0:' + table.cell(row_index, 0).value)
            logging.info('1:' + table.cell(row_index, 1).value)
            logging.info('4:' + table.cell(row_index, 4).value)
            logging.info('7:' + table.cell(row_index, 7).value)
            logging.info('8:' + table.cell(row_index, 8).value)
            logging.info('10:' + table.cell(row_index, 10).value)
            logging.info('11:' + table.cell(row_index, 11).value)
            logging.info('12:' + table.cell(row_index, 12).value)
            logging.info('13:' + table.cell(row_index, 13).value)
            logging.info('14:' + str(table.cell(row_index, 14).value) )
            logging.info('15' + table.cell(row_index, 15).value)
            
            OrderNo = str(table.cell(row_index, 1).value).split('.')[-1]
            #strTurnDate = str(table.cell(row_index, 9).value).replace("/", "-")
            #TurnDate = datetime.strptime(strTurnDate, '%Y-%m-%d %H:%M')
            #strShipmentDate = str(table.cell(row_index, 10).value).replace("/", "-")
            #ShipmentDate = datetime.strptime(strShipmentDate, '%Y-%m-%d')
            #strInvoiceDate = str(table.cell(row_index, 25).value).replace("/", "-")
            #InvoiceDate = datetime.strptime(strInvoiceDate, '%Y-%m-%d')
            logging.info('ORDERNO SUBSTRING:' + OrderNo[0:4])
            logging.info('ORDERNO SUBSTRING:' + OrderNo[0:4])

            TurnDate = datetime.datetime.strptime(str(OrderNo[0:4]) + '/' + str(table.cell_value(row_index, 0)),'%Y/%m/%d %H:%M')
            logging.info('TurnDate:' + str(TurnDate))
            #ShipmentDate = datetime.datetime.strptime(str(table.cell_value(row_index, 1)),'%Y/%m/%d %H:%M')
            #InvoiceDate = datetime.datetime.strptime(str(table.cell_value(row_index, 25)),'%Y/%m/%d')

            Name = table.cell(row_index, 10).value
            ClientName = aes.AESencrypt("p@ssw0rd", Name.encode('utf8'), True)
            Tel = table.cell(row_index, 12).value
            ClientTel = aes.AESencrypt("p@ssw0rd", Tel, True)
            Phone = table.cell(row_index, 11).value
            ClientPhone = aes.AESencrypt("p@ssw0rd", Phone, True)
            Add = table.cell(row_index, 13).value
            ClientAdd = aes.AESencrypt("p@ssw0rd", Add.encode('utf8'), True)
            PartNo = str(table.cell(row_index, 14).value).split('.')[-1]
            PartName = table.cell(row_index, 4).value
            PartQuility = table.cell(row_index, 7).value
            PartPrice = table.cell(row_index, 8).value
            #InvoiceNo = table.cell(row_index, 24).value

            GroupID = GroupID
            supplier = supplier
            UserID = UserID
                
            # SupplySQL = (str(uuid.uuid4()),GroupID, supplier,"","","","","","","","","","","")
            ProductSQL = (str(uuid.uuid4()), GroupID, PartNo, PartName,supplier, None,None,0,PartPrice,0,None,None,None,None)
            CustomereSQL = (str(uuid.uuid4()), GroupID, ClientName, ClientAdd, ClientTel, ClientPhone,None,None,None,None)

            # mysqlconnect.cursor.callproc('p_tb_supply', SupplySQL)
            mysqlconnect.cursor.callproc('p_tb_product', ProductSQL)

            SaleOrdersel="""select customer_id,name from tb_sale where order_no = '%s' and group_id ='%s' """ % (OrderNo,GroupID)
            mysqlconnect.cursor.execute(SaleOrdersel)
            orderexist = mysqlconnect.cursor.fetchall()
            if orderexist != []:
                logging.info('order exists')
                updateSaleSQL=""" update tb_sale set user_id=%s, product_name=%s, c_product_id=%s, quantity=%s, price=%s,
                                invoice=%s, invoice_date=%s, trans_list_date=%s, dis_date=%s, memo=%s,
                                sale_date=%s where customer_id = %s"""
                updateSaleValue = (UserID,PartName,PartNo,PartQuility,PartPrice,
                                   None,None,TurnDate,None,None,
                                   None,orderexist[0][0] )
                mysqlconnect.cursor.execute(updateSaleSQL,updateSaleValue)
                mysqlconnect.db.commit()

                logging.info('callproc p_tb_sale_momo:')
                logging.info('1:' + GroupID)
                logging.info('2:' + OrderNo)
                logging.info('3:' + UserID)
                logging.info('4:' + PartName)
                logging.info('5:' + PartNo)
                logging.info('6:' + orderexist[0][0])
                logging.info('7:' + orderexist[0][1])
                logging.info('8:' + PartQuility)
                logging.info('9:' + PartPrice)
                logging.info('12:' + str(TurnDate))
                logging.info('16:' + supplier)

                SaleSQL = (GroupID, OrderNo, UserID, PartName, PartNo,
                           orderexist[0][0], orderexist[0][1], PartQuility, PartPrice, None,
                           None, TurnDate, None, None, None,
                           supplier)

                try:
                    mysqlconnect.cursor.callproc('p_tb_sale_momo', SaleSQL)
                except Exception, e:
                    logging.info('callproc error:' + repr(e))
                    return
                mysqlconnect.db.commit()
                logging.info('after commit')
            else:
                logging.info('order NOT exists')
                CustomereSQLsel = """select group_id,name,address from tb_customer where group_id='%s'""" % (GroupID)
                CustomereSQLins = """ insert into tb_customer
                                    (customer_id,group_id ,name,address,phone,mobile,email,post,class,memo)
                                    values(%s,%s,%s,%s,%s,%s,%s,%s,%s,%s);"""
           
                mysqlconnect.cursor.execute(CustomereSQLins, CustomereSQL)
                mysqlconnect.db.commit()
                
                mysqlconnect.cursor.execute(CustomereSQLsel)
                Finalresult = mysqlconnect.cursor.fetchall()
                customer_id_temp=[]
                SalestrSQLsel_1="SELECT customer_id from tb_customer where name =%s and address= %s;"
                same_name=[]
                for x in Finalresult: 
                   Name_compare = aes.AESdecrypt("p@ssw0rd",x[1], True)
                   if Name.encode('utf-8') ==  Name_compare :
                       print "the same name data list"
                       same_name.append(x)
                       
                for r in same_name:
                    address_compare = aes.AESdecrypt("p@ssw0rd", r[2], True)     
                    if Add.encode('utf-8') == address_compare:
                        mysqlconnect.cursor.execute(SalestrSQLsel_1, (str(r[1]), str(r[2])))
                        customer_id_temp.append(mysqlconnect.cursor.fetchall()[0])

                print customer_id_temp
                for y in customer_id_temp:
                    for x in Finalresult:
                        if aes.AESdecrypt('p@ssw0rd', x[1], True) == Name.encode('utf-8'):
                            
                            SaleSQL = (
                            GroupID, OrderNo, UserID, PartName, PartNo, y[0],
                            x[1], PartQuility, PartPrice, None, None, TurnDate, None,
                            None, None,supplier)
                            mysqlconnect.cursor.callproc('p_tb_sale', SaleSQL)
                            mysqlconnect.db.commit()

            # mysqlconnect.cursor.callproc('p_tb_customer', CustomereSQL)

            if (Ordernum==OrderNo[0:13]):
                print 'update'
                logging.info('OrderNo[0:13] equal => update')
                self.updataOrder(mongoOrder,OrderNo,PartNo,PartName,PartQuility,PartPrice,GroupID,supplier)
            else:
                print 'insert'
                logging.info('OrderNo[0:13] not equal => insert')

                Ordernum=OrderNo
                self.insertOrder(mongoOrder,OrderNo,PartNo,PartName,PartQuility,PartPrice,GroupID,supplier)

            if (Clientnum == ClientName):
                print 'update'
                logging.info('ClientName equal => update')
                self.updataClient(mongodbClient, OrderNo,ClientName,ClientAdd,GroupID,supplier)
            else:
                print 'insert'
                logging.info('ClientName not equal => insert')
                Clientnum = ClientName
                self.insertClient(mongodbClient, OrderNo,ClientName,ClientAdd,GroupID,supplier)

        mysqlconnect.dbClose()
        mongoOrder.dbClose()
        mongodbClient.dbClose()
        logging.info('===ASAP_Data SUCCESS===')
        return 'success'

    # mongoDB storage   第�??��??�是丟�??��?mongoOrder or mongoClient
    def insertOrder(self,mongoOrder,_OrderNo,_PartNo,_PartName,_PartQuility,_PartPrice,_GroupID,_supplier):
        businessorder_doc={ 'OrderNo':_OrderNo,'PartNo':[_PartNo],'PartName':[_PartName],\
                           'PartQuility':[_PartQuility],'Price':[_PartPrice],\
                           'GroupID':[_GroupID],'supplier':[_supplier]
                            }

        mongoOrder.cursor.insert(businessorder_doc)
    def updataOrder(self,mongoOrder,_OrderNo,_PartNo,_PartName,_PartQuility,_PartPrice,_GroupID,_supplier):
        mongoOrder.cursor.update({ "OrderNo" : _OrderNo,'GroupID':_GroupID,'supplier':_supplier},{'$push':{'PartNo':_PartNo,'PartName':_PartName,\
                           'PartQuility':_PartQuility,'Price':_PartPrice}}
                                                )
    def insertClient(self,mongodbClient,_OrderNo,_ClientName,_ClientAdd,_GroupID,_supplier):

        businessorder_doc={ 'OrderNo':_OrderNo,'ClientName':[_ClientName],'ClientAdd':[_ClientAdd],'GroupID':[_GroupID],'supplier':[_supplier]}

        mongodbClient.cursor.insert(businessorder_doc)
    def updataClient(self,mongodbClient,_OrderNo,_ClientName,_ClientAdd,_GroupID,_supplier):
        mongodbClient.cursor.update({ 'ClientName':_ClientName,'ClientAdd':_ClientAdd,'GroupID':_GroupID,'supplier':_supplier}\
                           ,{'$push':{"OrderNo" : _OrderNo}})







