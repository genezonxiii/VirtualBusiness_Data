# -*-  coding: utf-8  -*-
__author__ = '10409003'
import xlrd
import uuid
import datetime
from aes_data import aes_data
from ToMongodb import ToMongodb
from ToMysql import ToMysql
import logging
import json

logger = logging.getLogger(__name__)

class ASAP_Data():
    Data=None

    # 預期要找出欄位的索引位置的欄位名稱
    TitleTuple = (u'接單時間', u'訂單編號', u'料號', u'商品名稱', u'數量',
                  u'總售價', u'總進貨價', u'收貨人', u'手機', u'地址')
    TitleList = []

    def __init__(self):
        pass
    def ASAP_Data(self,supplier,GroupID,path,UserID):
        try:

            logger.debug("ASAP")

            success = False
            resultinfo = ""
            totalRows = 0
        
            #mysql connector object
            mysqlconnect=ToMysql()
            mysqlconnect.connect()
            logger.debug("mysql connect OK")

            mongoOrder=ToMongodb()
            mongoOrder.setCollection('co_order')
            mongoOrder.connect()
            logger.debug("mongo Order Connect OK")

            mongodbClient=ToMongodb()
            mongodbClient.setCollection('co_client')
            mongodbClient.connect()
            logger.debug("mongo Client connect OK")

            Ordernum=""
            Clientnum=""

            data=xlrd.open_workbook(path)
            table=data.sheets()[0]
            num_cols=table.ncols

            totalRows = table.nrows - 1;

            # 存放excel中全部的欄位名稱
            self.TitleList = []
            for row_index in range(0, 1):
                for col_index in range(0, table.ncols):
                    self.TitleList.append(table.cell(row_index, col_index).value)

            logger.debug("TitleList OK")
            logger.debug(', '.join(self.TitleList))
            # for temp in self.TitleList:
            #     logger.debug(temp)

            print self.TitleList

            # 存放excel中對應TitleTuple欄位名稱的index
            for index in range(0, len(self.TitleTuple)):
                if self.TitleTuple[index] in self.TitleList:
                    logger.debug(str(index) + self.TitleTuple[index])
                    logger.debug(u'index in file - ' + str(self.TitleList.index(self.TitleTuple[index])))
                    # print str(index) + TitleTuple[index]
                    # print (TitleList.index(TitleTuple[index]))

            for row_index in range(1,table.nrows):
                aes=aes_data()

                OrderNo = str(table.cell(row_index, self.TitleList.index(self.TitleTuple[1])).value).split('.')[-1]
                #strTurnDate = str(table.cell(row_index, 9).value).replace("/", "-")
                #TurnDate = datetime.strptime(strTurnDate, '%Y-%m-%d %H:%M')
                #strShipmentDate = str(table.cell(row_index, 10).value).replace("/", "-")
                #ShipmentDate = datetime.strptime(strShipmentDate, '%Y-%m-%d')
                #strInvoiceDate = str(table.cell(row_index, 25).value).replace("/", "-")
                #InvoiceDate = datetime.strptime(strInvoiceDate, '%Y-%m-%d')
                logger.info('order_no substring(year):' + OrderNo[0:4])

                TurnDate = datetime.datetime.strptime(str(OrderNo[0:4]) + '/' + str(table.cell_value(row_index, self.TitleList.index(self.TitleTuple[0]))),'%Y/%m/%d %H:%M')
                logger.info('TurnDate:' + str(TurnDate))
                #ShipmentDate = datetime.datetime.strptime(str(table.cell_value(row_index, 1)),'%Y/%m/%d %H:%M')
                #InvoiceDate = datetime.datetime.strptime(str(table.cell_value(row_index, 25)),'%Y/%m/%d')

                Name = table.cell(row_index, self.TitleList.index(self.TitleTuple[7])).value
                ClientName = aes.AESencrypt("p@ssw0rd", Name.encode('utf8'), True)
                ClientTel = None
                Phone = table.cell(row_index, self.TitleList.index(self.TitleTuple[8])).value
                ClientPhone = aes.AESencrypt("p@ssw0rd", Phone, True)
                Add = table.cell(row_index, self.TitleList.index(self.TitleTuple[9])).value
                ClientAdd = aes.AESencrypt("p@ssw0rd", Add.encode('utf8'), True)
                PartNo = str(table.cell(row_index, self.TitleList.index(self.TitleTuple[2])).value).split('.')[-1]
                PartName = table.cell(row_index, self.TitleList.index(self.TitleTuple[3])).value
                PartQuility = table.cell(row_index, self.TitleList.index(self.TitleTuple[4])).value
                PartPrice = table.cell(row_index, self.TitleList.index(self.TitleTuple[6])).value
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
                    logger.debug("orderexist - update")
                    updateSaleSQL=""" update tb_sale set user_id=%s, product_name=%s, c_product_id=%s, quantity=%s, price=%s,
                                    invoice=%s, invoice_date=%s, trans_list_date=%s, dis_date=%s, memo=%s,
                                    sale_date=%s where customer_id = %s"""
                    updateSaleValue = (UserID,PartName,PartNo,PartQuility,PartPrice,
                                       None,None,TurnDate,None,None,
                                       None,orderexist[0][0] )
                    mysqlconnect.cursor.execute(updateSaleSQL,updateSaleValue)
                    mysqlconnect.db.commit()

                    SaleSQL = (GroupID, OrderNo, UserID, PartName, PartNo,
                               orderexist[0][0], orderexist[0][1], PartQuility, PartPrice, None,
                               None, TurnDate, None, None, None,
                               supplier)

                    try:
                        mysqlconnect.cursor.callproc('p_tb_sale_momo', SaleSQL)
                    except Exception, e:
                        logger.info('callproc error:' + repr(e))
                        return
                    mysqlconnect.db.commit()
                    logger.info('after commit')
                else:
                    logger.debug("orderexist - insert")
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

                # if (Ordernum==OrderNo[0:13]):
                #     print 'update'
                #     self.updataOrder(mongoOrder,OrderNo,PartNo,PartName,PartQuility,PartPrice,GroupID,supplier)
                # else:
                #     print 'insert'
                #
                #     Ordernum=OrderNo
                #     self.insertOrder(mongoOrder,OrderNo,PartNo,PartName,PartQuility,PartPrice,GroupID,supplier)
                #
                # if (Clientnum == ClientName):
                #     print 'update'
                #     self.updataClient(mongodbClient, OrderNo,ClientName,ClientAdd,GroupID,supplier)
                # else:
                #     print 'insert'
                #     Clientnum = ClientName
                #     self.insertClient(mongodbClient, OrderNo,ClientName,ClientAdd,GroupID,supplier)

            mysqlconnect.dbClose()
            mongoOrder.dbClose()
            mongodbClient.dbClose()

            success = True
        except Exception as inst:
            logger.error(inst.args)
            resultinfo = inst.args
        finally:
            logger.debug('===ASAP_Data finally===')
            return json.dumps({"success": success, "info": resultinfo, "total": totalRows}, sort_keys = False)


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







