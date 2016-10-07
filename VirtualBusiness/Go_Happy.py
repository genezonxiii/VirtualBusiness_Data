# -*-  coding: utf-8  -*-
__author__ = '10409003'
import xlrd
import uuid
from datetime import datetime
from aes_data import aes_data
from ToMongodb import ToMongodb
from ToMysql import ToMysql

class GoHappy_Data():
    Data=None
    def __init__(self):
        pass
    def GoHappy_Data(self,supplier,GroupID,path,UserID):
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

        for row_index in range(1,table.nrows):
            for col_index in range(0,num_cols):
                aes=aes_data()
                firm = GroupID
                firmNo = table.cell(row_index, 1).value
                ShipmentDateNo = table.cell(row_index, 3).value
                ReconciliationDate = table.cell(row_index, 4).value
                OrderNo = str(table.cell(row_index, 6).value[0:13])
                OrderStatus = table.cell(row_index, 7).value
                strOrderDate= xlrd.xldate_as_tuple(table.cell_value(row_index, 8), data.datemode)
                _OrderDate = datetime(*strOrderDate[0:6]).strftime('%Y-%m-%d')
                OrderDate = datetime.strptime(_OrderDate, '%Y-%m-%d')
                PartName = table.cell(row_index, 9).value
                PartType = str(table.cell(row_index, 10).value)
                PartNo = table.cell(row_index, 11).value
                FormatNo = table.cell(row_index, 12).value
                PartCost = table.cell(row_index, 13).value
                PartPrice = table.cell(row_index, 14).value
                PartQuility = table.cell(row_index, 15).value
                PartTotalPrice = table.cell(row_index, 16).value
                OrderName = table.cell(row_index, 17).value
                Name = table.cell(row_index, 18).value
                ClientName = aes.AESencrypt("p@ssw0rd", Name.encode('utf-8'), True)
                ClientZipCode = table.cell(row_index, 19).value
                Add = table.cell(row_index, 20).value
                ClientAdd = aes.AESencrypt("p@ssw0rd", Add.encode('utf-8'), True)
                Tel = table.cell(row_index, 21).value
                ClientTel = aes.AESencrypt("p@ssw0rd", Tel, True)
                Phone = table.cell(row_index, 22).value
                ClientPhone = aes.AESencrypt("p@ssw0rd", Phone, True)
                PartNote = table.cell(row_index, 23).value
                PartShopNo = table.cell(row_index, 24).value
                PartAction = table.cell(row_index, 25).value
                OrderType = table.cell(row_index, 26).value
                strShipmentdate = xlrd.xldate_as_tuple(table.cell_value(row_index, 27), data.datemode)
                _ShipmentDate = datetime(*strShipmentdate[0:6]).strftime('%Y-%m-%d')
                ShipmentDate = datetime.strptime(_ShipmentDate, '%Y-%m-%d')
                PartNum = str(table.cell(row_index, 28).value)
                supplier = supplier
                GroupID = GroupID
                UserID = UserID
            # SupplySQL = (str(uuid.uuid4()),GroupID, supplier,"","","","","","","","","","","")
            ProductSQL = (str(uuid.uuid4()), GroupID, PartNo, PartName,supplier, PartType,"",PartCost,PartPrice,0,None,None,None,None)
            CustomereSQL = (str(uuid.uuid4()), GroupID, ClientName, ClientAdd, ClientTel, ClientPhone,None,ClientZipCode,None,None)

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
                            GroupID, OrderNo, UserID, PartName, PartNo, y[0], x[1],
                            PartQuility,PartPrice, None, ShipmentDate, ShipmentDate, ShipmentDate, None, ShipmentDate,supplier)
                        mysqlconnect.cursor.callproc('p_tb_sale', SaleSQL)
                        # SalestrSQL = "INSERT INTO tb_sale (sale_id,seq_no,group_id,user_id,product_name,c_product_id,customer_id,name,quantity,price,trans_list_date,sale_date)"\
                        #                 "VALUES(%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)"
                        # mysqlconnect.cursor.execute(SalestrSQL,(str(uuid.uuid4()), OrderNo, GroupID, '22345', PartName ,PartNo, y[0], x[1],PartQuility, PartTotalPrice, OrderDate, ShipmentDate))
            mysqlconnect.db.commit()


            if (Ordernum==OrderNo):
                print 'update'
                self.updataOrder(mongoOrder,firm,firmNo,ShipmentDateNo,ReconciliationDate,OrderNo,OrderStatus,OrderDate,PartName,PartType,PartNo,FormatNo,\
                                PartCost,PartPrice,PartQuility,PartTotalPrice,PartNote,PartShopNo,
                                PartAction,OrderType,ShipmentDate,PartNum,supplier)

            else:
                print 'insert'
                Ordernum=OrderNo
                self.insertOrder(mongoOrder,firm,firmNo,ShipmentDateNo,ReconciliationDate,OrderNo,OrderStatus,OrderDate,PartName,PartType,PartNo,FormatNo,\
                                PartCost,PartPrice,PartQuility,PartTotalPrice,PartNote,PartShopNo,
                                PartAction,OrderType,ShipmentDate,PartNum,supplier)
            if (Clientnum == ClientName):
                print 'update'
                self.updataClient(mongodbClient, firm, OrderNo, OrderName, ClientName, ClientZipCode, ClientAdd, ClientTel,ClientPhone, supplier)
            else:
                print 'insert'
                Clientnum = ClientName
                self.insertClient(mongodbClient, firm, OrderNo, ClientName, ClientZipCode, ClientAdd, ClientTel,ClientPhone, supplier)

        mysqlconnect.dbClose()
        mongoOrder.dbClose()
        mongodbClient.dbClose()
        return 'success'

    def insertOrder(self,mongoOrder,_firm,_firmNo,_ShipmentDateNo,_ReconciliationDate,_OrderNo,_OrderStatus,_OrderDate,_PartName,_PartType,_PartNo,_FormatNo,\
                                _PartCost,_PartPrice,_PartQuility,_PartTotalPrice,_PartNote,_PartShopNo,\
                                _PartAction,_OrderType,_ShipmentDate,_PartNum,_supplier):
        businessorder_doc={ 'firm':[_firm],'firmNo':[_firmNo],'ShipmentDateNo':[_ShipmentDateNo],'ReconciliationDate':[_ReconciliationDate],\
                           '_OrderNo':_OrderNo,'OrderStatus':[_OrderStatus],'OrderDate':_OrderDate,\
                           'PartName':[_PartName],'PartType':[_PartType],'PartNo':[_PartNo],'FormatNo':[_FormatNo],\
                           'PartCost':[_PartCost],'Price':[_PartPrice],'PartQuility':[_PartQuility],'PartTotalPrice':[_PartTotalPrice],\
                           'PartNote':[_PartNote],'PartShopNo':[_PartShopNo],'PartAction':[_PartAction],'OrderType':[_OrderType],\
                           'ShipmentDate':_ShipmentDate,'PartNum':[_PartNum],'supplier':[_supplier]
                            }
        mongoOrder.cursor.insert(businessorder_doc)
    def updataOrder(self,mongoOrder,_firm,_firmNo,_ShipmentDateNo,_ReconciliationDate,_OrderNo,_OrderStatus,_OrderDate,_PartName,_PartType,_PartNo,_FormatNo,\
                                _PartCost,_PartPrice,_PartQuility,_PartTotalPrice,_PartNote,_PartShopNo,\
                                _PartAction,_OrderType,_ShipmentDate,_PartNum,_supplier):
        mongoOrder.cursor.update( { "OrderNo" : _OrderNo,'firm':_firm,'supplier':_supplier,'firmNo':[_firmNo]},\
                           {'$push':{'ShipmentDateNo':_ShipmentDateNo,'ReconciliationDate':_ReconciliationDate,\
                           'OrderStatus':_OrderStatus,'OrderDate':_OrderDate,\
                           'PartName':_PartName,'PartType':_PartType,'PartNo':_PartNo,'FormatNo':_FormatNo,\
                           'PartCost':_PartCost,'Price':_PartPrice,'PartQuility':_PartQuility,'PartTotalPrice':_PartTotalPrice,\
                           'PartNote':_PartNote,'PartShopNo':_PartShopNo,'PartAction':_PartAction,'OrderType':_OrderType,\
                           'ShipmentDate':_ShipmentDate,'PartNum':_PartNum}}
                                                )

    def insertClient(self,mongodbClient,_firm,_OrderNo,_ClientName,_ClientZipCode,_ClientAdd,_ClientTel,_ClientPhone,_supplier):
        businessorder_doc={ 'firm':[_firm],'OrderNo':_OrderNo,'ClientName':[_ClientName],'ClientZipCode':[_ClientZipCode],'ClientAdd':[_ClientAdd],'ClientTel':[_ClientTel],'ClientPhone':[_ClientPhone],'supplier':[_supplier]}

        mongodbClient.cursor.insert(businessorder_doc)

    def updataClient(self,mongodbClient,_firm,_OrderNo,_ClientName,_ClientZipCode,_ClientAdd,_ClientTel,_ClientPhone,_supplier):
        mongodbClient.cursor.update(  { 'firm':_firm,'ClientName':_ClientName,'ClientZipCode':_ClientZipCode,'ClientTel':_ClientTel,'ClientPhone':_ClientPhone,'supplier':_supplier}\
                           ,{'$push':{"OrderNo" : _OrderNo}})







