# -*-  coding: utf-8  -*-
__author__ = '10409003'
import xlrd
import uuid
from datetime import datetime
from aes_data import aes_data
from ToMongodb import ToMongodb
from ToMysql import ToMysql
import threading

class Test_Data():
    Data=None
    def __init__(self):
        pass
    def Test_Data(self,supplier,GroupID,path,UserID):
        #mysql connector object
        mysqlconnect=ToMysql()
        mysqlconnect.connect()

        mongoOrder=ToMongodb()
        mongoOrder.setCollection('co_OrderData')
        mongoOrder.connect()
        mongodbClient=ToMongodb()
        mongodbClient.setCollection('co_ClientData')
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
                strShipmentDate = xlrd.xldate_as_tuple(table.cell_value(row_index, 1), data.datemode)
                ShipmentDate = datetime(*strShipmentDate[0:6]).strftime('%Y-%m-%d')
                # ShipmentDate = datetime.strptime(_ShipmentDate, '%Y-%m-%d')
                strTurnDate = xlrd.xldate_as_tuple(table.cell_value(row_index, 0), data.datemode)
                TurnDate = datetime(*strTurnDate[0:6]).strftime('%Y-%m-%d')
                OrderNo = table.cell(row_index, 2).value[0:13]
                Name = table.cell(row_index, 5).value.encode('utf-8').strip()
                ClientName = aes.AESencrypt("p@ssw0rd", Name, True)
                Phone = table.cell(row_index, 7).value
                ClientPhone = aes.AESencrypt("p@ssw0rd", Phone, True)
                Tel = table.cell(row_index, 6).value
                ClientTel = aes.AESencrypt("p@ssw0rd", Tel, True)
                ClientZipCode = table.cell(row_index, 8).value
                Addr = table.cell(row_index, 9).value.encode('utf-8').strip()
                ClientAdd = aes.AESencrypt("p@ssw0rd", Addr, True)
                PartNo = str(table.cell(row_index, 10).value)
                PartName = table.cell(row_index, 11).value
                PartQuility = table.cell(row_index, 14).value
                PartTotalPrice = table.cell(row_index, 15).value
                PartCost = table.cell(row_index, 12).value
                GroupID = GroupID
                supplier = supplier
                UserID = UserID
            # SupplySQL = (str(uuid.uuid4()),GroupID, supplier,"","","","","","","","","","","")
            ProductSQL = (str(uuid.uuid4()), GroupID, PartNo, PartName,supplier, "",'',PartCost,PartCost,0,'NULL','NULL','NULL','NULL')
            CustomereSQL = (str(uuid.uuid4()), GroupID, ClientName, ClientAdd, ClientTel, ClientPhone,'NULL',ClientZipCode,'NULL','NULL')

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
                if Name in Name_compare :
                    print "update"
                    mysqlconnect.cursor.execute(CustomereSQLupd,(ClientAdd, ClientTel, ClientPhone,'NULL',ClientZipCode,'NULL','NULL',GroupID,x[1]))
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
                if aes.AESdecrypt('p@ssw0rd',x[1],True)==Name:
                    mysqlconnect.cursor.execute(SalestrSQLsel,(str(x[1]),))
                    customer_id_temp.append(mysqlconnect.cursor.fetchall()[0])
            print customer_id_temp
            for y in customer_id_temp:
                for x in result:
                    if aes.AESdecrypt('p@ssw0rd', x[1], True) == Name:
                        # seq_sql=(GroupID,ShipmentDate,0)
                        # seq_no=mysqlconnect.cursor.callproc('sp_get_sale_newseqno_withdate',seq_sql)
                        # print seq_no[2]
                        SaleSQL = (
                            GroupID, OrderNo, UserID, PartName, PartNo, y[0],
                            x[1], PartQuility, PartTotalPrice, 'NULL', ShipmentDate, TurnDate, TurnDate,
                            'NULL', ShipmentDate,supplier)

                        mysqlconnect.cursor.callproc('p_tb_sale', SaleSQL)
                        print SaleSQL

                        #threading._sleep(1)
                        # SalestrSQL = "INSERT INTO tb_sale (sale_id,seq_no,group_id,user_id,customer_id,name,c_product_id,quantity,price,trans_list_date,sale_date,product_name)"\
                        # "VALUES(%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)"
                        # mysqlconnect.cursor.execute(SalestrSQL,(str(uuid.uuid4()), OrderNo, GroupID, '12345', y[0], x[1], PartNo, PartQuility,PartTotalPrice, TurnDate,ShipmentDate,PartName))
            mysqlconnect.db.commit()
            # mysqlconnect.cursor.callproc('p_tb_customer', CustomereSQL)

            if (Ordernum==OrderNo[0:13]):
                print 'update'
                self.updataOrder(mongoOrder,ShipmentDate,OrderNo,PartNo,PartName,\
                                PartQuility,PartTotalPrice,PartCost,\
                                GroupID,supplier)

            else:
                print 'insert'
                Ordernum=OrderNo
                self.insertOrder(mongoOrder,ShipmentDate,OrderNo,PartNo,PartName,\
                                PartQuility,PartTotalPrice,PartCost,\
                                GroupID,supplier)
            if (Clientnum == ClientName):
                print 'update'
                self.updataClient(mongodbClient, OrderNo,ClientName,ClientZipCode,\
                                ClientAdd,ClientTel,ClientPhone,\
                               GroupID,supplier)
            else:
                print 'insert'
                Clientnum = ClientName
                self.insertClient(mongodbClient, OrderNo,ClientName,ClientZipCode,\
                                ClientAdd,ClientTel,ClientPhone,\
                               GroupID,supplier)

        mysqlconnect.dbClose()
        mongoOrder.dbClose()
        mongodbClient.dbClose()
        return 'Success'

    # mongoDB storage   ç¬¬ä??‹å??¸æ˜¯ä¸Ÿä??¢ç?mongoOrder or mongoClient
    def insertOrder(self,mongoOrder,_ShipmentDate,_OrderNo,_PartNo,_PartName,\
                                _PartQuility,_PartTotalPrice,_PartCost,\
                                _GroupID,_supplier):
        businessorder_doc={'ShipmentDate':_ShipmentDate,'OrderNo':_OrderNo,\
                           'PartNo':[_PartNo],'PartName':[_PartName],\
                           'PartQuility':[_PartQuility],'PartTotalPrice':[_PartTotalPrice],\
                           'Price': [_PartCost],'GroupID':[_GroupID],'supplier':[_supplier]}

        mongoOrder.cursor.insert(businessorder_doc)
    def updataOrder(self,mongoOrder,_ShipmentDate,_OrderNo,_PartNo,_PartName,\
                                _PartQuility,_PartTotalPrice,_PartCost,\
                                _GroupID,_supplier):
        mongoOrder.cursor.update( { "OrderNo" : _OrderNo,'GroupID':_GroupID,'supplier':_supplier},\
                           {'$push':{'ShipmentDate':_ShipmentDate,'PartNo':_PartNo ,'PartName':_PartName,\
                           'PartQuility':_PartQuility,'PartTotalPrice':_PartTotalPrice,'Price': _PartCost}})
    def insertClient(self,mongodbClient,_OrderNo,_ClientName,_ClientZipCode,_ClientAdd,_ClientTel,\
                   _ClientPhone,_GroupID,_supplier):
        businessorder_doc={'OrderNo':_OrderNo,'ClientName':[_ClientName],'ClientZipCode':[_ClientZipCode],'ClientTel':[_ClientTel],'ClientPhone':[_ClientPhone],'ClientAdd':[_ClientAdd],\
                           'GroupID':[_GroupID],'supplier':[_supplier]}

        mongodbClient.cursor.insert(businessorder_doc)
    def updataClient(self,mongodbClient,_OrderNo,_ClientName,_ClientZipCode,_ClientAdd,_ClientTel,\
                   _ClientPhone,_GroupID,_supplier):
        mongodbClient.cursor.update({ 'ClientName':_ClientName,'ClientZipCode':_ClientZipCode,'ClientAdd':_ClientAdd,'ClientTel':_ClientTel,'ClientPhone':_ClientPhone,\
                            'GroupID':_GroupID,'supplier':_supplier}\
                           ,{'$push':{"OrderNo" : _OrderNo}})







