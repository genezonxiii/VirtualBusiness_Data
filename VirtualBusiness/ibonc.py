# -*-  coding: utf-8  -*-
__author__ = '10409003'
import xlrd
import uuid
from datetime import datetime
from aes_data import aes_data
from ToMongodb import ToMongodb
from ToMysql import ToMysql

class IBON_DataC():
    Data=None
    def __init__(self):
        pass
    def IBON_DataC(self,supplier,GroupID,path,UserID):
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
                strTurntDate = xlrd.xldate_as_tuple(table.cell_value(row_index, 5), data.datemode)
                _TurnDate = datetime(*strTurntDate[0:6]).strftime('%Y-%m-%d')
                TurnDate = datetime.strptime(_TurnDate, '%Y-%m-%d')
                strShipmentDate = xlrd.xldate_as_tuple(table.cell_value(row_index, 7), data.datemode)
                _ShipmentDate = datetime(*strShipmentDate[0:6]).strftime('%Y-%m-%d')
                ShipmentDate = datetime.strptime(_ShipmentDate, '%Y-%m-%d')
                OrderNo = table.cell(row_index, 8).value[0:13]
                OrderName = table.cell(row_index, 11).value
                Phone = table.cell(row_index, 12).value
                ClientPhone = aes.AESencrypt("p@ssw0rd", Phone, True)
                Name = table.cell(row_index, 13).value
                ClientName = aes.AESencrypt("p@ssw0rd", Name, True)
                PartNo = table.cell(row_index, 21).value
                PartName = table.cell(row_index, 22).value
                PartSpec = table.cell(row_index, 23).value
                PartQuility = table.cell(row_index, 25).value
                PartCost = table.cell(row_index, 26).value
                firm = GroupID
                supplier = supplier
                UserID = UserID
            # SupplySQL = (str(uuid.uuid4()),GroupID, supplier,"","","","","","","","","","","")
            ProductSQL = (str(uuid.uuid4()), GroupID, PartNo, PartName,supplier,None,PartSpec,PartCost,0,0,None,None,None,None)
            CustomereSQL = (str(uuid.uuid4()), GroupID, ClientName,None, None, ClientPhone,None,None,None,None)

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
                    mysqlconnect.cursor.execute(CustomereSQLupd,(None,None, ClientPhone,None,None,None,None,GroupID,x[1]))
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
            # print x[1]
            for y in customer_id_temp:
                for x in result:
                    if aes.AESdecrypt('p@ssw0rd', x[1], True) == Name.encode('utf-8'):
                        print 'GroupID'
                        print PartName
                        SaleSQL = (
                            GroupID,OrderNo, UserID, PartName, PartNo, y[0], x[1],
                            PartQuility, PartCost, None, ShipmentDate, TurnDate, ShipmentDate,None,
                            ShipmentDate,supplier)
                        mysqlconnect.cursor.callproc('p_tb_sale', SaleSQL)
                        # SalestrSQL = "INSERT INTO tb_sale (sale_id,seq_no,group_id,user_id,product_name,c_product_id,customer_id,name,quantity,price,trans_list_date,sale_date)"\
                        #                 "VALUES(%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)"
                        # mysqlconnect.cursor.execute(SalestrSQL,(str(uuid.uuid4()), OrderNo, GroupID, '32345', PartName ,PartNo, y[0], x[1], PartQuility, PartCost, TurntDate, ShipmentDate))
            mysqlconnect.db.commit()

            if (Ordernum==OrderNo[0:13]):
                print 'update'
                self.updataOrder(mongoOrder,TurnDate,ShipmentDate,OrderNo,PartNo,PartName,PartSpec,\
                                PartQuility,PartCost,firm,supplier)

            else:
                print 'insert'
                Ordernum=OrderNo
                self.insertOrder(mongoOrder,TurnDate,ShipmentDate,OrderNo,PartNo,PartName,PartSpec,\
                                PartQuility,PartCost,firm,supplier)
            if (Clientnum == ClientName):
                print 'update'
                self.updataClient(mongodbClient, OrderNo,OrderName,ClientPhone,ClientName,\
                                firm,supplier)
            else:
                print 'insert'
                Clientnum = ClientName
                self.insertClient(mongodbClient, OrderNo,OrderName,ClientPhone,ClientName,\
                                firm,supplier)

        mysqlconnect.dbClose()
        mongoOrder.dbClose()
        mongodbClient.dbClose()
        return 'success'

    # mongoDB storage   第一個參數是丟上面的mongoOrder or mongoClient
    def insertOrder(self,mongoOrder,_TurnDate,_ShipmentDate,_OrderNo,_PartNo,_PartName,_PartSpec,\
                                _PartQuility,_PartCost,_firm,_supplier):
        businessorder_doc={'TurnDate':_TurnDate,'ShipmentDate':_ShipmentDate,'OrderNo':_OrderNo,'PartNo':[_PartNo],\
                           'PartName':[_PartName],'PartSpec':[_PartSpec],\
                           'PartQuility':[_PartQuility],'Price':[_PartCost],'firm':[_firm],'supplier':[_supplier]
                            }

        mongoOrder.cursor.insert(businessorder_doc)
    def updataOrder(self,mongoOrder,_TurnDate,_ShipmentDate,_OrderNo,_PartNo,_PartName,_PartSpec,\
                                _PartQuility,_PartCost,_firm,_supplier):
        mongoOrder.cursor.update( { "OrderNo" : _OrderNo,'firm':_firm,'supplier':_supplier},\
                           {'$push':{'TurnDate':_TurnDate,'ShipmentDate':_ShipmentDate,'PartNo':_PartNo ,'PartName':_PartName,\
                           'PartSpec':_PartSpec,'PartQuility':_PartQuility,'Price':_PartCost}}
                                                )
    def insertClient(self,mongodbClient,_OrderNo,_OrderName,_ClientPhone,_ClientName,\
                                _firm,_supplier):

        businessorder_doc={'OrderNo':_OrderNo,'OrderName':[_OrderName],'ClientPhone':[_ClientPhone],'ClientName':[_ClientName],\
                           'firm':[_firm],'supplier':[_supplier]}

        mongodbClient.cursor.insert(businessorder_doc)
    def updataClient(self,mongodbClient,_OrderNo,_OrderName,_ClientPhone,_ClientName,\
                                _firm,_supplier):
        mongodbClient.cursor.update(  { 'OrderName':[_OrderName],'ClientName':_ClientName,'ClientPhone':[_ClientPhone],\
                             'firm':_firm,'supplier':_supplier}\
                           ,{'$push':{"OrderNo" : _OrderNo}})







