# -*-  coding: utf-8  -*-
__author__ = '10409003'
import xlrd
import uuid
from datetime import datetime
from aes_data import aes_data
from ToMongodb import ToMongodb
from ToMysql import ToMysql

class IBON_Data():
    Data=None

    def __init__(self):
        pass
    def IBON_Data(self,supplier,GroupID,path,UserID):
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
                strTurntDate = xlrd.xldate_as_tuple(table.cell_value(row_index, 3), data.datemode)
                _TurnDate = datetime(*strTurntDate[0:6]).strftime('%Y-%m-%d')
                TurnDate = datetime.strptime(_TurnDate, '%Y-%m-%d')
                OrderNo = table.cell(row_index, 4).value[0:13]
                OrderName = table.cell(row_index, 6).value
                Name = table.cell(row_index, 7).value
                ClientName = aes.AESencrypt("p@ssw0rd", Name.encode('utf-8'), True)
                Phone = table.cell(row_index, 8).value
                ClientPhone = aes.AESencrypt("p@ssw0rd", Phone, True)
                Add = table.cell(row_index, 9).value
                ClientAdd = aes.AESencrypt("p@ssw0rd", Add.encode('utf-8'), True)
                Mail = table.cell(row_index, 10).value
                ClientMail = aes.AESencrypt("p@ssw0rd", Mail, True)
                ClientZipCode = table.cell(row_index, 11).value
                OrderType = table.cell(row_index, 13).value
                PartNo = str(table.cell(row_index, 15).value)
                PartName = table.cell(row_index, 17).value
                PartSpec = table.cell(row_index, 18).value
                PartQuility = table.cell(row_index, 19).value
                PartCost = table.cell(row_index, 20).value
                strShipmentDate = xlrd.xldate_as_tuple(table.cell_value(row_index, 25), data.datemode)
                _ShipmentDate = datetime(*strShipmentDate[0:6]).strftime('%Y-%m-%d')
                ShipmentDate = datetime.strptime(_ShipmentDate, '%Y-%m-%d')
                firm = GroupID
                supplier = supplier
                UserID = UserID
            # SupplySQL = (str(uuid.uuid4()),GroupID, supplier,"","","","","","","","","","","")
            ProductSQL = (str(uuid.uuid4()), GroupID, PartNo, PartName,supplier,"",PartSpec,PartCost,0,0,None,None,None,None)
            CustomereSQL = (str(uuid.uuid4()), GroupID, ClientName, ClientAdd, "", ClientPhone,'',ClientZipCode,None,None)
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
                    mysqlconnect.cursor.execute(CustomereSQLupd,(ClientAdd,None, ClientPhone,None,None,None,None,GroupID,x[1]))
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
                        SaleSQL = (
                            GroupID, OrderNo, UserID, PartName, PartNo, y[0], x[1],
                            PartQuility, PartCost, None, ShipmentDate, TurnDate, ShipmentDate, None,
                            ShipmentDate,supplier)
                        mysqlconnect.cursor.callproc('p_tb_sale', SaleSQL)
                        # SalestrSQL = "INSERT INTO tb_sale (sale_id,seq_no,group_id,user_id,product_name,c_product_id,customer_id,name,quantity,price,trans_list_date,sale_date)"\
                        #                 "VALUES(%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)"
                        # mysqlconnect.cursor.execute(SalestrSQL,(str(uuid.uuid4()), OrderNo, GroupID, '32345',PartName ,PartNo, y[0], x[1],  PartQuility, PartCost, TurntDate, ShipmentDate))
            mysqlconnect.db.commit()

            if (Ordernum==OrderNo):
                print 'update'
                self.updataOrder(mongoOrder,TurnDate,OrderNo,OrderType,PartNo,PartName,PartSpec,\
                                PartQuility,PartCost,ShipmentDate,firm,supplier)

            else:
                print 'insert'
                Ordernum=OrderNo
                self.insertOrder(mongoOrder,TurnDate,OrderNo,OrderType,PartNo,PartName,PartSpec,\
                                PartQuility,PartCost,ShipmentDate,firm,supplier)
            if (Clientnum == ClientName):
                print 'update'
                self.updataClient(mongodbClient, OrderNo,OrderName,ClientName,ClientPhone,ClientAdd,ClientMail,\
                                ClientZipCode,firm,supplier)
            else:
                print 'insert'
                Clientnum = ClientName
                self.insertClient(mongodbClient, OrderNo,OrderName,ClientName,ClientPhone,ClientAdd,ClientMail,\
                                ClientZipCode,firm,supplier)

        mysqlconnect.dbClose()
        mongoOrder.dbClose()
        mongodbClient.dbClose()
        return 'success'

    # mongoDB storage   第�??��??�是丟�??��?mongoOrder or mongoClient
    def insertOrder(self,mongoOrder,_TurnDate,_OrderNo,_OrderType,_PartNo,_PartName,_PartSpec,\
                                _PartQuility,_PartCost,_ShipmentDate,_firm,_supplier):
        businessorder_doc={ 'TurnDate':_TurnDate,'OrderNo':_OrderNo,'OrderType':[_OrderType],'PartNo':[_PartNo],\
                           'PartName':[_PartName],'PartSpec':[_PartSpec],\
                           'PartQuility':[_PartQuility],'Price':[_PartCost],'ShipmentDate':_ShipmentDate,'firm':[_firm],'supplier':[_supplier]
                            }
        mongoOrder.cursor.insert(businessorder_doc)

    def updataOrder(self,mongoOrder,_TurnDate,_OrderNo,_OrderType,_PartNum,_PartName,_PartSpec,\
                                _PartQuility,_PartCost,_ShipmentDate,_firm,_supplier):
        mongoOrder.cursor.update( { "OrderNo" : _OrderNo,'firm':_firm,'supplier':_supplier},\
                           {'$push':{'TurnDate':_TurnDate,'OrderType':_OrderType,'PartNum':_PartNum ,'PartName':_PartName,\
                           'PartSpec':_PartSpec,'PartQuility':_PartQuility,'Price':_PartCost,'ShipmentDate':_ShipmentDate}}
                                                )

    def insertClient(self,mongodbClient,_OrderNo,_OrderName,_ClientName,_ClientPhone,_ClientAdd,_ClientMail,\
                                _ClientZipCode,_firm,_supplier):
        businessorder_doc={'OrderNo':_OrderNo,'OrderName':[_OrderName],'ClientName':[_ClientName],'ClientPhone':[_ClientPhone],'ClientAdd':[_ClientAdd],\
                           'ClientMail':[_ClientMail],'ClientZipCode':[_ClientZipCode],'firm':[_firm],'supplier':[_supplier]}
        mongodbClient.cursor.insert(businessorder_doc)

    def updataClient(self,mongodbClient,_OrderNo,_OrderName,_ClientName,_ClientPhone,_ClientAdd,_ClientMail,\
                                _ClientZipCode,_firm,_supplier):
        mongodbClient.cursor.update(  {  'OrderName':_OrderName,'ClientName':_ClientName,'ClientPhone':[_ClientPhone],'ClientAdd':_ClientAdd,\
                             'ClientMail':_ClientMail,'ClientZipCode':_ClientZipCode,'firm':_firm,'supplier':_supplier}\
                           ,{'$push':{"OrderNo" : _OrderNo}})







