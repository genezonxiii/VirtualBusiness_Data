# -*-  coding: utf-8  -*-
__author__ = '10409003'
import xlrd
import uuid
import datetime
from aes_data import aes_data
from ToMongodb import ToMongodb
from ToMysql import ToMysql

class Yahoos_Data():
    Data=None
    def __init__(self):
        pass
    def Yahoos_Data(self,supplier,GroupID,path,UserID):
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
        for row_index in range(2,table.nrows):
            for col_index in range(0,num_cols):
                aes=aes_data()
                OrderNo = str(table.cell(row_index, 0).value)
                #strTurnDate = str(table.cell(row_index, 2).value)
                #_strTurnDate = datetime.datetime.strptime(strTurnDate,'%Y/%m/%d %H:%M ')
                #TurnDate = datetime.datetime.strftime(_strTurnDate,'%Y-%m-%d %H:%M')
                TurnDate = datetime.datetime.strptime(str(table.cell_value(row_index, 1)),'%Y/%m/%d %H:%M')
                ShipmentDate = datetime.datetime.strptime(str(table.cell_value(row_index, 9)),'%Y/%m/%d %H:%M')

                Name = table.cell(row_index, 4).value
                ClientName = aes.AESencrypt("p@ssw0rd", Name.encode('utf-8'), True)
                Phone = table.cell(row_index, 6).value
                ClientPhone = aes.AESencrypt("p@ssw0rd", Phone, True)

                PartNo = str(table.cell(row_index, 11).value).split('.')[0]
                PartName = table.cell(row_index,12).value

                PartQuility = table.cell(row_index, 15).value
                PartTotalPrice = table.cell(row_index, 16).value
                # strShipmentDate = xlrd.xldate_as_tuple(table.cell_value(row_index, 10), data.datemode)
                # ShipmentDate = datetime(*strShipmentDate[0:6]).strftime('%Y-%m-%d')

                firm = GroupID
                supplier = supplier
                UserID = UserID
            # SupplySQL = (str(uuid.uuid4()),GroupID, supplier,"","","","","","","","","","","")
            ProductSQL = (str(uuid.uuid4()), GroupID, PartNo, PartName,supplier, None,None,0,PartTotalPrice,0,None,None,None,None)
            CustomereSQL = (str(uuid.uuid4()), GroupID, ClientName, None, None, ClientPhone,None,None,None,None)

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
                    mysqlconnect.cursor.execute(CustomereSQLupd,(None, None, ClientPhone,None,None,None,None,GroupID,x[1]))
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
                            x[1], PartQuility, PartTotalPrice, None, None, TurnDate,None,
                            None, ShipmentDate,supplier)
                        mysqlconnect.cursor.callproc('p_tb_sale', SaleSQL)
                        # SalestrSQL = "INSERT INTO tb_sale (sale_id,seq_no,group_id,user_id,c_product_id,customer_id,name,quantity,price,trans_list_date,sale_date,product_name)"\
                        # "VALUES(%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)"
                        # mysqlconnect.cursor.execute(SalestrSQL,(str(uuid.uuid4()), OrderNo, GroupID,'73345',PartNo , y[0], x[1], PartQuility,PartTotalPrice, TurnDate,ShipmentDate,PartName))
            mysqlconnect.db.commit()
            # mysqlconnect.cursor.callproc('p_tb_customer', CustomereSQL)

            if (Ordernum==OrderNo[0:13]):
                print 'update'
                self.updataOrder(mongoOrder,TurnDate,ShipmentDate,OrderNo,PartName,\
                                PartQuility,PartTotalPrice,\
                                firm,supplier)

            else:
                print 'insert'
                Ordernum=OrderNo
                self.insertOrder(mongoOrder,TurnDate,ShipmentDate,OrderNo,PartName,\
                                PartQuility,PartTotalPrice,\
                                firm,supplier)
            if (Clientnum == ClientName):
                print 'update'
                self.updataClient(mongodbClient, OrderNo,ClientName,ClientPhone,firm,supplier)
            else:
                print 'insert'
                Clientnum = ClientName
                self.insertClient(mongodbClient, OrderNo,ClientName,\
                                ClientPhone,\
                               firm,supplier)

        mysqlconnect.dbClose()
        mongoOrder.dbClose()
        mongodbClient.dbClose()
        return 'success'

    # mongoDB storage   ?��??????豢�??��???�?mongoOrder or mongoClient
    def insertOrder(self,mongoOrder,_TurnDate,_ShipmentDate,_OrderNo,_PartName,\
                                _PartQuility,_PartTotalPrice,\
                                _firm,_supplier):
        businessorder_doc={'TurnDate':_TurnDate,'ShipmentDate':_ShipmentDate,'OrderNo':_OrderNo,\
                           'PartName':[_PartName],\
                           'PartQuility':[_PartQuility],'Price':[_PartTotalPrice],\
                           'firm':[_firm],'supplier':[_supplier]}

        mongoOrder.cursor.insert(businessorder_doc)
    def updataOrder(self,mongoOrder,_TurnDate,_ShipmentDate,_OrderNo,_PartName,\
                                _PartQuility,_PartTotalPrice,\
                                _firm,_supplier):
        mongoOrder.cursor.update( {'TurnDate':_TurnDate, "OrderNo" : _OrderNo,'firm':_firm,'supplier':_supplier},\
                           {'$push':{'ShipmentDate':_ShipmentDate,'PartName':_PartName,\
                           'PartQuility':_PartQuility,'Price':_PartTotalPrice}})
    def insertClient(self,mongodbClient,_OrderNo,_ClientName,\
                   _ClientPhone,_firm,_supplier):
        businessorder_doc={'OrderNo':_OrderNo,'ClientName':[_ClientName],'ClientPhone':[_ClientPhone],\
                           'firm':[_firm],'supplier':[_supplier]}

        mongodbClient.cursor.insert(businessorder_doc)
    def updataClient(self,mongodbClient,_OrderNo,_ClientName,\
                   _ClientPhone,_firm,_supplier):
        mongodbClient.cursor.update({ 'ClientName':_ClientName,'ClientPhone':_ClientPhone,\
                            'firm':_firm,'supplier':_supplier}\
                           ,{'$push':{"OrderNo" : _OrderNo}})







