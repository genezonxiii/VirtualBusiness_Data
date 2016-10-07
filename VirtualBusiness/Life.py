# -*-  coding: utf-8  -*-
__author__ = '10409003'
import xlrd
import uuid
from datetime import datetime
from aes_data import aes_data
from ToMongodb import ToMongodb
from ToMysql import ToMysql

class Life_Data():
    Data=None
    def __init__(self):
        pass
    def Life_Data(self,supplier,GroupID,path,UserID):
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
                OrderNo = table.cell(row_index, 0).value[0:13]
                StrPurchaseDate = xlrd.xldate_as_tuple(table.cell_value(row_index, 1), data.datemode)
                _PurchaseDate = datetime(*StrPurchaseDate[0:6]).strftime('%Y-%m-%d')
                PurchaseDate = datetime.strptime(_PurchaseDate, '%Y-%m-%d')
                Name = table.cell(row_index, 2).value
                ClientName = aes.AESencrypt("p@ssw0rd", Name.encode('utf8'), True)
                Phone = table.cell(row_index, 3).value
                ClientPhone = aes.AESencrypt("p@ssw0rd", Phone, True)
                Add = table.cell(row_index, 4).value
                ClientAdd = aes.AESencrypt("p@ssw0rd", Add.encode('utf8'), True)
                PartAction = table.cell(row_index, 5).value
                PartNum = table.cell(row_index, 6).value
                PartQuility = table.cell(row_index, 7).value
                firm = GroupID
                supplier = supplier
                UserID = UserID
            # SupplySQL = (str(uuid.uuid4()),GroupID, supplier,"","","","","","","","","","","")
            ProductSQL = (str(uuid.uuid4()), GroupID, None, PartAction,supplier, None,PartNum,0,0,0,None,None,None,None)
            CustomereSQL = (str(uuid.uuid4()), GroupID, ClientName, ClientAdd,None, ClientPhone,None,None,None,None)

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
                    mysqlconnect.cursor.execute(CustomereSQLupd,(ClientAdd, None, ClientPhone,None,None,None,None,GroupID,x[1]))
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
                            GroupID, OrderNo, UserID, PartAction,None, y[0], x[1],
                            PartQuility, 0,None, PurchaseDate, PurchaseDate, PurchaseDate, None,
                            PurchaseDate,supplier)
                        mysqlconnect.cursor.callproc('p_tb_sale', SaleSQL)
                        # SalestrSQL = "INSERT INTO tb_sale (sale_id,seq_no,group_id,user_id,product_name,customer_id,name,quantity,sale_date)"\
                        # "VALUES(%s,%s,%s,%s,%s,%s,%s,%s,%s)"
                        # mysqlconnect.cursor.execute(SalestrSQL,(str(uuid.uuid4()), OrderNo, GroupID, '52345',PartAction, y[0], x[1] , PartQuility,PurchaseDate))
            mysqlconnect.db.commit()
            # mysqlconnect.cursor.callproc('p_tb_customer', CustomereSQL)

            if (Ordernum==OrderNo[0:13]):
                print 'update'
                self.updataOrder(mongoOrder,OrderNo,PurchaseDate,PartAction,PartNum,PartQuility,firm,supplier)

            else:
                print 'insert'
                Ordernum=OrderNo
                self.insertOrder(mongoOrder,OrderNo,PurchaseDate,PartAction,PartNum,PartQuility,firm,supplier)
            if (Clientnum == ClientName):
                print 'update'
                self.updataClient(mongodbClient, OrderNo,ClientName,ClientPhone,ClientAdd,firm,supplier)
            else:
                print 'insert'
                Clientnum = ClientName
                self.insertClient(mongodbClient,OrderNo,ClientName,ClientPhone,ClientAdd,firm,supplier)

        mysqlconnect.dbClose()
        mongoOrder.dbClose()
        mongodbClient.dbClose()
        return 'success'

    # mongoDB storage   第�??��??�是丟�??��?mongoOrder or mongoClient
    def insertOrder(self,mongoOrder,_OrderNo,_PurchaseDate,_PartAction,_PartNum,_PartQuility,_firm,_supplier):
        businessorder_doc={ 'OrderNo':_OrderNo,'PurchaseDate':_PurchaseDate,'PartAction':[_PartAction],\
                           'PartNum':[_PartNum],'PartQuility':[_PartQuility],\
                           'firm':[_firm],'supplier':[_supplier]
                            }

        mongoOrder.cursor.insert(businessorder_doc)
    def updataOrder(self,mongoOrder,_OrderNo,_PurchaseDate,_PartAction,_PartNum,_PartQuility,_firm,_supplier):
        mongoOrder.cursor.update({ "OrderNo" : _OrderNo,'firm':_firm,'supplier':_supplier},\
                           {'$push':{'PurchaseDate':_PurchaseDate,'PartAction':_PartAction ,'PartNum':_PartNum,'PartQuility':_PartQuility}}
                                                )
    def insertClient(self,mongodbClient,_OrderNo,_ClientName,_ClientPhone,_ClientAdd,_firm,_supplier):

        businessorder_doc={'OrderNo':_OrderNo,'ClientName':[_ClientName],'ClientPhone':[_ClientPhone],'ClientAdd':[_ClientAdd],'firm':[_firm],'supplier':[_supplier]}

        mongodbClient.cursor.insert(businessorder_doc)
    def updataClient(self,mongodbClient,_OrderNo,_ClientName,_ClientPhone,_ClientAdd,_firm,_supplier):
        mongodbClient.cursor.update( { 'ClientName':_ClientName,'ClientPhone':_ClientPhone,'ClientAdd':_ClientAdd,'firm':_firm,'supplier':_supplier}\
                           ,{'$push':{"OrderNo" : _OrderNo}})







