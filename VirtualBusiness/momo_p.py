# -*-  coding: utf-8  -*-
__author__ = '10409003'
from ToMysql import ToMysql
from pdfminer.pdfinterp import PDFResourceManager, PDFPageInterpreter
from pdfminer.converter import TextConverter
from pdfminer.layout import LAParams
from pdfminer.pdfpage import PDFPage
from cStringIO import StringIO
import re
import uuid
from aes_data import aes_data


class Momo_Datap():
    Data=None

    def __init__(self):
        pass
    def Momo_Datap(self,supplier,GroupID,path,UserID):
        mysqlconnect = ToMysql()
        mysqlconnect.connect()
        rsrcmgr = PDFResourceManager()
        retstr = StringIO()
        codec = 'utf-8'
        laparams = LAParams()
        device = TextConverter(rsrcmgr, retstr, codec=codec, laparams=laparams)
        fp = file(path, 'rb')
        interpreter = PDFPageInterpreter(rsrcmgr, device)
        password = ""
        maxpages = 0
        caching = True
        pagenos=set()
        for page in PDFPage.get_pages(fp, pagenos, maxpages=maxpages, password=password,caching=caching, check_extractable=True):
            interpreter.process_page(page)
        text = retstr.getvalue()
        text= text.split("\n")
        # print len(text)
        singles = []
        for single in text:
            if single not in singles:
                singles.append(single)
        for sub in singles:
            if re.match('0\\d{1,3}-?(\\d{6,8})(#\\d{1,5}){0,1}',sub):
                aes=aes_data()
                mobile = int(singles.index(sub))
                Tel= singles[mobile].split('/')[0]
                ClientTel = aes.AESencrypt("p@ssw0rd", Tel, True)
                Phone = singles[mobile].split('/')[-1]
                ClientPhone = aes.AESencrypt("p@ssw0rd", Phone, True)
                name=int(singles.index(sub))+1
                Name= singles[name]
                ClientName= aes.AESencrypt("p@ssw0rd", Name, True)
                address=int(singles.index(sub))-5
                Add = singles[address]
                ClientAdd = aes.AESencrypt("p@ssw0rd", Add, True)
                ordernum = int(singles.index(sub))+13
                OrderNo = singles[ordernum].split(':')[-1]
                SaleOrdersel="""select customer_id,name from tb_sale where order_no = '%s' and group_id = '%s' """ % (OrderNo,GroupID)
                mysqlconnect.cursor.execute(SaleOrdersel)
                orderexist = mysqlconnect.cursor.fetchall()
                if orderexist != []:
                    
                
                    updateCustomer="""update tb_customer
                                   set name= %s ,
                                   address= %s ,
                                   phone= %s,
                                   mobile = %s
                                   where customer_id =%s;
                                   """
                    updateSaleValue = (ClientName,ClientAdd,ClientTel,ClientPhone,orderexist[0][0])
                    mysqlconnect.cursor.execute(updateCustomer,updateSaleValue)
                    mysqlconnect.db.commit()
                
                else:
                    CustomereSQLsel = """select group_id,name,address,mobile from tb_customer where group_id='%s'""" % (GroupID)
                    mysqlconnect.cursor.execute(CustomereSQLsel)
                    Finalresult = mysqlconnect.cursor.fetchall()
                    customer_id_temp=[]
                    SalestrSQLsel_1="SELECT customer_id from tb_customer where name =%s and address= %s;"
                    SalestrSQLsel_2 = "SELECT customer_id from tb_customer where name =%s and mobile=%s;"
                    same_name=[]
                    for x in Finalresult:
                        Name_compare = aes.AESdecrypt("p@ssw0rd",x[1], True)
                        if Name ==  Name_compare :
                            print "the same name data list"
                            same_name.append(x)
                    for r in same_name:
                        address_compare = aes.AESdecrypt("p@ssw0rd", r[2], True)
                        mobile_compare = aes.AESdecrypt("p@ssw0rd", r[3], True)
                        if Add == address_compare:
                            mysqlconnect.cursor.execute(SalestrSQLsel_1, (str(r[1]), str(r[2])))
                            customer_id_temp.append(mysqlconnect.cursor.fetchall()[0])
                        elif Phone == mobile_compare:
                            mysqlconnect.cursor.execute(SalestrSQLsel_2, (str(r[1]), str(r[3])))
                            customer_id_temp.append(mysqlconnect.cursor.fetchall()[0])

                    print customer_id_temp
                    mysqlconnect.db.commit()
                    if customer_id_temp !=[]:
                        for y in customer_id_temp:
                            for x in Finalresult:
                                if aes.AESdecrypt('p@ssw0rd', x[1], True) == Name:
                                    InsertSQLins = """INSERT INTO tb_sale (sale_id,group_id,order_no,user_id,product_id,customer_id,name,order_source)
                                            VALUES(%s,%s,%s,%s,%s,%s,%s,%s);"""
                                            
                                    SaleSQL = (str(uuid.uuid4()),GroupID, OrderNo, UserID,str(uuid.uuid4()),y[0],x[1],supplier)
                                    mysqlconnect.cursor.execute(InsertSQLins,SaleSQL)
                                    mysqlconnect.db.commit()
                    else:
                        CustomereSQL = (
                        str(uuid.uuid4()), GroupID, ClientName, ClientAdd,ClientTel,ClientPhone, None, None, None, None)
                        CustomereSQLsel = """select group_id,name,address,mobile from tb_customer where group_id='%s'""" % (GroupID)
                        CustomereSQLins = """ insert into tb_customer
                                                (customer_id,group_id ,name,address,phone,mobile,email,post,class,memo)
                                                values(%s,%s,%s,%s,%s,%s,%s,%s,%s,%s);"""

                        mysqlconnect.cursor.execute(CustomereSQLins, CustomereSQL)

                        mysqlconnect.cursor.execute(CustomereSQLsel)
                        Finalresult = mysqlconnect.cursor.fetchall()
                        customer_id_temp = []
                        SalestrSQLsel_1 = "SELECT customer_id from tb_customer where name =%s and address= %s;"
                        SalestrSQLsel_2 = "SELECT customer_id from tb_customer where name =%s and mobile= %s;"
                        same_name = []
                        for x in Finalresult:
                            Name_compare = aes.AESdecrypt("p@ssw0rd", x[1], True)
                            if Name == Name_compare:
                                print "the same name data list"
                                same_name.append(x)

                        for r in same_name:
                            address_compare = aes.AESdecrypt("p@ssw0rd", r[2], True)
                            mobile_compare = aes.AESdecrypt("p@ssw0rd", r[3], True)
                            if Add == address_compare:
                                mysqlconnect.cursor.execute(SalestrSQLsel_1, (str(r[1]), str(r[2])))
                                customer_id_temp.append(mysqlconnect.cursor.fetchall()[0])
                            elif Phone == mobile_compare:
                                mysqlconnect.cursor.execute(SalestrSQLsel_2, (str(r[1]), str(r[2])))
                                customer_id_temp.append(mysqlconnect.cursor.fetchall()[0])

                        print customer_id_temp
                        for y in customer_id_temp:
                            for x in Finalresult:
                                if aes.AESdecrypt('p@ssw0rd', x[1], True) == Name:
                                    InsertSQLins = """INSERT INTO tb_sale (sale_id,group_id,order_no,user_id,product_id,customer_id,name,order_source)
                                            VALUES(%s,%s,%s,%s,%s,%s,%s,%s);"""
                                    SaleSQL = (str(uuid.uuid4()),GroupID, OrderNo, UserID,str(uuid.uuid4()),y[0],x[1],supplier)
                                    mysqlconnect.cursor.execute(InsertSQLins,SaleSQL)
                        mysqlconnect.db.commit()
                return 'success'





#if __name__=='__main__':
    #data=Momo_Datap()
    #data.Momo_Datap('momo','123','C:/Users/10409003/Desktop/delivery.pdf','456')