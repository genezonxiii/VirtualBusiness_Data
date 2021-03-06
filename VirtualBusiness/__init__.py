# -*-  coding: utf-8  -*-
__author__ = '10409003'
from pymongo import MongoClient
import xlrd
import datetime
# from datetime import datetime
from aes_data import aes_data
import uuid
import ToMysql
import ToMongodb
import logging
import time , chardet
import mysql.connector
from setting import DBSetting
import re

logger = logging.getLogger(__name__)

class convertType():
    def __init__(self):
        pass

    def ToDateTime(self,value):
        try:
            if value== None :
                return None
            elif value=='':
                return None
            else:
                dateObj=datetime.datetime.strptime(str(value),'%Y/%m/%d %H:%M')
                return dateObj
        except Exception as e:
            logger.ERROR(e.message)

    def ToDateTimeMDHM(self,value):
        try:
            if value== None :
                return None
            elif value=='':
                return None
            else:
                dateObj=datetime.datetime.strptime(str(value),'%m/%d %H:%M')
                return dateObj
        except Exception as e:
            logger.ERROR(e.message)

    def ToDateTimeYYYYMMDD(self,value):
        try:
            if value== None :
                return None
            elif value=='':
                return None
            else:
                dateObj=datetime.datetime.strptime(str(value),'%Y/%m/%d')
                return dateObj
        except Exception as e:
            logger.ERROR(e.message)

    def ToDateTimeYYYYMMDDHHMM(self,value):
        try:
            if value== None :
                return None
            elif value=='':
                return None
            else:
                dateObj=datetime.datetime.strptime(str(value),'%Y/%m/%d/%H/%M')
                return dateObj
        except Exception as e:
            logger.ERROR(e.message)

    def ToDateTimeYMD(self,value):
        try:
            if value== None :
                return None
            elif value=='':
                return None
            else:
                dateObj=datetime.datetime.strptime(str(value),'%Y-%m-%d')
                return dateObj
        except Exception as e:
            logger.ERROR(e.message)

    def ToDateTimeYMDHM(self,value):
        try:
            if value== None :
                return None
            elif value=='':
                return None
            else:
                dateObj=datetime.datetime.strptime(str(value),'%Y-%m-%d %H:%M')
                return dateObj
        except Exception as e:
            logger.ERROR(e.message)

    def ToDateTimeYMDHMS(self,value):
        try:
            if value== None :
                return None
            elif value=='':
                return None
            else:
                dateObj=datetime.datetime.strptime(str(value),'%Y-%m-%d %H:%M:%S')
                return dateObj
        except Exception as e:
            logger.ERROR(e.message)

    def ToDateTimeYMDHMSF(self,value):
        try:
            if value== None :
                return None
            elif value=='':
                return None
            else:
                dateObj=datetime.datetime.strptime(str(value),'%Y-%m-%d %H:%M:%S.%f')
                return dateObj
        except Exception as e:
            logger.ERROR(e.message)

    def ToDateTimeYYYYMMDD_float(self,value):
        try:
            if value== None :
                return None
            elif value=='':
                return None
            else:
                ct = convertType()
                dateObj=datetime.datetime.strptime(ct.getdate(value),'%Y/%m/%d')
                return dateObj
        except Exception as e:
            logger.ERROR(e.message)

    def ToDateTimeYMDHMS_float(self,value):
        try:
            if value== None :
                return None
            elif value=='':
                return None
            else:
                ct = convertType()
                dateObj=datetime.datetime.strptime(ct.xldate_to_datetime(value),"%Y-%m-%d %H:%M:%S")
                return dateObj
        except Exception as e:
            logger.ERROR(e.message)

    def ToDateTime2(self,value):
        try:
            if value== None :
                return None
            elif value=='':
                return None
            else:
                dateObj=datetime.datetime.strptime(str(value),'%Y/%m/%d/%H/%M')
                return dateObj
        except Exception as e:
            logger.ERROR(e.message)

    def getdate(self, date):
        __s_date = datetime.date(1899, 12, 31).toordinal() - 1
        if isinstance(date, float):
            date = int(date)
        d = datetime.date.fromordinal(__s_date + date)
        return d.strftime("%Y/%m/%d")

    def xldate_to_datetime(self,xldate):
        tempDate = datetime.datetime(1899, 12, 30)
        deltaDays = datetime.timedelta(days=int(xldate))
        secs = (int((xldate % 1) * 86400) - 60)
        detlaSeconds = datetime.timedelta(seconds=secs)
        TheTime = (tempDate + deltaDays + detlaSeconds)
        return TheTime.strftime("%Y-%m-%d %H:%M:%S")

    def ToInt(self,value):
        try:
            if value== None :
                return 0
            elif value=='':
                return 0
            else:
                return int(float(value))
        except Exception as e:
            logger.ERROR(e.message)

    def ToFloat(self,value):
        try:
            if value== None :
                return 0.0
            elif value=='':
                return 0.0
            elif type(value) == unicode:
                value = str(value)
                if value.find(',') >= 0:
                    return float(value.replace(",",""))
                else:
                    return float(value)
            elif type(value) == float:
                return value
        except Exception as e:
            logger.ERROR(e.message)

    def ToString(self,value):
        try:
            if value== None :
                return None
            elif value=='':
                return ''
            else:
                return str(value)
        except Exception as e:
            logger.ERROR(e.message)

    def ToStringNoEncode(self,value):
        try:
            if value== None :
                return None
            elif value=='':
                return ''
            else:
                return value
        except Exception as e:
            logger.ERROR(e.message)

    def ToBoolean(self,value):
        try:
            if value == True:
                return True
            else:
                return False
        except Exception as e :
            logger.ERROR(e.message)

class Sale():
    p_sale_id, p_seq_no, p_group_id, p_order_no = None, None, None, None
    p_user_id, p_product_id, p_product_name, p_product_spec, p_c_product_id = None, None, None, None, None
    p_customer_id, p_name, p_invoice,o_name = None, None, None, None
    p_quantity, p_price = 0, 0.0
    p_invoice_date, p_trans_list_date, p_dis_date, p_sale_date, p_return_date = None, None, None, None, None
    p_memo, p_order_source, p_deliveryway = None, None, None
    p_total_amt, p_order_status, p_deliver_name, p_deliver_to, p_deliver_store = 0, None, None, None, None
    p_deliver_phone, p_deliver_mobile, p_pay_kind, p_pay_status = None, None, None, None
    p_user_def1, p_user_def2, p_user_def3, p_user_def4, p_user_def5 = None, None, None, None, None
    p_user_def6, p_user_def7, p_user_def8, p_user_def9, p_user_def10 = None, None, None, None, None
    p_isreturn = False
    aes,convert = None, None

    def __init__(self):
        self.convert = convertType()
        self.aes = aes_data()

    def __del__(self):
        self.convert = None
        self.aes = None

    def setSale_id(self, value):
        self.p_sale_id = self.convert.ToString(value.encode('utf-8'))

    def setSeq_no(self,value):
        self.p_seq_no = self.convert.ToString(value.encode('utf-8'))

    def setGroup_id(self,value):
        self.p_group_id = self.convert.ToString(value.encode('utf-8'))

    def setOrder_No(self,value):
        self.p_order_no = self.convert.ToString(value.encode('utf-8'))

    def setOrder_No_float(self,value):
        self.p_order_no = self.convert.ToString(value).split(".")[0]

    def setUser_id(self,value):
        self.p_user_id = self.convert.ToString(value.encode('utf-8'))

    def setProduct_id(self,value):
        self.p_product_id = self.convert.ToString(value.encode('utf-8'))

    def setProduct_name(self,value):
        self.p_product_name = self.convert.ToString(value.decode('utf-8').encode("utf-8"))

    def setProduct_spec(self,value):
        self.p_product_spec = self.convert.ToString(value.decode('utf-8').encode("utf-8"))

    def setProduct_name_NoEncode(self,value):
        self.p_product_name = self.convert.ToStringNoEncode(value)

    def setProduct_spec_NoEncode(self,value):
        self.p_product_spec = self.convert.ToStringNoEncode(value)

    def setC_Product_id(self,value):
        self.p_c_product_id = self.convert.ToString(value.encode('utf-8'))

    def setC_Product_id_float(self,value):
        self.p_c_product_id = self.convert.ToString(value).split(".")[0]

    def setCustomer_id(self,value):
        self.p_customer_id = self.convert.ToString(value.encode('utf-8'))

    def setName(self,value):
        if value == None or value == '' or value == 'NULL':
            self.p_name = None
            self.o_name = None
        else:
            self.p_name = self.aes.AESencrypt('p@ssw0rd', value.decode('utf-8').encode('utf-8'), True)
            self.o_name = self.convert.ToString(value.decode('utf-8').encode('utf-8'))

    def setNameNoEncode(self,value):
        if value == None or value == '' or value == 'NULL':
            self.p_name = None
            self.o_name = None
        else:
            self.p_name = self.aes.AESencrypt('p@ssw0rd', value.encode('utf-8'), True)
            self.o_name = self.convert.ToStringNoEncode(value)

    def setInvoice(self,value):
        self.p_invoice = self.convert.ToString(value.encode('utf-8'))

    def setQuantity(self,value):
        self.p_quantity = self.convert.ToInt(int(value))

    def setPrice(self,value):
        self.p_price = self.convert.ToFloat(value)

    def setPrice_str(self,value):
        self.p_price = self.convert.ToString(int(value))

    def setInvoice_date(self,value):
        self.p_invoice_date = self.convert.ToDateTime(value)

    def setTrans_list_date(self,value):
        self.p_trans_list_date = self.convert.ToDateTime(value)

    def setTrans_list_date_MDHM(self,value):
        self.p_trans_list_date = self.convert.ToDateTimeMDHM(value)

    def setTrans_list_date_YYYYMMDD(self,value):
        self.p_trans_list_date = self.convert.ToDateTimeYYYYMMDD(value)

    def setTrans_list_date_YYYYMMDDHHMM(self,value):
        self.p_trans_list_date = self.convert.ToDateTimeYYYYMMDDHHMM(value)

    def setTrans_list_date_YMD(self,value):
        self.p_trans_list_date = self.convert.ToDateTimeYMD(value)

    def setTrans_list_date_YMDHM(self,value):
        self.p_trans_list_date = self.convert.ToDateTimeYMDHM(value)

    def setTrans_list_date_YMDHMS(self,value):
        self.p_trans_list_date = self.convert.ToDateTimeYMDHMS(value)

    def setTrans_list_date_YMDHMSF(self,value):
        self.p_trans_list_date = self.convert.ToDateTimeYMDHMSF(value)

    def setTrans_list_date_YYYYMMDD_float(self,value):
        self.p_trans_list_date = self.convert.ToDateTimeYYYYMMDD_float(value)

    def setTrans_list_date_YMDHMS_float(self,value):
        self.p_trans_list_date = self.convert.ToDateTimeYMDHMS_float(value)

    def setTrans_list_date_udn(self,value):
        self.p_trans_list_date = self.convert.ToDateTime2(value)

    def setDis_date(self,value):
        self.p_dis_date = self.convert.ToDateTime(value)

    def setSale_date(self,value):
        self.p_sale_date = self.convert.ToDateTime(value)

    def setSale_date_MDHM(self,value):
        self.p_sale_date = self.convert.ToDateTimeMDHM(value)

    def setSale_date_YMD(self,value):
        self.p_sale_date = self.convert.ToDateTimeYMD(value)

    def setSale_date_YMDHMS(self,value):
        self.p_sale_date = self.convert.ToDateTimeYMDHMS(value)

    def setSale_date_YMDHM(self,value):
        self.p_sale_date = self.convert.ToDateTimeYMDHM(value)

    def setSale_date_YMDHMS_float(self,value):
        self.p_trans_list_date = self.convert.ToDateTimeYMDHMS_float(value)

    def setSale_date_YMDHMSF(self,value):
        self.p_sale_date = self.convert.ToDateTimeYMDHMSF(value)

    def setSale_date_YYYYMMDD(self,value):
        self.p_sale_date = self.convert.ToDateTimeYYYYMMDD(value)

    def setSale_date_YYYYMMDDHHMM(self,value):
        self.p_sale_date = self.convert.ToDateTimeYYYYMMDDHHMM(value)

    def setSale_date_YYYYMMDD_float(self,value):
        self.p_sale_date = self.convert.ToDateTimeYYYYMMDD_float(value)

    def setReturn_date(self,value):
        self.p_return_date = self.convert.ToDateTime(value)

    def setMemo(self,value):
        self.p_memo = self.convert.ToString(value.encode('utf-8'))

    def setOrder_source(self,value):
        # self.p_order_source = self.convert.ToString(value.encode('utf-8'))
        self.p_order_source = self.convert.ToString(value.decode('utf-8').encode('utf-8'))

    def setOrder_sourceNodecode(self,value):
        self.p_order_source = self.convert.ToString(value.encode('utf-8'))

    def setDeliveryway(self,value):
        self.p_deliveryway = self.convert.ToString(value.encode('utf-8'))

    def setIsreturn(self,value):
        self.p_isreturn = self.convert.ToBoolean(value)

    def setTotal_amt(self,value):
        self.p_total_amt = self.convert.ToFloat(value)

    def setOrder_status(self,value):
        self.p_order_status = self.convert.ToString(value.encode('utf-8'))

    def setDeliver_name(self,value):
        self.p_deliver_name = self.convert.ToString(value.encode('utf-8'))

    def setDeliver_to(self,value):
        self.p_deliver_to = self.convert.ToString(value.encode('utf-8'))

    def setDeliver_store(self,value):
        self.p_deliver_store = self.convert.ToString(value.encode('utf-8'))

    def setDeliver_phone(self,value):
        self.p_deliver_phone = self.convert.ToString(value.encode('utf-8'))

    def setDeliver_mobile(self,value):
        self.p_deliver_mobile = self.convert.ToString(value.encode('utf-8'))

    def setPay_kind(self,value):
        self.p_pay_kind = self.convert.ToString(value.encode('utf-8'))

    def setPay_status(self,value):
        self.p_pay_status = self.convert.ToString(value.encode('utf-8'))

    def setUser_def1(self,value):
        self.p_user_def1 = self.convert.ToString(value.encode('utf-8'))

    def setUser_def2(self,value):
        self.p_user_def2 = self.convert.ToString(value.encode('utf-8'))

    def setUser_def3(self,value):
        self.p_user_def3 = self.convert.ToString(value.encode('utf-8'))

    def setUser_def4(self,value):
        self.p_user_def4 = self.convert.ToString(value.encode('utf-8'))

    def setUser_def5(self,value):
        self.p_user_def5 = self.convert.ToString(value.encode('utf-8'))

    def setUser_def6(self,value):
        self.p_user_def6 = self.convert.ToString(value.encode('utf-8'))

    def setUser_def7(self,value):
        self.p_user_def7 = self.convert.ToString(value.encode('utf-8'))

    def setUser_def8(self,value):
        self.p_user_def8 = self.convert.ToString(value.encode('utf-8'))

    def setUser_def9(self,value):
        self.p_user_def9 = self.convert.ToString(value.encode('utf-8'))

    def setUser_def10(self,value):
        self.p_user_def10 = self.convert.ToString(value.encode('utf-8'))

    def getSale_id(self):
        return self.p_sale_id

    def getSeq_no(self):
        return self.p_seq_no

    def getGroup_id(self):
        return self.p_group_id

    def getOrder_No(self):
        return self.p_order_no

    def getUser_id(self):
        return self.p_user_id

    def getProduct_id(self):
        return self.p_product_id

    def getProduct_name(self):
        return self.p_product_name

    def getProduct_spec(self):
        return self.p_product_spec

    def getC_Product_id(self):
        return self.p_c_product_id

    def getCustomer_id(self):
        return self.p_customer_id

    def getName(self):
        return self.p_name

    def get_Name(self):
        return self.o_name

    def getInvoice(self):
        return self.p_invoice

    def getQuantity(self):
        return self.p_quantity

    def getPrice(self):
        return self.p_price

    def getInvoice_date(self):
        return self.p_invoice_date

    def getTrans_list_date(self):
        return self.p_trans_list_date

    def getDis_date(self):
        return self.p_dis_date

    def getSale_date(self):
        return self.p_sale_date

    def getReturn_date(self):
        return self.p_return_date

    def getMemo(self):
        return self.p_memo

    def getOrder_source(self):
        return self.p_order_source

    def getDeliveryway(self):
        return self.p_deliveryway

    def getIsreturn(self):
        return self.p_isreturn

    def getTotal_amt(self):
        return self.p_total_amt

    def getOrder_status(self):
        return self.p_order_status

    def getDeliver_name(self):
        return self.p_deliver_name

    def getDeliver_to(self):
        return self.p_deliver_to

    def getDeliver_store(self):
        return self.p_deliver_store

    def getDeliver_phone(self):
        return self.p_deliver_phone

    def getDeliver_mobile(self):
        return self.p_deliver_mobile

    def getPay_kind(self):
        return self.p_pay_kind

    def getPay_status(self):
        return self.p_pay_status

    def getUser_def1(self):
        return self.p_user_def1

    def getUser_def2(self):
        return self.p_user_def2

    def getUser_def3(self):
        return self.p_user_def3

    def getUser_def4(self):
        return self.p_user_def4

    def getUser_def5(self):
        return self.p_user_def5

    def getUser_def6(self):
        return self.p_user_def6

    def getUser_def7(self):
        return self.p_user_def7

    def getUser_def8(self):
        return self.p_user_def8

    def getUser_def9(self):
        return self.p_user_def9

    def getUser_def10(self):
        return self.p_user_def10

class Customer():
    p_customer_id,p_group_id,p_name=None,None,None
    p_address, p_phone, p_mobile=None,None,None
    p_email, p_post, p_class, p_memo = None,None,None,None
    o_name,o_address, o_phone, o_mobile,o_email = None,None,None,None,None
    aes, convert , p_user_id = None,None,None

    def __init__(self):
        self.aes = aes_data()
        self.convert = convertType()

    def __del__(self):
        self.aes = None
        self.convert = None

    def setCustomer_id(self,value):
        self.p_customer_id = self.convert.ToString(value)

    def setGroup_id(self,value):
        self.p_group_id = self.convert.ToString(value)

    def setUser_id(self,value):
        self.p_user_id = self.convert.ToString(value)

    def setName(self,value):
        if value == None or value == '' or value == 'NULL':
            self.p_name = None
            self.o_name = None
        else:
            self.p_name = self.aes.AESencrypt('p@ssw0rd', value.decode('utf-8').encode('utf-8'), True)
            self.o_name = self.convert.ToString(value.decode('utf-8').encode('utf-8'))

    def setNameNoEncode(self,value):
        if value == None or value == '' or value == 'NULL':
            self.p_name = None
            self.o_name = None
        else:
            self.p_name = self.aes.AESencrypt('p@ssw0rd', value.encode('utf-8'), True)
            self.o_name = self.convert.ToStringNoEncode(value)

    def setAddress(self,value):
        if value == None or value == '' or value == 'NULL':
            self.p_address = None
            self.o_address = None
        else:
            self.p_address = self.aes.AESencrypt('p@ssw0rd', value.decode('utf-8').encode('utf-8'), True)
            self.o_address = self.convert.ToString(value.decode('utf-8').encode('utf-8'))

    def setAddressNoEncode(self,value):
        if value == None or value == '' or value == 'NULL':
            self.p_address = None
            self.o_address = None
        else:
            self.p_address = self.aes.AESencrypt('p@ssw0rd', value.encode('utf-8'), True)
            self.o_address = self.convert.ToStringNoEncode(value)

    def setPhone(self,value):
        if value == None or value == '' or value == 'NULL':
            self.p_phone = None
            self.o_phone = None
        else:
            if type(value) <> unicode :
                if type(value) == float :
                    value = str(int(value))
                else:
                    value = str(value)
            self.p_phone = self.aes.AESencrypt('p@ssw0rd', value.encode('utf-8'), True)
            self.o_phone = self.convert.ToString(value.encode('utf-8'))

    def setMobile(self,value):
        if value == None or value == '' or value == 'NULL':
            self.p_mobile = None
            self.o_mobile = None
        else:
            if type(value) <> unicode :
                if type(value) == float :
                    value = str(int(value))
                else:
                    value = str(value)
            self.p_mobile = self.aes.AESencrypt('p@ssw0rd', value.encode('utf-8'), True)
            self.o_mobile = self.convert.ToString(value.encode('utf-8'))

    def setEmail(self,value):
        if value == None or value == '' or value == 'NULL':
            self.p_email = None
            self.o_email = None
        else:
            self.p_email = self.aes.AESencrypt('p@ssw0rd', value.encode('utf-8'), True)
            self.o_email = self.convert.ToString(value.encode('utf-8'))

    def setPost(self,value):
        self.p_post = self.convert.ToString(value)

    def setClass(self,value):
        self.p_class = self.convert.ToString(value)

    def setMemo(self,value):
        self.p_memo = self.convert.ToString(value)

    def getCustomer_id(self):
        return  self.p_customer_id

    def getGroup_id(self):
        return self.p_group_id

    def getUser_id(self):
        return self.p_user_id

    def getName(self):
        return self.p_name

    def getAddress(self):
        return self.p_address

    def getphone(self):
        return self.p_phone

    def getMobile(self):
        return self.p_mobile

    def getEmail(self):
        return self.p_email

    def get_Name(self):
        return self.o_name

    def get_Address(self):
        return self.o_address

    def get_phone(self):
        return self.o_phone

    def get_Mobile(self):
        return self.o_mobile

    def get_Email(self):
        return self.o_email

    def getPost(self):
        return self.p_post

    def getClass(self):
        return self.p_class

    def getMemo(self):
        return self.p_memo

class updateCustomer():
    conn = None
    def __init__(self):
        pass

    def checkCustomerid(self,p_group_id,p_name,p_address,p_phone,p_mobile,p_email):
        try:
            p_result=None
            paremeter = (p_group_id,p_name,p_address,p_phone,p_mobile,p_email,p_result)
            print paremeter
            if self.conn == None :
                self.getConnection()
            cursor = self.conn.cursor()
            result = cursor.callproc('sp_find_customer',paremeter)
            # self.conn.close()
            if result==None:
                return None
            else:
                return result[6]
        except mysql.connector.Error:
            logger.error("Connection DB Error")
            raise
        except Exception as e:
            print e.message
            logger.error(e.message)
            raise

    def updataData(self,parameter):
        try:
            if self.conn == None :
                self.getConnection()
            cursor = self.conn.cursor()
            cursor.callproc('sp_update_customer', parameter)
            cursor.close()
            self.conn.commit()
            return True
        except mysql.connector.Error:
            logger.error("Connection DB Error")
            raise
            return False
        except Exception as e:
            logger.error(e.message)
            raise
            return False

    def getConnection(self):
        try:
            config = DBSetting()
            self.conn = mysql.connector.connect(user=config.dbUser, password=config.dbPassword,
                                       host=config.dbHost, database=config.dbName)
        except mysql.connector.Error:
            logger.error("Connection DB Error")
            raise
        except Exception as e:
            logger.error(e.message)
            raise

    def __del__(self):
        if self.conn <> None :
            self.conn.close()
            self.conn = None

# class DBSetting():
#     dbHost ,dbName , dbUser , dbPassword = None , None ,None ,None
#     def __init__(self):
#         self.dbHost = '192.168.112.164'
#         self.dbName = 'tmp'
#         self.dbUser = 'root'
#         self.dbPassword = 'admin123'
#         # self.dbHost = 'localhost'
#         # self.dbName = 'tmp'
#         # self.dbUser = 'root'
#         # self.dbPassword = 'mysql'

class detectFile():
    def __init__(self):
        pass

    def detect(self,filename):
        try:
            File = open(filename).read()
            result = chardet.detect(File)
            return result.get('encoding')
        except Exception as e:
            logger.error(e.message)

class checkNum():
    def __init__(self):
        pass

    def getNumber(self, strWord):
        try:
            num = ""
            for i in range(len(strWord)):
                if strWord[i].isdigit():
                    num += strWord[i]
                else:
                    break
            return num
        except Exception as e:
            logger.error(e.message)

class ASAP():
    Data=None
    def __init__(self):
        pass

    def ASAP_Data(self,supplier,GroupID,path,collection_order,collection_client):
        logging.basicConfig(filename='pyupload.log', level=logging.DEBUG, format='%(asctime)s %(message)s', datefmt='%Y/%m/%d %I:%M:%S %p')
        logging.Formatter.converter = time.gmtime
        logger.info('===ASAP===')
        logger.debug('supplier:' + supplier)
        logger.debug('GroupID:' + GroupID)
        logger.debug('path:' + path)
        logger.debug('UserID:' + UserID)
        
        #mysql connector object
        mysqlconnect=ToMysql()
        mysqlCursor=mysqlconnect.connect()

        #mongodb connector object?��?  ?��??��??�object ???�???object.cursor.?��??��?�?��ＵＤ
        mongoOrder=ToMongodb()
        mongoOrder.setcollection('co_ordertest')
        mongoOrder.connect()
        mongoOrder.cursor.find()

        mongodbClient=ToMongodb()
        mongodbClient.setcollection('co_co_clienttest')
        mongodbClient.connect()
        mongodbClient.cursor.updateOne()

        Ordernum=""
        Clientnum=""

        data=xlrd.open_workbook(path)
        table=data.sheets()[0]
        num_cols=table.ncols
        #put the data into the corresponding variable
        for row_index in range(1,table.nrows):
            for col_index in range(0,num_cols):
                aes=aes_data()
                strTurntDate    =str(table.cell(row_index,0).value).replace("/","-")
                TurntDate       =datetime.strptime(strTurntDate,'%Y-%m-%d')
                OrderNo         =table.cell(row_index,1).value[0:13]
                PartNum         =str(table.cell(row_index,2).value)
                PartMaterial    =str(table.cell(row_index,3).value)
                PartName        =table.cell(row_index,4).value
                PartColor       =table.cell(row_index,5).value
                PartSize        =table.cell(row_index,6).value
                PartQuility     =table.cell(row_index,7).value
                Name            =table.cell(row_index, 8).value
                ClientName      = aes.AESencrypt("password", Name.encode('utf-8'), True)
                Phone           = table.cell(row_index, 9).value
                ClientPhone     = aes.AESencrypt("password", Phone.encode('utf-8'), True)
                Tel             = table.cell(row_index, 10).value
                ClientTel       = aes.AESencrypt("password", Tel.encode('utf-8'), True)
                Add             = table.cell(row_index, 11).value
                ClientAdd       = aes.AESencrypt("password", Add.encode('utf-8'), True)
                PartNo          =str(table.cell(row_index,12).value)
                firm            =table.cell(row_index,13).value
                supplier        =supplier
                GroupID         =GroupID

            # if (Ordernum==OrderNo[0:13]):
            #     print 'update'
            #     self.updataOrder(mongoOrder,TurntDate,OrderNo,PartNum,PartMaterial,PartName,PartColor,PartSize,PartQuility,PartNo,firm,supplier)

                mysqlCursor.execute(self.ToSql(GroupID,OrderNo,PartNo,PartQuility,TurntDate,PartName,PartSize,supplier,ClientName,ClientAdd,ClientPhone,ClientTel,1))
                mysqlCursor.callproc(self.ToSql())

            # else:
            #     print 'insert'
            #     Ordernum=OrderNo
            #     self.insertOrder(mongoOrder,TurntDate,OrderNo,PartNum,PartMaterial,PartName,PartColor,PartSize,PartQuility,PartNo,firm,supplier)
            # if (Clientnum == ClientName):
            #     print 'update'
            #     self.updataClient(mongodbClient, OrderNo, ClientName, ClientPhone, ClientTel, ClientAdd, firm, supplier)
            # else:
            #     print 'insert'
            #     Clientnum = ClientName
            #     self.insertClient(mongodbClient, OrderNo, ClientName, ClientPhone, ClientTel, ClientAdd, firm, supplier)

        mysqlCursor.close()
        mongoOrder.close()
        mongodbClient.close()

    #mongoDB storage   第�??��??�是丟�??��?mongoOrder or mongoClient
    # def insertOrder(self,mongoOrder,_TurntDate,_OrderNo,_PartNum,_PartMaterial,_PartName,_PartColor,_PartSize,_PartQuility,_PartNo,_firm,_supplier):
    #     businessorder_doc={ 'TurntDate':_TurntDate,
    #                         'OrderNo':_OrderNo,
    #                         'PartNum':[_PartNum],
    #                         'PartMaterial':[_PartMaterial],\
    #                         'PartName':[_PartName],
    #                         'PartColor':[_PartColor],
    #                         'PartSize':[_PartSize],\
    #                         'PartQuility':[_PartQuility],
    #                         'PartNo':[_PartNo],
    #                         'firm':[_firm],
    #                         'supplier':[_supplier]
    #                         }
    #
    #     mongoOrder.cursor.insert(businessorder_doc)
    # def updataOrder(self,mongoOrder,_TurntDate,_OrderNo,_PartNum,_PartMaterial,_PartName,_PartColor,_PartSize,_PartQuility,_PartNo,_firm,_supplier):
    #     mongoOrder.cursor.update( { "OrderNo" : _OrderNo,
    #                                 'firm':_firm,'supplier':_supplier},\
    #                                 {'$push':{  'TurntDate':_TurntDate,
    #                                             'PartNum':_PartNum ,
    #                                             'PartMaterial':_PartMaterial,
    #                                             'PartName':_PartName,\
    #                                             'PartColor':_PartColor,
    #                                             'PartSize':_PartSize,
    #                                             'PartNo':_PartNo,
    #                                             'PartQuility':_PartQuility}
    #                                             }
    #                                             )
    # def insertClient(self,mongodbClient,_OrderNo,_ClientName,_ClientPhone,_ClientTel,_ClientAdd,_firm,_supplier):
    #
    #     businessorder_doc={ 'OrderNo':_OrderNo,
    #                         'ClientName':[_ClientName],
    #                         'ClientPhone':[_ClientPhone],
    #                         'ClientTel':[_ClientTel],
    #                         'ClientAdd':[_ClientAdd],
    #                         'firm':[_firm],
    #                         'supplier':[_supplier]}
    #
    #     mongodbClient.cursor.insert(businessorder_doc)
    # def updataClient(self,mongodbClient,_OrderNo,_ClientName,_ClientPhone,_ClientTel,_ClientAdd,_firm,_supplier):
    #     mongodbClient.cursor.update( {  'ClientName':_ClientName,
    #                                     'ClientPhone':_ClientPhone,
    #                                     'ClientTel':[_ClientTel],
    #                                     'ClientAdd':_ClientAdd,
    #                                     'firm':_firm,
    #                                     'supplier':_supplier}\
    #                                     ,{'$push':{"OrderNo" : _OrderNo}})


    #The judgement of mysql sql by sqlselect parameter
    def ToSql(self,GroupID,OrderNo,PartNo,PartQuility,TurntDate,PartName,PartSize,supplier,ClientName,ClientAdd,ClientPhone,ClientTel,sqlselect):
        SupplySQL=("INSERT INTO tb_sale (supply_id,group_id,supply_name)"
                  "VALUES(%s,%s,%s)", str(uuid.uuid4()),GroupID,supplier)
        SalestrSQL =("INSERT INTO tb_sale (sale_id,group_id,seq_no,c_product_id,customer_id,name,quantity,trans_list_date,sale_date)"
                  "VALUES(%s,%s,%s,%s,%s,%s,%s)", str(uuid.uuid4()),GroupID,OrderNo, PartNo, PartQuility,TurntDate,TurntDate)
        ProductSQL=("INSERT INTO tb_product (product_id,group_id,c_product_id,product_name,unit_id,supply_name)"
                  "VALUES(%s,%s,%s,%s,%s)", str(uuid.uuid4()),GroupID, PartNo, PartName,PartSize,supplier)
        CustomereSQL=("INSERT INTO tb_customer (product_id,group_id,name,address,phone,mobile)"
                  "VALUES(%s,%s,%s,%s,%s,%s)", str(uuid.uuid4()),GroupID, ClientName, ClientAdd,ClientTel,ClientPhone)

        if sqlselect==1:
            sql = SupplySQL
        elif sqlselect==2:
            sql = SalestrSQL
        elif sqlselect == 3:
            sql=ProductSQL
        elif sqlselect==4:
            sql = CustomereSQL
        else:
            print "Sqlselect error!"
        return sql

