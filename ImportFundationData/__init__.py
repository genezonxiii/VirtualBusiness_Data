# -*-  coding: utf-8  -*-
#__author__ = '10509002'

import logging
import os
import time
from ImportFundationData.insert import SupplyData
from ImportFundationData.insert import ProductData
from ImportFundationData.insert import PackageData
from ImportFundationData.insert import ContrastData

logger = logging.getLogger(__name__)

class DataInsert():
    Data=None

    def __init__(self):
        pass
    def virtualbusiness_import(self,DataPath,userID):
        Importtype = str(DataPath).split('/')[3]  #供應商、商品、組合包、對照表
        Groupid = str(DataPath).split('/')[4]      #group id

        print Importtype
        print Groupid


        logging.basicConfig(filename='/data/VirtualBusiness_Data/pyupload.log',
			level=logging.DEBUG,
			format='%(asctime)s - %(levelname)s - %(filename)s:%(name)s:%(module)s/%(funcName)s/%(lineno)d - %(message)s',
            datefmt='%Y/%m/%d %I:%M:%S %p')
        logging.Formatter.converter = time.gmtime

        logger.info('Importtype')
        logger.info(Importtype)

        if Importtype == 'supply':
            logger.debug('supply')
            FinalData = SupplyData()
            return FinalData.Supply(Groupid, os.path.join(DataPath),userID)

        if Importtype == 'product':
            logger.debug('product')
            FinalData = ProductData()
            return FinalData.Product(Groupid, os.path.join(DataPath), userID)

        if Importtype == 'package':
            logger.debug('package')
            FinalData = PackageData()
            return FinalData.Package(Groupid, os.path.join(DataPath), userID)

        if Importtype == 'contrast':
            logger.debug('contrast')
            FinalData = ContrastData()
            return FinalData.Contrast(Groupid, os.path.join(DataPath), userID)

# class convertType():
#     def __init__(self):
#         pass
#
#     def ToDateTime(self,value):
#         try:
#             if value== None :
#                 return None
#             elif value=='':
#                 return None
#             else:
#                 dateObj=datetime.datetime.strptime(str(value),'%Y/%m/%d %H:%M')
#                 return dateObj
#         except Exception as e:
#             logger.ERROR(e.message)
#
#     def ToDateTimeMDHM(self,value):
#         try:
#             if value== None :
#                 return None
#             elif value=='':
#                 return None
#             else:
#                 dateObj=datetime.datetime.strptime(str(value),'%m/%d %H:%M')
#                 return dateObj
#         except Exception as e:
#             logger.ERROR(e.message)
#
#     def ToDateTimeYYYYMMDD(self,value):
#         try:
#             if value== None :
#                 return None
#             elif value=='':
#                 return None
#             else:
#                 dateObj=datetime.datetime.strptime(str(value),'%Y/%m/%d')
#                 return dateObj
#         except Exception as e:
#             logger.ERROR(e.message)
#
#     def ToDateTimeYYYYMMDDHHMM(self,value):
#         try:
#             if value== None :
#                 return None
#             elif value=='':
#                 return None
#             else:
#                 dateObj=datetime.datetime.strptime(str(value),'%Y/%m/%d/%H/%M')
#                 return dateObj
#         except Exception as e:
#             logger.ERROR(e.message)
#
#     def ToDateTimeYMD(self,value):
#         try:
#             if value== None :
#                 return None
#             elif value=='':
#                 return None
#             else:
#                 dateObj=datetime.datetime.strptime(str(value),'%Y-%m-%d')
#                 return dateObj
#         except Exception as e:
#             logger.ERROR(e.message)
#
#     def ToDateTimeYMDHM(self,value):
#         try:
#             if value== None :
#                 return None
#             elif value=='':
#                 return None
#             else:
#                 dateObj=datetime.datetime.strptime(str(value),'%Y-%m-%d %H:%M')
#                 return dateObj
#         except Exception as e:
#             logger.ERROR(e.message)
#
#     def ToDateTimeYMDHMS(self,value):
#         try:
#             if value== None :
#                 return None
#             elif value=='':
#                 return None
#             else:
#                 dateObj=datetime.datetime.strptime(str(value),'%Y-%m-%d %H:%M:%S')
#                 return dateObj
#         except Exception as e:
#             logger.ERROR(e.message)
#
#     def ToDateTimeYMDHMSF(self,value):
#         try:
#             if value== None :
#                 return None
#             elif value=='':
#                 return None
#             else:
#                 dateObj=datetime.datetime.strptime(str(value),'%Y-%m-%d %H:%M:%S.%f')
#                 return dateObj
#         except Exception as e:
#             logger.ERROR(e.message)
#
#     def ToDateTimeYYYYMMDD_float(self,value):
#         try:
#             if value== None :
#                 return None
#             elif value=='':
#                 return None
#             else:
#                 ct = convertType()
#                 dateObj=datetime.datetime.strptime(ct.getdate(value),'%Y/%m/%d')
#                 return dateObj
#         except Exception as e:
#             logger.ERROR(e.message)
#
#     def ToDateTimeYMDHMS_float(self,value):
#         try:
#             if value== None :
#                 return None
#             elif value=='':
#                 return None
#             else:
#                 ct = convertType()
#                 dateObj=datetime.datetime.strptime(ct.xldate_to_datetime(value),"%Y-%m-%d %H:%M:%S")
#                 return dateObj
#         except Exception as e:
#             logger.ERROR(e.message)
#
#     def ToDateTime2(self,value):
#         try:
#             if value== None :
#                 return None
#             elif value=='':
#                 return None
#             else:
#                 dateObj=datetime.datetime.strptime(str(value),'%Y/%m/%d/%H/%M')
#                 return dateObj
#         except Exception as e:
#             logger.ERROR(e.message)
#
#     def getdate(self, date):
#         __s_date = datetime.date(1899, 12, 31).toordinal() - 1
#         if isinstance(date, float):
#             date = int(date)
#         d = datetime.date.fromordinal(__s_date + date)
#         return d.strftime("%Y/%m/%d")
#
#     def xldate_to_datetime(self,xldate):
#         tempDate = datetime.datetime(1899, 12, 30)
#         deltaDays = datetime.timedelta(days=int(xldate))
#         secs = (int((xldate % 1) * 86400) - 60)
#         detlaSeconds = datetime.timedelta(seconds=secs)
#         TheTime = (tempDate + deltaDays + detlaSeconds)
#         return TheTime.strftime("%Y-%m-%d %H:%M:%S")
#
#     def ToInt(self,value):
#         try:
#             if value== None :
#                 return 0
#             elif value=='':
#                 return 0
#             else:
#                 return int(float(value))
#         except Exception as e:
#             logger.ERROR(e.message)
#
#     def ToFloat(self,value):
#         try:
#             if value== None :
#                 return 0.0
#             elif value=='':
#                 return 0.0
#             else:
#                 return float(value.replace(",",""))
#         except Exception as e:
#             logger.ERROR(e.message)
#
#     def ToString(self,value):
#         try:
#             if value== None :
#                 return None
#             elif value=='':
#                 return ''
#             else:
#                 return str(value)
#         except Exception as e:
#             logger.ERROR(e.message)
#
#     def ToStringNoEncode(self,value):
#         try:
#             if value== None :
#                 return None
#             elif value=='':
#                 return ''
#             else:
#                 return value
#         except Exception as e:
#             logger.ERROR(e.message)
#
#     def ToBoolean(self,value):
#         try:
#             if value == True:
#                 return True
#             else:
#                 return False
#         except Exception as e :
#             logger.ERROR(e.message)
#
# class Supply():
#     p_group_id, p_supply_name, p_supply_unicode, p_address, p_contact = None, None, None, None, None
#     p_phone, p_ext, p_mobile,p_email, p_phone1 = None, None, None, None, None
#     p_ext1, p_mobile1, p_email1, p_contact1, p_memo, p_user_id = None, None, None, None, None, None
#     aes, convert = None, None
#
#     def __init__(self):
#         self.convert = convertType()
#         self.aes = aes_data()
#     def __del__(self):
#         self.convert = None
#         self.aes = None
#
#     def setGroup_id(self,value):
#         self.p_group_id = self.convert.ToString(value.encode('utf-8'))
#
#     def setSupply_name(self,value):
#         self.p_supply_name = self.convert.ToString(value.encode('utf-8'))
#
#     def setSupply_unicode(self,value):
#         self.p_supply_unicode = self.convert.ToString(int(value))
#
#     def setAddress(self,value):
#         self.p_address = self.convert.ToString(value.encode('utf-8'))
#
#     def setContact(self,value):
#         self.p_contact = self.convert.ToString(value.encode('utf-8'))
#
#     def setPhone(self,value):
#         self.p_phone = self.convert.ToString(int(value))
#
#     def setExt(self,value):
#         self.p_ext = self.convert.ToString(int(value))
#
#     def setMobile(self,value):
#         self.p_mobile = self.convert.ToString(value.encode('utf-8'))
#
#     def setEmail(self,value):
#         self.p_email = self.convert.ToString(value.encode('utf-8'))
#
#     def setContact1(self, value):
#         self.p_contact1 = self.convert.ToString(value.encode('utf-8'))
#
#     def setPhone1(self, value):
#         self.p_phone1 = self.convert.ToString(value.encode('utf-8'))
#
#     def setExt1(self, value):
#         self.p_ext1 = self.convert.ToString(value.encode('utf-8'))
#
#     def setMobile1(self, value):
#         self.p_mobile1 = self.convert.ToString(value.encode('utf-8'))
#
#     def setEmail1(self, value):
#         self.p_email1 = self.convert.ToString(value.encode('utf-8'))
#
#     def setMemo(self, value):
#         self.p_memo = self.convert.ToString(value.encode('utf-8'))
#
#     def setUser_id(self,value):
#         self.p_user_id = self.convert.ToString(value.encode('utf-8'))
#
#     def getGroup_id(self):
#         return self.p_group_id
#
#     def getUser_id(self):
#         return self.p_user_id
#
#     def getSupply_name(self):
#         return self.p_supply_name
#
#     def getSupply_unicode(self):
#         return self.p_supply_unicode
#
#     def getAddress(self):
#         return  self.p_address
#
#     def getContact(self):
#         return self.p_contact
#
#     def getPhone(self):
#         return self.p_phone
#
#     def getExt(self):
#         return self.p_ext
#
#     def getMobile(self):
#         return self.p_mobile
#
#     def getEmail(self):
#         return self.p_email
#
#     def getContact1(self):
#         return self.p_contact1
#
#     def getPhone1(self):
#         return self.p_phone1
#
#     def getExt1(self):
#         return self.p_ext1
#
#     def getMobile1(self):
#         return self.p_mobile1
#
#     def getEmail1(self):
#         return self.p_email1
#
#     def getMemo(self):
#         return self.p_memo
#
# class Product():
#     p_group_id, p_c_product_id, p_product_name, p_supply_name, p_type_name = None, None, None, None, None
#     p_unit_name, p_cost, p_price, p_description, p_barcode = None, None, None, None, None
#     p_keep_stock, p_begging_stock, p_user_id = None, None, None
#
#     def __init__(self):
#         self.convert = convertType()
#         self.aes = aes_data()
#
#     def __del__(self):
#         self.convert = None
#         self.aes = None
#
#     def setGroup_id(self,value):
#         self.p_group_id = self.convert.ToString(value.encode('utf-8'))
#
#     def setC_Product_id(self,value):
#         self.p_c_product_id = self.convert.ToString(value.encode('utf-8'))
#
#     def setP_Product_name(self,value):
#         self.p_product_name = self.convert.ToString(value.encode('utf-8'))
#
#     def setSupply_name(self,value):
#         self.p_supply_name = self.convert.ToString(value.encode('utf-8'))
#
#     def setType_name(self,value):
#         self.p_type_name = self.convert.ToString(value.encode('utf-8'))
#
#     def setUnit_name(self,value):
#         self.p_unit_name = self.convert.ToString(value.encode('utf-8'))
#
#     def setCost(self,value):
#         self.p_cost = self.convert.ToFloat(value)
#
#     def setPrice(self,value):
#         self.p_cost = self.convert.ToFloat(value)
#
#     def setDescription(self,value):
#         self.p_description = self.convert.ToString(value.encode('utf-8'))
#
#     def setBarcode(self,value):
#         self.p_barcode = self.convert.ToString(value.encode('utf-8'))
#
#     def setKeep_stock(self,value):
#         self.p_keep_stock = self.convert.ToInt(value)
#
#     def setBegging_stock(self, value):
#         self.p_begging_stock = self.convert.ToInt(value)
#
#     def setUser_id(self,value):
#         self.p_user_id = self.convert.ToString(value.encode('utf-8'))
#
#     def getGroup_id(self):
#         return self.p_group_id
#
#     def getC_Product_id(self):
#         return self.p_c_product_id
#
#     def getP_Product_name(self):
#         return self.p_product_name
#
#     def getSupply_name(self):
#         return self.p_supply_name
#
#     def getType_name(self):
#         return self.p_type_name
#
#     def getUnit_name(self):
#         return self.p_unit_name
#
#     def getCost(self):
#         return self.p_cost
#
#     def getPrice(self):
#         return self.p_price
#
#     def getDescription(self):
#         return self.p_description
#
#     def getBarcode(self):
#         return self.p_barcode
#
#     def getKeep_stock(self):
#         return self.p_keep_stock
#
#     def getBegging_stock(self):
#         return self.p_begging_stock
#
#     def getUser_id(self):
#         return self.p_user_id
#
# class Package():
#
#     def __init__(self):
#         self.convert = convertType()
#         self.aes = aes_data()
#     def __del__(self):
#         self.convert = None
#         self.aes = None
#
# class Contrast():
#
#     def __init__(self):
#         self.convert = convertType()
#         self.aes = aes_data()
#     def __del__(self):
#         self.convert = None
#         self.aes = None