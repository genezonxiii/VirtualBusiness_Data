# -*-  coding: utf-8  -*-
__author__ = '10409003'
import web
import json
from UploadData import VirtualBusiness
from AesData import Insert_Client
# from SelectCustomer import Query_customer
from Shipper import ShipperData, VBsale_Analytics
from GroupBuy import FileProcess
from ImportFundationData import DataInsert
from SFexpress import FileProcess_sf
from SFexpress.sfexpress_api import sendDataSFexpress
import datetime
from time import mktime
import logging
import time
from xml.dom import minidom

# settings.configure()
urls = ("/upload/(.*)", "Uploaddata", \
        "/aes/(.*)", "Encrypt", \
        "/query/(.*)", "GetClientData", \
        "/ship/(.*)", "Shipper", \
        "/analytics/(.*)", "Analytics", \
        "/groupbuy/(.*)", "GroupBuy",\
        "/import/(.*)", "Import", \
        "/sfexpress/(.*)", "SFExpress", \
        "/sfexpressapi/(.*)", "SendSFExpress", \
        "/PurchaseOrderInboundPush/(.*)", "SFInboundDetail", \
        "/SaleOrderOutboundDetailPush/(.*)", "SFOutboundDetail", \
        "/SaleOrderStatusPush/(.*)", "SFStatus", \
        "/InventoryChange/(.*)", "SFChange", \
        "/InventoryBalance/(.*)", "SFBalance")

app = web.application(urls, globals())

logger = logging.getLogger('vb')

def vb_logger(name):
    logger = logging.getLogger(name)
    logger.setLevel(logging.DEBUG)

    # create file handler
    fh = logging.FileHandler('/data/VirtualBusiness_Data/sf_warehouse.log')
    fh.setLevel(logging.DEBUG)

    # create console handler
    ch = logging.StreamHandler()
    ch.setLevel(logging.DEBUG)

    # create formatter and add it to the handlers
    formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(name)s/%(funcName)s/%(lineno)d - %(message)s')

    # formatter.datefmt = '%Y/%m/%d %I:%M:%S %p'
    fh.setFormatter(formatter)
    ch.setFormatter(formatter)

    # add the handlers to the logger
    if not logger.handlers:
        logger.addHandler(fh)
        logger.addHandler(ch)

    return logger

class Uploaddata():
    def GET(self, name):
        data = name.split('&')
        for i in range(len(data)):
            data[i] = data[i][5:len(data[i])].decode('base64')
            print len(data[1])

        logging.basicConfig(filename='/data/VirtualBusiness_Data/pyupload.log',
                            level=logging.DEBUG,
                            format='%(asctime)s - %(levelname)s - %(filename)s:%(name)s:%(module)s/%(funcName)s/%(lineno)d - %(message)s',
                            datefmt='%Y/%m/%d %I:%M:%S %p')
        logging.Formatter.converter = time.gmtime

        logger.debug('===Uploaddata===')
        logger.debug('data[0]:' + data[0])
        logger.debug('data[1]:' + data[1])

        Upload = VirtualBusiness()
        return Upload.virtualbusiness(data[0], data[1])
        # return 'Success'

class Encrypt():
    def GET(self, name):
        data = name.split('&')
        for i in range(len(data)):
            data[i] = data[i][5:len(data[i])].decode('base64')
        customerData = Insert_Client()
        web.header('Content-Type', 'text/json; charset=utf-8', unique=True)
        result = json.dumps(customerData.AESen_CustomerData(data[0], data[1], data[2], data[3]))
        return result

class GetClientData():
    def GET(self, name):
        data = name.split('&')
        for i in range(len(data)):
            data[i] = data[i][6:len(data[i])].decode('base64')
        clientData = Query_customer()
        web.header('Content-Type', 'text/json; charset=utf-8', unique=True)
        result = json.dumps(clientData.GetDataContent(data[0]))
        return result

class MyEncoder(json.JSONEncoder):
    def default(self, obj):
        if isinstance(obj, datetime.datetime):
            return int(mktime(obj.timetuple()))

        return json.JSONEncoder.default(self, obj)

class GroupBuy():
    def GET(self, name):
        data = name.split('&')
        for i in range(len(data)):
            data[i] = data[i][5:len(data[i])].decode('base64')

        logging.basicConfig(filename='/data/VirtualBusiness_Data/pyupload.log',
                            level=logging.DEBUG,
                            format='%(asctime)s - %(levelname)s - %(filename)s:%(name)s:%(module)s/%(funcName)s/%(lineno)d - %(message)s',
                            datefmt='%Y/%m/%d %I:%M:%S %p')
        logging.Formatter.converter = time.gmtime

        logger.debug('===Uploaddata===')
        logger.debug('data[0]:' + data[0])
        logger.debug('data[1]:' + data[1])
        logger.debug('data[2]:' + data[2])
        logger.debug('data[3]:' + data[3])

        Upload = FileProcess()
        return Upload.transferFile(data[0], data[1],int(data[2]),data[3])

class Shipper():
    def GET(self, name):
        data = name.split('&')
        for i in range(len(data)):
            data[i] = data[i][8:len(data[i])].decode('base64')
        shipData = ShipperData()
        web.header('Content-Type', 'text/json; charset=utf-8', unique=True)
        result = json.dumps(shipData.GetDataContent(data[0], data[1], data[2]), cls=DatetimeEncoder)
        # result= json.dumps(shipData.GetDataContent(data[0],data[1],data[2]),cls=DjangoJSONEncoder)

        return result

class DatetimeEncoder(json.JSONEncoder):
    def default(self, obj):
        if isinstance(obj, datetime.date):
            return obj.strftime('%Y-%m-%dT%H:%M:%SZ')

        # Let the base class default method raise the TypeError
        return json.JSONEncoder.default(self, obj)


class Analytics():
    def GET(self, name):
        data = name.split('&')
        for i in range(len(data)):
            data[i] = data[i][5:len(data[i])].decode('base64')
            print data[i]
        result = []
        sp = VBsale_Analytics()
        if data[0] == 'TOP10':
            result = sp.get_buyer_top10(data[1], data[2], data[3])
        else:
            result = sp.get_buyer_channel(data[1], data[2], data[3], data[4])

        web.header('Content-Type', 'text/json; charset=utf-8', unique=True)
        resule = json.dumps(result)
        return resule

class Import():
    def GET(self, name):
        data = name.split('&')
        for i in range(len(data)):
            data[i] = data[i][5:len(data[i])].decode('base64')
            print len(data[1])

        logging.basicConfig(filename='/data/VirtualBusiness_Data/pyupload.log',
                            level=logging.DEBUG,
                            format='%(asctime)s - %(levelname)s - %(filename)s:%(name)s:%(module)s/%(funcName)s/%(lineno)d - %(message)s',
                            datefmt='%Y/%m/%d %I:%M:%S %p')
        logging.Formatter.converter = time.gmtime

        logger.debug('===Importdata===')
        logger.debug('data[0]:' + data[0])
        logger.debug('data[1]:' + data[1])

        Upload = DataInsert()
        return Upload.virtualbusiness_import(data[0], data[1])

class SFExpress():
    def GET(self, name):
        data = name.split('&')
        for i in range(len(data)):
            data[i] = data[i][5:len(data[i])].decode('base64')

        logging.basicConfig(filename='/data/VirtualBusiness_Data/pyupload.log',
                            level=logging.DEBUG,
                            format='%(asctime)s - %(levelname)s - %(filename)s:%(name)s:%(module)s/%(funcName)s/%(lineno)d - %(message)s',
                            datefmt='%Y/%m/%d %I:%M:%S %p')
        logging.Formatter.converter = time.gmtime

        logger.debug('===Uploaddata===')
        logger.debug('data[0]:' + data[0])
        logger.debug('data[1]:' + data[1])
        logger.debug('data[2]:' + data[2])
        logger.debug('data[3]:' + data[3])

        Upload = FileProcess_sf()
        return Upload.transferFile(data[0], data[1],int(data[2]),data[3])

class SendSFExpress():
    def GET(self,name):
        data = name.split('&')
        for i in range(len(data)):
            data[i] = data[i][5:len(data[i])].decode('base64')
        logging.basicConfig(filename='/data/VirtualBusiness_Data/pyupload.log',
                            level=logging.DEBUG,
                            format='%(asctime)s - %(levelname)s - %(filename)s:%(name)s:%(module)s/%(funcName)s/%(lineno)d - %(message)s',
                            datefmt='%Y/%m/%d %I:%M:%S %p')
        logging.Formatter.converter = time.gmtime

        logger.debug('===Senddata To SFExpress===')
        logger.debug('data[0]:' + data[0])
        SF = sendDataSFexpress()
        return SF.sendToSFexpress(data[0])
        # http: // localhost:8080 / sfexpressapi / data = PD94bWwgdmVyc2lvbj0iMS4wIiBlbmNvZGluZz0iVVRGLTgiIHN0YW5kYWxvbmU9InllcyI / Pg0KPFJlcXVlc3Qgc2VydmljZT0iSVRFTV9RVUVSWV9TRVJWSUNFIiBsYW5nPSJ6aC1UVyI + DQogICAgPEhlYWQ + DQogICAgICAgIDxBY2Nlc3NDb2RlPklUQ05DMWh0WFY5eHVPS3JodTI0b3c9PTwvQWNjZXNzQ29kZT4NCiAgICAgICAgPENoZWNrd29yZD5BTlUyVkh2VjVlcXNyMlBKSHUyem5XbVd0ejJDZEl2ajwvQ2hlY2t3b3JkPg0KICAgIDwvSGVhZD4NCiAgICA8Qm9keT4NCiAgICAgICAgPEl0ZW1RdWVyeVJlcXVlc3Q + DQogICAgICAgICAgICA8Q29tcGFueUNvZGU + V1lER0o8L0NvbXBhbnlDb2RlPg0KICAgICAgICAgICAgPFNrdU5vTGlzdD4NCiAgICAgICAgICAgICAgICA8U2t1Tm8 + UFkzMDAxQVNGPC9Ta3VObz4NCiAgICAgICAgICAgIDwvU2t1Tm9MaXN0Pg0KICAgICAgICA8L0l0ZW1RdWVyeVJlcXVlc3Q + DQogICAgPC9Cb2R5Pg0KPC9SZXF1ZXN0Pg ==


# 入庫單明細推送接口
class SFInboundDetail():
    def __init__(self):
        logger = vb_logger("vb")

    def POST(self,data):

        # dict_items
        input_data = web.input()

        logger = logging.getLogger("vb.SFInboundDetail")

        # 取出 dict
        for temp_param in input_data.items():
            print temp_param

            if temp_param[0] == "logistics_interface":

                logger.debug(temp_param[0])
                logger.debug(temp_param[1])
                key = temp_param[0]
                value = temp_param[1]
                print key
                print value
            elif temp_param[0] == "data_digest":

                logger.debug(temp_param[0])
                logger.debug(temp_param[1])

                key = temp_param[0]
                value = temp_param[1]

                print key
                print value
        return '<Response service="PURCHASE_ORDER_INBOUND_SERVICE" lang="zh-CN"><Head>OK</Head><Body><PurchaseOrderInboundResponse><Result>Success</Result><Note></Note></PurchaseOrderInboundResponse></Body></Response>'

# 出庫單明細推送接口
class SFOutboundDetail():
    def __init__(self):
        logger = vb_logger("vb")

    def POST(self,data):
        # dict_items
        input_data = web.input()
        logger = logging.getLogger("vb.SFOutboundDetail")

        for key, value in input_data.items():
            if key == 'logistics_interface':
                logger.debug(key)
                logger.debug(value)
                print key
                print value
            elif key == 'data_digest':
                logger.debug(key)
                logger.debug(value)
                print key
                print
        return '<Response service="SALE_ORDER_OUTBOUND_DETAIL_SERVICE" lang="zh-CN"><Head>OK</Head><Body><SaleOrderOutboundDetailResponse><Result>Success</Result><Note></Note></SaleOrderOutboundDetailResponse></Body></Response>'

# 出庫單狀態推送入口
class SFStatus():
    def __init__(self):
        logger = vb_logger("vb")

    def POST(self,data):
        # dict_items
        input_data = web.input()
        logger = logging.getLogger("vb.SFStatus")

        for key, value in input_data.items():
            if key == 'logistics_interface':
                logger.debug(key)
                logger.debug(value)
                print key
                print value
            elif key == 'data_digest':
                logger.debug(key)
                logger.debug(value)
                print key
                print value
        return '<Response service="SALE_ORDER_STATUS_PUSH_SERVICE" lang="zh-CN"><Head>OK</Head><Body><SaleOrderStatusResponse><Result>Success</Result><Note></Note></SaleOrderStatusResponse></Body></Response>'

# 庫存變化推送接口
class SFChange():
    def __init__(self):
        logger = vb_logger("vb")

    def POST(self,data):
        # dict_items
        input_data = web.input()
        logger = logging.getLogger("vb.SFChange")

        # dict
        for key, value in input_data.items():
            if key == 'logistics_interface':
                logger.debug(key)
                logger.debug(value)
                print key
                print value
            elif key == 'data_digest':
                logger.debug(key)
                logger.debug(value)
                print key
                print value
        return '<Response service="INVENTORY_CHANGE_SERVICE" lang="zh-CN"><Head>OK</Head><Body><InventoryChangeResponse><Result>Success</Result><Note></Note></InventoryChangeResponse></Body></Response>'

# 庫存對帳推送接口
class SFBalance():
    def __init__(self):
        logger = vb_logger("vb")

    def POST(self,data):
        # dict_items
        input_data = web.input()
        logger = logging.getLogger("vb.SFBalance")

        for key, value in input_data.items():
            if key == 'logistics_interface':
                logger.debug(key)
                logger.debug(value)
                print key
                print value
            elif key == 'data_digest':
                logger.debug(key)
                logger.debug(value)
                print key
                print value
        return '<Response service="INVENTORY_BALANCE_SERVICE" lang="zh-CN"><Head>OK</Head><Body><InventoryBalanceResponse><Result>Success</Result><Note></Note></InventoryBalanceResponse></Body></Response>'
                
if __name__ == "__main__":
    app.run()
