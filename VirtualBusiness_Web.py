# -*-  coding: utf-8  -*-
__author__ = '10409003'
import web
import json
from UploadData import VirtualBusiness
from AesData import Insert_Client,DEncrypt
from SelectCustomer import Query_customer
from Shipper import ShipperData,VBsale_Analytics
import datetime
from time import mktime
import logging
import time



#settings.configure()
urls = ("/upload/(.*)", "Uploaddata",\
        "/aes/(.*)", "Encrypt",\
        "/query/(.*)", "GetClientData",\
        "/ship/(.*)","Shipper",\
        "/analytics/(.*)", "Analytics",\
        "/dataEncrypt/(.*)", "DEcryptData")

app =web.application(urls, globals())
logger = logging.getLogger('VirtualBusiness_Web')

class Uploaddata():
    def GET(self, name):
        data=name.split('&')
        for i in range(len(data)):
            data[i]=data[i][5:len(data[i])].decode('base64')
            print len(data[1])
            
        logging.basicConfig(filename='/data/VirtualBusiness_Data/pyupload.log', level=logging.DEBUG, format='%(asctime)s - %(name)s - %(levelname)s - %(message)s', datefmt='%Y/%m/%d %I:%M:%S %p')
        logging.Formatter.converter = time.gmtime
        logging.info('===Uploaddata GO!!!!===')
        logging.info('data[0]:' + data[0])
        logging.info('data[1]:' + data[1])
        logging.info('===Uploaddata GO END!!!!===')

        Upload = VirtualBusiness()
        return Upload.virtualbusiness(data[0],data[1])
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

class DEcryptData():
    def GET(self,name):
        data = name.split('&')
        for i in range(len(data)):
            data[i] = data[i][5:len(data[i])].decode('base64')

        resultData = DEncrypt()
        result = None
        if data[0]=='1':
            result = json.dumps(resultData.encrypt(data[1]))
        else:
            result = json.dumps(resultData.decrypt(data[1]))
        return result

class GetClientData():
    def GET(self, name):
        data=name.split('&')
        for i in range(len(data)):
            data[i]=data[i][6:len(data[i])].decode('base64')
        clientData=Query_customer()
        web.header('Content-Type','text/json; charset=utf-8',unique=True)
        result= json.dumps(clientData.GetDataContent(data[0]))
        return result

class MyEncoder(json.JSONEncoder):

    def default(self, obj):
        if isinstance(obj, datetime.datetime):
            return int(mktime(obj.timetuple()))

        return json.JSONEncoder.default(self, obj)

class Shipper():

    def GET(self, name):
        data=name.split('&')
        for i in range(len(data)):
            data[i]=data[i][8:len(data[i])].decode('base64')
        shipData=ShipperData()
        web.header('Content-Type','text/json; charset=utf-8',unique=True)
        result= json.dumps(shipData.GetDataContent(data[0],data[1],data[2]),cls=DatetimeEncoder)
        #result= json.dumps(shipData.GetDataContent(data[0],data[1],data[2]),cls=DjangoJSONEncoder)

        return result

class DatetimeEncoder(json.JSONEncoder):
    def default(self, obj):
        if isinstance(obj, datetime.date):
            return obj.strftime('%Y-%m-%dT%H:%M:%SZ')

        # Let the base class default method raise the TypeError
        return json.JSONEncoder.default(self, obj)

class Analytics():
    def GET(self, name):
        data=name.split('&')
        for i in range(len(data)):
            data[i]=data[i][5:len(data[i])].decode('base64')
            print data[i]
        result=[]
        sp = VBsale_Analytics()
        if data[0]=='TOP10':
            result=sp.get_buyer_top10(data[1],data[2],data[3])
        else:
            result=sp.get_buyer_channel(data[1],data[2],data[3],data[4])
        

        web.header('Content-Type', 'text/json; charset=utf-8', unique=True)
        resule = json.dumps(result)
        return resule
        
if __name__ == "__main__":
    app.run()
