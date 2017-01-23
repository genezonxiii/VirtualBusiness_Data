# -*-  coding: utf-8  -*-
__author__ = '10409003'

import os
import csv
import xlrd
import logging
import time
from VirtualBusiness.test import Test_Data
# from VirtualBusiness.ASAP import ASAP_Data
# from VirtualBusiness.ibon import IBON_Data
# from VirtualBusiness.Life import Life_Data
# from VirtualBusiness.Line_mart import LineMart_Data
from VirtualBusiness.Momo25 import Momo25_Data
from VirtualBusiness.momo24_csv import Momo24csv_Data
from VirtualBusiness.Myfone22_table import Myfone22table_Data
# from VirtualBusiness.payeasy import payeasy_Data
# from VirtualBusiness.Pchome1 import PCHome_Data
from VirtualBusiness.UDN30 import Udn30_Data
from VirtualBusiness.UDN30_csv import UDN30csv_Data
# from VirtualBusiness.Nine import Nine_Data
# from VirtualBusiness.Tree import TREE_Data
# from VirtualBusiness.MaJi import MAJI_Data
from VirtualBusiness.etmall29_csv import Etmall29csv_Data
# from VirtualBusiness.Book import Book_Data
from VirtualBusiness.Umall30_csv import Umall30csv_Data
# from VirtualBusiness.LoveBuy import LoveBuy_Data
# from VirtualBusiness.Lotte import Lotte_Data
from VirtualBusiness.Yahoomall24_table import YahooS24table_Data
from VirtualBusiness.Savesafe22_table import Savesafe22table_Data
from VirtualBusiness.Babyhome17 import Babyhome17_Data
from VirtualBusiness.Friday16 import Friday16_Data
from VirtualBusiness.Momomall21 import Momomall21_Data
from VirtualBusiness.ihergo22 import Ihergo22_Data
from VirtualBusiness.Gohappy22_csv import Gohappy22csv_Data
from VirtualBusiness.Myfone19_csv import Myfone19csv_Data

logger = logging.getLogger(__name__)

class VirtualBusiness():
    Data=None

    def __init__(self):
        pass
    def virtualbusiness(self,DataPath,userID):
        Supplier = str(DataPath).split('/')[3]  #group id
        Firm = str(DataPath).split('/')[5]      #平台
        supplierType = str(DataPath).split('/')[4]  #訂單種類
        OutputFile = os.path.basename(DataPath).split(".")[0]  #輸出檔案名稱
        outPutPath = ExcelTemplate()
        OutputFile = outPutPath.T_Cat_OutputFilePath + Supplier + '/' + Firm +'/' + OutputFile +".xls" #輸出檔案路徑及名稱

        print Firm
        print Supplier
        print supplierType

        logging.basicConfig(filename='/data/VirtualBusiness_Data/pyupload.log', 
			level=logging.DEBUG, 
			format='%(asctime)s - %(levelname)s - %(filename)s:%(name)s:%(module)s/%(funcName)s/%(lineno)d - %(message)s',
            datefmt='%Y/%m/%d %I:%M:%S %p')
        logging.Formatter.converter = time.gmtime

        logger.info('Supplier')
        logger.info(Supplier)


        if Supplier == 'test':
            FinalData=Test_Data()
            return FinalData.Test_Data(Supplier,Firm,os.path.join(DataPath),userID)

        # elif Supplier == 'asap':
        #     logger.debug('asap')
        #     if supplierType == 'home-delivery':
        #         logger.debug( 'home-delivery')
        #         FinalData = ASAP_Data()
        #         return FinalData.ASAP_Data('ASAP', Firm, os.path.join(DataPath),userID, OutputFile)
        #     if supplierType == 'instore-pickup':
        #         logger.debug('instore-pickup')
        #         print 'instore-pickup'
        #         FinalData = ASAP_Data()
        #         return FinalData.ASAP_Data('ASAP', Firm, os.path.join(DataPath),userID, OutputFile)

        elif Supplier =='gohappy':
            logger.debug('gohappy')
            if supplierType == 'home-delivery':
                logger.debug('home-delivery')
                FinalData = Gohappy22csv_Data()
                return FinalData.Gohappy_22_Data('Gohappy', Firm, os.path.join(DataPath),userID, OutputFile)

        # elif Supplier =='linemart':
        #     logger.debug('Line Mart')
        #     if supplierType == 'home-delivery':
        #         logger.debug('home-delivery')
        #         FinalData = LineMart_Data()
        #         return FinalData.LineMart_Data('Line_Mart', Firm, os.path.join(DataPath),userID,OutputFile)
        #     if supplierType == 'instore-pickup':
        #         logger.debug('instore-pickup')
        #         print 'instore-pickup'
        #         FinalData = LineMart_Data()
        #         return FinalData.LineMart_Data('Line_Mart', Firm, os.path.join(DataPath),userID,OutputFile)

        elif Supplier =='momo':
            logger.debug("momo")
            if supplierType == 'home-delivery':
                logger.debug('home-delivery')
                if DataPath.endswith(".xls"):
                    FinalData = Momo25_Data()
                    return FinalData.Momo_25_Data('momo', Firm, os.path.join(DataPath),userID,OutputFile)
                if DataPath.endswith(".csv"):
                    momo = Momo24csv_Data()
                    return  momo.Momo_24_Data('momo', Firm,os.path.join(DataPath),userID,OutputFile)

        elif Supplier =='myfone':
            logger.debug('myfone')
            if supplierType == 'home-delivery':
                logger.debug('home-delivery')
                # DataPath = DataPath.split('.')[0] + '.html'
                if DataPath.endswith(".html"):
                    FinalData = Myfone22table_Data()
                    return FinalData.Myfone_22_Data('myfone', Firm, os.path.join(DataPath), userID,OutputFile)
                if DataPath.endswith(".csv"):
                    FinalData = Myfone19csv_Data()
                    return  FinalData.Myfone_19_Data('myfone', Firm, os.path.join(DataPath), userID,OutputFile)

        # elif Supplier == 'payeasy':
        #     logger.debug('payeasy')
        #     if supplierType == 'home-delivery':
        #         logger.debug( 'home-delivery')
        #         FinalData = payeasy_Data()
        #         return FinalData.payeasy_Data('payeasy', Firm, os.path.join(DataPath),userID,OutputFile)
        #     if supplierType == 'instore-pickup':
        #         logger.debug('instore-pickup')
        #         print 'instore-pickup'
        #         FinalData = payeasy_Data()
        #         return FinalData.payeasy_Data('payeasy', Firm, os.path.join(DataPath),userID,OutputFile)
        #
        # elif Supplier == 'pchome':
        #     logger.debug('pchome')
        #     if supplierType == 'home-delivery':
        #         logger.debug( 'home-delivery')
        #         FinalData = PCHome_Data()
        #         return FinalData.PCHome_Data('Pchome', Firm, os.path.join(DataPath),userID,OutputFile)
        #     if supplierType == 'instore-pickup':
        #         logger.debug('instore-pickup')
        #         print 'instore-pickup'
        #         FinalData = PCHome_Data()
        #         return FinalData.PCHome_Data('Pchome', Firm, os.path.join(DataPath),userID,OutputFile)

        elif Supplier == 'udn':
            logger.debug('udn')
            if supplierType == 'home-delivery':
                logger.debug('home-delivery')
                if DataPath.endswith(".xls"):
                    FinalData = Udn30_Data()
                    return FinalData.Udn_30_Data('udn', Firm, os.path.join(DataPath), userID, OutputFile)
                if DataPath.endswith(".csv"):
                    FinalData = UDN30csv_Data()
                    return FinalData.UDN_30_Data('udn', Firm, os.path.join(DataPath), userID, OutputFile)

        # elif Supplier == 'yahoo':
        #     logger.debug('yahoo')
        #     if supplierType == 'home-delivery':
        #         logger.debug( 'home-delivery')
        #         FinalData = Yahoo_Data()
        #         return FinalData.Yahood_Data('yahoo', Firm, os.path.join(DataPath),userID, OutputFile)
        #     if supplierType == 'instore-pickup':
        #         logger.debug('instore-pickup')
        #         print 'instore-pickup'
        #         FinalData = Yahoo_Data()
        #         return FinalData.Yahoos_Data('yahoo', Firm, os.path.join(DataPath),userID, OutputFile)

        # elif Supplier == '91mai':
        #     logger.debug('91mai')
        #     if supplierType == 'home-delivery':
        #         logger.debug( 'home-delivery')
        #         FinalData = Nine_Data()
        #         return FinalData.Nine_Data(u'九易', Firm, os.path.join(DataPath),userID, OutputFile)
        #     if supplierType == 'instore-pickup':
        #         logger.debug('instore-pickup')
        #         print 'instore-pickup'
        #         FinalData = Nine_Data()
        #         return FinalData.Nine_Data(u'九易'.encode('utf-8'), Firm, os.path.join(DataPath),userID, OutputFile)
        #
        # elif Supplier == 'treemall':
        #     logger.debug('treemall')
        #     if supplierType == 'home-delivery':
        #         logger.debug( 'home-delivery')
        #         FinalData = TREE_Data()
        #         return FinalData.TREE_Data(u'國泰Tree'.encode('utf-8'), Firm, os.path.join(DataPath),userID, OutputFile)
        #     if supplierType == 'instore-pickup':
        #         logger.debug('instore-pickup')
        #         print 'instore-pickup'
        #         FinalData = TREE_Data()
        #         return FinalData.TREE_Data(u'國泰Tree'.encode('utf-8'), Firm, os.path.join(DataPath),userID, OutputFile)

        elif Supplier == 'etmall':
            logger.debug('etmall')
            if supplierType == 'home-delivery':
                logger.debug( 'home-delivery')
                # if DataPath.endswith(".xls"):
                #     FinalData = GM16_Data()
                #     return FinalData.GM_16_Data(u'東森購物'.encode("utf-8"), Firm, os.path.join(DataPath), userID)
                if DataPath.endswith(".csv"):
                    FinalData = Umall30csv_Data()
                    return  FinalData.Umall_30_Data(u'東森購物'.encode("utf-8"), Firm, os.path.join(DataPath), userID, OutputFile)

        # elif Supplier == 'books':
        #     logger.debug('books')
        #     if supplierType == 'home-delivery':
        #         logger.debug( 'home-delivery')
        #         FinalData = Book_Data()
        #         return FinalData.Book_Data(u'博客來'.encode('utf-8'), Firm, os.path.join(DataPath),userID, OutputFile)
        #     if supplierType == 'instore-pickup':
        #         logger.debug('instore-pickup')
        #         print 'instore-pickup'
        #         FinalData = Book_Data()
        #         return FinalData.Book_Data(u'博客來'.encode('utf-8'), Firm, os.path.join(DataPath),userID, OutputFile)

        elif Supplier == 'umall':
            logger.debug('umall')
            if supplierType == 'home-delivery':
                logger.debug( 'home-delivery')
                FinalData = Etmall29csv_Data()
                return FinalData.Etmall_29_Data(u'森森購物'.encode('utf-8'), Firm, os.path.join(DataPath),userID, OutputFile)

        elif Supplier == 'yahoomall':
            logger.debug('yahoomall')
            # if supplierType == 'home-delivery':
            #     logger.debug( 'home-delivery')
            #     FinalData = YahooS29_Data()
            #     return FinalData.YahooS_29_Data(u'超級商城'.encode('utf-8'), Firm, os.path.join(DataPath),userID)
            if supplierType == 'instore-pickup':
                logger.debug('instore-pickup')
                print 'instore-pickup'
                FinalData = YahooS24table_Data()
                return FinalData.YahooS_24_Data(u'超級商城'.encode('utf-8'), Firm, os.path.join(DataPath),userID, OutputFile)

        # elif Supplier == 'amart':
        #     logger.debug('amart')
        #     if supplierType == 'home-delivery':
        #         logger.debug( 'home-delivery')
        #         FinalData = LoveBuy_Data()
        #         return FinalData.LoveBuy_Data(u'愛買'.encode('utf-8'), Firm, os.path.join(DataPath),userID, OutputFile)
        #     if supplierType == 'instore-pickup':
        #         logger.debug('instore-pickup')
        #         print 'instore-pickup'
        #         FinalData = LoveBuy_Data()
        #         return FinalData.LoveBuy_Data(u'愛買'.encode('utf-8'), Firm, os.path.join(DataPath),userID, OutputFile)
        #
        # elif Supplier == 'rakuten':
        #     logger.debug('rakuten')
        #     if supplierType == 'home-delivery':
        #         logger.debug( 'home-delivery')
        #         FinalData = Lotte_Data()
        #         return FinalData.Lotte_Data(u'樂天'.encode('utf-8'), Firm, os.path.join(DataPath),userID, OutputFile)
        #     if supplierType == 'instore-pickup':
        #         logger.debug('instore-pickup')
        #         print 'instore-pickup'
        #         FinalData = Lotte_Data()
        #         return FinalData.Lotte_Data(u'樂天'.encode('utf-8'), Firm, os.path.join(DataPath),userID, OutputFile)

        elif Supplier == 'bigbuy':
            logger.debug('bigbuy')
            if supplierType == 'home-delivery':
                logger.debug('home-delivery')
                FinalData = Savesafe22table_Data()
                return FinalData.Savesafe_22_Data(u'大買家', Firm, os.path.join(DataPath),userID, OutputFile)
            # if supplierType == 'instore-pickup':
            #     logger.debug('instore-pickup')
            #     print 'instore-pickup'
            #     FinalData = Savesafe22table_Data()
            #     return FinalData.Savesafe_22_Data(u'大買家', Firm, os.path.join(DataPath),userID)

        elif Supplier == 'babyhome':
            logger.debug('babyhome')
            if supplierType == 'home-delivery':
                logger.debug( 'home-delivery')
                FinalData = Babyhome17_Data()
                return FinalData.Babyhome_17_Data('babyhome', Firm, os.path.join(DataPath), userID, OutputFile)
            # if supplierType == 'instore-pickup':
            #     logger.debug('instore-pickup')
            #     print 'instore-pickup'
            #     FinalData = Babyhome17_Data()
            #     return FinalData.Babyhome_17_Date('babyhome', Firm, os.path.join(DataPath),userID)

        elif Supplier == 'friday':
            logger.debug('friday')
            if supplierType == 'home-delivery':
                logger.debug( 'home-delivery')
                FinalData = Friday16_Data()
                return FinalData.Friday_16_Data('friday', Firm, os.path.join(DataPath), userID, OutputFile)
            # if supplierType == 'instore-pickup':
            #     logger.debug('instore-pickup')
            #     print 'instore-pickup'
            #     FinalData = Friday17_Data()
            #     return FinalData.Friday_17_Data('friday', Firm, os.path.join(DataPath),userID)

        elif Supplier == 'momomall':
            logger.debug('momomall')
            if supplierType == 'instore-pickup':
                logger.debug('instore-pickup')
                FinalData = Momomall21_Data()
                return FinalData.Momomall_21_Data(u'摩天商城', Firm, os.path.join(DataPath), userID, OutputFile)

        elif Supplier == 'ihergo':
            logger.debug('ihergo')
            if supplierType == 'home-delivery':
                logger.debug( 'home-delivery')
                FinalData = Ihergo22_Data()
                return FinalData.Ihergo_22_Data('Ihergo', Firm, os.path.join(DataPath), userID, OutputFile)
            # if supplierType == 'instore-pickup':
            #     logger.debug('instore-pickup')
            #     print 'instore-pickup'
            #     FinalData = Friday17_Data()
            #     return FinalData.Friday_17_Data('friday', Firm, os.path.join(DataPath),userID)

if __name__ == '__main__':
    Business = VirtualBusiness()
    # Business.virtualbusiness("/data/vbupload/test/", "19647356")
    Business.virtualbusiness("D:/vbdata/vbupload/momo/cbcc3138-5603-11e6-a532-000d3a800878/A1102_3_1_010031_20160301110142.xls", "19647356")
#     Business=VirtualBusiness()
#     Business.virtualbusiness("C:/vbdata/vbupload/17life/7dcc2045-472e-11e6-806e-000c29c1d000/Data2.xls")
