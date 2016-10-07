# -*-  coding: utf-8  -*-
__author__ = '10409003'
import os
import csv
import xlrd
from VirtualBusiness.test import Test_Data
from VirtualBusiness.ASAP import ASAP_Data
from VirtualBusiness.Go_Happy import GoHappy_Data
from VirtualBusiness.ibon import IBON_Data
from VirtualBusiness.ibonc import IBON_DataC
from VirtualBusiness.Life import Life_Data
from VirtualBusiness.Line_mart import LineMart_Data
from VirtualBusiness.Momo1 import Momo_Data
from VirtualBusiness.Momo2 import Momo_Data2
from VirtualBusiness.Momo3 import Momo3_Data
from VirtualBusiness.Myfone import Myfone_Data
from VirtualBusiness.payeasy import payeasy_Data
from VirtualBusiness.payeasy2 import payeasy_Data2
from VirtualBusiness.Pchome1 import PCHome_Data
from VirtualBusiness.Pchome2 import PCHome2_Data
from VirtualBusiness.Pchome3 import PCHome3_Data
from VirtualBusiness.UDN import UDN_Data
from VirtualBusiness.UDN_2 import UDN_Data2
from VirtualBusiness.UDN_3 import UDN_Data3
from VirtualBusiness.Yahoo1 import Yahoo1_Data
from VirtualBusiness.Yahoo2 import Yahoo2_Data
from VirtualBusiness.Yahoo3 import Yahoo3_Data
from VirtualBusiness.Nine import Nine_Data
from VirtualBusiness.Tree import TREE_Data
from VirtualBusiness.MaJi import MAJI_Data
from VirtualBusiness.gm import GM_Data
from VirtualBusiness.Book import Book_Data
from VirtualBusiness.Umall import UMall_Data
from VirtualBusiness.YahooS1 import Yahoo_DataS1
from VirtualBusiness.YahooS2 import Yahoo_DataS2
from VirtualBusiness.YahooS3 import Yahoo_DataS3
from VirtualBusiness.LoveBuy import LoveBuy_Data
from VirtualBusiness.Lotte import Lotte_Data
from VirtualBusiness.Yahoo_d import Yahood_Data
from VirtualBusiness.Yahoo_s import Yahoos_Data
from VirtualBusiness.momo_p import Momo_Datap

class VirtualBusiness():
    Data=None

    def __init__(self):
        pass
    def virtualbusiness(self,DataPath,userID):
        Supplier = str(DataPath).split('/')[3]
        Firm = str(DataPath).split('/')[4]
        print Firm
        print Supplier

        if Supplier == 'test':
            FinalData=Test_Data()
            return FinalData.Test_Data(Supplier,Firm,os.path.join(DataPath),userID)
        elif Supplier == 'asap':
            FinalData = ASAP_Data()
            return FinalData.ASAP_Data('ASAP', Firm, os.path.join(DataPath), userID)
        elif Supplier =='gohappy':
            FinalData = GoHappy_Data()
            return FinalData.GoHappy_Data('GoHappy', Firm, os.path.join(DataPath),userID)
        elif Supplier =='ibon':
            FinalData = IBON_Data()
            return FinalData.IBON_Data('ibon', Firm, os.path.join(DataPath),userID)
        elif Supplier =='ibonc':
            FinalData = IBON_DataC()
            return FinalData.IBON_DataC('ibon', Firm, os.path.join(DataPath),userID)
        elif Supplier =='17life':
            FinalData = Life_Data()
            return FinalData.Life_Data('17Life', Firm, os.path.join(DataPath),userID)
        elif Supplier =='Line_Mart':
            FinalData = LineMart_Data()
            return FinalData.LineMart_Data('Line_Mart', Firm, os.path.join(DataPath),userID)
        elif Supplier =='momo':
            if DataPath.split('.')[-1] != 'pdf':
                data = xlrd.open_workbook(os.path.join(DataPath))
                table = data.sheets()[0]
                num_cols = table.ncols
                if num_cols == 25:
                    FinalData = Momo_Data()
                    return FinalData.Momo_Data('momo', Firm, os.path.join(DataPath), userID)
                elif num_cols == 28:
                    FinalData = Momo_Data2()
                    return FinalData.Momo_Data2('momo', Firm, os.path.join(DataPath), userID)
                else:
                    FinalData = Momo_Data3()
                    return FinalData.Momo_Data3('momo', Firm, os.path.join(DataPath), userID)
            else:
                FinalData = Momo_Datap()
                return FinalData.Momo_Datap('momo', Firm, os.path.join(DataPath),userID)
        elif Supplier =='myfone':
            FinalData = Myfone_Data()
            return FinalData.Myfone_Data('myfone', Firm, os.path.join(DataPath),userID)
        elif Supplier == 'payeasy':
            if DataPath.split('.')[-1] != 'csv':
                data = xlrd.open_workbook(os.path.join(DataPath))
                table = data.sheets()[0]
                num_cols = table.ncols
                if num_cols == 22:
                    FinalData = payeasy_Data()
                    return FinalData.payeasy_Data('payeasy', Firm, os.path.join(DataPath), userID)
            else:
              FinalData = payeasy_Data2()
              return FinalData.payeasy_Data2('payeasy', Firm, os.path.join(DataPath),userID)
        elif Supplier == 'pchome':
            FinalData = PCHome_Data()
            return FinalData.PCHome_Data('Pchome', Firm, os.path.join(DataPath),userID)
        elif Supplier == 'pchome1':
            FinalData = PCHome2_Data()
            return FinalData.PCHome2_Data('Pchome', Firm, os.path.join(DataPath),userID)
        elif Supplier == 'pchome2':
            FinalData = PCHome3_Data()
            return FinalData.PCHome3_Data('Pchome',Firm, os.path.join(DataPath),userID)
        elif Supplier == 'udn':
            if DataPath.split('.')[-1] == 'xls':
                data = xlrd.open_workbook(os.path.join(DataPath))
                table = data.sheets()[0]
                num_cols = table.ncols
                if num_cols == 13:
                    FinalData = UDN_Data()
                    return FinalData.UDN_Data('UDN', Firm, os.path.join(DataPath), userID)
                elif num_cols == 27:
                    FinalData = UDN_Data2()
                    return FinalData.UDN_Data2('UDN', Firm, os.path.join(DataPath), userID)
                elif num_cols == 28:
                    FinalData = UDN_Data3()
                    return FinalData.UDN_Data3('UDN', Firm, os.path.join(DataPath), userID)
        elif Supplier == 'yahoo':
            if DataPath.split('.')[-1] == 'xls':
                data=xlrd.open_workbook(os.path.join(DataPath))
                table=data.sheets()[0]
                num_cols=table.ncols
                if num_cols == 26:
                    FinalData = Yahood_Data()
                    return FinalData.Yahood_Data('yahoo', Firm, os.path.join(DataPath),userID)
                elif num_cols == 22:
                    FinalData = Yahoos_Data()
                    return FinalData.Yahoos_Data('yahoo', Firm, os.path.join(DataPath),userID)
                else:
                    FinalData = Yahoo3_Data()
                    return FinalData.Yahoo3_Data('yahoo', Firm, os.path.join(DataPath),userID)
            else:
                with open(os.path.join(DataPath), 'rb') as f:
                    reader = csv.reader(f, delimiter=',',skipinitialspace=True)
                    first_row = next(reader)
                    num_cols = len(first_row)
                    if num_cols == 13 :
                        FinalData = Yahoo1_Data()
                        return FinalData.Yahoo1_Data('yahoo', Firm, os.path.join(DataPath),userID)
                    else:
                        FinalData = Yahoo2_Data()
                        return FinalData.Yahoo2_Data('yahoo', Firm, os.path.join(DataPath),userID)
        elif Supplier == '91mai':
            FinalData = Nine_Data()
            return FinalData.Nine_Data(u'九易'.encode("utf-8"), Firm, os.path.join(DataPath),userID)
        elif Supplier == 'treemall':
            FinalData = TREE_Data()
            return FinalData.TREE_Data(u'國泰Tree'.encode("utf-8"), Firm, os.path.join(DataPath),userID)
        elif Supplier == 'gomaji':
            FinalData = MAJI_Data()
            return FinalData.MAJI_Data(u'夠麻吉'.encode("utf-8"), Firm, os.path.join(DataPath),userID)
        elif Supplier == 'etmall':
            FinalData = GM_Data()
            return FinalData.GM_Data(u'東森購物'.encode("utf-8"), Firm, os.path.join(DataPath),userID)
        elif Supplier == 'books':
            FinalData = Book_Data()
            return FinalData.Book_Data(u'博客來'.encode("utf-8"), Firm, os.path.join(DataPath),userID)
        elif Supplier == 'umall':
            FinalData = UMall_Data()
            return FinalData.UMall_Data(u'森森購物'.encode("utf-8"), Firm, os.path.join(DataPath),userID)
        elif Supplier == 'yahoomall':
            if DataPath.split('.')[-1] == 'xls' or DataPath.split('.')[-1] == 'xlsx':
                data = xlrd.open_workbook(os.path.join(DataPath))
                table = data.sheets()[0]
                num_cols = table.ncols
                if num_cols == 16:
                    FinalData = Yahoo_DataS1()
                    return FinalData.Yahoo_DataS1(u'超級商城'.encode("utf-8"), Firm, os.path.join(DataPath), userID)
                elif num_cols == 24:
                    FinalData = Yahoo_DataS2()
                    return FinalData.Yahoo_DataS2(u'超級商城'.encode("utf-8"), Firm, os.path.join(DataPath), userID)
                elif num_cols == 29:
                    FinalData = Yahoo_DataS3()
                    return FinalData.Yahoo_DataS3(u'超級商城'.encode("utf-8"), Firm, os.path.join(DataPath), userID)
        elif Supplier == 'amart':
            FinalData = LoveBuy_Data()
            return FinalData.LoveBuy_Data(u'愛買'.encode("utf-8"), Firm, os.path.join(DataPath),userID)
        elif Supplier == 'rakuten':
            FinalData = Lotte_Data()
            return FinalData.Lotte_Data(u'樂天'.encode("utf-8"), Firm, os.path.join(DataPath),userID)







if __name__ == '__main__':
    Business = VirtualBusiness()
    Business.virtualbusiness("/data/vbupload/test/", "19647356")
#     Business=VirtualBusiness()
#     Business.virtualbusiness("C:/vbdata/vbupload/17life/7dcc2045-472e-11e6-806e-000c29c1d000/Data2.xls")
