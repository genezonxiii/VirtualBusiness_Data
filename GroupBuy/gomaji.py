# -*-  coding: utf-8  -*-
# __author__ = '10408001'
import datetime,time
import logging
import xlrd,xlwt
from ToMysql import ToMysql
import uuid
from VirtualBusiness import Customer,updateCustomer
from xlutils.copy import copy
from GroupBuy import ExcelTemplate
from GroupBuy.buy123 import buy123

class gomaji(buy123):
    # 解析原始檔
    def parserXls(self, GroupID, UserID, ProductCode=None, inputFile=None, outputFile=None):
        logging.basicConfig(filename='pyupload.log', level=logging.DEBUG, format='%(asctime)s %(message)s',
                            datefmt='%Y/%m/%d %I:%M:%S %p')
        logging.Formatter.converter = time.gmtime
        logging.info('===Gomaji_Data===')
        logging.debug('GroupID:' + GroupID)
        logging.debug('path:' + inputFile)
        logging.debug('UserID:' + UserID)
        self.GroupID = GroupID
        self.customer = Customer()
        self.UserID = UserID
        self.ProductCode = ProductCode
        try:
            data = xlrd.open_workbook(inputFile)
            table = data.sheets()[0]
            result = []
            # 讀 excel 檔
            for row_index in range(1, table.nrows):
                tmp = []
                tmp.append(self.ConvertText(str(table.cell(row_index, 2))))  # 代碼
                tmp.append(table.cell(row_index, 6).value)  # 收件人
                tmp.append(table.cell(row_index, 11).value)  # 收件地址
                tmp.append(str(table.cell(row_index, 8).value).replace('-',''))  # 電話
                tmp.append(table.cell(row_index, 10).value)  # 方案名稱
                tmp.append(self.getResultForDigit(table.cell(row_index, 10).value)) # 方案名稱盒數
                tmp.append(table.cell(row_index, 9).value)  # 訂單份數
                tmp.append(table.cell(row_index, 5).value)  # 訂購人
                result.append(tmp)
            self.writeT_catXls(result, outputFile)
        except Exception as e:
            logging.error(e.message)
            return 'failure'


if __name__ == '__main__':
    buy = gomaji()
    buy.parserXls('robintest', 'test', 'MS',
                  inputFile=u'C:/Users/10408001/Desktop/團購平台訂單資訊/gomaji/原始檔/2016.12.27/142642_204347_s.xls', \
                  outputFile=u'C:/Users/10408001/Desktop/20161229-gomaji出貨單.xls')