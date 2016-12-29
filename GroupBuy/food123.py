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

class food123(buy123):
    # 解析原始檔
    def parserXls(self, GroupID, UserID, LogisticsID=2, ProductCode=None, inputFile=None, outputFile=None):
        self.init_log('Food123_Data', GroupID, UserID, ProductCode, inputFile)
        try:
            data = xlrd.open_workbook(inputFile)
            table = data.sheets()[0]
            result = []
            # 讀 excel 檔
            for row_index in range(1, table.nrows):
                tmp = []
                tmp.append(self.ReplaceField(str(table.cell(row_index, 0).value), '.')) #訂單編號
                tmp.append(self.ReplaceField(table.cell(row_index, 1).value, '(', 1))  # 收件人
                tmp.append(table.cell(row_index, 2).value)  # 收件地址
                tmp.append('0' + self.ReplaceField(str(table.cell(row_index, 3).value), '.'))  # 電話
                tmp.append(table.cell(row_index, 5).value)  # 檔次名稱
                tmp.append(self.ReplaceField(table.cell(row_index, 6).value, u'盒'))  # 訂購方案
                tmp.append(table.cell(row_index, 7).value)  # 訂單份數
                tmp.append("")  # 訂購人
                result.append(tmp)
            self.writeXls(LogisticsID, result, outputFile)
        except Exception as e:
            logging.error(e.message)
            return 'failure'

if __name__ == '__main__':
    buy = food123()
    buy.parserXls('robintest', 'test',2, 'MS',
                  inputFile=u'C:/Users/10408001/Desktop/團購平台訂單資訊/好吃宅配網\原始檔/2016.12.05/2016-12-05_好吃宅配網_FD12377160F_悠活原力有限公司_(0822食品高毛利)欣敏立清益生菌-蔓越莓多多(32點5%策略性商品)_未出貨.xls', \
                  outputFile=u'C:/Users/10408001/Desktop/20161229-food123出貨單.xls')