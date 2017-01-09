# -*-  coding: utf-8  -*-
# __author__ = '10408001'

import logging
import xlrd
import json
from GroupBuy.buy123 import buy123

#中華優購
class Mybenefit(buy123):
    # 解析原始檔
    def parserFile(self, GroupID, UserID, LogisticsID=2, ProductCode=None, inputFile=None, outputFile=None):
        self.init_log('Mybenefit_Data', GroupID, UserID, ProductCode, inputFile)
        success = False
        try:
            data = xlrd.open_workbook(inputFile)
            table = data.sheets()[0]
            result = []
            resultinfo = ""
            # 讀 excel 檔
            for row_index in range(1, table.nrows):
                tmp = []
                tmp.append(table.cell(row_index, 2).value)  # 訂單編號
                tmp.append(table.cell(row_index, 3).value)  # 收件人
                tmp.append(table.cell(row_index, 5).value)  # 收件地址
                tmp.append('0' + self.ReplaceField(str(table.cell(row_index, 13).value), '.'))  # 電話
                tmp.append(table.cell(row_index, 9).value)  # 方案名稱
                tmp.append(1)  # 訂購方案
                tmp.append(table.cell(row_index, 11).value)  # 訂單份數
                tmp.append(table.cell(row_index, 1).value)  # 訂購人
                result.append(tmp)
            self.writeXls(LogisticsID, result, outputFile)
            success = True
        except Exception as e:
            logging.error(e.message)
            resultinfo = e.message
            return 'failure'
        finally:
            return json.dumps({"success": success, "info": resultinfo,"download": outputFile}, sort_keys=False)


if __name__ == '__main__':
    buy = Mybenefit()
    buy.parserFile('robintest', 'test', 2, 'DS',
                   inputFile=u'C:/Users/10408001/Desktop/團購平台訂單資訊/中華優購/2016.03.07/2016.03.07 by Jack.xlsx', \
                   outputFile=u'C:/Users/10408001/Desktop/20161229-Mybenefit出貨單.xls')
