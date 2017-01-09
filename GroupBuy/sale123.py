# -*-  coding: utf-8  -*-
# __author__ = '10408001'

import logging
import xlrd
import json
from GroupBuy.buy123 import buy123

#姊妹購物網
class Sale123(buy123):
    # 解析原始檔
    def parserFile(self, GroupID, UserID, LogisticsID=2, ProductCode=None, inputFile=None, outputFile=None):
        self.init_log('Sale123_Data', GroupID, UserID, ProductCode, inputFile)
        success = False
        try:
            data = xlrd.open_workbook(inputFile)
            table = data.sheets()[0]
            result = []
            resultinfo = ""
            # 讀 excel 檔
            for row_index in range(1, table.nrows):
                tmp = []
                tmp.append(self.ReplaceField(str(table.cell(row_index, 0).value),'.'))  # 訂單編號
                tmp.append(self.ReplaceField(table.cell(row_index, 1).value,'(',1))  # 收件人
                tmp.append(table.cell(row_index, 2).value)  # 收件地址
                tmp.append('0' + self.ReplaceField(str(table.cell(row_index, 3).value), '.'))  # 電話
                tmp.append(table.cell(row_index, 5).value)  # 檔次名稱
                tmp.append(self.ReplaceField(table.cell(row_index,6).value,u"盒"))  # 訂購方案
                tmp.append(table.cell(row_index, 7).value)  # 訂單份數
                tmp.append('')  # 訂購人
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
    buy = Sale123()
    buy.parserFile('robintest', 'test', 2, 'DS',
                   inputFile=u'C:/Users/10408001/Desktop/團購平台訂單資訊/姊妹購物網/原始檔/2016.12.02/2016-12-02_姊妹購物網_AY12390480F_悠活原力有限公司_欣敏立清益生菌-草莓多多_未出貨.xls', \
                   outputFile=u'C:/Users/10408001/Desktop/20170104-姊妹購物網出貨單.xls')
