# -*-  coding: utf-8  -*-
# __author__ = '10408001'

import logging
import xlrd
import json
from GroupBuy.buy123 import buy123

#松果購物
class Pcone(buy123):
    # 解析原始檔
    def parserFile(self, GroupID, UserID, LogisticsID=2, ProductCode=None, inputFile=None, outputFile=None):
        self.init_log('Sale123_Data', GroupID, UserID, ProductCode, inputFile)
        success = False
        if self.getRegularEx(GroupID) == False:
            return json.dumps({"success": success}, sort_keys=False)
        else:
            try:
                data = xlrd.open_workbook(inputFile)
                table = data.sheets()[0]
                result = []
                resultinfo = ""
                # 讀 excel 檔
                for row_index in range(1, table.nrows):
                    tmp = []
                    tmp.append(self.ReplaceField(str(table.cell(row_index, 0).value), '.'))  # 訂單編號
                    tmp.append(self.ReplaceField(table.cell(row_index, 6).value, '(', 1))  # 收件人
                    tmp.append(table.cell(row_index, 8).value)  # 收件地址
                    tmp.append('0' + self.ReplaceField(str(table.cell(row_index, 7).value), '.'))  # 電話
                    tmp.append(table.cell(row_index, 9).value)  # 商品
                    order = self.getResultForDigit(self.parserRegularEx(table.cell(row_index, 11).value))  # 方案
                    # order = self.ReplaceField(table.cell(row_index, 11).value, u'入')
                    # if order in ".":
                    #     order = order.split(".")[1]
                    tmp.append(order)
                    tmp.append(1)  # 訂單份數
                    tmp.append(self.ReplaceField(table.cell(row_index, 2).value,'/'))  # 訂購人
                    result.append(tmp)
                success = self.writeXls(LogisticsID, result, outputFile)
            except Exception as e:
                logging.error(e.message)
                resultinfo = e.message
                success = False
            finally:
                if success == False :
                    Message = UserID + u' 轉檔錯誤，檔案路徑為 ：'
                    self.sendMailToPSC(Message,inputFile)
                return json.dumps({"success": success, "info": resultinfo,"download": outputFile}, sort_keys=False)

if __name__ == '__main__':
    buy = Pcone()
    print buy.parserFile('cbcc3138-5603-11e6-a532-000d3a800878', 'test', 2, 'OS',
                   inputFile=u'C:\\Users\\10509002\\Desktop\\test\\松果\\2017-01-16_松果購物_L1814016_出貨單.xls', \
                   outputFile=u'C:\\Users\\10509002\\Desktop\\2017-01-16_松果購物_L1814016_出貨單.xls')
