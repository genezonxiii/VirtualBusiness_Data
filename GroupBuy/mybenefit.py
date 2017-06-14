# -*-  coding: utf-8  -*-
# __author__ = '10408001'

import logging
import xlrd
import json
from GroupBuy.buy123 import buy123

#中華優購
class Mybenefit(buy123):
    # 解析原始檔
    def parserFile(self, GroupID, UserID, Platform = None, LogisticsID=2, ProductCode=None, inputFile=None, outputFile=None):
        self.init_log('Mybenefit_Data', GroupID, UserID, Platform, ProductCode, inputFile)
        success = False
        if self.getRegularEx(GroupID) == False:
            return json.dumps({"success": success}, sort_keys=False)
        else:
            try:
                self.dup_order_no = []
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
                    tmp.append(table.cell(row_index, 17).value) # 商品料號
                    result.append(tmp)
                success = self.writeXls(LogisticsID, result, outputFile)
            except Exception as e:
                logging.error(e.message)
                resultinfo = e.message
                success = False
            finally:
                dup_str = ','.join(self.dup_order_no)
                self.dup_order_no = []
                if success == False :
                    Message = UserID + u' 轉檔錯誤，檔案路徑為 ：'
                    self.sendMailToPSC(Message,inputFile)
                return json.dumps({"success": success, "info": resultinfo,"download": outputFile, "duplicate": dup_str}, sort_keys=False)


if __name__ == '__main__':
    buy = Mybenefit()
    buy.parserFile('cbcc3138-5603-11e6-a532-000d3a800878', 'test', 'chinaugo',2, 'DS',
                   inputFile=u'C:/Users/10509002/Desktop/report.xls', \
                   outputFile=u'C:/Users/10509002/Desktop/20161229-Mybenefit出貨單.xls')
