# -*-  coding: utf-8  -*-
# __author__ = '10408001'

import logging
import xlrd
import json
from GroupBuy.buy123 import buy123
# 好吃宅配網
class food123(buy123):
    # 解析原始檔
    def parserFile(self, GroupID, UserID, LogisticsID=2, ProductCode=None, inputFile=None, outputFile=None):
        self.init_log('Food123_Data', GroupID, UserID, ProductCode, inputFile)
        success = False
        try:
            self.getRegularEx(GroupID)
            data = xlrd.open_workbook(inputFile)
            table = data.sheets()[0]
            result = []
            resultinfo = ""
            # 讀 excel 檔
            for row_index in range(1, table.nrows):
                tmp = []
                tmp.append(self.ReplaceField(str(table.cell(row_index, 0).value), '.')) #訂單編號
                tmp.append(self.ReplaceField(table.cell(row_index, 1).value, '(', 1))  # 收件人
                tmp.append(table.cell(row_index, 2).value)  # 收件地址
                tmp.append('0' + self.ReplaceField(str(table.cell(row_index, 3).value), '.'))  # 電話
                tmp.append(table.cell(row_index, 5).value)  # 檔次名稱
                # tmp.append(self.ReplaceField(table.cell(row_index, 6).value, u'盒'))  # 訂購方案
                tmp.append(self.getResultForDigit(self.parserRegularEx(table.cell(row_index, 6).value)))
                # order = self.ReplaceField(table.cell(row_index, 6).value, u'盒')
                # if "." in order :
                #     order = order.split(".")[1].strip()
                # tmp.append(order)
                tmp.append(table.cell(row_index, 7).value)  # 訂單份數
                tmp.append("")  # 訂購人
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
    buy = food123()
    print buy.parserFile('robintest', 'test',2, 'MS',
                  inputFile=u'C:/Users/10509002/Desktop/test/好吃/2017-01-16_好吃宅配網_FD12381462F_悠活原力有限公司_欣敏立清益生菌-紅蘋果多多_未出貨(找不到檔案).xls', \
                  outputFile=u'C:/Users/10509002/Desktop/test/2017-01-16_好吃宅配網_FD12381462F_悠活原力有限公司_欣敏立清益生菌-紅蘋果多多_未出貨(找不到檔案).xls')