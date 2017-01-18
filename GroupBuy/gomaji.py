# -*-  coding: utf-8  -*-
# __author__ = '10408001'
import logging
import xlrd
import json
from GroupBuy.buy123 import buy123
# gomaji
class gomaji(buy123):
    # 解析原始檔
    def parserFile(self, GroupID, UserID, LogisticsID=2, ProductCode=None, inputFile=None, outputFile=None):
        self.init_log('Gomaji_Data', GroupID, UserID, ProductCode, inputFile)
        success = False
        try:
            data = xlrd.open_workbook(inputFile)
            table = data.sheets()[0]
            result = []
            resultinfo = ""
            # 讀 excel 檔
            for row_index in range(1, table.nrows):
                tmp = []
                tmp.append(self.ConvertText(str(table.cell(row_index, 2))))  # 代碼
                tmp.append(table.cell(row_index, 6).value)  # 收件人
                tmp.append(table.cell(row_index, 11).value)  # 收件地址
                tmp.append(str(table.cell(row_index, 8).value).replace('-',''))  # 電話
                tmp.append(table.cell(row_index, 10).value)  # 方案名稱
                # tmp.append(self.getResultForDigit(self.ReplaceField(table.cell(row_index, 10).value.split('.')[1], u'盒'))) # 方案名稱盒數
                order = self.ReplaceField(table.cell(row_index, 10).value, u'盒')
                if order in ".":
                    order = order.split(".")[1]
                tmp.append(order)
                # tmp.append(self.getResultForDigit(self.ReplaceField(table.cell(row_index, 10).value, u'盒')))  # 方案名稱盒
                tmp.append(table.cell(row_index, 9).value)  # 訂單份數
                tmp.append(table.cell(row_index, 5).value)  # 訂購人
                result.append(tmp)
            success = self.writeXls(LogisticsID, result, outputFile)
        except Exception as e:
            logging.error(e.message)
            resultinfo = e.message
            success = False
        finally:
            return json.dumps({"success": success, "info": resultinfo,"download": outputFile}, sort_keys=False)


if __name__ == '__main__':
    buy = gomaji()
    buy.parserFile('robintest', 'test', 2 , 'MS',
                  inputFile=u'C:/Users/10509002/Desktop/test/GOMAJI/146917_209521_s.xls', \
                  outputFile=u'C:/Users/10509002/Desktop/test/146917_209521_s.xls')