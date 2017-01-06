# -*-  coding: utf-8  -*-
# __author__ = '10408001'
import logging
import xlrd
from GroupBuy.buy123 import buy123
# gomaji
class gomaji(buy123):
    # 解析原始檔
    def parserFile(self, GroupID, UserID, LogisticsID=2, ProductCode=None, inputFile=None, outputFile=None):
        self.init_log('Gomaji_Data', GroupID, UserID, ProductCode, inputFile)
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
            self.writeXls(LogisticsID,result, outputFile)
        except Exception as e:
            logging.error(e.message)
            return 'failure'


if __name__ == '__main__':
    buy = gomaji()
    buy.parserFile('robintest', 'test', 2 , 'MS',
                  inputFile=u'C:/Users/10408001/Desktop/團購平台訂單資訊/gomaji/原始檔/2016.12.27/142642_204347_s.xls', \
                  outputFile=u'C:/Users/10408001/Desktop/20161229-gomaji出貨單.xls')