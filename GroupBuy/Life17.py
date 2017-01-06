# -*-  coding: utf-8  -*-
# __author__ = '10408001'

import logging
import xlrd
import json
from GroupBuy.buy123 import buy123
#17P團購
class Life17(buy123):
    #拆解字串中的包數
    def parserString(self,SourceString,keyWord):
        return SourceString.find(keyWord)

    # 解析原始檔
    def parserFile(self, GroupID, UserID, LogisticsID=2, ProductCode=None, inputFile=None, outputFile=None):
        self.init_log('17P_Data', GroupID, UserID, ProductCode, inputFile)
        try:
            data = xlrd.open_workbook(inputFile)
            table = data.sheets()[0]
            result = []
            success = False
            resultinfo = ""
            # 讀 excel 檔
            for row_index in range(1, table.nrows):
                tmp = []
                tmp.append(table.cell(row_index, 1).value) #訂單編號
                tmp.append(table.cell(row_index, 3).value)  # 收件人
                tmp.append(table.cell(row_index, 5).value)  # 收件地址
                tmp.append('0' + self.ReplaceField(str(table.cell(row_index, 4).value), '.'))  # 電話
                tmp.append(table.cell(row_index, 6).value)  # 方案名稱
                #暫時以 "菌" 及"包" 來判斷,需要跟悠活原力確認
                word = self.ReplaceField(table.cell(row_index, 6).value,u"包")
                tmp.append(int(word[word.find(u"菌")+1:])/30) # 訂購方案
                tmp.append(table.cell(row_index, 9).value)  # 訂單份數
                tmp.append("")  # 訂購人
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
    buy = Life17()
    # word = buy.ReplaceField("[24H出貨]欣敏立清-草莓多多益生菌120包+贈德國Purafit-維他命C發泡錠Vitamin C","包")
    # print int(word[word.find("菌")+3:])/30
    buy.parserFile('robintest', 'test',2, 'DS',
                  inputFile=u'C:/Users/10408001/Desktop/團購平台訂單資訊/17P\原始檔/2016.12.26/10397319_[24H出貨]欣敏立清-草莓多多益生菌_出貨清冊 (1).xls', \
                  outputFile=u'C:/Users/10408001/Desktop/20161229-17P出貨單.xls')