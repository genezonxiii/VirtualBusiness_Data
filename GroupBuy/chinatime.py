# -*-  coding: utf-8  -*-
# __author__ = '10408001'

import logging
import xlrd
import json
from GroupBuy.buy123 import buy123
#中時團購
class chinatime(buy123):
    # 解析盒數
    def parserBox(self,SourceString):
        boxList=[u'盒',u'瓶']
        parameter = ""
        for i in range(len(boxList)):
            if SourceString.find(boxList[i]) <> -1 :
                parameter = boxList[i]
                break
        box = self.ReplaceField(SourceString, parameter)
        box = box[box.find("/")+1:]
        if box == "" :
            return 1
        else:
            return int(box)

    # 解析原始檔
    def parserFile(self, GroupID, UserID, LogisticsID=2, ProductCode=None, inputFile=None, outputFile=None):
        self.init_log('chinatime_Data', GroupID, UserID, ProductCode, inputFile)
        success = False
        try:
            data = xlrd.open_workbook(inputFile)
            table = data.sheets()[0]
            result = []
            resultinfo = ""
            orderNo,Name,Address,CellPhone="","","",""
            # 讀 excel 檔
            for row_index in range(1, table.nrows):
                tmp = []
                if table.cell(row_index, 0).value <> "":
                    orderNo = table.cell(row_index, 0).value
                tmp.append(orderNo) #訂單編號
                if table.cell(row_index, 2).value <> "":
                    Name = table.cell(row_index, 2).value
                tmp.append(Name)  # 收件人
                if table.cell(row_index, 4).value <> "":
                    Address = table.cell(row_index, 4).value
                tmp.append(Address)  # 收件地址
                if table.cell(row_index, 3).value <> "":
                    CellPhone = table.cell(row_index, 3).value
                tmp.append(CellPhone)  # 電話
                tmp.append(table.cell(row_index, 5).value)  # 商品名稱
                # tmp.append(self.parserBox(table.cell(row_index, 5).value)) # 訂購方案(規則太亂)
                tmp.append(1) # 訂購方案
                tmp.append(table.cell(row_index, 6).value)  # 訂單份數
                tmp.append("")  # 訂購人
                result.append(tmp)
            success = self.writeXls(LogisticsID, result, outputFile)
        except Exception as e:
            logging.error(e.message)
            resultinfo = e.message
            success = False
        finally:
            return json.dumps({"success": success, "info": resultinfo,"download": outputFile}, sort_keys=False)

if __name__ == '__main__':
    buy = chinatime()
    buy.parserFile('robintest', 'test',2, 'MS',
                  inputFile=u'/Users/csi/Desktop/團購/中時團購/general/396a2df8-472e-11e6-806e-000c29c1d067/20161020待出貨訂單列表.xlsx', \
                  outputFile=u'/Users/csi/Desktop/20161229-chinatime出貨單.xls')