# -*-  coding: utf-8  -*-
# __author__ = '10408001'
import logging
import xlrd
import json
from GroupBuy.buy123 import buy123
# gomaji
class gomaji(buy123):
    # 解析原始檔
    def parserFile(self, GroupID, UserID, Platform = None, LogisticsID=2, ProductCode=None, inputFile=None, outputFile=None):
        self.init_log('Gomaji_Data', GroupID, UserID, Platform, ProductCode, inputFile)
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
                    tmp.append(self.ConvertText(str(table.cell(row_index, 2))))  # 代碼
                    tmp.append(table.cell(row_index, 6).value)  # 收件人
                    tmp.append(table.cell(row_index, 11).value)  # 收件地址
                    tmp.append(str(table.cell(row_index, 8).value).replace('-',''))  # 電話
                    tmp.append(table.cell(row_index, 10).value)  # 方案名稱
                    order = self.getResultForDigit(self.parserRegularEx(table.cell(row_index, 10).value))   # 方案名稱盒數
                    # order = self.ReplaceField(table.cell(row_index, 10).value, u'盒')
                    # if "." in order :
                    #     order = order.split(".")[1].strip()
                    # tmp.append(order)
                    tmp.append(order)
                    # tmp.append(self.getResultForDigit(self.ReplaceField(table.cell(row_index, 10).value, u'盒')))  # 方案名稱盒
                    tmp.append(table.cell(row_index, 9).value)  # 訂單份數
                    tmp.append(table.cell(row_index, 5).value)  # 訂購人
                    tmp.append(table.cell(row_index, 24).value) # 商品料號
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
    buy = gomaji()
    print buy.parserFile('cbcc3138-5603-11e6-a532-000d3a800878', 'test','gomaji', 2 , 'MS',
                  inputFile=u'C:/Users/10509002/Desktop/146917_209521_s.xls', \
                  outputFile=u'C:/Users/10509002/Desktop/2017-0612_夠麻吉.xls')