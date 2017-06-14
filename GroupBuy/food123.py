# -*-  coding: utf-8  -*-
# __author__ = '10408001'

import logging
import xlrd
import json
from GroupBuy.buy123 import buy123
# 好吃宅配網
class food123(buy123):
    # 解析原始檔
    def parserFile(self, GroupID, UserID, Platform = None, LogisticsID=2, ProductCode=None, inputFile=None, outputFile=None):
        self.init_log('Food123_Data', GroupID, UserID, Platform, ProductCode, inputFile)
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
                    tmp.append(table.cell(row_index,11).value)  # 商品料號
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
    buy = food123()
    print buy.parserFile('cbcc3138-5603-11e6-a532-000d3a800878', 'test', 'delicious',2, 'MS',
                  inputFile=u'C:/Users/10509002/Desktop/2016-12-05_好吃宅配網_FD12377160F_悠活原力有限公司_(0822食品高毛利)欣敏立清益生菌-蔓越莓多多(32點5%策略性商品)_未出貨.xls', \
                  outputFile=u'C:/Users/10509002/Desktop/2016-12-15_Food.xls')