# -*-  coding: utf-8  -*-
# __author__ = '10408001'
import logging
import json
from GroupBuy.buy123 import buy123
from bs4 import BeautifulSoup

# HerBuy
class HerBuy(buy123):
    # 解析原始檔
    def parserFile(self, GroupID, UserID, Platform = None, LogisticsID=2, ProductCode=None, inputFile=None, outputFile=None):
        self.init_log('HerBuy_Data', GroupID, UserID, Platform, ProductCode, inputFile)
        print inputFile
        success = False
        if self.getRegularEx(GroupID) == False:
            return json.dumps({"success": success}, sort_keys=False)
        else:
            try:
                self.dup_order_no = []
                file = open(inputFile).read()
                soup = BeautifulSoup(file, 'xml')
                workbook = []
                resultinfo = ""
                for sheet in soup.findAll('Worksheet'):
                    sheet_as_list = []
                    for row in sheet.findAll('Row'):
                        row_as_list = []
                        for cell in row.findAll('Cell'):
                            if not cell.Data is None:
                                row_as_list.append(cell.Data.text)
                            else:
                                pass
                        sheet_as_list.append(row_as_list)
                    workbook.append(sheet_as_list)
                print  "testt"
                result = []
                for row in workbook:
                    for i in range(len(row)):
                        if i >= 6 :
                            print row[i]
                            if row[i]<> [] and row[i][0] <> u'物流商列表':
                                tmp = []
                                tmp.append(row[i][1])  # 訂單編號
                                tmp.append(row[i][14])  # 收件人
                                tmp.append(row[i][16])  # 收件地址
                                tmp.append(row[i][15])  # 電話
                                tmp.append(row[i][7])  # 檔次名稱
                                tmp.append(1)  # 盒數
                                tmp.append(row[i][10])  # 數量
                                tmp.append("")  # 訂購人
                                tmp.append(row[i][22])  # 商品料號
                                result.append(tmp)
                            else:
                                break
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
    buy = HerBuy()
    buy.parserFile('cbcc3138-5603-11e6-a532-000d3a800878', 'test', 'herbuy',2, 'MS',
                  inputFile=u'C:/Users/10509002/Desktop/all_20161213__229_orders_20161213142038.xml', \
                  outputFile=u'C:/Users/10509002/Desktop/20170103-HerBuy出貨單.xls')