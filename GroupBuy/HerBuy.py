# -*-  coding: utf-8  -*-
# __author__ = '10408001'
import logging
import json
from GroupBuy.buy123 import buy123
from bs4 import BeautifulSoup

# HerBuy
class HerBuy(buy123):
    # 解析原始檔
    def parserFile(self, GroupID, UserID, LogisticsID=2, ProductCode=None, inputFile=None, outputFile=None):
        self.init_log('HerBuy_Data', GroupID, UserID, ProductCode, inputFile)
        print inputFile
        success = False
        if self.getRegularEx(GroupID) == False:
            return json.dumps({"success": success}, sort_keys=False)
        else:
            try:
                print 'test'
                file = open(inputFile).read()
                print 'stop'
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
                                tmp.append('')  # 訂單編號
                                tmp.append(row[i][14])  # 收件人
                                tmp.append(row[i][16])  # 收件地址
                                tmp.append(row[i][15])  # 電話
                                tmp.append(row[i][7])  # 檔次名稱
                                tmp.append(1)  # 盒數
                                tmp.append(row[i][10])  # 數量
                                tmp.append("")  # 訂購人
                                result.append(tmp)
                            else:
                                break
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
    buy = HerBuy()
    buy.parserFile('robintest', 'test',2, 'MS',
                  inputFile=u'C:/Users/10408001/Desktop/團購平台訂單資訊/HerBuy/2016.08.12/all_20160812_0_229_orders_20160812113116.xml', \
                  outputFile=u'C:/Users/10408001/Desktop/20170103-HerBuy出貨單-2.xls')