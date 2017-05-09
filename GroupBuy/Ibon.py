# -*-  coding: utf-8  -*-
# __author__ = '10408001'
import logging
import codecs
import json
from GroupBuy.buy123 import buy123
# I bon
class Ibon(buy123):
    # 解析原始檔
    def parserFile(self, GroupID, UserID, LogisticsID=2, ProductCode=None, inputFile=None, outputFile=None):
        self.init_log('Ibon_Data', GroupID, UserID, ProductCode, inputFile)
        success = False
        if self.getRegularEx(GroupID) == False:
            return json.dumps({"success": success}, sort_keys=False)
        else:
            try:
                data = codecs.open(inputFile,'rb',encoding='big5')
                tmpdata,result = [],[]
                resultinfo = ""
                # 讀 csv 檔
                i = 0
                for row in data:
                    if i > 0 :
                        rowResult = row.split(',')
                        tmp = []
                        for j in rowResult:
                            j = j.replace("'","")
                            tmp.append(j)
                        tmpdata.append(tmp)
                    i +=1

                for row in tmpdata:
                    tmp = []
                    tmp.append(row[4])  # 訂單編號
                    tmp.append(row[6])  # 收件人
                    tmp.append(row[8])  # 收件地址
                    tmp.append(row[7])  # 電話
                    tmp.append(row[17])  # 商品名稱
                    tmp.append(1) #盒數
                    # word = self.ReplaceField(rowResult[17], u'盒')
                    # print self.getResultForDigit(word)
                    # tmp.append(self.getResultForDigit(word[word.find(u"多多")+2])) # 商品名稱
                    tmp.append(row[19]) #出貨數量
                    tmp.append(row[6])  # 訂購人
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
    buy = Ibon()
    buy.parserFile('robintest', 'test', 2 , 'MS',
                  inputFile=u'C:/Users/10408001/Desktop/團購平台訂單資訊/I bon/宅配出貨單_20161027_14155.csv', \
                  outputFile=u'C:/Users/10408001/Desktop/20170103-Ibon出貨單.xls')