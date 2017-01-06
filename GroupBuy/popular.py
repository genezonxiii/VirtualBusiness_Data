# -*-  coding: utf-8  -*-
# __author__ = '10408001'
import logging
import codecs
import json
from GroupBuy.buy123 import buy123
# 小P團購
class popular(buy123):
    # 解析原始檔
    def parserFile(self, GroupID, UserID, LogisticsID=2, ProductCode=None, inputFile=None, outputFile=None):
        self.init_log('Popular_Data', GroupID, UserID, ProductCode, inputFile)
        try:
            data = codecs.open(inputFile,'rb',encoding='big5')
            tmpdata,result = [],[]
            success = False
            resultinfo = ""
            # 讀 csv 檔
            i = 0
            for row in data:
                if i > 0 :
                    rowResult = row.split(',')
                    tmp = []
                    for j in rowResult:
                        if rowResult[0]=="" or rowResult[0]==" ":
                            break
                        j = j.replace("'","")
                        tmp.append(j)
                    if tmp <> []:
                        tmpdata.append(tmp)
                i +=1

            for row in tmpdata:
                tmp = []
                tmp.append(row[0])  # 訂單編號
                tmp.append(row[12])  # 收件人
                tmp.append(row[13])  # 收件地址
                tmp.append(row[14])  # 電話
                tmp.append(row[3])  # 檔次名稱
                tmp.append(1) #盒數
                # word = self.ReplaceField(rowResult[17], u'盒')
                # print self.getResultForDigit(word)
                # tmp.append(self.getResultForDigit(word[word.find(u"多多")+2])) # 商品名稱
                tmp.append(row[8])   #訂購總數
                tmp.append(row[11])  # 訂購人
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
    buy = popular()
    buy.parserFile('robintest', 'test', 2 , 'MS',
                  inputFile=u'C:/Users/10408001/Desktop/團購平台訂單資訊/小P/原始檔/2016.10.28/20161028.csv', \
                  outputFile=u'C:/Users/10408001/Desktop/20170104-popular出貨單.xls')