# -*-  coding: utf-8  -*-
# __author__ = '10408001'
import logging
import codecs
from GroupBuy.buy123 import buy123
# 瘋狂賣客
class Crazymike(buy123):
    # 取 CSV 欄位中某些值
    def ReplaceField(self, SourceString, ReplaceName, intLen=0):
        location = SourceString.find(ReplaceName)
        return SourceString[location- intLen: ]

    # 解析原始檔
    def parserFile(self, GroupID, UserID, LogisticsID=2, ProductCode=None, inputFile=None, outputFile=None):
        self.init_log('Popular_Data', GroupID, UserID, ProductCode, inputFile)
        try:
            data = codecs.open(inputFile,'rb',encoding='big5')
            tmpdata,result = [],[]
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
                tmp.append(row[9])  # 收件人
                tmp.append(row[10])  # 收件地址
                tmp.append(row[11])  # 電話
                tmp.append(row[3])  # 商品名稱
                word = self.ReplaceField(rowResult[5], u'x',-1)
                tmp.append(int(word))   #盒數
                tmp.append(1)  # 訂購份數
                tmp.append('')  # 訂購人
                result.append(tmp)
            self.writeXls(LogisticsID, result, outputFile)
        except Exception as e:
            logging.error(e.message)
            return 'failure'


if __name__ == '__main__':
    buy = Crazymike()
    buy.parserFile('robintest', 'test', 2 , 'MS',
                  inputFile=u'C:/Users/10408001/Desktop/團購平台訂單資訊/瘋狂賣客/原始檔/2016.10.21/orders-1048-ex_2 (2).csv', \
                  outputFile=u'C:/Users/10408001/Desktop/20170104-Crazymike出貨單.xls')