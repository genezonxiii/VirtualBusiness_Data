# -*-  coding: utf-8  -*-
# __author__ = '10408001'
import logging
import codecs
import json
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
        success = False
        try:
            self.getRegularEx(GroupID)
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
                #word = self.ReplaceField(rowResult[5], u'x',-1)
                word = self.getResultForDigit(self.parserRegularEx(rowResult[5]))
                tmp.append(int(word))   #盒數
                tmp.append(1)  # 訂購份數
                tmp.append('')  # 訂購人
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
            return json.dumps({"success": success, "info": resultinfo, "download": outputFile}, sort_keys=False)


if __name__ == '__main__':
    buy = Crazymike()
    print buy.parserFile('cbcc3138-5603-11e6-a532-000d3a800878', 'test', 2 , 'MS',
                  inputFile=u'C:/Users/10509002/Desktop/for_Joe_test/團購/瘋狂賣客/orders-1048-ex_2 (2).csv', \
                  outputFile=u'C:/Users/10509002/Desktop/20170104-Crazymike出貨單.xls')