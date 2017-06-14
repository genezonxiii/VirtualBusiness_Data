# -*-  coding: utf-8  -*-
# __author__ = '10408001'
import logging
import codecs
import json
from GroupBuy.buy123 import buy123
# 小P團購
class popular(buy123):
    # 解析原始檔
    def parserFile(self, GroupID, UserID, Platform = None, LogisticsID=2, ProductCode=None, inputFile=None, outputFile=None):
        self.init_log('Popular_Data', GroupID, UserID, Platform, ProductCode, inputFile)
        success = False
        if self.getRegularEx(GroupID) == False:
            return json.dumps({"success": success}, sort_keys=False)
        else:
            try:
                self.dup_order_no = []
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
                    tmp.append(row[25].split('\r\n')[0]) # 商品料號
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
    buy = popular()
    buy.parserFile('cbcc3138-5603-11e6-a532-000d3a800878', 'test', 'popular',2 , 'MS',
                  inputFile=u'C:/Users/10509002/Desktop/20161028.csv', \
                  outputFile=u'C:/Users/10509002/Desktop/20170104-popular出貨單.xls')