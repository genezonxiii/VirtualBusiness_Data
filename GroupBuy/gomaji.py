# -*-  coding: utf-8  -*-
# __author__ = '10408001'
import datetime,time
import logging
import xlrd,xlwt
from ToMysql import ToMysql
import uuid
from VirtualBusiness import Customer,updateCustomer
from xlutils.copy import copy
from GroupBuy import ExcelTemplate
from GroupBuy.buy123 import buy123

class gomaji(buy123):

    #中文數字轉換成阿拉伯數字
    def getResultForDigit(self,a, encoding="utf-8"):
        Numdict = {u'零': 0, u'一': 1, u'二': 2, u'三': 3, u'四': 4, u'五': 5, u'六': 6, u'七': 7, u'八': 8, u'九': 9, u'十': 10, \
                   u'百': 100, u'千': 1000, u'万': 10000, u'０': 0, u'１': 1, u'２': 2, u'３': 3, u'４': 4, u'５': 5, u'６': 6, \
                   u'７': 7, u'８': 8, u'９': 9, u'壹': 1, u'貳': 2, u'參': 3, u'肆': 4, u'伍': 5, u'陸': 6, u'柒': 7, u'捌': 8, \
                   u'玖': 9, u'拾': 10, u'佰': 100, u'仟': 1000, u'萬': 10000, u'億': 100000000}
        if isinstance(a, str):
            a = a.decode(encoding)
        count = 0
        result = 0
        tmp = 0
        Billion = 0
        while count < len(a):
            tmpChr = a[count]
            # print tmpChr
            tmpNum = Numdict.get(tmpChr, None)
            # 如果等于1亿
            if tmpNum == 100000000:
                result = result + tmp
                result = result * tmpNum
                # 获得亿以上的数量，将其保存在中间变量Billion中并清空result
                Billion = Billion * 100000000 + result
                result = 0
                tmp = 0
            # 如果等于1万
            elif tmpNum == 10000:
                result = result + tmp
                result = result * tmpNum
                tmp = 0
            # 如果等于十或者百，千
            elif tmpNum >= 10:
                if tmp == 0:
                    tmp = 1
                result = result + tmpNum * tmp
                tmp = 0
            # 如果是个位数
            elif tmpNum is not None:
                tmp = tmp * 10 + tmpNum
            count += 1
        result = result + tmp
        result = result + Billion
        return result

    # 取 Excel 欄位中某些值
    def ReplaceField(self, SourceString, ReplaceName, intLen=0):
        location = SourceString.find(ReplaceName)
        SourceString = SourceString[:location - intLen]
        return self.getResultForDigit(SourceString)

    def ConvertText(self,SourceString):
        return SourceString.replace("text:u","").replace("'","")

    # 解析原始檔
    def parserXls(self, GroupID, UserID, ProductCode=None, inputFile=None, outputFile=None):
        logging.basicConfig(filename='pyupload.log', level=logging.DEBUG, format='%(asctime)s %(message)s',
                            datefmt='%Y/%m/%d %I:%M:%S %p')
        logging.Formatter.converter = time.gmtime
        logging.info('===Gomaji_Data===')
        logging.debug('GroupID:' + GroupID)
        logging.debug('path:' + inputFile)
        logging.debug('UserID:' + UserID)
        self.GroupID = GroupID
        self.customer = Customer()
        self.UserID = UserID
        self.ProductCode = ProductCode
        try:
            data = xlrd.open_workbook(inputFile)
            table = data.sheets()[0]
            result = []
            # 讀 excel 檔
            for row_index in range(1, table.nrows):
                tmp = []
                tmp.append(self.ConvertText(str(table.cell(row_index, 2))))  # 代碼
                tmp.append(table.cell(row_index, 6).value)  # 收件人
                tmp.append(table.cell(row_index, 11).value)  # 收件地址
                tmp.append(str(table.cell(row_index, 8).value).replace('-',''))  # 電話
                tmp.append(table.cell(row_index, 10).value)  # 方案名稱
                tmp.append(self.ReplaceField(table.cell(row_index, 10).value, u'盒'))  # 方案名稱
                tmp.append(table.cell(row_index, 9).value)  # 訂單份數
                tmp.append(table.cell(row_index, 5).value)  # 訂購人
                result.append(tmp)
            self.writeT_catXls(result, outputFile)
        except Exception as e:
            logging.error(e.message)
            return 'failure'


if __name__ == '__main__':
    buy = gomaji()
    buy.parserXls('robintest', 'test', 'MS',
                  inputFile=u'C:/Users/10408001/Desktop/團購平台訂單資訊/gomaji/原始檔/2016.12.27/142642_204347_s.xls', \
                  outputFile=u'C:/Users/10408001/Desktop/20161229-gomaji出貨單.xls')