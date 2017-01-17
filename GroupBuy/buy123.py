# -*-  coding: utf-8  -*-
# __author__ = '10408001'
import datetime,time,json
import logging
import xlrd,xlwt
from ToMysql import ToMysql
import uuid
from VirtualBusiness import Customer,updateCustomer
from xlutils.copy import copy
# from GroupBuy import ExcelTemplate
#FilePath
class ExcelTemplate():
    T_Cat_OutputFilePath, T_Cat_TemplateFile = None, None
    def __init__(self):
        self.T_Cat_TemplateFile = '/data/vbGroupbuy/Logistics_Tcat.xls'
        self.T_Cat_OutputFilePath = '/data/vbGroupbuy_output/'

#生活市集
class buy123():
    mysqlconnect = None
    customer = None
    GroupID , UserID , ProductCode = None , None , None
    def __init__(self):
        self.mysqlconnect = ToMysql()
        self.mysqlconnect.connect()

    # 把 Excel 欄位中的 text:u 等字元 replace
    def ConvertText(self, SourceString):
        return SourceString.replace("text:u", "").replace("'", "")

    # 中文數字轉換成阿拉伯數字
    def getResultForDigit(self, a, encoding="utf-8"):
        Numdict = {u'零': 0, u'一': 1, u'二': 2, u'三': 3, u'四': 4, u'五': 5, u'六': 6, u'七': 7, u'八': 8, u'九': 9,
                   u'十': 10, \
                   u'百': 100, u'千': 1000, u'万': 10000, u'０': 0, u'１': 1, u'２': 2, u'３': 3, u'４': 4, u'５': 5,
                   u'６': 6, \
                   u'７': 7, u'８': 8, u'９': 9, u'壹': 1, u'貳': 2, u'參': 3, u'肆': 4, u'伍': 5, u'陸': 6, u'柒': 7,
                   u'捌': 8, \
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

    #取 Excel 欄位中某些值
    def ReplaceField(self,SourceString,ReplaceName,intLen=0):
        location = SourceString.find(ReplaceName)
        return SourceString[:location-intLen]

    #初始記錄解析 Log
    def init_log(self,dataName,GroupID,UserID ,ProductCode,inputFile):
        logging.basicConfig(filename='pyupload.log', level=logging.DEBUG, format='%(asctime)s %(message)s',
                            datefmt='%Y/%m/%d %I:%M:%S %p')
        logging.Formatter.converter = time.gmtime
        logging.info('===' + dataName + '===')
        logging.debug('GroupID:' + GroupID)
        logging.debug('path:' + inputFile)
        logging.debug('UserID:' + UserID)
        self.GroupID = GroupID
        self.customer = Customer()
        self.UserID = UserID
        self.ProductCode = ProductCode

    #解析原始檔
    def parserFile(self, GroupID, UserID, LogisticsID=2, ProductCode=None, inputFile=None, outputFile=None):
        self.init_log('Buy123_Data',GroupID, UserID , ProductCode , inputFile)
        success = False
        try:
            data = xlrd.open_workbook(inputFile)
            table = data.sheets()[0]
            result = []
            resultinfo = ""
            # 讀 excel 檔
            for row_index in range(1, table.nrows):
                tmp=[]
                tmp.append(self.ReplaceField(str(table.cell(row_index, 0).value),'.'))           #訂單編號
                tmp.append(self.ReplaceField(table.cell(row_index, 1).value,'(',1))              # 收件人
                tmp.append(table.cell(row_index, 2).value)                                      #收件地址
                tmp.append('0' + self.ReplaceField(str(table.cell(row_index, 3).value),'.'))      #電話
                tmp.append(table.cell(row_index, 5).value)                                      #檔次名稱
                tmp.append(self.ReplaceField(table.cell(row_index, 6).value,u'盒'))              #訂購方案
                tmp.append(table.cell(row_index, 7).value)                                      #組數
                tmp.append(self.ReplaceField(table.cell(row_index, 8).value,'/'))                #訂購人                                                   #訂購人
                result.append(tmp)
            success = self.writeXls(LogisticsID,result,outputFile)
        except Exception as e :
            logging.error(e.message)
            resultinfo = e.message
            success = False
        finally:
            return json.dumps({"success": success, "info": resultinfo,"download": outputFile}, sort_keys=False)

    def writeXls(self,LogisticsID,data,outputFile):
        if LogisticsID == 2 :
            return self.writeT_catXls(data,outputFile)

    #寫入黑貓出貨單內容
    def writeT_catXls(self,data,outputFile):
        try:
            Template = ExcelTemplate()
            fileTemplate =  Template.T_Cat_TemplateFile
            rb = xlrd.open_workbook(fileTemplate)
            file = copy(rb)
            table = file.get_sheet(0)
            i = 1
            success = False
            for row in data:
                d1 = datetime.datetime.strftime(datetime.date.today(),'%Y/%m/%d')
                d2 = datetime.datetime.strftime(datetime.date.today()+ datetime.timedelta(days=1), '%Y/%m/%d')
                table.write(i,0,d1)             # 收件日
                table.write(i, 1,d2)            # 配達日
                table.write(i,3,row[0])         # 訂單編號
                table.write(i, 5, row[7])       # 訂購人
                table.write(i, 6, row[1])       # 收件人
                table.write(i, 7, row[2])       # 收件地址
                table.write(i, 8, row[3])       # 收件人手機1
                table.write(i, 10, str(int(row[5])*int(row[6])) + str(self.ProductCode) )  #交易備註
                table.write(i, 11,'1')           # 送貨時段
                table.write(i, 12, row[4])      # 訂購品項
                table.write(i, 13, row[6])      # 訂購份數
                table.write(i, 14, int(row[5]))      # 盒數
                table.write(i, 15, int(row[5])*int(row[6]))  # 總數量
                self.customer.setGroup_id(self.GroupID)
                self.customer.setNameNoEncode(row[1])
                self.customer.setMobile(row[3])
                self.customer.setAddressNoEncode(row[2])
                # insert or update table tb_customer
                self.updateDB_Customer()
                i += 1
            self.mysqlconnect.db.commit()
            self.mysqlconnect.dbClose()
            file.save(outputFile)
            success = True
        except Exception as e :
            logging.error(e.message)
            return False
        finally:
            return success
    #將訂購人資料寫入 DB
    def updateDB_Customer(self):
        try:
            # insert or update table tb_customer
            updatecustomer = updateCustomer()
            self.customer.setCustomer_id(
                updatecustomer.checkCustomerid(self.customer.getGroup_id(), self.customer.get_Name(), self.customer.get_Address(), \
                                               self.customer.get_phone(), self.customer.get_Mobile(), self.customer.get_Email()))
            if self.customer.getCustomer_id() == None:
                self.customer.setCustomer_id(uuid.uuid4())
                CustomereSQL = (
                    self.customer.getCustomer_id(), self.customer.getGroup_id(), self.customer.getName(), \
                    self.customer.getAddress(), self.customer.getphone(), self.customer.getMobile(), \
                    self.customer.getEmail(), self.customer.getPost(), self.customer.getClass(), self.customer.getMemo(), self.UserID)
                self.mysqlconnect.cursor.callproc('sp_insert_customer_bysys', CustomereSQL)
            else:
                CustomereSQL = (self.customer.getCustomer_id(), self.customer.getGroup_id(), self.customer.getName(), \
                                self.customer.getAddress(), self.customer.getphone(), self.customer.getMobile(), \
                                self.customer.getEmail(), self.customer.getPost(), self.customer.getClass(), \
                                self.customer.getMemo(),self.UserID)
                self.mysqlconnect.cursor.callproc('sp_update_customer', CustomereSQL)

            CustomereSQL = (self.customer.getCustomer_id(), self.customer.getGroup_id(), self.customer.get_Name(), \
                            self.customer.get_Address(), self.customer.get_phone(), self.customer.get_Mobile(), \
                            self.customer.get_Email())
            updatecustomer.updataData(CustomereSQL)
        except Exception as e :
            print e.message
            logging.error(e.message)
            raise

if __name__ == '__main__':
    buy = buy123()
    print buy.parserFile('robintest', 'test', 2, 'MS',
                  inputFile=u'C:/Users/10408001/Desktop/團購平台訂單資訊/生活市集/原始檔/2016.12.05/2016-12-05_生活市集_BY123375489F_悠活原力有限公司_(0822食品高毛利)欣敏立清益生菌-蔓越莓多多(32點5%策略性商品)_未出貨.xls', \
                  outputFile=u'C:/Users/10408001/Desktop/20161228-1出貨單.xls')