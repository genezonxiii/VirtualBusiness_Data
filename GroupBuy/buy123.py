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

class buy123():
    mysqlconnect = None
    customer = None
    GroupID , UserID = None , None
    def __init__(self):
        self.mysqlconnect = ToMysql()
        self.mysqlconnect.connect()

    #取 Excel 欄位中某些值
    def ReplaceField(self,SourceString,ReplaceName,intLen=0):
        location = SourceString.find(ReplaceName)
        return SourceString[:location-intLen]

    #解析原始檔
    def parserXls(self,GroupID,UserID,filename,output):
        logging.basicConfig(filename='pyupload.log', level=logging.DEBUG, format='%(asctime)s %(message)s',
                            datefmt='%Y/%m/%d %I:%M:%S %p')
        logging.Formatter.converter = time.gmtime
        logging.info('===Buy123_Data===')
        logging.debug('GroupID:' + GroupID)
        logging.debug('path:' + filename)
        logging.debug('UserID:' + UserID)
        self.GroupID = GroupID
        self.customer = Customer()
        self.UserID = UserID
        try:
            data = xlrd.open_workbook(filename)
            table = data.sheets()[0]
            result = []
            # 讀 excel 檔
            for row_index in range(1, table.nrows):
                tmp=[]
                tmp.append(self.ReplaceField(str(table.cell(row_index, 0).value),'.'))
                tmp.append(self.ReplaceField(table.cell(row_index, 1).value,'(',1))
                tmp.append(table.cell(row_index, 2).value)
                tmp.append('0' + self.ReplaceField(str(table.cell(row_index, 3).value),'.'))
                tmp.append(table.cell(row_index, 5).value)
                tmp.append(self.ReplaceField(table.cell(row_index, 6).value,u'盒'))
                tmp.append(table.cell(row_index, 7).value)
                result.append(tmp)
            self.writeT_catXls(result,output)
        except Exception as e :
            logging.error(e.message)
            return 'failure'

    #寫入黑貓出貨單內容
    def writeT_catXls(self,data,output):
        try:
            outputFile = ExcelTemplate()
            fileTemplate =  outputFile.T_Cat_OutputFile
            rb = xlrd.open_workbook(fileTemplate)
            file = copy(rb)
            table = file.get_sheet(0)
            i = 1
            for row in data:
                d1 = datetime.datetime.strftime(datetime.date.today(),'%Y/%m/%d')
                d2 = datetime.datetime.strftime(datetime.date.today()+ datetime.timedelta(days=1), '%Y/%m/%d')
                table.write(i,0,d1)             # 收件日
                table.write(i, 1,d2)            # 配達日
                table.write(i,3,row[0])         # 訂單編號
                table.write(i, 6, row[1])       # 收件人
                table.write(i, 7, row[2])       # 收件地址
                table.write(i, 8, row[3])       # 收件人手機1
                table.write(i, 11,'1')           # 送貨時段
                table.write(i, 12, row[4])      # 訂購品項
                table.write(i, 13, row[6])      # 訂購份數
                table.write(i, 14, int(row[5]))      # 盒數
                table.write(i, 15, int(row[5])*int(row[6]))  # 總數量
                self.customer.setGroup_id(self.GroupID)
                self.customer.setName(row[1])
                self.customer.setMobile(row[3])
                self.customer.setAddress(row[2])
                # insert or update table tb_customer
                self.updateDB_Customer()
                i += 1
            self.mysqlconnect.db.commit()
            self.mysqlconnect.dbClose()
            file.save(output)
        except Exception as e :
            logging.error(e.message)
            return 'failure'

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
    buy.parserXls('robintest','test',u'C:/Users/10408001/Desktop/團購平台訂單資訊/生活市集/原始檔/2016.12.05/2016-12-05_生活市集_BY123375367F_悠活原力有限公司_欣敏立清益生菌-原味多多_未出貨.xls', \
                  u'C:/Users/10408001/Desktop/20161228出貨單.xls')