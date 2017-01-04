# -*- coding: utf-8 -*-
#__author__ = '10408001'
import datetime,time
import logging
import json
import xlrd
from ToMysql import ToMysql
import uuid
from VirtualBusiness import Sale,Customer,updateCustomer

logger = logging.getLogger(__name__)

class Yahoo_Data():
    Data = None
    mysqlconnect = None
    sale , customer = None, None

    # 預期要找出欄位的索引位置的欄位名稱
    TitleTuple = (u'訂單編號', u'轉單日', u'應出貨日', u'出貨日期', u'最晚出貨日',
                  u'收件人姓名', u'收件人電話(日)', u'收件人電話(夜)', u'收件人手機', u'收件人郵遞區號',
                  u'收件人地址', u'供應商料號', u'商品名稱', u'數量', u'商品成本')
    TitleList = []

    def __init__(self):
        # mysql connector object
        self.mysqlconnect = ToMysql()
        self.mysqlconnect.connect()

    def Yahood_Data(self, supplier, GroupID, path, UserID):

        try:

            logger.debug("===Yahood_Data===")

            success = False
            resultinfo = ""
            totalRows = 0

            data = xlrd.open_workbook(path)
            table = data.sheets()[0]

            totalRows = table.nrows - 2

            # 存放excel中全部的欄位名稱
            self.TitleList = []
            for row_index in range(1, 2):
                for col_index in range(0, table.ncols):
                    self.TitleList.append(table.cell(row_index, col_index).value)

            print self.TitleList

            # 存放excel中對應TitleTuple欄位名稱的index
            for index in range(0, len(self.TitleTuple)):
                if self.TitleTuple[index] in self.TitleList:
                    logger.debug(str(index) + self.TitleTuple[index])
                    logger.debug(u'index in file - ' + str(self.TitleList.index(self.TitleTuple[index])) )
                    # print str(index) + TitleTuple[index]
                    # print (TitleList.index(TitleTuple[index]))

            for row_index in range(2, table.nrows):
                self.sale = Sale()
                self.customer = Customer()
                #Parser Data from xls
                self.parserData(table, row_index, GroupID, UserID, supplier)
                # insert or update table tb_customer
                self.updateDB_Customer()
                # insert table tb_sale
                self.updateDB_Sale()
                self.sale = None
                self.customer = None
            self.mysqlconnect.db.commit()
            self.mysqlconnect.dbClose()

            success = True

        except Exception as inst:
            logger.error(inst.args)
            resultinfo = inst.args
        finally:
            logger.debug('===Yahoo_Data finally===')
            return json.dumps({"success": success, "info": resultinfo, "total": totalRows}, sort_keys=False)

    def parserData(self,table,row_index,GroupID,UserID,supplier):
        try:
            self.sale.setGroup_id(GroupID)
            self.sale.setUser_id(UserID)
            self.sale.setOrder_source(supplier)
            self.sale.setOrder_No(table.cell(row_index, self.TitleList.index(self.TitleTuple[0])).value)
            self.sale.setTrans_list_date(table.cell(row_index, self.TitleList.index(self.TitleTuple[1])).value)
            self.sale.setSale_date(table.cell(row_index, self.TitleList.index(self.TitleTuple[1])).value)
            self.sale.setC_Product_id(str(table.cell(row_index, self.TitleList.index(self.TitleTuple[11])).value).split('.')[0])
            self.sale.setProduct_name(table.cell(row_index, self.TitleList.index(self.TitleTuple[12])).value)
            self.sale.setQuantity(table.cell(row_index, self.TitleList.index(self.TitleTuple[13])).value)
            self.sale.setPrice(table.cell(row_index, self.TitleList.index(self.TitleTuple[14])).value)
            self.sale.setName(table.cell(row_index, self.TitleList.index(self.TitleTuple[5])).value)

            self.customer.setGroup_id(GroupID)
            self.customer.setName(table.cell(row_index, self.TitleList.index(self.TitleTuple[5])).value)
            self.customer.setPhone(table.cell(row_index, self.TitleList.index(self.TitleTuple[6])).value)
            self.customer.setMobile(table.cell(row_index, self.TitleList.index(self.TitleTuple[8])).value)
            self.customer.setPost(table.cell(row_index, self.TitleList.index(self.TitleTuple[9])).value)
            self.customer.setAddress(table.cell(row_index, self.TitleList.index(self.TitleTuple[10])).value)
        except Exception as e :
            print e.message
            logging.error(e.message)

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
                    self.customer.getEmail(), self.customer.getPost(), self.customer.getClass(), self.customer.getMemo(), self.sale.getUser_id())
                self.mysqlconnect.cursor.callproc('sp_insert_customer_bysys', CustomereSQL)
            else:
                CustomereSQL = (self.customer.getCustomer_id(), self.customer.getGroup_id(), self.customer.getName(), \
                                self.customer.getAddress(), self.customer.getphone(), self.customer.getMobile(), \
                                self.customer.getEmail(), self.customer.getPost(), self.customer.getClass(), \
                                self.customer.getMemo(),self.sale.getUser_id())
                self.mysqlconnect.cursor.callproc('sp_update_customer', CustomereSQL)

            CustomereSQL = (self.customer.getCustomer_id(), self.customer.getGroup_id(), self.customer.get_Name(), \
                            self.customer.get_Address(), self.customer.get_phone(), self.customer.get_Mobile(), \
                            self.customer.get_Email())
            updatecustomer.updataData(CustomereSQL)
        except Exception as e :
            print e.message
            logging.error(e.message)
            raise

    def updateDB_Sale(self):
        try:
            SaleSQL = (self.sale.getGroup_id(), self.sale.getOrder_No(), self.sale.getUser_id(), self.sale.getProduct_name(), \
                       self.sale.getC_Product_id(), self.customer.getCustomer_id(), self.sale.getName(), self.sale.getQuantity(), \
                       self.sale.getPrice(), self.sale.getInvoice(), self.sale.getInvoice_date(), self.sale.getTrans_list_date(), \
                       self.sale.getDis_date(), self.sale.getMemo(), self.sale.getSale_date(), self.sale.getOrder_source())
            self.mysqlconnect.cursor.callproc('p_tb_sale', SaleSQL)
            return
        except Exception as e :
            print e.message
            logging.error(e.message)
            raise

if __name__ == '__main__':
    yahoo = Yahoo_Data()
    groupid = ""
    groupid='cbcc3138-5603-11e6-a532-000d3a800878'
    print yahoo.Yahood_Data('yahoo',groupid,u'C:\\Users\\10509002\\Desktop\\新增資料夾 (2)\\delivery-Y購-new-4.xls','system')
    # print yahoo.checkCustomerid('data_09221433(test).xlsx','鍾妮',\
    #                       '111台北市士林區中山北路六段77號','02-24609497','0966056315',None)