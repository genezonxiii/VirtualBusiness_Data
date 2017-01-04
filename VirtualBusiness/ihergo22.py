# -*- coding: utf-8 -*-
#__author__ = '10409003'

import logging
import json
import xlrd
from ToMysql import ToMysql
import uuid
from VirtualBusiness import Sale,Customer,updateCustomer

logger = logging.getLogger(__name__)

class Ihergo22_Data():
    Data = None
    mysqlconnect = None
    sale , customer = None, None

    # 預期要找出欄位的索引位置的欄位名稱
    TitleTuple = (u'訂單編號', u'收件人', u'收件人電話', u'收件人手機',u'郵遞區號',
                  u'商品編號', u'商品名稱', u'單價(含稅)', u'數量',u'小計(含稅)',
                  u'總金額(含稅)', u'到貨日期',u'收貨地址')
    TitleList = []

    def __init__(self):
        # mysql connector object
        self.mysqlconnect = ToMysql()
        self.mysqlconnect.connect()

    def Ihergo_22_Data(self, supplier, GroupID, path, UserID):

        try:

            logger.debug("===Ihergo22_Data===")

            success = False
            resultinfo = ""
            totalRows = 0

            data = xlrd.open_workbook(path)
            table = data.sheets()[0]

            totalRows = table.nrows - 2

            # 存放excel中全部的欄位名稱
            self.TitleList = []
            for row_index in range(3, 4):
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

            for row_index in range(4, table.nrows-1):
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
            logger.debug('===Ihergo22_Data finally===')
            return json.dumps({"success": success, "info": resultinfo, "total": totalRows}, sort_keys=False)

    def parserData(self,table,row_index,GroupID,UserID,supplier):
        try:
            self.sale.setGroup_id(GroupID)
            self.sale.setUser_id(UserID)
            self.sale.setOrder_source(supplier)
            self.sale.setOrder_No_float(table.cell(row_index, self.TitleList.index(self.TitleTuple[0])).value)
            self.sale.setTrans_list_date_YYYYMMDD_float(table.cell(row_index, self.TitleList.index(self.TitleTuple[11])).value)
            self.sale.setSale_date_YYYYMMDD_float(table.cell(row_index, self.TitleList.index(self.TitleTuple[11])).value)
            self.sale.setC_Product_id_float(table.cell(row_index, self.TitleList.index(self.TitleTuple[5])).value)
            self.sale.setProduct_name_NoEncode(table.cell(row_index, self.TitleList.index(self.TitleTuple[6])).value)
            self.sale.setQuantity(table.cell(row_index, self.TitleList.index(self.TitleTuple[8])).value)
            self.sale.setPrice(str(table.cell(row_index, self.TitleList.index(self.TitleTuple[10])).value).split('.')[0])
            self.sale.setNameNoEncode(table.cell(row_index, self.TitleList.index(self.TitleTuple[1])).value)

            self.customer.setGroup_id(GroupID)
            self.customer.setNameNoEncode(table.cell(row_index, self.TitleList.index(self.TitleTuple[1])).value)
            self.customer.setPhone(table.cell(row_index, self.TitleList.index(self.TitleTuple[2])).value)
            self.customer.setMobile(table.cell(row_index, self.TitleList.index(self.TitleTuple[2])).value)
            self.customer.setPost(table.cell(row_index, self.TitleList.index(self.TitleTuple[4])).value)
            self.customer.setAddressNoEncode(table.cell(row_index, self.TitleList.index(self.TitleTuple[12])).value)
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
    ihergo =Ihergo22_Data()
    groupid = ""
    groupid='cbcc3138-5603-11e6-a532-000d3a800878'
    print ihergo.Ihergo_22_Data('ihergo',groupid,u'C:\\Users\\10509002\\Documents\\電商檔案\\網購平台訂單資訊\\愛合購\\原始檔\\2014\\2014.07.11\\ihergo_861683_1405058433202.xls','system')