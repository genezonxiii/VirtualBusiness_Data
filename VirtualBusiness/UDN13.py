# -*-  coding: utf-8  -*-
__author__ = '10409003'
# import xlrd
# import uuid
# from datetime import datetime
# from aes_data import aes_data
# from ToMongodb import ToMongodb
# from ToMysql import ToMysql
# import logging
# import time
import datetime,time
import logging
import json
import xlrd
from ToMysql import ToMysql
import uuid
from VirtualBusiness import Sale,Customer,updateCustomer


logger = logging.getLogger(__name__)

class UDN_Data():
    Data = None
    mysqlconnect = None
    sale, customer = None, None

    # 預期要找出欄位的索引位置的欄位名稱
    TitleTuple = (u'訂單編號', u'訂購日期', u'商品編號', u'商科料號', u'商品型號',
                  u'國際條碼', u'商品名稱+規格尺寸', u'訂購數量', u'原售價', u'原售價-小計',
                  u'進貨價', u'進貨價-小計', u'廠商回押時間')
    TitleList = []

    def __init__(self):
        # mysql connector object
        self.mysqlconnect = ToMysql()
        self.mysqlconnect.connect()

    def Udn1_Data(self, supplier, GroupID, path, UserID):
        """
        檔案格式 xls，欄位數為13，缺少收件人資訊


        """


        try:

            logger.debug("===Udn1_Data===")

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
                    logger.debug(u'index in file - ' + str(self.TitleList.index(self.TitleTuple[index])))
                    # print str(index) + TitleTuple[index]
                    # print (TitleList.index(TitleTuple[index]))

            for row_index in range(2, table.nrows):
                self.sale = Sale()
                self.customer = Customer()
                # Parser Data from xls
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
            logger.debug('===Udn_Data finally===')
            return json.dumps({"success": success, "info": resultinfo, "total": totalRows}, sort_keys=False)

    def parserData(self, table, row_index, GroupID, UserID, supplier):
        try:
            self.sale.setGroup_id(GroupID)
            self.sale.setUser_id(UserID)
            self.sale.setOrder_source(supplier)
            self.sale.setOrder_No(table.cell(row_index, self.TitleList.index(self.TitleTuple[0])).value)
            self.sale.setTrans_list_date(None)
            self.sale.setSale_date(table.cell(row_index, self.TitleList.index(self.TitleTuple[1])).value)
            self.sale.setC_Product_id(
                str(table.cell(row_index, self.TitleList.index(self.TitleTuple[2])).value).split('.')[0])
            self.sale.setProduct_name(table.cell(row_index, self.TitleList.index(self.TitleTuple[6])).value)
            self.sale.setQuantity(table.cell(row_index, self.TitleList.index(self.TitleTuple[7])).value)
            self.sale.setPrice(table.cell(row_index, self.TitleList.index(self.TitleTuple[8])).value)
            self.sale.setName("")

            self.customer.setGroup_id(GroupID)
            self.customer.setName("")
            self.customer.setPhone(None)
            self.customer.setMobile(None)
            self.customer.setPost(None)
            self.customer.setAddress(None)
        except Exception as e:
            print e.message
            logging.error(e.message)

    def updateDB_Customer(self):
        try:
            # insert or update table tb_customer
            updatecustomer = updateCustomer()
            self.customer.setCustomer_id(
                updatecustomer.checkCustomerid(self.customer.getGroup_id(), self.customer.get_Name(),
                                               self.customer.get_Address(), \
                                               self.customer.get_phone(), self.customer.get_Mobile(),
                                               self.customer.get_Email()))

            if self.customer.getCustomer_id() == None:
                self.customer.setCustomer_id(uuid.uuid4())
                CustomereSQL = (
                    self.customer.getCustomer_id(), self.customer.getGroup_id(), self.customer.getName(), \
                    self.customer.getAddress(), self.customer.getphone(), self.customer.getMobile(), \
                    self.customer.getEmail(), self.customer.getPost(), self.customer.getClass(),
                    self.customer.getMemo(), self.sale.getUser_id())
                self.mysqlconnect.cursor.callproc('sp_insert_customer_bysys', CustomereSQL)
            else:
                CustomereSQL = (self.customer.getCustomer_id(), self.customer.getGroup_id(), self.customer.getName(), \
                                self.customer.getAddress(), self.customer.getphone(), self.customer.getMobile(), \
                                self.customer.getEmail(), self.customer.getPost(), self.customer.getClass(), \
                                self.customer.getMemo(), self.sale.getUser_id())
                self.mysqlconnect.cursor.callproc('sp_update_customer', CustomereSQL)

            CustomereSQL = (self.customer.getCustomer_id(), self.customer.getGroup_id(), self.customer.get_Name(), \
                            self.customer.get_Address(), self.customer.get_phone(), self.customer.get_Mobile(), \
                            self.customer.get_Email())
            updatecustomer.updataData(CustomereSQL)
        except Exception as e:
            print e.message
            logging.error(e.message)
            raise

    def updateDB_Sale(self):
        try:
            SaleSQL = (
            self.sale.getGroup_id(), self.sale.getOrder_No(), self.sale.getUser_id(), self.sale.getProduct_name(), \
            self.sale.getC_Product_id(), self.customer.getCustomer_id(), self.sale.getName(), self.sale.getQuantity(), \
            self.sale.getPrice(), self.sale.getInvoice(), self.sale.getInvoice_date(), self.sale.getTrans_list_date(), \
            self.sale.getDis_date(), self.sale.getMemo(), self.sale.getSale_date(), self.sale.getOrder_source())
            self.mysqlconnect.cursor.callproc('p_tb_sale', SaleSQL)
            return
        except Exception as e:
            print e.message
            logging.error(e.message)
            raise


if __name__ == '__main__':
    udn = UDN_Data()

    groupid = 'cbcc3138-5603-11e6-a532-000d3a800878'
    print udn.Udn1_Data('udn', groupid, u'C:\\Users\\10507001\\Downloads\\Order_20160608104154204.xls', 'system')
    # print yahoo.checkCustomerid('data_09221433(test).xlsx','鍾妮',\
    #                       '111台北市士林區中山北路六段77號','02-24609497','0966056315',None)