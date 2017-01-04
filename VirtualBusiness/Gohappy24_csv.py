# -*- coding: utf-8 -*-
#__author__ = '10408001'
import datetime,time
import logging
import csv
from ToMysql import ToMysql
import uuid
from VirtualBusiness import Sale,Customer,updateCustomer

logger = logging.getLogger(__name__)

class Gohappy23_Data():
    Data = None
    mysqlconnect = None
    sale , customer = None, None
    header = []
    content = []

    def __init__(self):
        # mysql connector object
        self.mysqlconnect = ToMysql()
        self.mysqlconnect.connect()

    def readFile(self, _file):
        cr = open(_file, 'rb')

        i = 0
        for row in cr:
            str = row.split(',')

            if i == 0:
                self.header.append([r for r in str])
            else:
                # print "content"
                temp = [r for r in str]
                self.content.append(temp)

            i += 1

    def Gohappy_23_Data(self, supplier, GroupID, path, UserID):
        logging.basicConfig(filename='/data/VirtualBusiness_Data/pyupload.log',
                            level=logging.DEBUG,
                            format='%(asctime)s - %(levelname)s - %(filename)s:%(name)s:%(module)s/%(funcName)s/%(lineno)d - %(message)s',
                            datefmt='%Y/%m/%d %I:%M:%S %p')
        logging.Formatter.converter = time.gmtime

        logger.info('===Gohappy23_Data===')
        logger.debug('supplier:' + supplier)
        logger.debug('GroupID:' + GroupID)
        logger.debug('path:' + path)
        logger.debug('UserID:' + UserID)
        try:
            self.readFile(path)
            logger.debug("header:")
            logger.debug(self.header)
            print len(self.content)

            for row_index in range(0, len(self.content)):
                self.sale = Sale()
                self.customer = Customer()
                #Parser Data from xls
                self.parserData(self.content, row_index, GroupID, UserID, supplier)
                # insert or update table tb_customer
                self.updateDB_Customer()
                # insert table tb_sale
                self.updateDB_Sale()
                self.sale = None
                self.customer = None
            self.mysqlconnect.db.commit()
            self.mysqlconnect.dbClose()
            logger.info('===Gohappy23_Data SUCCESS===')
            return 'success'
        except Exception as e :
            logger.error(e.message)
            return 'failure'

    def parserData(self,table,row_index,GroupID,UserID,supplier):
        try:
            row = self.content[row_index]

            self.sale.setGroup_id(GroupID)
            self.sale.setUser_id(UserID)
            self.sale.setOrder_source(supplier)
            self.sale.setOrder_No(row[3].lstrip("\\'"))
            self.sale.setTrans_list_date_YMDHMSF(row[0].lstrip("\\'"))
            self.sale.setSale_date_YMDHMSF(row[0].lstrip("\\'"))
            self.sale.setC_Product_id(row[23].lstrip("\\'").split('\n')[0])
            self.sale.setProduct_name(row[5].decode('big5').encode('utf-8').lstrip("\\'"))
            self.sale.setQuantity(row[6][row[6].find("("):].strip("()"))
            self.sale.setPrice(row[6][:row[6].find("(")].strip("'").lstrip("\\'"))
            self.sale.setName(row[9].decode('big5').encode('utf-8').lstrip("\\'"))

            self.customer.setGroup_id(GroupID)
            self.customer.setName(row[9].decode('big5').encode('utf-8').lstrip("\\'"))
            self.customer.setPhone(row[11].lstrip("'").lstrip("\\'"))
            self.customer.setMobile(row[12].lstrip("'").lstrip("\\'"))
            self.customer.setPost(None)
            self.customer.setAddress(row[10].decode('big5').encode('utf-8').lstrip("'"))
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
    gohappy = Gohappy23_Data()
    groupid = ""
    groupid='cbcc3138-5603-11e6-a532-000d3a800878'
    print gohappy.Gohappy_23_Data('gohappy',groupid, u'C:\\Users\\10509002\\Documents\\電商檔案\\網購平台訂單資訊\\GoHappy\\2016.04.08\\OrderData_42242.csv','system')
