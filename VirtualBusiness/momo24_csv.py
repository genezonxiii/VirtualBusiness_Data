# -*- coding: utf-8 -*-
#__author__ = '10408001'
import datetime,time
import logging
import xlrd
from ToMysql import ToMysql
import uuid
from VirtualBusiness import Sale,Customer,updateCustomer,detectFile
import codecs
import json

logger = logging.getLogger(__name__)

class Momo24csv_Data():
    Data = None
    mysqlconnect = None
    sale , customer = None, None
    header = []
    content = []
    dup_order_no = []

    def __init__(self):
        # mysql connector object
        self.mysqlconnect = ToMysql()
        self.mysqlconnect.connect()

    def readFile(self, _file):
        try:
            cr = None
            detect = detectFile()
            result = detect.detect(_file)
            if result == 'Big5':
                cr = codecs.open(_file, 'rb',encoding='Big5')
            else:
                cr = codecs.open(_file, 'rb',encoding='utf-8')
        except Exception as e:
            logger.error(e.message)

        del self.header[:]
        del self.content[:]

        i = 0
        for row in cr:
            str = row.split(',')

            if i == 0:
                self.header.append([r for r in str])
            else:
                # print "content"
                temp = [r for r in str]
                if temp[0] <> "":
                    self.content.append(temp)

            i += 1



    def Momo_24_Data(self, supplier, GroupID, path, UserID):
        logging.basicConfig(filename='/data/VirtualBusiness_Data/pyupload.log',
                            level=logging.DEBUG,
                            format='%(asctime)s - %(levelname)s - %(filename)s:%(name)s:%(module)s/%(funcName)s/%(lineno)d - %(message)s',
                            datefmt='%Y/%m/%d %I:%M:%S %p')
        logging.Formatter.converter = time.gmtime

        logger.info('===Momo24_Data===')
        logger.debug('supplier:' + supplier)
        logger.debug('GroupID:' + GroupID)
        logger.debug('path:' + path)
        logger.debug('UserID:' + UserID)
        try:
            self.dup_order_no = []
            self.readFile(path)
            logger.debug("header:")
            logger.debug(self.header)
            print len(self.content)
            resultinfo = ""
            totalRows = len(self.content)

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

            success = True
        except Exception as inst:
            logger.error(inst.args)
            resultinfo = inst.args

        finally:
            dup_str = ','.join(self.dup_order_no)
            self.dup_order_no = []
            logger.debug('===Momo24csv_Data finally===')
            return json.dumps({"success": success, "info": resultinfo, "duplicate": dup_str, "total": totalRows}, sort_keys=False)

    def parserData(self,table,row_index,GroupID,UserID,supplier):
        try:
            row = self.content[row_index]

            self.sale.setGroup_id(GroupID)
            self.sale.setUser_id(UserID)
            self.sale.setOrder_source(supplier)
            self.sale.setOrder_No(row[2][0:18])
            self.sale.setTrans_list_date(row[9])
            self.sale.setSale_date(row[9])
            self.sale.setC_Product_id(row[12])
            self.sale.setProduct_name_NoEncode(row[13])
            self.sale.setProduct_spec_NoEncode(row[15])
            self.sale.setQuantity(row[16])
            self.sale.setPrice(row[17])
            self.sale.setNameNoEncode(row[3])
            self.sale.setDeliveryway('1') #宅配: 1, 超取711: 2, 超取全家: 3
            self.sale.setOrder_status('A0')

            self.customer.setGroup_id(GroupID)
            self.customer.setNameNoEncode(row[3])
            self.customer.setPhone(None)
            self.customer.setMobile(None)
            self.customer.setPost(None)
            self.customer.setAddressNoEncode(row[4])
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
            SaleSQL = (self.sale.getGroup_id(), self.sale.getOrder_No(), self.sale.getUser_id(), self.sale.getProduct_name(), self.sale.getProduct_spec(), \
                       self.sale.getC_Product_id(), self.customer.getCustomer_id(), self.sale.getName(), self.sale.getQuantity(), \
                       self.sale.getPrice(), self.sale.getInvoice(), self.sale.getInvoice_date(), self.sale.getTrans_list_date(), \
                       self.sale.getDis_date(), self.sale.getMemo(), self.sale.getSale_date(), self.sale.getOrder_source(),\
                       self.sale.getDeliveryway(), self.sale.getTotal_amt(), self.sale.getOrder_status(), self.sale.getDeliver_name(), \
                       self.sale.getDeliver_to(), self.sale.getDeliver_store(), self.sale.getDeliver_phone(), self.sale.getDeliver_mobile(), \
                       self.sale.getPay_kind(), self.sale.getPay_status(), "")
            result = self.mysqlconnect.cursor.callproc('p_tb_sale_new', SaleSQL)
            if result[27] != None:
                self.dup_order_no.append(result[27])
            return
        except Exception as e :
            print e.message
            logging.error(e.message)
            raise

if __name__ == '__main__':
    momo = Momo24csv_Data()
    # groupid = ""
    groupid='cbcc3138-5603-11e6-a532-000d3a800878'
    print momo.Momo_24_Data('momo',groupid, u'C:/Users/10509002/Desktop/A1102_3_2_008992_20160513102705.csv','system')