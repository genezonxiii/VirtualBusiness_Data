# -*- coding: utf-8 -*-
#__author__ = '10408001'
import time, BeautifulSoup
import logging
from ToMysql import ToMysql
import uuid
from VirtualBusiness import Sale,Customer,updateCustomer
import string

logger = logging.getLogger(__name__)

class YahooS37_Data():
    Data = None
    mysqlconnect = None
    sale , customer = None, None
    header = []
    content = []

    def __init__(self):
        # mysql connector object
        self.mysqlconnect = ToMysql()
        self.mysqlconnect.connect()

    def makelist(self, table):
        result = []
        allrows = table.findAll('tr')
        for row in allrows:
            result.append([])
            allcols = row.findAll('td')
            for col in allcols:
                thestrings = [unicode(s) for s in col.findAll(text=True)]
                thetext = ''.join(thestrings)
                result[-1].append(thetext)
        return result

    def YahooS_37_Data(self, supplier, GroupID, path, UserID):
        logging.basicConfig(filename='/data/VirtualBusiness_Data/pyupload.log',
                            level=logging.DEBUG,
                            format='%(asctime)s - %(levelname)s - %(filename)s:%(name)s:%(module)s/%(funcName)s/%(lineno)d - %(message)s',
                            datefmt='%Y/%m/%d %I:%M:%S %p')
        logging.Formatter.converter = time.gmtime

        logger.info('===YahooS37_Data===')
        logger.debug('supplier:' + supplier)
        logger.debug('GroupID:' + GroupID)
        logger.debug('path:' + path)
        logger.debug('UserID:' + UserID)
        try:
            content = ""
            with open(path, 'rb') as f:
                for line in f:
                    content += line

            table = BeautifulSoup.BeautifulSoup(content)
            list = self.makelist(table)
            rows = iter(list)
            headers = [col for col in next(rows)]

            dict_list = []
            for row in rows:
                values = [col for col in row]
                d = dict(zip(headers, values))

                dict_list.append(d)

            logger.debug(dict_list)

            # for i in range(0, 2):
            #     print dict_list[i][u'訂單編號']

            for row_index in range(0, len(dict_list)):
                self.sale = Sale()
                self.customer = Customer()
                #Parser Data from xls
                self.parserData(dict_list, row_index, GroupID, UserID, supplier)
                # insert or update table tb_customer
                self.updateDB_Customer()
                # insert table tb_sale
                self.updateDB_Sale()
                self.sale = None
                self.customer = None
            self.mysqlconnect.db.commit()
            self.mysqlconnect.dbClose()
            logger.info('===YahooS37_Data SUCCESS===')
            return 'success'
        except Exception as e :
            logger.error(e.message)
            return 'failure'

    def parserData(self,dict_list,row_index,GroupID,UserID,supplier):
        try:
            # row = self.content[row_index]

            # print dict_list[row_index][u'訂單編號']
            self.sale.setGroup_id(GroupID)
            self.sale.setUser_id(UserID)
            self.sale.setOrder_source(supplier)
            self.sale.setOrder_No(dict_list[row_index][u'訂單編號'])
            self.sale.setTrans_list_date(dict_list[row_index][u'轉單日'])
            self.sale.setSale_date(dict_list[row_index][u'轉單日'])
            self.sale.setC_Product_id(dict_list[row_index][u'商品編號'])
            # print dict_list[row_index][u'商品名稱']
            self.sale.setProduct_name_NoEncode(dict_list[row_index][u'商品名稱'])
            self.sale.setQuantity(dict_list[row_index][u'數量'])
            self.sale.setPrice(dict_list[row_index][u'金額小計'])
            self.sale.setNameNoEncode(dict_list[row_index][u'收件人姓名'])

            self.customer.setGroup_id(GroupID)
            self.customer.setNameNoEncode(dict_list[row_index][u'收件人姓名'])
            self.customer.setPhone(dict_list[row_index][u'收件人電話'])
            self.customer.setMobile(dict_list[row_index][u'收件人行動電話'])
            self.customer.setPost(None)
            self.customer.setAddressNoEncode(dict_list[row_index][u'收件人地址'])
        except Exception as e:
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
    yahooS = YahooS37_Data()
    # groupid = ""
    groupid='cbcc3138-5603-11e6-a532-000d3a800878'
    print yahooS.YahooS_37_Data('Yahoomall',groupid, u'C:\\Users\\10509002\\Documents\\電商檔案\\網購平台訂單資訊\\Yahoo商城\\2015.06.18\\storders (2).xls','system')
