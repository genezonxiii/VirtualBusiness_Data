# -*- coding: utf-8 -*-
#__author__ = '10408001'
import datetime,time
import logging
import xlrd
from ToMysql import ToMysql
import uuid
from VirtualBusiness import Sale ,Customer,updateCustomer

class Yahoo_Data():
    Data = None
    mysqlconnect = None
    sale , customer =None, None
    def __init__(self):
        # mysql connector object
        self.mysqlconnect = ToMysql()
        self.mysqlconnect.connect()

    def Yahood_Data(self, supplier, GroupID, path, UserID):
        logging.basicConfig(filename='pyupload.log', level=logging.DEBUG, format='%(asctime)s %(message)s',
                            datefmt='%Y/%m/%d %I:%M:%S %p')
        logging.Formatter.converter = time.gmtime
        logging.info('===Yahood_Data===')
        logging.debug('supplier:' + supplier)
        logging.debug('GroupID:' + GroupID)
        logging.debug('path:' + path)
        logging.debug('UserID:' + UserID)
        try:
            data = xlrd.open_workbook(path)
            table = data.sheets()[0]
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
            logging.info('===Yahood_Data SUCCESS===')
            return 'success'
        except Exception as e :
            logging.error(e.message)
            return 'failure'

    def parserData(self,table,row_index,GroupID,UserID,supplier):
        try:
            self.sale.setGroup_id(GroupID)
            self.sale.setOrder_No(table.cell(row_index, 1).value)
            self.sale.setUser_id(UserID)
            self.sale.setTrans_list_date(table.cell(row_index, 2).value)
            self.sale.setSale_date(table.cell(row_index, 3).value)
            self.sale.setC_Product_id(str(table.cell(row_index, 13).value).split('.')[0])
            self.sale.setProduct_name(table.cell(row_index, 14).value)
            self.sale.setQuantity(table.cell(row_index, 18).value)
            self.sale.setPrice(table.cell(row_index, 20).value)
            self.sale.setName(table.cell(row_index, 6).value)
            self.sale.setOrder_source(supplier)

            self.customer.setGroup_id(GroupID)
            self.customer.setName(table.cell(row_index, 6).value)
            self.customer.setPhone(table.cell(row_index, 7).value)
            self.customer.setMobile(table.cell(row_index, 9).value)
            self.customer.setPost(table.cell(row_index, 10).value)
            self.customer.setAddress(table.cell(row_index, 11).value)
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
    groupid='cbcc3138-5603-11e6-a532-000d3a800878'
    print yahoo.Yahood_Data('yahoo',groupid,u'C:\\Users\\10408001\\Desktop\\delivery-Y購-new-4.xls','system')
    # print yahoo.checkCustomerid('data_09221433(test).xlsx','鍾妮',\
    #                       '111台北市士林區中山北路六段77號','02-24609497','0966056315',None)