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
    def __init__(self):
        pass

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
            # mysql connector object
            mysqlconnect = ToMysql()
            mysqlconnect.connect()
            updatecustomer = updateCustomer()

            data = xlrd.open_workbook(path)
            table = data.sheets()[0]
            for row_index in range(2, table.nrows):
                sale = Sale()
                customer = Customer()
                sale.setGroup_id(GroupID)
                sale.setOrder_No(table.cell(row_index, 1).value)
                sale.setUser_id(UserID)
                sale.setTrans_list_date(table.cell(row_index, 2).value)
                sale.setSale_date(table.cell(row_index, 3).value)
                sale.setC_Product_id(str(table.cell(row_index, 13).value).split('.')[0])
                sale.setProduct_name(table.cell(row_index, 14).value)
                sale.setQuantity(table.cell(row_index, 18).value)
                sale.setPrice(table.cell(row_index, 20).value)
                sale.setName(table.cell(row_index, 6).value)
                sale.setOrder_source(supplier)

                customer.setGroup_id(GroupID)
                customer.setName(table.cell(row_index, 6).value)
                customer.setPhone(table.cell(row_index, 7).value)
                customer.setMobile(table.cell(row_index, 9).value)
                customer.setPost(table.cell(row_index, 10).value)
                customer.setAddress(table.cell(row_index, 11).value)
                supplier = supplier

                # insert or update table tb_customer
                customer.setCustomer_id(
                    updatecustomer.checkCustomerid(customer.getGroup_id(), customer.getName(), customer.getAddress(),\
                                         customer.getphone(),customer.getMobile(), customer.getEmail()))
                if customer.getCustomer_id() == None:
                    customer.setCustomer_id(uuid.uuid4())
                    CustomereSQL = (
                    customer.getCustomer_id(), customer.getGroup_id(), customer.getName(),\
                    customer.getAddress(),customer.getphone(), customer.getMobile(),\
                    customer.getEmail(),customer.getPost(), customer.getClass(), customer.getMemo(), UserID)
                    mysqlconnect.cursor.callproc('sp_insert_customer_bysys', CustomereSQL)
                else:
                    CustomereSQL = (customer.getCustomer_id(), customer.getGroup_id(), customer.getName(),\
                                    customer.getAddress(),customer.getphone(), customer.getMobile(),\
                                    customer.getEmail(),customer.getPost(), customer.getClass(),\
                                    customer.getMemo(), UserID)
                    mysqlconnect.cursor.callproc('sp_update_customer', CustomereSQL)

                CustomereSQL = (customer.getCustomer_id(), customer.getGroup_id(), customer.get_Name(),\
                                customer.get_Address(),customer.get_phone(), customer.get_Mobile(),\
                                customer.get_Email())
                updatecustomer.updataData(CustomereSQL)

                # insert table tb_sale
                SaleSQL = (sale.getGroup_id(), sale.getOrder_No(), sale.getUser_id(), sale.getProduct_name(), \
                           sale.getC_Product_id(), customer.getCustomer_id(), sale.getName(), sale.getQuantity(), \
                           sale.getPrice(), sale.getInvoice(), sale.getInvoice_date(), sale.getTrans_list_date(), \
                           sale.getDis_date(), sale.getMemo(), sale.getSale_date(), sale.getOrder_source())
                print SaleSQL
                mysqlconnect.cursor.callproc('p_tb_sale', SaleSQL)
                mysqlconnect.db.commit()
                sale = None
                customer = None

            mysqlconnect.dbClose()
            logging.info('===Yahood_Data SUCCESS===')
            return 'success'
        except Exception as e :
            logging.error(e.message)
            return 'failure'

if __name__ == '__main__':
    yahoo = Yahoo_Data()
    groupid='cbcc3138-5603-11e6-a532-000d3a800878'
    print yahoo.Yahood_Data('yahoo',groupid,u'C:\\Users\\10408001\\Desktop\\delivery-Y購-new-4.xls','system')
    # print yahoo.checkCustomerid('data_09221433(test).xlsx','鍾妮',\
    #                       '111台北市士林區中山北路六段77號','02-24609497','0966056315',None)