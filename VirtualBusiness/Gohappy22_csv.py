# -*- coding: utf-8 -*-
#__author__ = '10408001'
import datetime,time
import logging
import csv
from ToMysql import ToMysql
import uuid
from VirtualBusiness import Sale,Customer,updateCustomer
from VirtualBusiness.momo24_csv import Momo24csv_Data
import codecs
import json

logger = logging.getLogger(__name__)

class Gohappy22csv_Data(Momo24csv_Data):
    # Data = None
    # mysqlconnect = None
    # sale , customer = None, None
    # header = []
    # content = []
    #
    # def __init__(self):
    #     # mysql connector object
    #     self.mysqlconnect = ToMysql()
    #     self.mysqlconnect.connect()
    #
    # def readFile(self, _file):
    #     cr = open(_file, 'rb')
    #
    #     del self.header[:]
    #     del self.content[:]
    #
    #     i = 0
    #     for row in cr:
    #         str = row.split(',')
    #
    #         if i == 0:
    #             self.header.append([r for r in str])
    #         else:
    #             # print "content"
    #             temp = [r for r in str]
    #             self.content.append(temp)
    #
    #         i += 1

    def Gohappy_22_Data(self, supplier, GroupID, path, UserID):
        logging.basicConfig(filename='/data/VirtualBusiness_Data/pyupload.log',
                            level=logging.DEBUG,
                            format='%(asctime)s - %(levelname)s - %(filename)s:%(name)s:%(module)s/%(funcName)s/%(lineno)d - %(message)s',
                            datefmt='%Y/%m/%d %I:%M:%S %p')
        logging.Formatter.converter = time.gmtime

        logger.info('===Gohappy22_Data===')
        logger.debug('supplier:' + supplier)
        logger.debug('GroupID:' + GroupID)
        logger.debug('path:' + path)
        logger.debug('UserID:' + UserID)
        try:
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
            logger.debug('===Gohappy22_Data SUCCESS===')
            return json.dumps({"success": success, "info": resultinfo, "total": totalRows}, sort_keys=False)

    def parserData(self,table,row_index,GroupID,UserID,supplier):
        try:
            row = self.content[row_index]

            self.sale.setGroup_id(GroupID)
            self.sale.setUser_id(UserID)
            self.sale.setOrder_source(supplier)
            self.sale.setOrder_No(row[3].lstrip("\\'"))
            self.sale.setTrans_list_date_YMDHMSF(row[0].lstrip("\\'"))
            self.sale.setSale_date_YMDHMSF(row[0].lstrip("\\'"))
            self.sale.setC_Product_id(row[21].lstrip("\\'").split('\n')[0])
            self.sale.setProduct_name_NoEncode(row[5].lstrip("\\'"))
            self.sale.setProduct_spec('')
            self.sale.setQuantity(row[6][row[6].find("("):].strip("()").lstrip("\\'"))
            self.sale.setPrice(row[6][:row[6].find("(")].strip("'").lstrip("\\'"))
            self.sale.setNameNoEncode(row[9].lstrip("\\'"))
            self.sale.setDeliveryway('1')   #宅配: 1, 超取711: 2, 超取全家: 3

            self.customer.setGroup_id(GroupID)
            self.customer.setNameNoEncode(row[9].lstrip("\\'"))
            self.customer.setPhone(row[11].lstrip("'").lstrip("\\'"))
            self.customer.setMobile(row[12].lstrip("'").lstrip("\\'"))
            self.customer.setPost(None)
            self.customer.setAddressNoEncode(row[10].lstrip("'"))
        except Exception as e :
            print e.message
            logging.error(e.message)


if __name__ == '__main__':
    gohappy = Gohappy22csv_Data()
    groupid = ""
    groupid='cbcc3138-5603-11e6-a532-000d3a800878'
    print gohappy.Gohappy_22_Data('gohappy',groupid, '/Users/csi/Desktop/for_Joe_test/網購/gohappy/宅配/OrderData_42242.csv','system')
