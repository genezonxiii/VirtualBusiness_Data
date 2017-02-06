# -*- coding: utf-8 -*-
#__author__ = '10408001'
import datetime,time
import json
import logging
import csv
from ToMysql import ToMysql
import uuid
from VirtualBusiness import Sale,Customer,updateCustomer
from VirtualBusiness.momo24_csv import Momo24csv_Data

logger = logging.getLogger(__name__)

class Yahoo22csv_Data(Momo24csv_Data):

    def Yahoo_22_Data(self, supplier, GroupID, path, UserID):
        logging.basicConfig(filename='/data/VirtualBusiness_Data/pyupload.log',
                            level=logging.DEBUG,
                            format='%(asctime)s - %(levelname)s - %(filename)s:%(name)s:%(module)s/%(funcName)s/%(lineno)d - %(message)s',
                            datefmt='%Y/%m/%d %I:%M:%S %p')
        logging.Formatter.converter = time.gmtime

        logger.info('===Yahoo22_Data===')
        logger.debug('supplier:' + supplier)
        logger.debug('GroupID:' + GroupID)
        logger.debug('path:' + path)
        logger.debug('UserID:' + UserID)
        try:
            self.readFile(path)
            logger.debug("header:")
            logger.debug(self.header)
            print len(self.content)
            success = False
            resultinfo = ""


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
            totalRows = len(self.content)

        except Exception as inst:
            logger.error(inst.args)
            resultinfo = inst.args
        finally:
            logger.debug('===Yahoo22_Data finally===')
            return json.dumps({"success": success, "info": resultinfo, "total": totalRows}, sort_keys=False)

    def parserData(self,table,row_index,GroupID,UserID,supplier):
        try:
            row = self.content[row_index]

            self.sale.setGroup_id(GroupID)
            self.sale.setUser_id(UserID)
            self.sale.setOrder_source(supplier)
            self.sale.setOrder_No(row[0])
            self.sale.setTrans_list_date_YMDHM(row[1])
            self.sale.setSale_date_YMDHM(row[1])
            self.sale.setC_Product_id(row[13])
            self.sale.setProduct_name_NoEncode(row[12])
            self.sale.setQuantity(row[15])
            self.sale.setPrice(row[16])
            self.sale.setNameNoEncode(row[4])
            self.sale.setDeliveryway('2')   #宅配: 1, 超取711: 2, 超取全家: 3

            self.customer.setGroup_id(GroupID)
            self.customer.setNameNoEncode(row[4])
            self.customer.setPhone(row[5])
            self.customer.setMobile(row[6])
            self.customer.setPost(None)
            self.customer.setAddressNoEncode(None)
        except Exception as e :
            print e.message
            logging.error(e.message)


if __name__ == '__main__':
    buy = Yahoo22csv_Data()
    groupid = ""
    groupid='cbcc3138-5603-11e6-a532-000d3a800878'
    print buy.Yahoo_22_Data('yahoo',groupid, u'C:\\Users\\10509002\\Desktop\\新增資料夾 (2)\\0407\\yahoo 購物中心\\spstorders.csv','system')
