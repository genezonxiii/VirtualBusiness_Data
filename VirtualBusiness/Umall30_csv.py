# -*- coding: utf-8 -*-
#__author__ = '10408001'
import datetime,time
import logging
import csv
from ToMysql import ToMysql
import uuid
from VirtualBusiness import Sale,Customer,updateCustomer
from VirtualBusiness.momo24_csv import Momo24csv_Data
import json

logger = logging.getLogger(__name__)

class Umall30csv_Data(Momo24csv_Data):

    def Umall_30_Data(self, supplier, GroupID, path, UserID):
        logging.basicConfig(filename='/data/VirtualBusiness_Data/pyupload.log',
                            level=logging.DEBUG,
                            format='%(asctime)s - %(levelname)s - %(filename)s:%(name)s:%(module)s/%(funcName)s/%(lineno)d - %(message)s',
                            datefmt='%Y/%m/%d %I:%M:%S %p')
        logging.Formatter.converter = time.gmtime

        logger.info('===Umall30_Data===')
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
            logger.debug('===Umall30_Data finally===')
            return json.dumps({"success": success, "info": resultinfo, "duplicate": dup_str, "total": totalRows}, sort_keys=False)

    def parserData(self,table,row_index,GroupID,UserID,supplier):
        try:
            row = self.content[row_index]

            self.sale.setGroup_id(GroupID)
            self.sale.setUser_id(UserID)
            self.sale.setOrder_source(supplier)
            self.sale.setOrder_No(row[1])
            self.sale.setTrans_list_date_YYYYMMDD(row[20])
            self.sale.setSale_date_YYYYMMDD(row[20])
            self.sale.setC_Product_id(row[6])
            self.sale.setProduct_name_NoEncode(row[7])
            self.sale.setProduct_spec('')
            self.sale.setQuantity(row[12])
            self.sale.setPrice_str(row[13])
            self.sale.setNameNoEncode(row[14])
            self.sale.setDeliveryway('1')   #宅配: 1, 超取711: 2, 超取全家: 3
            self.sale.setOrder_status('A0')

            self.customer.setGroup_id(GroupID)
            self.customer.setNameNoEncode(row[14])
            self.customer.setPhone(row[16].lstrip("'"))
            self.customer.setMobile(row[15].lstrip("'"))
            self.customer.setPost(None)
            self.customer.setAddressNoEncode(row[17])
        except Exception as e :
            print e.message
            logging.error(e.message)


if __name__ == '__main__':
    umall = Umall30csv_Data()
    groupid = ""
    groupid='cbcc3138-5603-11e6-a532-000d3a800878'
    print umall.Umall_30_Data('umall',groupid, 'C:/Users/10509002/Desktop/cso_export_1405320008748.csv','system')
