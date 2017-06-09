# -*- coding: utf-8 -*-
#__author__ = '10408001'
import time, BeautifulSoup
import logging
from ToMysql import ToMysql
import uuid
from VirtualBusiness import Sale,Customer,updateCustomer
from VirtualBusiness.Myfone22_table import Myfone22table_Data
import json

logger = logging.getLogger(__name__)

class YahooS24table_Data(Myfone22table_Data):

    def YahooS_24_Data(self, supplier, GroupID, path, UserID):
        logging.basicConfig(filename='/data/VirtualBusiness_Data/pyupload.log',
                            level=logging.DEBUG,
                            format='%(asctime)s - %(levelname)s - %(filename)s:%(name)s:%(module)s/%(funcName)s/%(lineno)d - %(message)s',
                            datefmt='%Y/%m/%d %I:%M:%S %p')
        logging.Formatter.converter = time.gmtime

        logger.info('===YahooS24_Data===')
        logger.debug('supplier:' + supplier)
        logger.debug('GroupID:' + GroupID)
        logger.debug('path:' + path)
        logger.debug('UserID:' + UserID)
        try:
            self.dup_order_no = []
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
            resultinfo = ''
            totalRows = len(dict_list)
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

            success = True

        except Exception as inst:
            logger.error(inst.args)
            resultinfo = inst.args

        finally:
            dup_str = ','.join(self.dup_order_no)
            self.dup_order_no = []
            logger.debug('===YahooS24_Data finally===')
            return json.dumps({"success": success, "info": resultinfo, "duplicate": dup_str, "total": totalRows}, sort_keys=False)

    def parserData(self,dict_list,row_index,GroupID,UserID,supplier):
        try:

            self.sale.setGroup_id(GroupID)
            self.sale.setUser_id(UserID)
            self.sale.setOrder_source(supplier)
            self.sale.setOrder_No(dict_list[row_index][u'訂單編號'])
            self.sale.setTrans_list_date(dict_list[row_index][u'轉單日'])
            self.sale.setSale_date(dict_list[row_index][u'轉單日'])
            self.sale.setC_Product_id(dict_list[row_index][u'商品編號'])
            # print dict_list[row_index][u'商品名稱']
            self.sale.setProduct_name_NoEncode(dict_list[row_index][u'商品名稱'])
            self.sale.setProduct_spec(dict_list[row_index][u'商品規格'])
            self.sale.setQuantity(dict_list[row_index][u'數量'])
            self.sale.setPrice_str(dict_list[row_index][u'金額小計'])
            self.sale.setNameNoEncode(dict_list[row_index][u'收件人姓名'])
            self.sale.setDeliveryway('3') #宅配: 1, 超取711: 2, 超取全家: 3
            self.sale.setOrder_status('A0')

            self.customer.setGroup_id(GroupID)
            self.customer.setNameNoEncode(dict_list[row_index][u'收件人姓名'])
            self.customer.setPhone(dict_list[row_index][u'收件人電話'])
            self.customer.setMobile(dict_list[row_index][u'收件人行動電話'])
            self.customer.setPost(None)
            self.customer.setAddressNoEncode(dict_list[row_index][u'收件人地址'])
        except Exception as e:
            print e.message
            logging.error(e.message)


if __name__ == '__main__':
    yahooS = YahooS24table_Data()
    # groupid = ""
    groupid='cbcc3138-5603-11e6-a532-000d3a800878'
    print yahooS.YahooS_24_Data('yahoomall',groupid, 'C:/Users/10509002/Desktop/storders.xls','system')
