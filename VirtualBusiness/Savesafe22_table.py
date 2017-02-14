# -*- coding: utf-8 -*-
#__author__ = '10408001'
import time, BeautifulSoup
import logging
from ToMysql import ToMysql
import uuid
from VirtualBusiness import Sale,Customer,updateCustomer,detectFile
import codecs
from VirtualBusiness.Myfone22_table import Myfone22table_Data
import json

logger = logging.getLogger(__name__)

class Savesafe22table_Data(Myfone22table_Data):

    def Savesafe_22_Data(self, supplier, GroupID, path, UserID):
        logging.basicConfig(filename='/data/VirtualBusiness_Data/pyupload.log',
                            level=logging.DEBUG,
                            format='%(asctime)s - %(levelname)s - %(filename)s:%(name)s:%(module)s/%(funcName)s/%(lineno)d - %(message)s',
                            datefmt='%Y/%m/%d %I:%M:%S %p')
        logging.Formatter.converter = time.gmtime

        logger.info('===Savesafe22_Data===')
        logger.debug('supplier:' + supplier)
        logger.debug('GroupID:' + GroupID)
        logger.debug('path:' + path)
        logger.debug('UserID:' + UserID)
        try:
            content = ""
            detect = detectFile()

            if detect.detect(path) == "Big5":
                with codecs.open(path, 'rb',encoding="Big5") as f:
                    for line in f:
                        content += line
            else:
                with codecs.open(path, 'rb',encoding="utf-8") as f:
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

            # for i in range(0, 2):
            #     print dict_list[i][u'訂單編號']
            resultinfo = ''
            totalRows = len(d)
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
            logger.debug('===Savesafe22_Data finally===')
            return json.dumps({"success": success, "info": resultinfo, "total": totalRows}, sort_keys=False)

    def parserData(self,dict_list,row_index,GroupID,UserID,supplier):
        try:
            # row = self.content[row_index]

            # print dict_list[row_index][u'訂單編號']
            self.sale.setGroup_id(GroupID)
            self.sale.setUser_id(UserID)
            self.sale.setOrder_sourceNodecode(supplier)
            self.sale.setOrder_No(dict_list[row_index][u'訂單編號'])
            self.sale.setTrans_list_date_YMDHMS(dict_list[row_index][u'接單時間'])
            self.sale.setSale_date_YMDHMS(dict_list[row_index][u'接單時間'])
            self.sale.setC_Product_id(dict_list[row_index][u'商品貨號'])
            self.sale.setProduct_name_NoEncode(dict_list[row_index][u'商品名稱'])
            self.sale.setProduct_spec('')
            self.sale.setQuantity(dict_list[row_index][u'商品數量'])
            self.sale.setPrice(dict_list[row_index][u'供貨成本(未稅)'])
            self.sale.setNameNoEncode(dict_list[row_index][u'收貨人'])
            self.sale.setDeliveryway('1')   #宅配: 1, 超取711: 2, 超取全家: 3

            self.customer.setGroup_id(GroupID)
            self.customer.setNameNoEncode(dict_list[row_index][u'收貨人'])
            self.customer.setPhone(dict_list[row_index][u'聯絡電話'])
            self.customer.setMobile(dict_list[row_index][u'行動電話'])
            self.customer.setPost(None)
            self.customer.setAddressNoEncode(dict_list[row_index][u'地址'])
        except Exception as e:
            print e.message
            logging.error(e.message)


if __name__ == '__main__':
    savesafe = Savesafe22table_Data()
    # groupid = ""
    groupid='cbcc3138-5603-11e6-a532-000d3a800878'
    print savesafe.Savesafe_22_Data('Savesafe',groupid, '/Users/csi/Documents/6f27e56d-019e-4fa4-8a56-f0585572fb14.html','system')
