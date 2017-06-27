# -*-  coding: utf-8  -*-
#__author__ = '10509002'
import logging
import json
import xlrd
from ToMysql import ToMysql
import uuid
from VirtualBusiness import Sale,Customer,updateCustomer
from VirtualBusiness.Momo25 import Momo25_Data
import re

logger = logging.getLogger(__name__)

class Shopee_Data(Momo25_Data):
    Data = None
    mysqlconnect = None
    sale , customer = None, None


    # 預期要找出欄位的索引位置的欄位名稱
    # 悠活原力
    TitleTuple = (u'訂單編號', u'訂單成立時間', u'商品資訊', u'收件地址',u'收件者姓名',
                  u'電話號碼', u'寄送方式')

    TitleList = []

    def Shopee_26_Data(self, supplier, GroupID, path, UserID):

        try:

            logger.debug("===Shopee_Data===")

            self.dup_order_no = []
            success = False
            resultinfo = ""
            totalRows = 0

            data = xlrd.open_workbook(path)
            table = data.sheets()[0]

            totalRows = table.nrows - 1

            # 存放excel中全部的欄位名稱
            self.TitleList = []
            for row_index in range(0, 1):
                for col_index in range(0, table.ncols):
                    self.TitleList.append(table.cell(row_index, col_index).value)

            print self.TitleList

            # 存放excel中對應TitleTuple欄位名稱的index
            for index in range(0, len(self.TitleTuple)):
                if self.TitleTuple[index] in self.TitleList:
                    logger.debug(str(index) + self.TitleTuple[index])
                    logger.debug(u'index in file - ' + str(self.TitleList.index(self.TitleTuple[index])) )
                    # print str(index) + TitleTuple[index]
                    # print (TitleList.index(TitleTuple[index]))

            for row_index in range(1, table.nrows):
                # self.sale = Sale()
                # self.customer = Customer()
                #Parser Data from xls
                self.parserData(table, row_index, GroupID, UserID, supplier)
                # insert or update table tb_customer
                # self.updateDB_Customer()
                # insert table tb_sale
                # self.updateDB_Sale()
            #     self.sale = None
            #     self.customer = None
            # self.mysqlconnect.db.commit()
            # self.mysqlconnect.dbClose()

            success = True

        except Exception as inst:
            logger.error(inst.args)
            resultinfo = inst.args
        finally:
            dup_str = ','.join(self.dup_order_no)
            self.dup_order_no = []
            logger.debug('===Shopee_Data finally===')
            return json.dumps({"success": success, "info": resultinfo,  "duplicate": dup_str, "total": totalRows}, sort_keys=False)

    def parserData(self,table,row_index,GroupID,UserID,supplier):
        try:
            # 解析商品資訊
            product = table.cell(row_index, self.TitleList.index(self.TitleTuple[2])).value
            product_list = re.split('[[0-9]+]', product)
            count = len(product_list)
            detail = []
            for i in product_list:
                product_list2 = i.split(';')
                detail.append(product_list2)
            # product_list2 = [information.split(';') for information in product_list]

            product_name = []
            price = []
            quantity = []
            c_product_id = []

            # 將商品資訊的細項分開
            for i in range(1, count):
                product_name.append(detail[i][0])
                price.append(detail[i][2])
                quantity.append(detail[i][3])
                c_product_id.append(detail[i][4])

            for j in range(count - 1):
                self.sale = Sale()
                self.customer = Customer()
                product_name2 = product_name[j].split(':')[1]
                price2 = price[j].split('$ ')[1]
                quantity2 = quantity[j].split(': ')[1]
                c_product_id2 = c_product_id[j].split(': ')[1]

                self.sale.setGroup_id(GroupID)
                self.sale.setUser_id(UserID)
                self.sale.setOrder_source(supplier)
                self.sale.setOrder_No(table.cell(row_index, self.TitleList.index(self.TitleTuple[0])).value)
                self.sale.setTrans_list_date_YMDHM(table.cell(row_index, self.TitleList.index(self.TitleTuple[1])).value)
                self.sale.setSale_date_YMDHM(table.cell(row_index, self.TitleList.index(self.TitleTuple[1])).value)
                self.sale.setC_Product_id(c_product_id2)
                self.sale.setProduct_name_NoEncode(product_name2)
                self.sale.setProduct_spec('')
                self.sale.setQuantity(quantity2)
                self.sale.setPrice(price2)
                self.sale.setNameNoEncode(table.cell(row_index, self.TitleList.index(self.TitleTuple[4])).value)
                if table.cell(row_index, self.TitleList.index(self.TitleTuple[6])).value == '7-11':
                    self.sale.setDeliveryway('2')   #宅配: 1, 超取711: 2, 超取全家: 3
                elif table.cell(row_index, self.TitleList.index(self.TitleTuple[6])).value == u'全家':
                    self.sale.setDeliveryway('3')
                else:
                    self.sale.setDeliveryway('1')

                self.sale.setOrder_status('A0')

                self.customer.setGroup_id(GroupID)
                self.customer.setNameNoEncode(table.cell(row_index, self.TitleList.index(self.TitleTuple[4])).value)
                phone_number = table.cell(row_index, self.TitleList.index(self.TitleTuple[5])).value
                self.customer.setPhone('0' + phone_number.lstrip('886'))
                self.customer.setMobile('0' + phone_number.lstrip('886'))
                self.customer.setPost('')
                self.customer.setAddressNoEncode(table.cell(row_index, self.TitleList.index(self.TitleTuple[3])).value)
                # insert or update table tb_customer
                self.updateDB_Customer()
                # insert table tb_sale
                self.updateDB_Sale()
                self.sale = None
                self.customer = None
            self.mysqlconnect.db.commit()
            self.mysqlconnect.dbClose()

        except Exception as e :
            print e.message
            logging.error(e.message)


if __name__ == '__main__':
    udn = Shopee_Data()
    # groupid = ""
    groupid='cbcc3138-5603-11e6-a532-000d3a800878'
    print udn.Shopee_26_Data('udn',groupid,u'C:/Users/10509002/Desktop/yohopower.shopee-order.20170601-20170620.xls','system')
