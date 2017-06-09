# -*- coding: utf-8 -*-
#__author__ = '10509002'
import datetime,time
import logging
import json
import xlrd
from ToMysql import ToMysql
import uuid
from VirtualBusiness import Sale,Customer,updateCustomer
from VirtualBusiness.Momo25 import Momo25_Data

logger = logging.getLogger(__name__)

class Friday16_Data(Momo25_Data):
    Data = None
    mysqlconnect = None
    sale , customer = None, None

    # 預期要找出欄位的索引位置的欄位名稱
    TitleTuple = (u'訂單編號', u'訂購日', u'應出貨日', u'最晚出貨日', u'收件人',
                  u'收件人手機', u'收件地址', u'(商品編號)\n商品名稱',u'成本', u'數量')
    TitleList = []


    def Friday_16_Data(self, supplier, GroupID, path, UserID):

        try:

            logger.debug("===Friday16_Data===")

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

            success = True

        except Exception as inst:
            logger.error(inst.args)
            resultinfo = inst.args
        finally:
            logger.debug('===Friday16_Data finally===')
            return json.dumps({"success": success, "info": resultinfo, "total": totalRows}, sort_keys=False)

    def parserData(self,table,row_index,GroupID,UserID,supplier):
        try:
            self.sale.setGroup_id(GroupID)
            self.sale.setUser_id(UserID)
            self.sale.setOrder_source(supplier)
            self.sale.setOrder_No(table.cell(row_index, self.TitleList.index(self.TitleTuple[0])).value)
            self.sale.setTrans_list_date_YMD(table.cell(row_index, self.TitleList.index(self.TitleTuple[2])).value)
            self.sale.setSale_date_YMD(table.cell(row_index, self.TitleList.index(self.TitleTuple[1])).value)
            self.sale.setC_Product_id(str(table.cell(row_index, self.TitleList.index(self.TitleTuple[7])).value[1:13]).split('.')[0])
            self.sale.setProduct_spec('')
            self.sale.setProduct_name_NoEncode((table.cell(row_index, self.TitleList.index(self.TitleTuple[7])).value).split('\n')[1])
            self.sale.setQuantity(table.cell(row_index, self.TitleList.index(self.TitleTuple[9])).value)
            self.sale.setPrice(table.cell(row_index, self.TitleList.index(self.TitleTuple[8])).value)
            self.sale.setNameNoEncode(table.cell(row_index, self.TitleList.index(self.TitleTuple[4])).value)
            self.sale.setDeliveryway('1') #宅配: 1, 超取711: 2, 超取全家: 3
            self.sale.setOrder_status('A0')

            self.customer.setGroup_id(GroupID)
            self.customer.setNameNoEncode(table.cell(row_index, self.TitleList.index(self.TitleTuple[4])).value)
            self.customer.setPhone(table.cell(row_index, self.TitleList.index(self.TitleTuple[5])).value)
            self.customer.setMobile(table.cell(row_index, self.TitleList.index(self.TitleTuple[5])).value)
            self.customer.setPost(None)
            self.customer.setAddressNoEncode(table.cell(row_index, self.TitleList.index(self.TitleTuple[6])).value)
        except Exception as e :
            print e.message
            logging.error(e.message)


if __name__ == '__main__':
    friday = Friday16_Data()
    # groupid = ""
    groupid='cbcc3138-5603-11e6-a532-000d3a800878'
    print friday.Friday_16_Data('friday',groupid,u'C:\\Users\\10509002\\Desktop\\訂單處理_20141215.xls','system')
