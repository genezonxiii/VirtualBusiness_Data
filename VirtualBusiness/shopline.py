# -*-  coding: utf-8  -*-
# __author__ = '10509002'

import logging
import json
import xlrd
from VirtualBusiness import Sale,Customer,checkNum
from VirtualBusiness.Momo25 import Momo25_Data

logger = logging.getLogger(__name__)

#shopline
class Shopline(Momo25_Data):
    Data = None
    mysqlconnect = None
    sale, customer = None, None
    # 預期要找出欄位的索引位置的欄位名稱
    # 悠活原力
    TitleTuple = (u'訂單號碼', u'訂單日期', u'電話號碼', u'小計', u'收件人',
                  u'地址 1', u'郵政編號（如適用)', u'商品名稱', u'數量',u'商店貨號')

    TitleList = []

    #解析原始檔
    def Shopline_Data(self, supplier, GroupID, path, UserID):

        try:

            logger.debug("===shopline_Data===")

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
                # print "Melvin debug"
                # print row_index

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
            logger.debug('===Shopline_Data finally===')
            return json.dumps({"success": success, "info": resultinfo, "total": totalRows}, sort_keys=False)

    def parserData(self,table,row_index,GroupID,UserID,supplier):
        result = []
        try:
            self.sale.setGroup_id(GroupID)
            self.sale.setUser_id(UserID)
            self.sale.setOrder_source(supplier)
            self.sale.setOrder_No(table.cell(row_index, self.TitleList.index(self.TitleTuple[0])).value)
            self.sale.setTrans_list_date_YMDHMS_float(table.cell(row_index, self.TitleList.index(self.TitleTuple[1])).value)
            self.sale.setSale_date_YMDHMS_float(table.cell(row_index, self.TitleList.index(self.TitleTuple[1])).value)
            self.sale.setC_Product_id(str(table.cell(row_index, self.TitleList.index(self.TitleTuple[9])).value))
            self.sale.setProduct_name(table.cell(row_index, self.TitleList.index(self.TitleTuple[7])).value.encode('utf-8'))
            self.sale.setQuantity(table.cell(row_index, self.TitleList.index(self.TitleTuple[8])).value)
            self.sale.setPrice_str(table.cell(row_index, int(self.TitleList.index(self.TitleTuple[3]))).value)
            self.sale.setName(table.cell(row_index, self.TitleList.index(self.TitleTuple[4])).value.encode('utf-8'))
            self.sale.setDeliveryway('1')   #宅配: 1, 超取711: 2, 超取全家: 3

            self.customer.setGroup_id(GroupID)
            self.customer.setNameNoEncode(table.cell(row_index, self.TitleList.index(self.TitleTuple[4])).value)
            self.customer.setPhone(table.cell(row_index, self.TitleList.index(self.TitleTuple[2])).value)
            self.customer.setMobile(table.cell(row_index, self.TitleList.index(self.TitleTuple[2])).value)
            number = checkNum()
            numm = number.getNumber(table.cell(row_index, self.TitleList.index(self.TitleTuple[6])).value)
            self.customer.setPost(numm)
            #self.customer.setPost(table.cell(row_index, self.TitleList.index(self.TitleTuple[6])).value)
            self.customer.setAddressNoEncode(table.cell(row_index, self.TitleList.index(self.TitleTuple[5])).value)
        except Exception as e :
            print e.message
            logging.error(e.message)

if __name__ == '__main__':
    buy = Shopline()
    groupid = 'cbcc3138-5603-11e6-a532-000d3a800878'
    print buy.Shopline_Data('udn',groupid, u'C:\\Users\\10509002\\Desktop\\shopline相關資料\\yohopower_orders_20170117_20170117.xls','system')