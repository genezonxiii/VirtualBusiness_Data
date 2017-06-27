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

class CTI_Data(Momo25_Data):
    Data = None
    mysqlconnect = None
    sale , customer = None, None

    # 預期要找出欄位的索引位置的欄位名稱
    TitleTuple = (u'訂單明細檔編號', u'訂單成立時間', u'到貨時間', u'產品編碼', u'產品名稱',
                  u'數量', u'訂單總金額', u'備註', u'訂購人姓名',u'收貨人姓名',
                  u'收貨人電話(日)', u'收貨人縣市(1)', u'收貨人縣市(1)+ 收貨人地址(1)',u'建檔者姓名')
    TitleList = []


    def CTI_yoho_Data(self, supplier, GroupID, path, UserID):

        try:

            logger.debug("===CTI_Data===")

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
            dup_str = ','.join(self.dup_order_no)
            self.dup_order_no = []
            logger.debug('===CTI_Data finally===')
            return json.dumps({"success": success, "info": resultinfo, "duplicate": dup_str, "total": totalRows}, sort_keys=False)

    def parserData(self,table,row_index,GroupID,UserID,supplier):
        try:
            self.sale.setGroup_id(GroupID)
            self.sale.setUser_id(UserID)
            self.sale.setOrder_source(supplier)
            self.sale.setOrder_No(str(table.cell(row_index, self.TitleList.index(self.TitleTuple[0])).value).split('.')[0])
            self.sale.setTrans_list_date_YYYYMMDD_float(table.cell(row_index, self.TitleList.index(self.TitleTuple[1])).value)
            self.sale.setSale_date_YYYYMMDD_float(table.cell(row_index, self.TitleList.index(self.TitleTuple[1])).value)
            self.sale.setC_Product_id(str(table.cell(row_index, self.TitleList.index(self.TitleTuple[3])).value))
            self.sale.setProduct_spec('')
            self.sale.setProduct_name_NoEncode((table.cell(row_index, self.TitleList.index(self.TitleTuple[4])).value))
            self.sale.setQuantity(table.cell(row_index, self.TitleList.index(self.TitleTuple[5])).value)
            self.sale.setPrice(table.cell(row_index, self.TitleList.index(self.TitleTuple[6])).value)
            self.sale.setNameNoEncode(table.cell(row_index, self.TitleList.index(self.TitleTuple[8])).value)
            self.sale.setDeliveryway('1') #宅配: 1, 超取711: 2, 超取全家: 3
            self.sale.setOrder_status('A0')
            self.sale.setMemo(table.cell(row_index, self.TitleList.index(self.TitleTuple[7])).value)
            # 到貨時間
            self.sale.setUser_def1(table.cell(row_index, self.TitleList.index(self.TitleTuple[2])).value)
            # 建檔者姓名
            self.sale.setUser_def2(table.cell(row_index, self.TitleList.index(self.TitleTuple[13])).value)

            self.customer.setGroup_id(GroupID)
            self.customer.setNameNoEncode(table.cell(row_index, self.TitleList.index(self.TitleTuple[9])).value)
            self.customer.setPhone(table.cell(row_index, self.TitleList.index(self.TitleTuple[10])).value)
            self.customer.setMobile(table.cell(row_index, self.TitleList.index(self.TitleTuple[10])).value)
            self.customer.setPost(None)
            address_detail = table.cell(row_index, self.TitleList.index(self.TitleTuple[11])).value + table.cell(row_index, self.TitleList.index(self.TitleTuple[12])).value
            self.customer.setAddressNoEncode(address_detail)
        except Exception as e :
            print e.message
            logging.error(e.message)


    def updateDB_Sale(self):
        try:
            SaleSQL = (self.sale.getGroup_id(), self.sale.getOrder_No(), self.sale.getUser_id(), self.sale.getProduct_name(), self.sale.getProduct_spec(), \
                       self.sale.getC_Product_id(), self.customer.getCustomer_id(), self.sale.getName(), self.sale.getQuantity(), \
                       self.sale.getPrice(), self.sale.getInvoice(), self.sale.getInvoice_date(), self.sale.getTrans_list_date(), \
                       self.sale.getDis_date(), self.sale.getMemo(), self.sale.getSale_date(), self.sale.getOrder_source(),\
                       self.sale.getDeliveryway(), self.sale.getTotal_amt(), self.sale.getOrder_status(), self.sale.getDeliver_name(), \
                       self.sale.getDeliver_to(), self.sale.getDeliver_store(), self.sale.getDeliver_phone(), self.sale.getDeliver_mobile(), \
                       self.sale.getPay_kind(), self.sale.getPay_status(), self.sale.getUser_def1(), self.sale.getUser_def2(), \
                       self.sale.getUser_def3(), self.sale.getUser_def4(), self.sale.getUser_def5(), self.sale.getUser_def6(), \
                       self.sale.getUser_def7(), self.sale.getUser_def8(), self.sale.getUser_def9(), self.sale.getUser_def10(), "")
            result = self.mysqlconnect.cursor.callproc('p_tb_sale_upower_cti', SaleSQL)
            if result[36] != None:
                self.dup_order_no.append(result[36])

            return
        except Exception as e :
            print e.message
            logging.error(e.message)
            raise

if __name__ == '__main__':
    friday = CTI_Data()
    # groupid = ""
    groupid='cbcc3138-5603-11e6-a532-000d3a800878'
    print friday.Friday_16_Data('friday',groupid,u'C:\\Users\\10509002\\Desktop\\CTI檔案.xlsx','test')
