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
    TitleTuple = (u'訂單號碼', u'訂單日期', u'訂單狀態', u'送貨方式', u'付款方式',
                  u'付款狀態', u'合計', u'收件人', u'收件人電話號碼',u'地址 1',
                  u'門市名稱', u'商品貨號', u'商品名稱', u'數量',u'顧客',
                  u'郵政編號（如適用)', u'城市')

    TitleList = []

    Citylist = [u'台北市', u'新北市', u'桃園市', u'台中市', u'台南市', u'高雄市	',
                u'基隆市', u'新竹市', u'嘉義市', u'新竹縣', u'苗栗縣', u'彰化縣',
                u'南投縣', u'雲林縣', u'嘉義縣', u'屏東縣', u'	宜蘭縣', u'花蓮縣',
                u'台東縣',u'澎湖縣',u'金門縣', u'馬祖縣']

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
                # print row_index

                self.sale = Sale()
                self.customer = Customer()
                #Parser Data from xls
                self.parserData(table, row_index, GroupID, UserID, supplier)

                if self.sale.getOrder_status() == 'DD':
                    continue
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
            if table.cell(row_index, self.TitleList.index(self.TitleTuple[11])).value != None and table.cell(row_index, self.TitleList.index(self.TitleTuple[11])).value != u'':
                print 'notequal'
                self.sale.setC_Product_id(str(table.cell(row_index, self.TitleList.index(self.TitleTuple[11])).value))
            self.sale.setProduct_name_NoEncode(table.cell(row_index, self.TitleList.index(self.TitleTuple[12])).value)
            self.sale.setQuantity(table.cell(row_index, self.TitleList.index(self.TitleTuple[13])).value)
            self.sale.setPrice_str(0)
            self.sale.setNameNoEncode(table.cell(row_index, self.TitleList.index(self.TitleTuple[14])).value)
            if table.cell(row_index, self.TitleList.index(self.TitleTuple[3])).value[0:4] == u'7-11':
                self.sale.setDeliveryway('2')   #宅配: 1, 超取711: 2, 超取全家: 3
            elif table.cell(row_index, self.TitleList.index(self.TitleTuple[3])).value[0:2] == u'全家':
                self.sale.setDeliveryway('3')
            else:
                self.sale.setDeliveryway('1')

            self.sale.setTotal_amt(table.cell(row_index, self.TitleList.index(self.TitleTuple[6])).value)
            if table.cell(row_index, self.TitleList.index(self.TitleTuple[2])).value == u'處理中':
                self.sale.setOrder_status('A0')
            else:
                self.sale.setOrder_status('DD')
            self.sale.setDeliver_name(table.cell(row_index, self.TitleList.index(self.TitleTuple[7])).value)

            if table.cell(row_index, self.TitleList.index(self.TitleTuple[9])).value[0:3] in self.Citylist:
                self.sale.setDeliver_to(table.cell(row_index, self.TitleList.index(self.TitleTuple[9])).value)
            else:
                self.sale.setDeliver_to(table.cell(row_index, self.TitleList.index(self.TitleTuple[16])).value + \
                                        table.cell(row_index, self.TitleList.index(self.TitleTuple[9])).value)
            self.sale.setDeliver_store(table.cell(row_index, self.TitleList.index(self.TitleTuple[10])).value)
            self.sale.setDeliver_phone(table.cell(row_index, self.TitleList.index(self.TitleTuple[8])).value)
            self.sale.setDeliver_mobile(table.cell(row_index, self.TitleList.index(self.TitleTuple[8])).value)
            self.sale.setPay_kind(table.cell(row_index, self.TitleList.index(self.TitleTuple[4])).value)
            self.sale.setPay_status(table.cell(row_index, self.TitleList.index(self.TitleTuple[5])).value)


            self.customer.setGroup_id(GroupID)
            self.customer.setNameNoEncode(table.cell(row_index, self.TitleList.index(self.TitleTuple[14])).value)
            self.customer.setPhone(table.cell(row_index, self.TitleList.index(self.TitleTuple[8])).value)
            self.customer.setMobile(table.cell(row_index, self.TitleList.index(self.TitleTuple[8])).value)
            number = checkNum()
            numm = number.getNumber(table.cell(row_index, self.TitleList.index(self.TitleTuple[15])).value)
            self.customer.setPost(numm)
            #self.customer.setPost(table.cell(row_index, self.TitleList.index(self.TitleTuple[6])).value)
            if table.cell(row_index, self.TitleList.index(self.TitleTuple[9])).value[0:3] in self.Citylist:
                self.customer.setAddressNoEncode(table.cell(row_index, self.TitleList.index(self.TitleTuple[9])).value)
            else:
                self.customer.setAddressNoEncode(table.cell(row_index, self.TitleList.index(self.TitleTuple[16])).value + \
                                                 table.cell(row_index, self.TitleList.index(self.TitleTuple[9])).value)
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
                       self.sale.getPay_kind(), self.sale.getPay_status(), "")
            result = self.mysqlconnect.cursor.callproc('p_tb_sale_new', SaleSQL)
            if result[27] != None:
                self.dup_order_no.append(result[27])

            return
        except Exception as e :
            print e.message
            logging.error(e.message)
            raise



if __name__ == '__main__':
    buy = Shopline()
    groupid = 'cbcc3138-5603-11e6-a532-000d3a800878'
    print buy.Shopline_Data('shopline',groupid, u'C:\\Users\\10509002\\Desktop\\8ceb0f3b-4b34-479a-8743-87e1652b7abf.xls','system')