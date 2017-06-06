# -*-  coding: utf-8  -*-
# __author__ = '10409003'
import logging, time
import json
import xlrd
from ToMysql import ToMysql
import uuid
from VirtualBusiness import Sale,Customer,updateCustomer

logger = logging.getLogger(__name__)

#FilePath
# class ExcelTemplate():
#     T_Cat_OutputFilePath, T_Cat_TemplateFile = None, None
#     def __init__(self):
#         self.T_Cat_TemplateFile = '/data/vbupload/Logistics_Tcat.xls'
#         self.T_Cat_OutputFilePath = '/data/vbupload_output/'

#Momo_xls檔(25欄)
class Momo25_Data():
    Data = None
    mysqlconnect = None
    sale , customer = None, None
    dup_order_no = []

    # 預期要找出欄位的索引位置的欄位名稱
    # \0407\momo摩天商城\1458525585016.xls
    # TitleTuple = (u'訂單編號', u'付款日', u'最晚出貨日', u'出貨日期', u'收件人姓名',
    #               u'電話', u'行動電話',u'地址', u'商店品號', u'商品名稱',
    #               u'數量', u'成交價')

    # 悠活原力
    # \momo\2016.03.16\A1102_3_1_008992_20160316114348.xls
    TitleTuple = (u'訂單編號', u'收件人姓名', u'收件人地址', u'轉單日', u'品號',
                  u'商品原廠編號',u'品名', u'數量',u'進價(含稅)', u'發票號碼',
                  u'單品詳細')
    # \momo\2016.10.05\A1106_008992_EXL.xls
    # TitleTuple = (u'訂單編號', u'收件人姓名', u'發票號碼', u'轉單日', u'品號',
    #               u'商品原廠編號', u'品名', u'數量', u'進價(含稅)')
    TitleList = []

    def __init__(self):
        # mysql connector object
        self.mysqlconnect = ToMysql()
        self.mysqlconnect.connect()

    def Momo_25_Data(self, supplier, GroupID, path, UserID):
        logging.basicConfig(filename='/data/VirtualBusiness_Data/pyupload.log',
                            level=logging.DEBUG,
                            format='%(asctime)s - %(levelname)s - %(filename)s:%(name)s:%(module)s/%(funcName)s/%(lineno)d - %(message)s',
                            datefmt='%Y/%m/%d %I:%M:%S %p')
        logging.Formatter.converter = time.gmtime

        logger.info('===Momo25_Data===')
        logger.debug('supplier:' + supplier)
        logger.debug('GroupID:' + GroupID)
        logger.debug('path:' + path)
        logger.debug('UserID:' + UserID)

        try:

            logger.debug("===Momo25_Data===")

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
            logger.debug('===Momo25_Data finally===')
            return json.dumps({"success": success, "info": resultinfo, "duplicate": dup_str, "total": totalRows}, sort_keys=False)

    def parserData(self,table,row_index,GroupID,UserID,supplier):
        try:
            self.sale.setGroup_id(GroupID)
            self.sale.setUser_id(UserID)
            self.sale.setOrder_source(supplier)
            self.sale.setOrder_No(table.cell(row_index, self.TitleList.index(self.TitleTuple[0])).value[0:18])
            self.sale.setTrans_list_date(table.cell(row_index, self.TitleList.index(self.TitleTuple[3])).value)
            self.sale.setSale_date(table.cell(row_index, self.TitleList.index(self.TitleTuple[3])).value)
            self.sale.setC_Product_id(str(table.cell(row_index, self.TitleList.index(self.TitleTuple[5])).value).split('.')[0])
            self.sale.setProduct_name_NoEncode(table.cell(row_index, self.TitleList.index(self.TitleTuple[6])).value)
            self.sale.setProduct_spec_NoEncode(table.cell(row_index, self.TitleList.index(self.TitleTuple[10])).value) #規格
            self.sale.setQuantity(table.cell(row_index, self.TitleList.index(self.TitleTuple[7])).value)
            self.sale.setPrice(table.cell(row_index, self.TitleList.index(self.TitleTuple[8])).value)
            self.sale.setNameNoEncode(table.cell(row_index, self.TitleList.index(self.TitleTuple[1])).value)
            # 訂單類型
            self.sale.setDeliveryway("1")   #宅配: 1, 超取711: 2, 超取全家: 3
            self.sale.setOrder_status('A0')

            self.customer.setGroup_id(GroupID)
            self.customer.setNameNoEncode(table.cell(row_index, self.TitleList.index(self.TitleTuple[1])).value)
            self.customer.setPhone(None)
            self.customer.setMobile(None)
            self.customer.setPost(None)
            self.customer.setAddressNoEncode(table.cell(row_index, self.TitleList.index(self.TitleTuple[2])).value)
        except Exception as e :
            print e.message
            logging.error(e.message)

    def updateDB_Customer(self):
        try:
            # insert or update table tb_customer
            updatecustomer = updateCustomer()
            self.customer.setCustomer_id(
                updatecustomer.checkCustomerid(self.customer.getGroup_id(), self.customer.get_Name(), self.customer.get_Address(), \
                                               self.customer.get_phone(), self.customer.get_Mobile(), self.customer.get_Email()))

            if self.customer.getCustomer_id() == None:
                self.customer.setCustomer_id(uuid.uuid4())
                CustomereSQL = (
                    self.customer.getCustomer_id(), self.customer.getGroup_id(), self.customer.getName(), \
                    self.customer.getAddress(), self.customer.getphone(), self.customer.getMobile(), \
                    self.customer.getEmail(), self.customer.getPost(), self.customer.getClass(), self.customer.getMemo(), self.sale.getUser_id())
                self.mysqlconnect.cursor.callproc('sp_insert_customer_bysys', CustomereSQL)
            else:
                CustomereSQL = (self.customer.getCustomer_id(), self.customer.getGroup_id(), self.customer.getName(), \
                                self.customer.getAddress(), self.customer.getphone(), self.customer.getMobile(), \
                                self.customer.getEmail(), self.customer.getPost(), self.customer.getClass(), \
                                self.customer.getMemo(),self.sale.getUser_id())
                self.mysqlconnect.cursor.callproc('sp_update_customer', CustomereSQL)

            CustomereSQL = (self.customer.getCustomer_id(), self.customer.getGroup_id(), self.customer.get_Name(), \
                            self.customer.get_Address(), self.customer.get_phone(), self.customer.get_Mobile(), \
                            self.customer.get_Email())
            updatecustomer.updataData(CustomereSQL)
        except Exception as e :
            print e.message
            logging.error(e.message)
            raise

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
    momo = Momo25_Data()
    # groupid = ""
    groupid='cbcc3138-5603-11e6-a532-000d3a800878'
    print momo.Momo_25_Data('momo',groupid,u'C:\\Users\\10509002\\Desktop\\A1102_3_1_008992_20160316114348.xls','system')
