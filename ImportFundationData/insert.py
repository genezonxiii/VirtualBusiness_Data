# -*-  coding: utf-8  -*-
#__author__ = '10509002'

import logging
import json
import uuid
import xlrd
from ToMysql import ToMysql
from ImportFundationData.entertype import Supply
from ImportFundationData.entertype import Product
from ImportFundationData.entertype import Package_master
from ImportFundationData.entertype import Package_detail
from ImportFundationData.entertype import Contrast

logger = logging.getLogger(__name__)


class SupplyData():
    Data = None
    mysqlconnect = None
    supply = None
    dup_order_no = []

    TitleTuple = (u'供應商名稱', u'統一編號', u'地址', u'聯絡人1', u'電話',
                  u'分機', u'行動電話', 'email', u'聯絡人2', u'電話2',
                  u'分機2', u'行動電話2','email1')

    TitleList = []

    def __init__(self):
        # mysql connector object
        self.mysqlconnect = ToMysql()
        self.mysqlconnect.connect()

    def Supply(self, GroupID, path, UserID):

        try:

            logger.debug("===Supply_Data===")

            success = False
            resultinfo = ""
            totalRows = 0

            data = xlrd.open_workbook(path)
            table = data.sheets()[0]

            totalRows = table.nrows - 2

            # 存放excel中全部的欄位名稱
            self.TitleList = []
            for row_index in range(1, 2):
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

            for row_index in range(2, table.nrows):
                self.supply = Supply()

                #Parser Data from xls
                self.parserData_Supply(table, row_index, GroupID, UserID)
                # insert or update table tb_customer
                self.updateDB_Supply()

                self.supply = None

            self.mysqlconnect.db.commit()
            self.mysqlconnect.dbClose()

            success = True

        except Exception as inst:
            logger.error(inst.args)
            resultinfo = inst.args
        finally:
            dup_str = ','.join(self.dup_order_no)
            self.dup_order_no = []
            logger.debug('===SupplyData finally===')
            return json.dumps({"success": success, "info": resultinfo,"duplicate": dup_str, "total": totalRows}, sort_keys=False)

    def parserData_Supply(self,table,row_index,GroupID,UserID):
        try:
            self.supply.setGroup_id(GroupID)
            self.supply.setUser_id(UserID)
            self.supply.setSupply_name(table.cell(row_index, self.TitleList.index(self.TitleTuple[0])).value)
            self.supply.setSupply_unicode(table.cell(row_index, self.TitleList.index(self.TitleTuple[1])).value)
            self.supply.setAddress(table.cell(row_index, self.TitleList.index(self.TitleTuple[2])).value)
            self.supply.setContact(table.cell(row_index, self.TitleList.index(self.TitleTuple[3])).value)
            self.supply.setPhone(table.cell(row_index, self.TitleList.index(self.TitleTuple[4])).value)
            self.supply.setExt(table.cell(row_index, self.TitleList.index(self.TitleTuple[5])).value)
            self.supply.setMobile(table.cell(row_index, self.TitleList.index(self.TitleTuple[6])).value)
            self.supply.setEmail(table.cell(row_index, self.TitleList.index(self.TitleTuple[7])).value)
            self.supply.setContact1(table.cell(row_index, self.TitleList.index(self.TitleTuple[8])).value)
            self.supply.setPhone1(table.cell(row_index, self.TitleList.index(self.TitleTuple[9])).value)
            self.supply.setExt1(table.cell(row_index, self.TitleList.index(self.TitleTuple[10])).value)
            self.supply.setMobile1(table.cell(row_index, self.TitleList.index(self.TitleTuple[11])).value)
            self.supply.setEmail1(table.cell(row_index, self.TitleList.index(self.TitleTuple[12])).value)

        except Exception as e :
            print e.message
            logging.error(e.message)

    def updateDB_Supply(self):
        try:
            SupplySQL = (self.supply.getGroup_id(), self.supply.getSupply_name(), self.supply.getSupply_unicode(), self.supply.getAddress(), self.supply.getContact(), \
                         self.supply.getPhone(), self.supply.getExt(), self.supply.getMobile(),self.supply.getContact1(), \
                         self.supply.getPhone1(), self.supply.getExt1(), self.supply.getMobile1(), self.supply.getEmail(), \
                         self.supply.getEmail1(),self.supply.getMemo(), self.supply.getUser_id(),"")
            result = self.mysqlconnect.cursor.callproc('sp_insert_supply', SupplySQL)
            if result[16] != None:
                self.dup_order_no.append(result[16])
            return
        except Exception as e :
            print e.message
            logging.error(e.message)
            raise


class ProductData():
    Data = None
    mysqlconnect = None
    product = None
    dup_order_no = []

    TitleTuple = (u'自訂產品ID', u'產品名稱', u'供應商名稱', u'產品類型', u'產品單位',
                  u'成本', u'售價', u'產品說明', u'條碼', u'安全庫存',
                  u'期初庫存')

    TitleList = []

    def __init__(self):
        # mysql connector object
        self.mysqlconnect = ToMysql()
        self.mysqlconnect.connect()

    def Product(self,GroupID, path, UserID):

        try:

            logger.debug("===Product_Data===")

            success = False
            resultinfo = ""
            totalRows = 0

            data = xlrd.open_workbook(path)
            table = data.sheets()[0]

            totalRows = table.nrows - 2

            # 存放excel中全部的欄位名稱
            self.TitleList = []
            for row_index in range(1, 2):
                for col_index in range(0, table.ncols):
                    self.TitleList.append(table.cell(row_index, col_index).value)

            print self.TitleList

            # 存放excel中對應TitleTuple欄位名稱的index
            for index in range(0, len(self.TitleTuple)):
                if self.TitleTuple[index] in self.TitleList:
                    logger.debug(str(index) + self.TitleTuple[index])
                    logger.debug(u'index in file - ' + str(self.TitleList.index(self.TitleTuple[index])))
                    # print str(index) + TitleTuple[index]
                    # print (TitleList.index(TitleTuple[index]))

            for row_index in range(2, table.nrows):
                self.product = Product()
                # Parser Data from xls
                self.parserData_Product(table, row_index, GroupID, UserID)
                # insert or update table tb_customer
                self.updateDB_Product()
                self.product = None

            self.mysqlconnect.db.commit()
            self.mysqlconnect.dbClose()

            success = True

        except Exception as inst:
            logger.error(inst.args)
            resultinfo = inst.args
        finally:
            dup_str = ','.join(self.dup_order_no)
            self.dup_order_no = []
            logger.debug('===ProductData finally===')
            return json.dumps({"success": success, "info": resultinfo,"duplicate": dup_str, "total": totalRows}, sort_keys=False)

    def parserData_Product(self,table,row_index, GroupID, UserID):
        try:
            self.product.setGroup_id(GroupID)
            self.product.setUser_id(UserID)
            self.product.setC_Product_id(table.cell(row_index, self.TitleList.index(self.TitleTuple[0])).value)
            self.product.setP_Product_name(table.cell(row_index, self.TitleList.index(self.TitleTuple[1])).value)
            self.product.setSupply_name(table.cell(row_index, self.TitleList.index(self.TitleTuple[2])).value)
            self.product.setType_name(table.cell(row_index, self.TitleList.index(self.TitleTuple[3])).value)
            self.product.setUnit_name(table.cell(row_index, self.TitleList.index(self.TitleTuple[4])).value)
            self.product.setCost(table.cell(row_index, self.TitleList.index(self.TitleTuple[5])).value)
            self.product.setPrice(table.cell(row_index, self.TitleList.index(self.TitleTuple[6])).value)
            self.product.setDescription(table.cell(row_index, self.TitleList.index(self.TitleTuple[7])).value)
            self.product.setBarcode(table.cell(row_index, self.TitleList.index(self.TitleTuple[8])).value)
            self.product.setKeep_stock(table.cell(row_index, self.TitleList.index(self.TitleTuple[9])).value)
            self.product.setBegging_stock(table.cell(row_index, self.TitleList.index(self.TitleTuple[10])).value)

        except Exception as e:
            print e.message
            logging.error(e.message)

    def updateDB_Product(self):
        try:
            ProductSQL = (self.product.getGroup_id(), self.product.getC_Product_id(), self.product.getP_Product_name(), self.product.getSupply_name(), self.product.getType_name(), \
                         self.product.getUnit_name(), self.product.getCost(), self.product.getPrice(),self.product.getDescription(), \
                         self.product.getBarcode(), self.product.getKeep_stock(), self.product.getBegging_stock(), self.product.getUser_id(),"")
            result = self.mysqlconnect.cursor.callproc('sp_insert_product_xls', ProductSQL)
            if result[13] != None:
                self.dup_order_no.append(result[13])
            return
        except Exception as e :
            print e.message
            logging.error(e.message)
            raise


class PackageData():
    Data = None
    mysqlconnect = None
    package_M = None
    package_D = None
    packageMList = []
    packageDList = []
    dup_order_no = []

    TitleTupleM = (u'自訂組合包ID', u'組合包名稱', u'組合包規格', u'售價', u'條碼',
                  u'說明')

    TitleTupleD = (u'自訂組合包ID', u'自訂產品ID', u'產品名稱', u'產品說明', u'數量')

    TitleList = []
    TitleListD = []

    def __init__(self):
        # mysql connector object
        self.mysqlconnect = ToMysql()
        self.mysqlconnect.connect()

    def Package(self, GroupID, path, UserID):
        try:
            logger.debug("===Package_Data===")

            success = False
            resultinfo = ""
            totalRows = 0
            detailRows = 0

            data = xlrd.open_workbook(path)
            table = data.sheets()[0] # Master
            table_detail = data.sheets()[1] # detail

            totalRows = table.nrows - 2
            detailRows = table_detail.nrows - 2

            # 存放excel中全部的欄位名稱
            self.TitleList = []
            for row_index in range(1, 2):
                for col_index in range(0, table.ncols):
                    self.TitleList.append(table.cell(row_index, col_index).value)

            self.TitleListD = []
            for row_index in range(1, 2):
                for col_index in range(0, table_detail.ncols):
                    self.TitleListD.append(table_detail.cell(row_index, col_index).value)

            print self.TitleList
            print self.TitleListD

            # 存放excel中對應TitleTuple欄位名稱的index
            for index in range(0, len(self.TitleTupleM)):
                if self.TitleTupleM[index] in self.TitleList:
                    logger.debug(str(index) + self.TitleTupleM[index])
                    logger.debug(u'index in file - ' + str(self.TitleList.index(self.TitleTupleM[index])))
                    # print str(index) + TitleTuple[index]
                    # print (TitleList.index(TitleTuple[index]))

            for index2 in range(0, len(self.TitleTupleD)):
                if self.TitleTupleD[index2] in self.TitleList:
                    logger.debug(str(index2) + self.TitleTupleD[index2])
                    logger.debug(u'index in file - ' + str(self.TitleListD.index(self.TitleTupleD[index2])))
            # 讀Master的資料
            for row_index in range(2, table.nrows):
                self.package_M = Package_master()
                # Parser Data from xls
                self.parserData_Package(table, row_index, GroupID, UserID)
                # insert or update table tb_customer
                # self.updateDB_Package()
                self.package_M = None

            # 讀detail的資料
            for row_index in range(2,table_detail.nrows):
                self.package_D = Package_detail()
                # Parser Data from xls
                self.parserData_Package_detail(table_detail, row_index, GroupID, UserID)
                # insert or update table tb_customer
                # self.updateDB_Package_detail()
                self.package_D = None

            # for Master :
            #     c_package_id = Master.getC_package_id()
            #     package_id = Master.getPackage_id()
            #     Detail.filter("c_package_id", c_package_id)
            #     for Detail:
            #         Detail.setParent_id(package_id)

            # master 跟 detail 要 match
            for master in self.packageMList:
                logger.debug("===master_list get detail_list same sale_id===")
                logger.debug("master_id:" + master.getC_Package_id())
                g = master.getGroup_id()
                package_id = master.getPackage_id()
                master_detail_list = [detail for detail in self.packageDList if detail.getC_Package_id() == master.getC_Package_id()]

                self.package_M = master
                result8 = self.updateDB_Package()

                if result8 == None:
                    for detail in master_detail_list:
                        detail.setParent_id(package_id)
                        self.package_D = detail
                        self.updateDB_Package_detail()


            self.package_M = None
            self.package_D = None

            self.mysqlconnect.db.commit()
            self.packageMList = []
            self.packageDList = []
            self.mysqlconnect.dbClose()

            success = True

        except Exception as inst:
            logger.error(inst.args)
            resultinfo = inst.args
        finally:
            dup_str = ','.join(self.dup_order_no)
            self.dup_order_no = []
            logger.debug('===PackageData finally===')
            return json.dumps({"success": success, "info": resultinfo, "duplicate": dup_str, "totalM": totalRows, "totalD": detailRows}, sort_keys=False)

    def parserData_Package(self,table,row_index, GroupID, UserID):
        try:
            self.package_M.setGroup_id(GroupID)
            # self.package_M.setUser_id(UserID)
            self.package_M.setPackage_id(str(uuid.uuid4()))
            self.package_M.setC_Package_id(table.cell(row_index, self.TitleList.index(self.TitleTupleM[0])).value)
            self.package_M.setPackage_name(table.cell(row_index, self.TitleList.index(self.TitleTupleM[1])).value)
            self.package_M.setPackage_spec(table.cell(row_index, self.TitleList.index(self.TitleTupleM[2])).value)
            self.package_M.setAmount(table.cell(row_index, self.TitleList.index(self.TitleTupleM[3])).value)
            self.package_M.setBarcode(table.cell(row_index, self.TitleList.index(self.TitleTupleM[4])).value)
            self.package_M.setDescription(table.cell(row_index, self.TitleList.index(self.TitleTupleM[5])).value)
            self.packageMList.append(self.package_M)    # 放Master資料到一個新的list裡

        except Exception as e:
            print e.message
            logging.error(e.message)

    def parserData_Package_detail(self,table_detail,row_index, GroupID, UserID):
        try:
            self.package_D.setGroup_id(GroupID)
            # self.package.setUser_id(UserID)
            self.package_D.setParent_id("")
            self.package_D.setC_Package_id(table_detail.cell(row_index, self.TitleListD.index(self.TitleTupleD[0])).value)
            self.package_D.setC_Product_id(table_detail.cell(row_index, self.TitleListD.index(self.TitleTupleD[1])).value)
            self.package_D.setProduct_name(table_detail.cell(row_index, self.TitleListD.index(self.TitleTupleD[2])).value)
            self.package_D.setDescription(table_detail.cell(row_index, self.TitleListD.index(self.TitleTupleD[3])).value)
            self.package_D.setQuantity(table_detail.cell(row_index, self.TitleListD.index(self.TitleTupleD[4])).value)
            self.packageDList.append(self.package_D)    # 放Detail資料到一個新的list裡

        except Exception as e:
            print e.message
            logging.error(e.message)

    def updateDB_Package(self):
        try:
            PackageSQL = (self.package_M.getPackage_id(), self.package_M.getGroup_id(), self.package_M.getC_Package_id(), self.package_M.getPackage_name(), self.package_M.getPackage_spec(), \
                         self.package_M.getAmout(), self.package_M.getBarcode(), self.package_M.getDescription(),"")
            result = self.mysqlconnect.cursor.callproc('sp_insert_package_master_xls', PackageSQL)
            if result[8] != None:
                self.dup_order_no.append(result[8])
                return result[8]
            else:
                return

        except Exception as e :
            print e.message
            logging.error(e.message)
            raise

    def updateDB_Package_detail(self):
        try:
            PackageSQL = (self.package_D.getParent_id(), self.package_D.getGroup_id(), self.package_D.getC_Product_id(), \
                          self.package_D.getProduct_name(), self.package_D.getDescription(),self.package_D.getQuantity())
            self.mysqlconnect.cursor.callproc('sp_insert_package_detail_xls', PackageSQL)
            return

        except Exception as e:
            print e.message
            logging.error(e.message)
            raise

class ContrastData():
    Data = None
    mysqlconnect = None
    contrast = None
    dup_order_no = []

    TitleTuple = (u'對照類別', u'產品名稱/組合包名稱', u'產品說明/組合包規格', u'平台', u'平台用產品名稱',
                  u'平台用產品規格', u'應收金額')

    TitleList = []

    def __init__(self):
        # mysql connector object
        self.mysqlconnect = ToMysql()
        self.mysqlconnect.connect()

    def Contrast(self, GroupID, path, UserID):
        try:
            logger.debug("===Contrast_Data===")

            success = False
            resultinfo = ""
            totalRows = 0

            data = xlrd.open_workbook(path)
            table = data.sheets()[0]

            totalRows = table.nrows - 2

            # 存放excel中全部的欄位名稱
            self.TitleList = []
            for row_index in range(1, 2):
                for col_index in range(0, table.ncols):
                    self.TitleList.append(table.cell(row_index, col_index).value)

            print self.TitleList

            # 存放excel中對應TitleTuple欄位名稱的index
            for index in range(0, len(self.TitleTuple)):
                if self.TitleTuple[index] in self.TitleList:
                    logger.debug(str(index) + self.TitleTuple[index])
                    logger.debug(u'index in file - ' + str(self.TitleList.index(self.TitleTuple[index])))
                    # print str(index) + TitleTuple[index]
                    # print (TitleList.index(TitleTuple[index]))

            for row_index in range(2, table.nrows):
                self.contrast = Contrast()
                # Parser Data from xls
                self.parserData_contrast(table, row_index, GroupID, UserID)
                # insert or update table tb_customer
                self.updateDB_Contrast()
                self.contrast = None

            self.mysqlconnect.db.commit()
            self.mysqlconnect.dbClose()

            success = True

        except Exception as inst:
            logger.error(inst.args)
            resultinfo = inst.args
        finally:
            dup_str = ','.join(self.dup_order_no)
            self.dup_order_no = []
            logger.debug('===ContrastData finally===')
            return json.dumps({"success": success, "info": resultinfo,"duplicate": dup_str, "total": totalRows}, sort_keys=False)

    def parserData_contrast(self,table,row_index, GroupID, UserID):
        try:
            self.contrast.setGroup_id(GroupID)
            self.contrast.setContrast_type(table.cell(row_index, self.TitleList.index(self.TitleTuple[0])).value)
            self.contrast.setProduct_name(table.cell(row_index, self.TitleList.index(self.TitleTuple[1])).value)
            self.contrast.setDescription(table.cell(row_index, self.TitleList.index(self.TitleTuple[2])).value)
            self.contrast.setPlatform_name(table.cell(row_index, self.TitleList.index(self.TitleTuple[3])).value)
            self.contrast.setProduct_name_platform(table.cell(row_index, self.TitleList.index(self.TitleTuple[4])).value)
            self.contrast.setProduct_spec_platform(table.cell(row_index, self.TitleList.index(self.TitleTuple[5])).value)
            self.contrast.setAmount(table.cell(row_index, self.TitleList.index(self.TitleTuple[6])).value)

        except Exception as e:
            print e.message
            logging.error(e.message)

    def updateDB_Contrast(self):
        try:
            ContrastSQL = (self.contrast.getGroup_id(), self.contrast.getContrast_type(), self.contrast.getProduct_name(), self.contrast.getDescription(), self.contrast.getPlatform_name(), \
                         self.contrast.getProduct_name_platform(), self.contrast.getProduct_spec_platform(), self.contrast.getAmount(), "")
            result = self.mysqlconnect.cursor.callproc('sp_insert_product_contrast_xls', ContrastSQL)
            if result[8] != None:
                self.dup_order_no.append(result[8])
            return

        except Exception as e :
            print e.message
            logging.error(e.message)
            raise




if __name__ == '__main__':
    data =PackageData()
    # groupid = ""
    groupid='cbcc3138-5603-11e6-a532-000d3a800878'
    print data.Package(groupid, u'C:\\Users\\10509002\\Desktop\\packageTemplate.xls','system')
