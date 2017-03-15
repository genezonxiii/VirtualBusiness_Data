# -*-  coding: utf-8  -*-
#__author__ = '10509002'

import logging
import os
import time
from ImportFundationData.insert import SupplyData
from ImportFundationData.insert import ProductData
from ImportFundationData.insert import PackageData
from ImportFundationData.insert import ContrastData

logger = logging.getLogger(__name__)

class DataInsert():
    Data=None

    def __init__(self):
        pass
    def virtualbusiness_import(self,DataPath,userID):
        Importtype = str(DataPath).split('/')[3]  #供應商、商品、組合包、對照表
        Groupid = str(DataPath).split('/')[4]      #group id

        print Importtype
        print Groupid


        logging.basicConfig(filename='/data/VirtualBusiness_Data/pyupload.log',
			level=logging.DEBUG,
			format='%(asctime)s - %(levelname)s - %(filename)s:%(name)s:%(module)s/%(funcName)s/%(lineno)d - %(message)s',
            datefmt='%Y/%m/%d %I:%M:%S %p')
        logging.Formatter.converter = time.gmtime

        logger.info('Importtype')
        logger.info(Importtype)

        if Importtype == 'supply':
            logger.debug('supply')
            FinalData = SupplyData()
            return FinalData.Supply(Groupid, os.path.join(DataPath),userID)

        if Importtype == 'product':
            logger.debug('product')
            FinalData = ProductData()
            return FinalData.Product(Groupid, os.path.join(DataPath), userID)

        if Importtype == 'package':
            logger.debug('package')
            FinalData = PackageData()
            return FinalData.Package(Groupid, os.path.join(DataPath), userID)

        if Importtype == 'contrast':
            logger.debug('contrast')
            FinalData = ContrastData()
            return FinalData.Contrast(Groupid, os.path.join(DataPath), userID)