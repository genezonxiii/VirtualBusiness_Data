# -*-  coding: utf-8  -*-
# __author__ = '10408001'
import logging
from GroupBuy.buy123 import buy123,ExcelTemplate
from GroupBuy.chinatime import chinatime
from GroupBuy.crazymike import Crazymike
from GroupBuy.food123 import food123
from GroupBuy.gomaji import gomaji
from GroupBuy.HerBuy import HerBuy
from GroupBuy.Ibon import Ibon
from GroupBuy.Life17 import Life17
from GroupBuy.mybenefit import Mybenefit
from GroupBuy.pcone import Pcone
from GroupBuy.popular import popular
from GroupBuy.sale123 import Sale123
import os

logger = logging.getLogger(__name__)

class FileProcess():
    def __init__(self):
        pass

    def transferFile(self, DataPath, UserID , LogisticsID=2, ProductCode=None ):
        GroupID = DataPath.split('/')[5]                        #group id
        Platform = DataPath.split("/")[3]                      #平台
        OutputFile = os.path.basename(DataPath).split(".")[0]  #輸出檔案名稱
        outPutPath = ExcelTemplate()
        OutputFile = outPutPath.T_Cat_OutputFilePath + OutputFile +".xls" #輸出檔案路徑及名稱
        GB = None
        if Platform == 'lifemarket' :
            logger.debug("生活市集")
            GB = buy123()
        elif Platform == 'chinatime' :
            logger.debug("中時團購")
            GB = chinatime()
        elif Platform == 'pcone' :
            logger.debug("松果購物")
            GB = Pcone()
        elif Platform == 'popular' :
            logger.debug("小P大團購")
            GB = popular()
        elif Platform == 'sister' :
            logger.debug("姊妹購物網")
            GB = Sale123()
        elif Platform == 'healthmike' :
            logger.debug("健康麥克")
            GB = Crazymike()
        elif Platform == 'delicious' :
            logger.debug("好吃宅配網")
            GB = food123()
        elif Platform == 'chinaugo' :
            logger.debug("中華優購")
            GB = Mybenefit()
        elif Platform == 'ibon' :
            logger.debug("Ibon")
            GB = Ibon()
        elif Platform == '17life' :
            logger.debug("17life")
            GB = Life17()
        elif Platform == 'herbuy' :
            logger.debug("好買寶貝")
            GB = HerBuy()
        elif Platform == 'gomaji' :
            logger.debug("夠麻吉")
            GB = gomaji()
        # elif Platform == 'shopline' :
        #     logger.debug("Shopline")
        #     GB =
        # elif Platform == 'charmingmall' :
        #     logger.debug("俏美魔")
        #     GB =
        # elif Platform == 'mintyday' :
        #     logger.debug("清新好日")
        #     GB =

        return GB.parserFile(GroupID,UserID,LogisticsID,ProductCode,DataPath,OutputFile)


        # LogisticsID  :  1	中華郵政
        # LogisticsID  :  2	黑貓宅急便
        # LogisticsID  :  3	台灣宅配通
        # LogisticsID  :  4	新竹物流
        # LogisticsID  :  5	嘉里大榮物流（大榮貨運）
        # LogisticsID  :  6	便利帶
        # LogisticsID  :  7	日通快遞
        # LogisticsID  :  8	全速配
        # LogisticsID  :  9	同同宅配
        # LogisticsID  :  10	國陽郵遞
        # LogisticsID  :  11	第一郵控
        # LogisticsID  :  12	統群物流
        # LogisticsID  :  13	通盈貨運
        # LogisticsID  :  14	速遞家
        # LogisticsID  :  15	超峰快遞
        # LogisticsID  :  16	錸乾物流
        # LogisticsID  :  17	其他
        # LogisticsID  :  18	尚發運通
        # LogisticsID  :  19	雅仕快遞
        # LogisticsID  :  20	驥騰物流
        # LogisticsID  :  21	巨航快遞
        # LogisticsID  :  22	台快快遞
        # LogisticsID  :  23	亞風速遞
        # LogisticsID  :  24	中連貨運
        # LogisticsID  :  25	祥億貨運
        # LogisticsID  :  26	順豐速運
        # LogisticsID  :  27	永騰物流
        # LogisticsID  :  28	速利貨運
        # LogisticsID  :  29	咏高交通
        # LogisticsID  :  30	速遞物流
        # LogisticsID  :  31	揚績國際有限公司(揚績物流)
        # LogisticsID  :  32	聯新快遞
        # LogisticsID  :  33	飛際快遞
        # LogisticsID  :  34	日鎰快遞
        # LogisticsID  :  35	信速通運
        # LogisticsID  :  36	大誠/新航線快遞
        # LogisticsID  :  37	聯盟物流(駿通企業)
        # LogisticsID  :  38	優速快遞
        # LogisticsID  :  39	新航快遞
        # LogisticsID  :  40	來來快遞
        # LogisticsID  :  41	丞琦貨運
        # LogisticsID  :  42	棟樑物流
        # LogisticsID  :  43	明箭快遞
        # LogisticsID  :  44	神行(中國)快遞
        # LogisticsID  :  45	協合展業
        # LogisticsID  :  46	強訊/佑訊郵通(上大郵通)
        # LogisticsID  :  47	進南貨運有限公司
        # LogisticsID  :  48	迪比翼快遞DPEX
        # LogisticsID  :  49	CTC全統物流
        # LogisticsID  :  50	詹姆士國際快遞
        # LogisticsID  :  51	天成快遞
        # LogisticsID  :  52	全球快遞
        # LogisticsID  :  53	好運袋
        # LogisticsID  :  54	聯合統配股份有限公司


if __name__ == '__main__':
    test = FileProcess()
    print test.transferFile(u'/data/vbGroupbuy/lifemarket/general/396a2df8-472e-11e6-806e-000c29c1d067/2016-12-05_生活市集_BY123377302F.xls','robintest', 2, 'DS')
    # print 'finish'