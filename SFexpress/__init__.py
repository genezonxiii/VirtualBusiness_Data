# -*-  coding: utf-8  -*-

import logging
from SFexpress.sfexpress_out import sfexpressout, ExcelTemplate
from SFexpress.sfexpress_in import sfexpressin
import os


logger = logging.getLogger(__name__)

class FileProcess_sf():
    def __init__(self):
        pass

    def transferFile(self, DataPath, UserID , LogisticsID=26, ProductCode=None ):
        logger.debug(DataPath)
        GroupID = DataPath.split('/')[4]                        #group id
        Platform = DataPath.split("/")[3]                      #出庫或入庫
        logger.debug("Outputfile")
        OutputFile = DataPath.split("/")[5]  #輸出檔案名稱
        logger.debug(OutputFile)
        outPutPath = ExcelTemplate()
        logger.debug("again")
        OutputFile = outPutPath.T_SF_OutputFilePath + OutputFile +".xls" #輸出檔案路徑及名稱
        logger.debug(OutputFile)
        GB = None
        if Platform == 'inbound' :
            logger.debug("入庫")
            GB = sfexpressin()
        elif Platform == 'outbound' :
            logger.debug("出庫")
            GB = sfexpressout()

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
    test = FileProcess_sf()
    print test.transferFile(u'C:/data/vbSF/inbound/cbcc3138-5603-11e6-a532-000d3a800878/6412fe6b-8e5b-474d-8898-c4c553e11d03.xls','robintest', 2, 'DS')
    # print 'finish'