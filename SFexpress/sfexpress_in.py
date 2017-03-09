# -*-  coding: utf-8  -*-

import logging
import xlrd
import json
from xlutils.copy import copy
from SFexpress.sfexpress_out import sfexpressout

logger = logging.getLogger(__name__)

#FilePath
class ExcelTemplate():
    T_SF_OutputFilePath, T_SF_TemplateFile = None, None
    MailSender , MailReceiver , SMTPServer = None , None , None
    def __init__(self):
        self.T_SF_TemplateFile = '/data/vbGroupbuy/Logistics_SF_In.xls'
        self.T_SF_OutputFilePath = '/data/vbSF_output/'
        # Aber 正式用
        self.MailSender = 'pscaber@cloud.pershing.com.tw'
        self.MailReceiver =['joeyang@pershing.com.tw','hsuanmeng@pershing.com.tw','christinewei@pershing.com.tw']
        self.SMTPServer = 'cloud-pershing-com-tw.mail.protection.outlook.com'
        # Local 測試用
        # self.MailSender = 'hsuanmeng@pershing.com.tw'
        # self.MailReceiver = ['joeyang@pershing.com.tw', 'hsuanmeng@pershing.com.tw', 'christinewei@pershing.com.tw']
        # self.SMTPServer = 'ms1.pershing.com.tw'

class sfexpressin(sfexpressout):
    # 解析原始檔
    def parserFile(self, GroupID, UserID, LogisticsID=26, ProductCode=None, inputFile=None, outputFile=None):
        self.init_log('SFexpressIn_Data', GroupID, UserID, ProductCode, inputFile)
        success = False
        try:
            self.getRegularEx(GroupID)
            data = xlrd.open_workbook(inputFile)
            table = data.sheets()[0]
            rows = table.nrows
            result = []
            resultinfo = ""

            order_no = table.cell(4, 10).value
            receiver = table.cell(4, 4).value
            address = table.cell(8, 1).value
            phone = table.cell(5, 4).value
            client = table.cell(4, 1).value

            row_index = 11
            page_idx = 1
            while row_index < rows:
                page_row_start = (page_idx - 1) * 43 + 11
                page_row_end = (page_idx - 1) * 43 + 35

                # 取資料ROW 1
                row_column1 = table.cell(row_index, 0).value

                if row_column1 == u'付款條件':
                    break

                row_column2 = table.cell(row_index, 1).value
                row_column3 = int(table.cell(row_index, 4).value)
                row_column4 = table.cell(row_index, 6).value
                row_column5 = table.cell(row_index, 10).value

                # 取資料ROW 2
                row2_idx = row_index + 1
                if (row2_idx < page_row_start or row2_idx > page_row_end):
                    row2_idx = page_idx * 43 + 11
                    row_shift = 1
                row2_column1 = table.cell(row2_idx, 0).value
                if row2_column1 == '':
                    row2_column2 = table.cell(row2_idx, 1).value
                else:
                    row2_column2 = ''
                    print 'DONT GET ROW2 DATA'

                if row2_column1 == '':
                    row_index = row_index + 2
                else:
                    row_index = row_index + 1

                # 超出範圍時，CHANGE PAGE_INDEX, ROW_INDEX
                if (row_index < page_row_start or row_index > page_row_end):
                    page_idx = page_idx + 1
                    row_index = (page_idx - 1) * 43 + 11 + row_shift
                    row_shift = 0

                # if row_column1 == '9001':
                #     continue

                # 讀 excel 檔轉 入庫單
                tmp = []
                tmp.append(row_column1)  # 商品編號
                tmp.append(row2_column2)  # 參考號
                tmp.append(row_column2)  # 商品名稱
                tmp.append(row_column3)  # 數量
                result.append(tmp)

            success = self.writeXls(LogisticsID, result, outputFile)
        except Exception as e:
            logging.error(e.message)
            resultinfo = e.message
            success = False
        finally:
            if success == False:
                Message = UserID + ' Transfer File Failure , File Path is :'
                self.sendMailToPSC(Message, inputFile)
            return json.dumps({"success": success, "info": resultinfo, "download": outputFile}, sort_keys=False)

    # 判斷轉出出貨格式
    def writeXls(self, LogisticsID, data, outputFile):
        if LogisticsID == 26:
            return self.writeT_catXls(data, outputFile)

    # 寫入出貨單內容
    def writeT_catXls(self, data, outputFile):
        try:
            Template = ExcelTemplate()
            fileTemplate = Template.T_SF_TemplateFile
            rb = xlrd.open_workbook(fileTemplate)
            file = copy(rb)
            table = file.get_sheet(0)
            i = 3
            j = 1
            success = False
            for row in data:

                table.write(i, 0, j)  # 序號
                table.write(i, 1, row[0])  # 貨號
                table.write(i, 2, row[1])  # 參考號
                table.write(i, 3, row[2])  # 商品名稱(中文)
                table.write(i, 7, row[3])  # 預收數量

                i += 1
                j += 1
            logger.debug(outputFile)
            file.save(outputFile)
            success = True
        except Exception as e:
            logging.error(e.message)
            return False
        finally:
            return success

if __name__ == '__main__':
    buy = sfexpressin()
    print buy.parserFile('cbcc3138-5603-11e6-a532-000d3a800878', 'test', 26, 'MS',
                  inputFile=u'C:/Users/10509002/Desktop/鮪魚肚/0120-入庫（PLAY)/0120-入庫（PLAY)/0120進貨單-PLAY.xls', \
                  outputFile=u'C:/Users/10509002/Desktop/0120進貨單-PLAY.xls')