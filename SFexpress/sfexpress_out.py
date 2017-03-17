# -*-  coding: utf-8  -*-
# __author__ = '10408001'
import datetime,time,json
import logging
import xlrd,xlwt
from ToMysql import ToMysql
import uuid
from VirtualBusiness import Customer,updateCustomer
from xlutils.copy import copy
import smtplib
from email.mime.application import MIMEApplication
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from os.path import basename
import re
from os import listdir
from os.path import isfile, join

logger = logging.getLogger(__name__)
# from GroupBuy import ExcelTemplate
#FilePath
class ExcelTemplate():
    T_SF_OutputFilePath, T_SF_TemplateFile = None, None
    MailSender , MailReceiver , SMTPServer = None , None , None
    def __init__(self):
        self.T_SF_TemplateFile = '/data/vbGroupbuy/Logistics_SF_Out.xls'
        self.T_SF_OutputFilePath = '/data/vbSF_output/'
        # Aber 正式用
        self.MailSender = 'pscaber@cloud.pershing.com.tw'
        self.MailReceiver =['joeyang@pershing.com.tw','hsuanmeng@pershing.com.tw','christinewei@pershing.com.tw']
        self.SMTPServer = 'cloud-pershing-com-tw.mail.protection.outlook.com'
        # Local 測試用
        # self.MailSender = 'hsuanmeng@pershing.com.tw'
        # self.MailReceiver = ['joeyang@pershing.com.tw', 'hsuanmeng@pershing.com.tw', 'christinewei@pershing.com.tw']
        # self.SMTPServer = 'ms1.pershing.com.tw'

#Send Mail
class SendMail():
    def __init__(self):
        pass

    def sendmail(self, server, sender, receivers, Message, filename):
        try:
            msg = MIMEMultipart()
            msg['Subject'] = u'錯誤檔案'
            msg["From"] = sender
            msg["To"] = ', '.join(receivers)
            msg.preamble = 'This is test mail'
            body = Message + filename
            # body = "Python test mail"
            msg.attach(MIMEText(body, _charset='utf-8'))
            # msg.attach(MIMEText(Message, 'plain'))
            with open(filename, "rb") as fil:
                part = MIMEApplication(
                    fil.read(),
                    Name=basename(filename)
                )
                part['Content-Disposition'] = 'attachment; filename="%s"' % basename(filename)
                msg.attach(part)

            smtpObj = smtplib.SMTP(server)
            smtpObj.sendmail(sender, receivers, msg.as_string())
            logger.info("Successfully sent email")
        except Exception as e:
            logger.error(e.message)
            logger.error("Error: unable to send email")

#順豐
class sfexpressout():
    mysqlconnect = None
    customer = None
    GroupID , UserID , ProductCode = None , None , None
    RegularEX = None
    def __init__(self):
        self.mysqlconnect = ToMysql()
        self.mysqlconnect.connect()
        self.RegularEX = []
    # 把 Excel 欄位中的 text:u 等字元 replace
    def ConvertText(self, SourceString):
        return SourceString.replace("text:u", "").replace("'", "")

    # 中文數字轉換成阿拉伯數字
    def getResultForDigit(self, a, encoding="utf-8"):
        Numdict = {u'零': 0, u'一': 1, u'二': 2, u'三': 3, u'四': 4, u'五': 5, u'六': 6, u'七': 7, u'八': 8, u'九': 9,\
                   u'十': 10, u'0': 0, u'1': 1, u'2': 2, u'3': 3, u'4': 4, u'5': 5, u'6': 6, u'7': 7, u'8': 8, u'9': 9,\
                   u'10': 10, u'百': 100, u'千': 1000, u'万': 10000, u'０': 0, u'１': 1, u'２': 2, u'３': 3, u'４': 4,\
                   u'５': 5, u'６': 6,  u'７': 7, u'８': 8, u'９': 9, u'壹': 1, u'貳': 2, u'參': 3, u'肆': 4, u'伍': 5,\
                   u'陸': 6, u'柒': 7, u'捌': 8, u'玖': 9, u'拾': 10, u'佰': 100, u'仟': 1000, u'萬': 10000, u'億': 100000000}
        if isinstance(a, str):
            a = a.decode(encoding)
        count = 0
        result = 0
        tmp = 0
        Billion = 0
        while count < len(a):
            tmpChr = a[count]
            # print tmpChr
            tmpNum = Numdict.get(tmpChr, None)
            # 如果等于1亿
            if tmpNum == 100000000:
                result = result + tmp
                result = result * tmpNum
                # 获得亿以上的数量，将其保存在中间变量Billion中并清空result
                Billion = Billion * 100000000 + result
                result = 0
                tmp = 0
            # 如果等于1万
            elif tmpNum == 10000:
                result = result + tmp
                result = result * tmpNum
                tmp = 0
            # 如果等于十或者百，千
            elif tmpNum >= 10:
                if tmp == 0:
                    tmp = 1
                result = result + tmpNum * tmp
                tmp = 0
            # 如果是个位数
            elif tmpNum is not None:
                tmp = tmp * 10 + tmpNum
            count += 1
        result = result + tmp
        result = result + Billion
        return result

    #取 Excel 欄位中某些值
    def ReplaceField(self,SourceString,ReplaceName,intLen=0):
        location = SourceString.find(ReplaceName)
        return SourceString[:location-intLen]

    #初始記錄解析 Log
    def init_log(self,dataName,GroupID,UserID ,ProductCode,inputFile):
        logging.basicConfig(filename='/data/VirtualBusiness_Data/pyupload-group.log',
                            level=logging.DEBUG,
                            format='%(asctime)s - %(levelname)s - %(filename)s:%(name)s:%(module)s/%(funcName)s/%(lineno)d - %(message)s',
                            datefmt='%Y/%m/%d %I:%M:%S %p')
        logging.Formatter.converter = time.gmtime

        logging.info('===' + dataName + '===')
        logging.debug('GroupID:' + GroupID)
        logging.debug('path:' + inputFile)
        logging.debug('UserID:' + UserID)
        self.GroupID = GroupID
        self.customer = Customer()
        self.UserID = UserID
        self.ProductCode = ProductCode

    #解析原始檔
    def parserFile(self, GroupID, UserID, LogisticsID=26, ProductCode=None, inputFile=None, outputFile=None):
        self.init_log('SFexpressOut_Data',GroupID, UserID , ProductCode , inputFile)
        success = False
        try:
            self.getRegularEx(GroupID)
            onlyfiles = [f for f in listdir(inputFile) if isfile(join(inputFile, f))]
            result = []
            for filename in onlyfiles:
                fullpath = inputFile + "/" + filename
                data = xlrd.open_workbook(fullpath)
                table = data.sheets()[0]
                rows = table.nrows
                resultinfo = ""

                order_no = table.cell(4, 10).value
                receiver = table.cell(4, 4).value
                address = table.cell(8, 1).value
                phone = table.cell(5, 4).value
                client = table.cell(4, 1).value
                mobile = None
                if phone[0:2] == '09' :
                    mobile = phone
                    phone = None

                row_index = 11
                page_idx = 1
                while row_index < rows:
                    page_row_start = (page_idx - 1) * 43 + 11
                    page_row_end = (page_idx - 1) * 43 + 35

                    # 取資料ROW 1
                    row_column1 = table.cell(row_index, 0).value

                    if row_column1 == u'付款條件':
                        # pay = row_index + 1
                        # column2 = table.cell(pay, 1).value
                        # if column2.find(u'順豐') != -1:
                        #     price = int(table.cell(pay, 10).value)
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

                    if row_column1 == '9001':
                        pay = 'N'
                        for temp in result:
                            temp.append(pay) # 付款
                            temp.append("")  # 金額
                        continue
                    elif row_column1 == '9002':
                        pay = 'Y'
                        price = table.cell(row_index+2, 10).value
                        for temp in result:
                            temp.append(pay) # 付款
                            temp.append(price)  # 金額
                        continue

                    # 讀 excel 檔轉 出庫單
                    tmp=[]
                    tmp.append(order_no)           #訂單編號
                    tmp.append(client)  # 客戶
                    tmp.append(address)  # 送貨地址
                    tmp.append(row_column1) # 商品編號
                    tmp.append(row_column2) # 名稱
                    tmp.append(row_column3) # 數量
                    tmp.append(receiver)  # 收件人
                    tmp.append(phone)  #　電話
                    tmp.append(mobile) # 手機
                    # tmp.append(price) # 金額
                    # tmp.append(pay)
                    result.append(tmp)

            success = self.writeXls(LogisticsID,result,outputFile)
        except Exception as e :
            logging.error(e.message)
            resultinfo = e.message
            success = False
        finally:
            if success == False :
                Message = UserID + ' Transfer File Failure , File Path is :'
                self.sendMailToPSC(Message,inputFile)
            return json.dumps({"success": success, "info": resultinfo,"download": outputFile}, sort_keys=False)

    # 判斷轉出出貨格式
    def writeXls(self,LogisticsID,data,outputFile):
        if LogisticsID == 26 :
            return self.writeT_catXls(data,outputFile)

    #寫入出貨單內容
    def writeT_catXls(self,data,outputFile):
        try:
            Template = ExcelTemplate()
            fileTemplate = Template.T_SF_TemplateFile
            rb = xlrd.open_workbook(fileTemplate,formatting_info=True)
            logger.debug("copy start")
            file = copy(rb)
            logger.debug("copy end")
            table = file.get_sheet(0)
            i = 1
            success = False
            for row in data:
                d1 = '886DCA'
                d2 = '8860528883'
                table.write(i,0,d1)             # 倉庫ID
                table.write(i, 1,d2)            # 貨主ID
                table.write(i, 2, d2)             # 月結帳號
                table.write(i,6,row[0])         # 訂單編號
                table.write(i, 9, row[1])       # 收件公司
                table.write(i, 13, row[2])       # 地址
                # table.write(i, 14, row[9])      #是否貨到付款
                # table.write(i, 15, row[10])      # 代收貨款金額
                table.write(i, 18, row[3])       # 商品編號
                table.write(i, 19, row[4])       # 商品名稱
                table.write(i, 20, row[5])  # 商品出庫數量
                table.write(i, 23, row[1])      # 收件人
                table.write(i, 24, row[7])      # 電話
                table.write(i, 25, row[8])      # 手機

                i += 1
            logger.debug(outputFile)
            file.save(outputFile)
            success = True
        except Exception as e :
            logging.error(e.message)
            return False
        finally:
            return success

    #將訂購人資料寫入 DB
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
                    self.customer.getEmail(), self.customer.getPost(), self.customer.getClass(), self.customer.getMemo(), self.UserID)
                self.mysqlconnect.cursor.callproc('sp_insert_customer_bysys', CustomereSQL)
            else:
                CustomereSQL = (self.customer.getCustomer_id(), self.customer.getGroup_id(), self.customer.getName(), \
                                self.customer.getAddress(), self.customer.getphone(), self.customer.getMobile(), \
                                self.customer.getEmail(), self.customer.getPost(), self.customer.getClass(), \
                                self.customer.getMemo(),self.UserID)
                self.mysqlconnect.cursor.callproc('sp_update_customer', CustomereSQL)

            CustomereSQL = (self.customer.getCustomer_id(), self.customer.getGroup_id(), self.customer.get_Name(), \
                            self.customer.get_Address(), self.customer.get_phone(), self.customer.get_Mobile(), \
                            self.customer.get_Email())
            updatecustomer.updataData(CustomereSQL)
        except Exception as e :
            print e.message
            logging.error(e.message)
            raise

    #將錯誤檔案寄給 Joe，Avery,Christine
    def sendMailToPSC(self, MailContent , inputFile):
        Template = ExcelTemplate()
        Mail = SendMail()
        Mail.sendmail(Template.SMTPServer,Template.MailSender,Template.MailReceiver,MailContent,inputFile)

    # 取得 Regular Express
    def getRegularEx(self,GroupID):
        try:
            logger.debug("getRegularEx")
            parameter = [GroupID]
            self.mysqlconnect.cursor.callproc('sp_get_regularexpress',parameter)
            # tmp = None
            for result in self.mysqlconnect.cursor.stored_results():
                tmp = result.fetchall()
            for row in tmp:
                value = {"compile": row[0],"search":row[1]}
                self.RegularEX.append(value)
        except Exception as e:
            logging.error(e.message)

    # 使用 RegularExpress 解析"方案名稱"中的數量
    def parserRegularEx(self,value):
        logger.debug("parserRegularEx")
        matchWord = None
        try:
            for row in self.RegularEX:
                logger.debug("for loop")
                logger.debug(row.get("compile"))
                logger.debug(row.get("search"))
                # print row.get("compile"),row.get("search")
                if re.compile(row.get("compile")).match(value):
                    matchWord = re.search(row.get("search"), value)
                if matchWord <> None :
                    break
            logger.debug(matchWord)
            if matchWord:
                found = matchWord.group(1)
                return found
            logger.debug("parserRegularEx end")
        except Exception as e :
            logging.error(e.message)
            print e.message
            raise e

if __name__ == '__main__':
    buy = sfexpressout()
    print buy.parserFile('cbcc3138-5603-11e6-a532-000d3a800878', 'test', 26, 'MS',
                  inputFile=u'C:/Users/10509002/Desktop/鮪魚肚/0301出庫/0301出庫', \
                  outputFile=u'C:/Users/10509002/Desktop/test0315.xls')