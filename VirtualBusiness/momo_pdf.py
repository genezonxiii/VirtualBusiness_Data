# coding: utf-8

from ToMysql import ToMysql
from aes_data import aes_data

from pdfminer.pdfdocument import PDFDocument
from pdfminer.pdfpage import PDFPage
from pdfminer.pdfparser import PDFParser
from pdfminer.pdfinterp import PDFResourceManager, PDFPageInterpreter
from pdfminer.converter import PDFPageAggregator
from pdfminer.layout import LAParams, LTTextBox, LTTextLine, LTFigure
import logging
import json

logger = logging.getLogger(__name__)

class Momo_Pdf():
    Data = None

    def __init__(self):
        pass

    def parse_layout(self, layout, list):
        """
            Function to recursively parse the layout tree.
            return list (1: address, 2: phone, mobile, customer name, 3: invoice_no, invoice_date, unique_number, order_no)
        """

        for lt_obj in layout:
            # 確認每個物件的位置
            # print(lt_obj.__class__.__name__)
            # print(lt_obj.bbox)

            # 可撈取各屬性的值
            # print(lt_obj.width, lt_obj.height, )
            # print(lt_obj.x0, lt_obj.y0, lt_obj.x1, lt_obj.y1, )

            if isinstance(lt_obj, LTTextBox) or isinstance(lt_obj, LTTextLine):
                # print(lt_obj.get_text())

                # receiver block: address
                # 左半邊
                if (lt_obj.x0 >= 45 and lt_obj.x0 <= 70) \
                        and ((lt_obj.y0 >= 630 and lt_obj.y0 <= 650) or (lt_obj.y0 >= 210 and lt_obj.y0 <= 220)):
                    # print(lt_obj.get_text())
                    list.append(lt_obj.get_text().replace("\n", ""))
                    logger.debug("receiver block: address")

                # receiver block: phone, mobile, customer name,
                if (lt_obj.x0 >= 60 and lt_obj.x0 <= 70) \
                        and ((lt_obj.y0 >= 600 and lt_obj.y0 <= 620) or (lt_obj.y0 >= 180 and lt_obj.y0 <= 190)):
                    # print(lt_obj.get_text())
                    list.append(lt_obj.get_text())
                    logger.debug("receiver block: phone, mobile, customer name,")

                # invoice_no, invoice_date, unique_number, order_no,
                if (lt_obj.x0 >= 35 and lt_obj.x0 <= 40) \
                        and ((lt_obj.y0 >= 30 and lt_obj.y0 <= 50) or (lt_obj.y0 >= 460 and lt_obj.y0 <= 480)):
                    # print(lt_obj.get_text())
                    list.append(lt_obj.get_text())
                    logger.debug("invoice_no, invoice_date, unique_number, order_no,")

            elif isinstance(lt_obj, LTFigure):
                self.parse_layout(lt_obj, list)  # Recursive

    def modifyOrder(self, order_List, encode = True):
        """
            傳入 order_list [order_no, user_name, tel, mobile, address]
            由傳入的訂單編號找出customer_id，更新 tb_customer, tb_sale
        """

        logger.debug("modifyOrder")

        mysqlconnect = ToMysql()
        mysqlconnect.connect()

        getCustomerFromSale = """
                select customer_id, name
                from tb_sale
                where order_no = '%s'
                and group_id = '%s'
            """ % (order_List[0], self.__groupId)


        # 將撈出資料轉為dict, 可以欄位名稱撈取
        # mysqlconnect.cursor.execute(getCustomerFromSale)
        # row = dict(zip(mysqlconnect.cursor.column_names, mysqlconnect.cursor.fetchone()))
        # print (row['customer_id'], row['name'])

        if encode:
            aes = aes_data()
            for i in range(1, 4, 1):
                order_List[i] = aes.AESencrypt("p@ssw0rd", order_List[i], True)
            order_List[4] = aes.AESencrypt("p@ssw0rd", order_List[4].encode("utf-8"), True)

        mysqlconnect.cursor.execute(getCustomerFromSale)
        customerResult = mysqlconnect.cursor.fetchall()

        if customerResult != []:
            #update tb_customer
            logger.debug("update tb_customer")
            updateCustomer = """update tb_customer set
                name = %s, phone = %s, mobile = %s, address = %s
                where customer_id = %s; """
            updateCustomerParam = (order_List[1], order_List[2], order_List[3], order_List[4], customerResult[0][0])
            mysqlconnect.cursor.execute(updateCustomer, updateCustomerParam)

            # update tb_sale
            logger.debug("update tb_sale")
            updateCustomerInSale = """update tb_sale set
                            name = %s where customer_id = %s; """
            updateCustomerInSaleParam = (order_List[1], customerResult[0][0])
            mysqlconnect.cursor.execute(updateCustomerInSale, updateCustomerInSaleParam)

            mysqlconnect.db.commit()
        else :
            logger.debug("tb_sale NOT FOUND order_no=" + order_List[0])

    def getOrder(self, groupId, path):
        """
            parse_layout後，整理客戶資料
            return list format [order_no, user_name, tel, mobile, address]
        """

        try:
            logger.debug("===momo pdf===")
            logger.debug("getOrder")

            success = False
            resultinfo = ""
            totalRows = 0

            self.__groupId = groupId

            fp = open(path, 'rb')

            parser = PDFParser(fp)
            doc = PDFDocument(parser)

            rsrcmgr = PDFResourceManager()
            laparams = LAParams()
            device = PDFPageAggregator(rsrcmgr, laparams=laparams)
            interpreter = PDFPageInterpreter(rsrcmgr, device)

            for page in PDFPage.create_pages(doc):
                interpreter.process_page(page)
                layout = device.get_result()
                list = []
                self.parse_layout(layout, list)

                # for temp in list:
                #     print temp

                size = len(list)

                # pdf分頁，所以總筆數需累加
                totalRows += size / 3
                # print size
                for i in range(0, size / 3, 1):
                    address = ""
                    # print 'first>>>'
                    # print (list[3 * i])
                    address = list[3 * i]

                    # print 'second>>>'
                    # print (list[3 * i + 1])
                    temp = list[3 * i + 1]
                    temp_list = temp.split("\n")

                    tel = ""
                    mobile = ""
                    user_name = ""

                    if len(temp_list) == 3:
                        temp_list2 = temp_list[0].encode('utf8').split("/")
                        if len(temp_list2) == 2:
                            tel = temp_list2[0]
                            mobile = temp_list2[1]

                    if len(temp_list) == 3:
                        user_name = temp_list[1].encode('utf8')

                    # print 'third>>>'
                    # print (list[3 * i + 2])

                    temp = list[3 * i + 2]
                    temp_list = temp.split("\n")
                    order_no = ""
                    if len(temp_list) == 5:
                        order_no = temp_list[-2].encode('utf8').replace("訂單編號:", "")

                    order_info_list = [order_no, user_name, tel, mobile, address]
                    self.modifyOrder(order_info_list, True)

            success = True

        except Exception as inst:
            logger.error(inst.args)
            resultinfo = inst.args
        finally:
            logger.debug('===momo pdf finally===')
            return json.dumps({"success": success, "info": resultinfo, "total": totalRows}, sort_keys=False)

if __name__ == '__main__':
    # print 'start'
    path = u'D:\\vbdata\\vbupload\\momo\\A1102_3_1_010031_20160301110142.pdf'
    momo = Momo_Pdf()
    momo.getOrder("cbcc3138-5603-11e6-a532-000d3a800878", path)
    # print 'end'


