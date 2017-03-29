# -*- coding:utf-8 -*-
import requests
import hashlib
from urllib import quote_plus
import chardet

class sendDataSFexpress():
    sfUrl = None
    def __init__(self):
        self.sfUrl = "http://bsp.sit.sf-express.com:8080/bsp-wms/OmsCommons"

    def sendToSFexpress(self,sendData):
        try:
            sendData = sendData.replace("\r\n","").replace("\r","").replace("\n","")
            if chardet.detect(sendData)['encoding'] == "utf-8":
                sendData = sendData.decode("utf-8").encode("utf-8")
            elif chardet.detect(sendData)['encoding'] == "ascii":
                sendData = sendData.encode("utf-8")
            Certification = sendData + '123456'
            m = hashlib.md5()
            m.update(Certification)
            data_digest = m.digest()
            data_digest = quote_plus(data_digest.encode('base64'))
            urlEncode = quote_plus(sendData)
            url = self.sfUrl
            parameter = "logistics_interface=" + urlEncode + "&data_digest=" + data_digest
            result = requests.get(url, parameter)
            return result.text
        except Exception as e :
            print e.message
            raise

if __name__ == "__main__":
    Task = [True]
    data = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Request service="ITEM_QUERY_SERVICE" lang="zh-TW">\
<Head><AccessCode>ITCNC1htXV9xuOKrhu24ow==</AccessCode><Checkword>ANU2VHvV5eqsr2PJHu2znWmWtz2CdIvj</Checkword></Head>\
<Body><ItemQueryRequest><CompanyCode>WYDGJ</CompanyCode><SkuNoList><SkuNo>PY3001ASF</SkuNo></SkuNoList></ItemQueryRequest>\
</Body></Request>'

    if Task[0]:
        sf=sendDataSFexpress()
        print sf.sendToSFexpress(data)