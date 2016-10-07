# -*- coding: utf-8 -*-
import web
import json
from Shipper import ShipperData,VBsale_Analytics

urls = ("/analytics/(.*)", "Analytics")
app = web.application(urls,globals())

class Analytics():
    def GET(self, name):
        data=name.split('&')
        for i in range(len(data)):
            data[i]=data[i][5:len(data[i])].decode('base64')
            print data[i]
        result=[]
        sp = VBsale_Analytics()
        if data[0]=='TOP10':
            result=sp.get_buyer_top10(data[1],data[2],data[3])
        else:
            result=sp.get_buyer_channel(data[1],data[2],data[3],data[4])
        
        print data[1],data[2],data[3]
        web.header('Content-Type', 'text/json; charset=utf-8', unique=True)
        resule = json.dumps(result)
        return resule
        
if __name__ == "__main__":
    app.run()

