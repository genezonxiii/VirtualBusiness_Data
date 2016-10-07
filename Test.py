from Shipper import ShipperData,VBsale_Analytics
import datetime
from time import mktime


if __name__=='__main__':
    
    sp=VBsale_Analytics()
    data = sp.get_buyer_top10('cbcc3138-5603-11e6-a532-000d3a800878','2016-05-01','2016-08-31')
    for row in data:
        print row
    data = sp.get_buyer_channel('cbcc3138-5603-11e6-a532-000d3a800878','2016-05-01','2016-08-31','test')
    print 'sp_buyer_channel'
    for row in data:
        print row
