# -*-  coding: utf-8  -*-
__author__ = '10409003'

from aes_data import aes_data

class Insert_Client():
    Data=None
    def __init__(self):
        pass

    def AESen_CustomerData(self, _Name,_address,_phone,_mobile):
        aes = aes_data()
        if _Name== '':
            ClientName=''
        else:
            ClientName = (aes.AESencrypt("p@ssw0rd", _Name, True))
        if _address == '':
            ClientAdd=''
        else:
            ClientAdd = (aes.AESencrypt("p@ssw0rd", _address, True))
        if _phone == '':
            ClientPhone =''
        else:
            ClientPhone = (aes.AESencrypt("p@ssw0rd", _phone, True))
        if _mobile =='':
            ClientMobile=''
        else:
            ClientMobile = (aes.AESencrypt("p@ssw0rd", _mobile, True))
        r = {"name": ClientName, "address": ClientAdd, "phone": ClientPhone, "mobile": ClientMobile}
        return r

class DEncrypt():
    def __init__(self):
        pass

    def encrypt(self,value):
        aes = aes_data()
        result = aes.AESencrypt("p@ssw0rd", value, True)
        r = {"result" : result}
        return r

    def decrypt(self,value):
        aes = aes_data()
        result = aes.AESdecrypt("p@ssw0rd", value, True)
        r = {"result": result}
        return r