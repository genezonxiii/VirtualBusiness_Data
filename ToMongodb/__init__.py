# -*-  coding: utf-8  -*-
__author__ = '10409003'

from pymongo import MongoClient
class ToMongodb():
    database="db_product4"
    collection=""
    host="localhost"
    port=27017

    def __init__(self):
        pass

    def connect(self):
        self.client = MongoClient(self.host, self.port)
        db  = self.client[self.database]
        self.cursor   = db[self.collection]


    def setDatabase(self,dbname):
        self.database=dbname
        pass
    def setCollection(self,tbname):
        self.collection=tbname
        pass
    def setuser(self,username,pw):
        self.user=username
        self.password=pw
        pass
    def dbClose(self):
        self.client.close()
