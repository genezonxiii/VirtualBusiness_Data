class DBSetting():
    dbHost ,dbName , dbUser , dbPassword = None , None ,None ,None
    def __init__(self):
        self.dbHost = '192.168.112.164'
        self.dbName = 'tmp'
        self.dbUser = 'root'
        self.dbPassword = 'admin123'
        # self.dbHost = 'localhost'
        # self.dbName = 'tmp'
        # self.dbUser = 'root'
        # self.dbPassword = 'mysql'