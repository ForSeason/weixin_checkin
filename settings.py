import time
from datetime import datetime

class settings():

    def __init__(self):
        self.week0      = 7
        self.faketime   = time.time() - 43200
        self.fakeday    = datetime.fromtimestamp(self.faketime).weekday()
        self.fakeweek   = int(time.strftime("%W", time.localtime(self.faketime))) - self.week0
        self.pattern1   = '(光一|光二|应物|严班)([Cc][14])(\d{3})'
        self.pattern2   = '全员归寝无异常|全员回寝无异常|^全?齐$'
        self.pattern3   = '(.*?)(\d{11})'
        self.path       = '/var/www/html/documents/dormitory.xls'
        # self.path       = 'dormitory.xls'
        self.sheetname  = 'template'
        self.sheetname2 = str(self.fakeweek)
        self.filename   = '/var/www/html/documents/template.xls'
        # self.filename   = 'template.xls'
        self.filename2  = self.path

    def reflesh(self):
        self.faketime   = time.time() - 43200
        self.fakeday    = datetime.fromtimestamp(self.faketime).weekday()
        self.fakeweek   = int(time.strftime("%W", time.localtime(self.faketime))) - self.week0
        self.sheetname2 = str(self.fakeweek)
