import itchat
import os
import time
import datetime
import re
import xlrd
import xlwt
from xlutils.copy import copy

global row
global path
global nowTime


row  = 0
path = 'output.xls'
nowTime = datetime.datetime.now().strftime('%Y%m%d')


@itchat.msg_register(itchat.content.TEXT)
def main(msg):
    text  = (msg.text)
    text  = text.replace(' ', '')
    d1    = re.findall('光一|光二|应物|严班|[Cc][14]|\d{3}', text)
    if len(d1) == 3:
        d1[1] = d1[1].upper()
        d2    = re.split('光一|光二|应物|严班|[Cc][14]|\d{3}', text)
        d1.append(d2[-1])
        writln(d1)

def init():

    global path
    global nowTime
    global row

    if os.path.exists(path):
        book1 = xlrd.open_workbook(path)
        book  = copy(book1)
        if nowTime in book1.sheet_names():
           sheet = book1.sheet_by_name(nowTime)
           cols  = sheet.col_values(0)
           row   = len(cols) - 1
           return
    else:
        book  = xlwt.Workbook()
    sheet = book.add_sheet(nowTime)
    d   = ['班级', '宿舍楼', '宿舍号', '回寝情况']
    row = 0
    col = 0
    for title in d:
        sheet.write(row, col, title)
        col += 1
    book.save(path)

def writln(d):

    global path
    global row
    global nowTime

    book1  = xlrd.open_workbook(path)
    book2  = copy(book1)
    sheet1 = book2.get_sheet(nowTime)
    sheet2 = book1.sheet_by_name(nowTime)
    repl   = False
    for i in range(row):
        rows  = sheet2.row_values(i)
        if (rows[0] == d[0] and rows[1] == d[1] and rows[2] == d[2]):
            rowr = i
            repl = True
    if repl:
        sheet1.write(rowr, 3, d[3])
    else:
        row += 1
        col =  0
        for data in d:
            sheet1.write(row, col, data)
            col += 1
    book2.save(path)


init()
itchat.auto_login(enableCmdQR=False)
itchat.run()
