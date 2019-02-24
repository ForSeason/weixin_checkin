import itchat
import re
import xlrd
import xlwt
from xlutils.copy import copy

row = 0

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
    d     = ['班级', '宿舍楼', '宿舍号', '回寝情况']
    book  = xlwt.Workbook()
    sheet = book.add_sheet('s1')
    global row
    row   = 0
    col   = 0
    for title in d:
        sheet.write(row, col, title)
        col += 1
    book.save('output.xls')

def writln(d):
    book1 = xlrd.open_workbook('output.xls')
    book2 = copy(book1)
    sheet = book2.get_sheet(0)
    global row
    row += 1
    col =  0
    for data in d:
        sheet.write(row, col, data)
        col += 1
    book2.save('output.xls')


init()
itchat.auto_login(enableCmdQR=False)
itchat.run()
