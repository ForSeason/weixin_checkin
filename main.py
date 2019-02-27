import itchat, os, re, xlrd, xlwt
from xlutils.copy import copy
from settings import const

@itchat.msg_register(itchat.content.TEXT, isGroupChat = True)
def main(msg):
    text  = (msg.text)
    text  = text.replace(' ', '')
    d1    = re.findall(const.pattern1, text)
    if len(d1) == 3:
        d1[1] = d1[1].upper()
        d1[2] = d1[1] + '-' + d1[2]
        d2    = re.split(const.pattern1, text)
        if re.match(const.pattern2, d2[-1]):
            d2[-1] = '√'
        d1.append(d2[-1])
        writln(d1)

def init():
    if os.path.exists(const.path):
        book1 = xlrd.open_workbook(const.path)
        book  = copy(book1)
        if const.sheetname2 not in book1.sheet_names():
            book.add_sheet(const.sheetname2)
    else:
        book  = xlwt.Workbook()
        book.add_sheet(const.sheetname2)
    book.save(const.path)
    replace_xls()

def replace_xls():
    book  = xlrd.open_workbook(const.filename)
    sheet1 = book.sheet_by_name(const.sheetname)
    book1  = xlrd.open_workbook(const.filename2)
    book2  = copy(book1)
    sheet2 = book2.get_sheet(const.sheetname2)
    rows = sheet1.nrows
    cols = sheet1.ncols
    for i in range(rows):
        for j in range(cols):
            sheet2.write(i, j, sheet1.cell(i ,j).value)
    book2.save(const.filename2)


def writln(d):
    book1  = xlrd.open_workbook(const.path)
    book2  = copy(book1)
    sheet1 = book1.sheet_by_name(const.sheetname2)
    sheet2 = book2.get_sheet(const.sheetname2)
    rows = sheet1.nrows
    cols = sheet1.ncols
    repl   = False
    for i in range(rows):
        row = sheet1.row_values(i)
        if (row[0] == d[0] and row[1] == d[1] and row[2] == d[2]):
            rowr = i
            repl = True        
    if repl:
        sheet2.write(rowr, const.day+4, d[3])
        print('writing successfully')
    print(d)
    print('====================================================')
    book2.save(const.path)


init()
itchat.auto_login(enableCmdQR=False, hotReload=True)
itchat.run()
