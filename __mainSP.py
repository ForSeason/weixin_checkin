import itchat, os, re, xlrd, xlwt, time
from xlutils.copy import copy
from settings import const

@itchat.msg_register(itchat.content.TEXT, isGroupChat = True)
def main(msg):
    if int(time.strftime("%H", time.localtime())) == 0:
        init()
    text  = (msg.text)
    text  = text.replace(' ', '')
    d1    = list(re.findall(const.pattern1, text)[0]) if (re.match(const.pattern1, text)) else []
    repl  = False
    if len(d1) == 3:
        d1[1] = d1[1].upper()
        d1[2] = d1[1] + '-' + d1[2]
        d2    = re.split(const.pattern1, text)[-1]
        if re.match(const.pattern3, d2):
            d2    = re.findall(const.pattern3, re.split(const.pattern1, text)[-1])[0]
            d2    = list(d2)
            if len(d2) == 2:
                if re.match(const.pattern2, d2[0]):
                    d2[0] = '√'
                d1.append(d2[0])
                d1.append(d2[1])
                repl = writln(d1)
    if repl:
        itchat.send(d1[0] + d1[2] + '打卡成功', toUserName = msg['FromUserName'])

def init():
    print('Initializing...')
    if os.path.exists(const.path): # check const.path if exists
        print('File ', const.path, ' exists')
        book1 = xlrd.open_workbook(const.path)
        book  = copy(book1)
        if const.sheetname2 not in book1.sheet_names():
            book.add_sheet(const.sheetname2)
            print('Sheet ', const.sheetname2, ' is not in ', const.path)
            book.save(const.path)
            initializeSheet()
    else:
        print('File ', const.path, 'does not exist')
        book  = xlwt.Workbook()
        book.add_sheet(const.sheetname2)
        book.save(const.path)
        initializeSheet()
    print('Initialization done')

def initializeSheet():
    book   = xlrd.open_workbook(const.filename)
    sheet1 = book.sheet_by_name(const.sheetname)
    book1  = xlrd.open_workbook(const.filename2)
    book2  = copy(book1)
    sheet2 = book2.get_sheet(const.sheetname2)
    rows   = sheet1.nrows
    cols   = sheet1.ncols
    for i in range(rows):
        for j in range(cols):
            sheet2.write(i, j, sheet1.cell(i ,j).value)
    book2.save(const.filename2)
    print('Sheet ', const.sheetname2, 'has been initialized')


def writln(d):
    book1  = xlrd.open_workbook(const.path)
    book2  = copy(book1)
    sheet1 = book1.sheet_by_name(const.sheetname2)
    sheet2 = book2.get_sheet(const.sheetname2)
    rows   = sheet1.nrows
    cols   = sheet1.ncols
    repl   = False
    for i in range(rows):
        row = sheet1.row_values(i)
        if (row[0] == d[0] and row[1] == d[1] and row[2] == d[2]):
            rowr = i
            repl = True        
    if repl:
        sheet2.write(rowr, const.day+4, d[3])
        sheet2.write(rowr, 11, d[4])
        print('writing successfully')
    print(d)
    print('====================================================')
    book2.save(const.path)
    return repl


init()
itchat.auto_login(enableCmdQR=2)
itchat.run()
