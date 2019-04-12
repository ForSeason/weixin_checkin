import itchat, os, re, xlrd, xlwt, time
from xlutils.copy import copy
from settings import settings

@itchat.msg_register(itchat.content.TEXT, isGroupChat = True)
def main(msg):
    reply = False
    text  = (msg.text)
    text  = text.replace(' ', '')
    d1    = list(re.findall(const.pattern1, text)[0]) if (re.match(const.pattern1, text)) else []
    if len(d1) == 3:
        d1[1] = d1[1].upper()
        d1[2] = d1[1] + '-' + d1[2]
        d2    = re.split(const.pattern1, text)
        if re.match(const.pattern2, d2[-1]):
            d2[-1] = '√'
        d1.append(d2[-1])
        init()
        reply = writln(d1)
    if reply:
        response = d1[0] + d1[2] + '打卡成功'
        if int(time.strftime("%H", time.localtime())) not in range(12, 23):
            if len(lateList()) > 0:
                response += '\n'
                response += '截至' + time.strftime("%H:%M", time.localtime()) + ', 以下宿舍仍未上报:\n'
                response += ', '.join(lateList())
        itchat.send(response, toUserName = msg['FromUserName'])
        

def init():
    print('Initializing...')
    const.reflesh()
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
        sheet2.write(rowr, const.fakeday + 4, d[3])
        print('writing successfully')
    print(d)
    print('====================================================')
    book2.save(const.path)
    return True

def lateList():
    book1  = xlrd.open_workbook(const.path)
    sheet1 = book1.sheet_by_name(const.sheetname2)
    rows   = sheet1.nrows
    cols   = sheet1.ncols
    res    = []
    for i in range(rows):
        row = sheet1.row_values(i)
        if (row[const.fakeday + 4] == ''):
            res.append(row[0] + row[2])
    return res
const = settings()
init()
itchat.auto_login()
itchat.run()
