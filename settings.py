import const
import time
from datetime import datetime

const.pattern1 = '光一|光二|应物|严班|[Cc][14]|\d{3}'
const.pattern2 = '全员归寝无异常|全员回寝无异常|^全?齐$'
const.path = 'data.xls'
const.week0 = 7
const.sheetname = 'template'
# const.sheetname2 defined in main.py
const.filename = 'template.xls'
const.filename2 = const.path



const.day = datetime.now().weekday()
const.week = int(time.strftime("%W")) - const.week0
const.sheetname2 = str(const.week)
