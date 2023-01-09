import configparser
import datetime
from datetime import timedelta
from chinese_calendar import is_workday
import SH_stock
import SH_fund
import SZ

def lwd_find():
    for i in range(1, 11):                              #  往前10天内寻找最近工作日
        lwd = datetime.datetime.now() - timedelta(i)    # lwd最近工作日last work date
        if is_workday(lwd):
            lwd = lwd.strftime('%Y-%m-%d')
            print("往前", i, "天", lwd, "是最近工作日")
            print(lwd)
            break
    return lwd

def getdata(date):
    # 上海股票 =========================================================================
    text = SH_stock.getHTMLText(date)
    上海主板A_成交金额, 上海科创版_成交金额 = SH_stock.textParse(text)
    print('上海主板A_成交金额   = ', 上海主板A_成交金额)
    print('上海科创版_成交金额  = ', 上海科创版_成交金额)

    # 上海基金 =========================================================================
    text = SH_fund.getHTMLText(date)
    上海基金_成交金额 = SH_fund.textParse(text)
    print('上海基金_成交金额    = ', 上海基金_成交金额)
    print('--------------------------------------------------')


    # 深圳市场 =========================================================================
    text = SZ.getHTMLText(date)
    深圳股票_成交金额, 深圳主板B股_成交金额, 深圳A股_成交金额, 深圳基金_成交金额 = SZ.textParse(text)
    print('深圳股票_成交金额    = ', 深圳股票_成交金额)
    print('深圳主板B股_成交金额 = ', 深圳主板B股_成交金额)
    print('深圳A股_成交金额     = ', 深圳A股_成交金额)
    print('深圳基金_成交金额    = ', 深圳基金_成交金额)

config = configparser.ConfigParser()
config.read('./config.cfg', encoding='utf-8')
dates = eval(config['dates']['dates'])
if dates:
    print("有自定义日期")
    for date in dates:
        print('==================================================')
        print(date)
        print('==================================================')
        getdata(date)
else:
    print("无自定义日期")
    date = None
    date = lwd_find()
    getdata(date)
