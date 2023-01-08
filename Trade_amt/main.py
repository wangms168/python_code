import configparser
from chinese_calendar import is_workday
import SH_stock
import SH_fund
import SZ

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
dates = config['dates']['dates']
if dates == '':
    print("无自定义日期")
    date = None
    getdata(date)
else:
    print("有自定义日期")
    dates = config['dates']['dates'].split(',')
    print(dates)
    for date in dates:
        print('==================================================')
        print(date)
        print('==================================================')
        getdata(date)