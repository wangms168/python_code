import requests
import re
import json
import os
import pprint

pp = pprint.PrettyPrinter(indent=2)         # indent：定义几个空格的缩进

headers = {
    'Accept': 'application/json, text/javascript, */*; q=0.01',
    'Accept-Language': 'zh-CN,zh;q=0.9,en;q=0.8,en-GB;q=0.7,en-US;q=0.6',
    'Content-Type': 'application/json',
    'Proxy-Connection': 'keep-alive',
    'Referer': 'http://www.szse.cn/market/overview/index.html',
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/108.0.0.0 Safari/537.36 Edg/108.0.1462.54',
    'X-Request-Type': 'ajax',
    'X-Requested-With': 'XMLHttpRequest',
}

params = {
    'SHOWTYPE': 'JSON',
    'CATALOGID': '1803_sczm',
    'TABKEY': 'tab1',
    'txtQueryDate': '2023-01-05',
    # 'random': '0.08277726577763445',
}

def getHTMLText(date):
    params.update({"txtQueryDate": date})
    # print(params)
    try:
        r = requests.get('http://www.szse.cn/api/report/ShowReport/data', params=params, headers=headers, verify=False, timeout=30)
        r.raise_for_status()                # 如果状态不是200.引发HTTPError异常
        r.encoding = r.apparent_encoding
        return r.text
    except:
        print("网页异常!")
        return "网页异常"

def textParse(text):
    if text == "网页异常":
        os._exit(0)
    lst = json.loads(text)
    证券类别统计 = lst[0]
    # pp.pprint(证券类别统计)
    data = 证券类别统计['data']
    # pp.pprint(data)

    股票_成交金额    = float(data[0]['cjje'].replace(',', ''))
    主板B股_成交金额 = float(data[2]['cjje'].replace(',', ''))
    A股_成交金额     = 股票_成交金额 - 主板B股_成交金额
    基金_成交金额    = float(data[4]['cjje'].replace(',', ''))

    return 股票_成交金额, 主板B股_成交金额, A股_成交金额, 基金_成交金额


if __name__ == "__main__":
    dates = ['2023-01-05', '2023-01-04', '2023-01-03']
    for date in dates:
        print(date)
        text = getHTMLText(date)
        股票_成交金额, 主板B股_成交金额, A股_成交金额, 基金_成交金额 = textParse(text)
        print('股票_成交金额    = ', 股票_成交金额)
        print('主板B股_成交金额 = ', 主板B股_成交金额)
        print('A股_成交金额     = ', A股_成交金额)
        print('基金_成交金额    = ', 基金_成交金额)
        print('===========================================================================')
