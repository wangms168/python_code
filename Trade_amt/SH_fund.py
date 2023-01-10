import requests
import re
import json
import os
import pprint

pp = pprint.PrettyPrinter(indent=2)         # indent：定义几个空格的缩进

cookies = {
    'ba17301551dcbaf9_gdp_user_key': '',
    'gdp_user_id': 'gioenc-01bg0e2e%2C1754%2C59a1%2C9cad%2C1c4g6c7gc833',
    'ba17301551dcbaf9_gdp_session_id_bf8cb27c-af55-4bdf-8aa9-26288cf95c74': 'true',
    'yfx_c_g_u_id_10000042': '_ck23010612135116893708725750227',
    'ba17301551dcbaf9_gdp_session_id_2adaff12-b03b-4afc-9b39-6143a589bc64': 'true',
    'ba17301551dcbaf9_gdp_session_id_c7928f5e-a161-4f14-a0ed-80799c67140f': 'true',
    'ba17301551dcbaf9_gdp_session_id_1ec9e589-720f-4290-9800-d227769a4e7b': 'true',
    'ba17301551dcbaf9_gdp_session_id': 'a73cb5bf-ff4c-4760-abb1-e3535d6a11ef',
    'ba17301551dcbaf9_gdp_session_id_a73cb5bf-ff4c-4760-abb1-e3535d6a11ef': 'true',
    'yfx_f_l_v_t_10000042': 'f_t_1672978431664__r_t_1673058380499__v_t_1673060859162__r_c_1',
    'ba17301551dcbaf9_gdp_sequence_ids': '{%22globalKey%22:54%2C%22VISIT%22:6%2C%22PAGE%22:10%2C%22VIEW_CLICK%22:40}',
}

headers = {
    'Accept': '*/*',
    'Accept-Language': 'zh-CN,zh;q=0.9,en;q=0.8,en-GB;q=0.7,en-US;q=0.6',
    'Connection': 'keep-alive',
    # 'Cookie': 'ba17301551dcbaf9_gdp_user_key=; gdp_user_id=gioenc-01bg0e2e%2C1754%2C59a1%2C9cad%2C1c4g6c7gc833; ba17301551dcbaf9_gdp_session_id_bf8cb27c-af55-4bdf-8aa9-26288cf95c74=true; yfx_c_g_u_id_10000042=_ck23010612135116893708725750227; ba17301551dcbaf9_gdp_session_id_2adaff12-b03b-4afc-9b39-6143a589bc64=true; ba17301551dcbaf9_gdp_session_id_c7928f5e-a161-4f14-a0ed-80799c67140f=true; ba17301551dcbaf9_gdp_session_id_1ec9e589-720f-4290-9800-d227769a4e7b=true; ba17301551dcbaf9_gdp_session_id=a73cb5bf-ff4c-4760-abb1-e3535d6a11ef; ba17301551dcbaf9_gdp_session_id_a73cb5bf-ff4c-4760-abb1-e3535d6a11ef=true; yfx_f_l_v_t_10000042=f_t_1672978431664__r_t_1673058380499__v_t_1673060859162__r_c_1; ba17301551dcbaf9_gdp_sequence_ids={%22globalKey%22:54%2C%22VISIT%22:6%2C%22PAGE%22:10%2C%22VIEW_CLICK%22:40}',
    'Referer': 'http://www.sse.com.cn/',
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/108.0.0.0 Safari/537.36 Edg/108.0.1462.54',
}

params = {
    'jsonCallBack': 'jsonpCallback33752623',
    'sqlId': 'COMMON_SSE_SJ_GPSJ_CJGK_MRGK_C',
    # 'SEARCH_DATE': '2023-01-05',
    'PRODUCT_CODE': '05,13,16,14,15,12',
    'type': 'inParams',
    '_': '1673060859305',
}

def getHTMLText(date):
    params.update({"SEARCH_DATE": date})
    try:
        r = requests.get('http://query.sse.com.cn/commonQuery.do', params=params, cookies=cookies, headers=headers, verify=False, timeout=30)
        r.raise_for_status()                # 如果状态不是200.引发HTTPError异常
        r.encoding = r.apparent_encoding
        return r.text
    except:
        print("上交所网站异常!")
        return "网页异常"

def textParse(text):
    if text == "网页异常":
        os._exit(0)
    dict = json.loads(re.findall(r"[(](.*?)[)]", text)[0])
    # pp.pprint(dict)
    result = dict['result']
    # pp.pprint(result)

    Funds = result[0]
    基金_成交金额 = float(Funds['TRADE_AMT'].replace(',', ''))

    return 基金_成交金额


if __name__ == "__main__":
    dates = ['2023-01-05', '2023-01-04', '2023-01-03']
    # for date in dates:
    #     print(date)
    #     text = getHTMLText(date)
    #     基金_成交金额 = textParse(text)
    #     print('基金_成交金额    = ', 基金_成交金额)
    #     print('===========================================================================')
