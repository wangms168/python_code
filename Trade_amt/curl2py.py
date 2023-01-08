import uncurl

curlcode = "curl 'http://query.sse.com.cn/commonQuery.do?jsonCallBack=jsonpCallback78283267&sqlId=COMMON_SSE_SJ_GPSJ_CJGK_MRGK_C&PRODUCT_CODE=01%2C02%2C03%2C11%2C17&type=inParams&SEARCH_DATE=2023-01-05&_=1673158819638' \
  -H 'Accept: */*' \
  -H 'Accept-Language: zh-CN,zh;q=0.9,en;q=0.8,en-GB;q=0.7,en-US;q=0.6' \
  -H 'Connection: keep-alive' \
  -H 'Cookie: ba17301551dcbaf9_gdp_user_key=; gdp_user_id=gioenc-01bg0e2e%2C1754%2C59a1%2C9cad%2C1c4g6c7gc833; ba17301551dcbaf9_gdp_session_id_bf8cb27c-af55-4bdf-8aa9-26288cf95c74=true; yfx_c_g_u_id_10000042=_ck23010612135116893708725750227; ba17301551dcbaf9_gdp_session_id_2adaff12-b03b-4afc-9b39-6143a589bc64=true; ba17301551dcbaf9_gdp_session_id_c7928f5e-a161-4f14-a0ed-80799c67140f=true; ba17301551dcbaf9_gdp_session_id_1ec9e589-720f-4290-9800-d227769a4e7b=true; ba17301551dcbaf9_gdp_session_id_a73cb5bf-ff4c-4760-abb1-e3535d6a11ef=true; ba17301551dcbaf9_gdp_session_id_5f569a3e-b049-4fd0-820b-a1ed74b470aa=true; ba17301551dcbaf9_gdp_session_id_685a4f25-490a-4cb6-95e2-6247fb4110f0=true; ba17301551dcbaf9_gdp_session_id_d696b732-7a7b-4517-bb5c-ea742a374900=true; ba17301551dcbaf9_gdp_session_id_70ad4fd1-7c14-4124-ac55-dd5798d6fca4=true; ba17301551dcbaf9_gdp_session_id_0b787749-5bcc-41e1-9b2a-b417efbb6e86=true; ba17301551dcbaf9_gdp_session_id_33c720fa-8348-48eb-8eaf-78c20a5a93e1=true; ba17301551dcbaf9_gdp_session_id=2bd2f88c-d269-469e-b217-82d6fc0cf979; ba17301551dcbaf9_gdp_session_id_2bd2f88c-d269-469e-b217-82d6fc0cf979=true; yfx_f_l_v_t_10000042=f_t_1672978431664__r_t_1673155256335__v_t_1673158819503__r_c_2; ba17301551dcbaf9_gdp_sequence_ids={%22globalKey%22:116%2C%22VISIT%22:13%2C%22PAGE%22:19%2C%22VIEW_CLICK%22:86}' \
  -H 'Referer: http://www.sse.com.cn/' \
  -H 'User-Agent: Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/108.0.0.0 Safari/537.36 Edg/108.0.1462.76' \
  --compressed \
  --insecure"
# pycode= uncurl.parse(curlcode)
# print(pycode)
context = uncurl.parse_context(curlcode)
print(context.headers)
