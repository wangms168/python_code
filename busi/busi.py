# -*- coding: utf-8 -*-
"""
Created on Mon May 24 14:55:50 2021
@author: Administrator
需pip安装 pandas sqlalchemy，pywin32，chinese_calendar，cx_Oracle，openpyxl(pandas.read_excel依赖)6个模块
"""

import datetime
import logging
import os
import time
import warnings
from datetime import timedelta

import pandas as pd
from chinese_calendar import is_workday
from sqlalchemy import create_engine
import xlwings as xw


def lwd_find(xlfile):
    for i in range(1, 11):
        lwd = datetime.datetime.now() - timedelta(i)  # lwd最近工作日
        if is_workday(lwd):
            lwd = lwd.strftime('%Y/%m/%d')
            print("往前", i, "天", lwd, "是最近工作日, 下面将更新excel估值表")
            xl_update(xlfile, lwd)
            break


def xl_update(xlfile, lwd):

    xlapp = xw.App(visible=False, add_book=False)
    xlapp.display_alerts = False  # 不显示Excel消息框
    xlapp.screen_updating = False  # 关闭屏幕更新,可加快宏的执行速度
    wb = xlapp.books.open(xlfile)
    # sht = wb.sheets[0]  # 按表名获取工作表
    sht = wb.sheets['波动率与流动性折扣率']  # 按表名获取工作表

    num = 2  # 从第2行开始
    while True:
        cell1_value = sht['A'+str(num)].value
        if cell1_value:  # 单元格不为空
            sht['A' + str(num)].value = lwd
            # sht.range(num, 1).value = lwd
            num = num + 1
        else:
            break  # 单元格为空，则 break 跳出 while 循环

    time.sleep(1)
    wb.save()
    wb.close()
    xlapp.quit()


def db_init():
    con_string = "oracle+cx_oracle://otc:otc_2021@10.168.4.193:1521/headdb"
    engine = create_engine(con_string)
    conn = engine.connect()
    sql_col = "select * from rdmods.t_table_col_info"
    t_table_col = pd.read_sql(sql_col, engine)
    t_table_col = t_table_col.loc[t_table_col['py_insert_flag'] == '1']
    pd.set_option('display.max_columns', None)
    # pd.set_option('display.max_rows', None)
    # print(type(t_table_col), t_table_col)
    sql_del = 'delete from  rdmods.{table_name}'
    return engine, conn, t_table_col, sql_del


def df_dict(t_table_col, table_name, key, value):
    t = t_table_col.loc[t_table_col['table_name'].str.endswith(table_name)]
    return t[[key, value]].set_index(key).to_dict()[value]
    # https://blog.csdn.net/zx1245773445/article/details/103480750


def write_data(xlfile):
    table_name = 'SFN11_LIMIT_STOCK_LOMD'
    engine, conn, t_table_col, sql_del = db_init()
    value_key = df_dict(t_table_col, table_name, 'column_desc', 'column_name')
    col_type = df_dict(t_table_col, table_name, 'column_name', 'py_datetype')
    # sql_del0 = sql_del.format(table_name=table_name)
    # conn.execute(sql_del0)     
    col_type = {key: eval(value) for key, value in col_type.items()}
    t = pd.read_excel(xlfile, dtype=str)
    t.rename(columns=value_key, inplace=True)
    t = t.loc[:, :'lomd']
    t.to_sql(table_name, engine, schema='rdmods', if_exists='append', index=False, dtype=col_type)


def run(xlfile):
    print('现在处理"沪硅产业估值表"......')
    logging.info('现在处理"沪硅产业估值表"......')
    lwd_find(xlfile)
    write_data(xlfile)
    print('"沪硅产业估值表"处理结束')
    logging.info('"沪硅产业估值表"处理结束\n')


if __name__ == '__main__':
    start_time = time.time()
    warnings.filterwarnings("ignore")
    XLfile = './沪硅产业估值.xlsx'
    run(XLfile)
    print("程序运行完毕，耗时(秒)：", time.time() - start_time)
