# -*- coding: utf-8 -*-
"""
Created on Mon May 24 14:55:50 2021

@author: Administrator

需pip安装pandas\sqlalchemy\pywin32\chinese_calendar\cx_Oracle\openpyxl(pandas.read_excel依赖)6个模块
"""

from hashlib import sha1
import time,logging
import pandas as pd
from sqlalchemy import create_engine,String,Float,Integer, true
from os.path import join
from win32com.client import Dispatch
import os
import datetime
from datetime import date, timedelta
from chinese_calendar import is_workday
import warnings

warnings.filterwarnings("ignore")

path='./沪硅产业估值.xlsx'


def lwd_find(xlfile):

    for i in range(1,11):
        # lwd = datetime.datetime(2022, 5, 5)- timedelta(i)      # lwd最近工作日
        lwd = datetime.datetime.now()- timedelta(i)          # lwd最近工作日
        if is_workday(lwd):
            lwd = lwd.strftime('%Y/%m/%d')
            print("往前", i, "天", lwd, "是最近工作日, 下面将更新excel估值表")
            xl_update(xlfile,lwd)
            break 


def xl_update(xlfile,lwd):                            # 用win32com组件将out_xls打开再保存，这样用友才能识别其中内容。

    xl = Dispatch("Excel.Application")
    # 后台运行，不显示，不警告
    xl.Visible = False
    xl.DisplayAlerts = False

    xlfile = os.path.abspath(xlfile)
    wb = xl.Workbooks.Open(xlfile)                    # win32不认识相对路径，故需上一句转换为绝对路径。
    sh = wb.Worksheets(1)                             # wb.Worksheets[0]  wb.Worksheets("sheet name")  

    num = 2                                           # 从第2行开始
    while true:
        cell = sh.Cells(num, 1).value
        if cell:                                      # 单元格不为空
            sh.Cells(num,1).Value = lwd               # Cells(2,1) 第2行第1列     
            num = num +1
        else:   
            break                                     # 单元格为空，则 break 跳出 while 循环
    time.sleep(1)
    wb.Save()
    wb.Close()                                        # 关闭表格和excel对象     
    xl.Quit()


def db_init():

    con_string = "oracle+cx_oracle://otc:otc_2021@10.168.4.193:1521/headdb"
    engine=create_engine(con_string) 
    conn=engine.connect()
    sql_col="select * from rdmods.t_table_col_info"
    t_table_col=pd.read_sql(sql_col,engine)
    t_table_col=t_table_col.loc[t_table_col['py_insert_flag']=='1']
    pd.set_option('display.max_columns', None)
    pd.set_option('display.max_rows', None)
    sql_del='delete from  rdmods.{table_name}'  
    return engine, conn, t_table_col, sql_del


def df_dict(t_table_col,table_name,key,value):
    
    t=t_table_col.loc[t_table_col['table_name'].str.endswith(table_name)]    
    return t[[key,value]].set_index(key).to_dict()[value]                           # https://blog.csdn.net/zx1245773445/article/details/103480750


def write_data():

    table_name='SFN11_LIMIT_STOCK_LOMD'
    engine, conn, t_table_col, sql_del = db_init()
    value_key=df_dict(t_table_col,table_name,'column_desc','column_name') 
    col_type=df_dict(t_table_col,table_name,'column_name','py_datetype')    
    sql_del0=sql_del.format(table_name=table_name)
    # conn.execute(sql_del0)     
    col_type={key:eval(value) for key,value in col_type.items()}
    t=pd.read_excel(path,dtype=str)
    t.rename(columns=value_key,inplace=True) 
    t=t.loc[:, :'lomd']
    t.to_sql(table_name,engine,schema='rdmods'
                    ,if_exists='append',index=False,dtype=col_type)
  
            
def run():
    
    print('现在处理"沪硅产业估值表"......')
    logging.info('现在处理"沪硅产业估值表"......')     
    lwd_find(path)
    write_data()   
    print('"沪硅产业估值表"处理结束')
    logging.info('"沪硅产业估值表"处理结束\n')  
    
    
if __name__ == '__main__':
    
    start_time= time.time()
    run()
    print("程序运行完毕，耗时(秒)：",time.time()-start_time)