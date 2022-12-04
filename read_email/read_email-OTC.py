# -*- coding: utf-8 -*-
"""
Created on Thu Feb 24 16:17:02 2022

@author: admin
"""

# -*- coding: utf-8 -*-
"""
Created on Mon May 24 14:55:50 2021

@author: Administrator
"""

import imaplib,email,time,re,os
import pandas as pd
from dateutil.parser import parse,ParserError
from datetime import date,timedelta,datetime
from os.path import  isdir,join,exists,dirname,basename
from os import  mkdir,listdir,remove


pd.set_option('display.max_columns', None)
pd.set_option('display.max_rows', None)
date0=date.today()
path = '//10.100.131.159/shared/数据/估值表'

def decode_str(s):
    
    # print(s)    
    try:
        subject = email.header.decode_header(s)
        # print(subject)
    except:       
        return None    
    sub_bytes = subject[0][0] 
    sub_charset = subject[0][1]
    
    # print(sub_charset)
    if sub_charset in (None,'unknown-8bit'):        
        try:        
             subject = str(sub_bytes,'utf8')             
        except TypeError:
            try:
                subject = str(sub_bytes,'gbk')
            except TypeError:               
                subject=sub_bytes  
    else:
        subject = str(sub_bytes, sub_charset,errors='replace')
    return subject





def get_att(message,name_cp):    
    
    for part in message.walk():                      
        if not part.is_multipart():           
            name = part.get_filename()         
            if name:                                               
                fname=decode_str(name)
                attach_data = part.get_payload(decode=True)
                dir_cp_name=join(path,name_cp)
                if not exists(dir_cp_name):
                    mkdir(dir_cp_name)
                if fname.endswith('.pdf'):
                    print('    '+fname)                    
                    fname_path=join(path,name_cp,fname)
                    with open(fname_path, 'wb') as f:                
                        f.write(attach_data)                 
                        
                        

def get_email():
    
    host = "mail.cgws.com" 
    username = ""
    password = ""
        
    with imaplib.IMAP4(host, 143) as serv:
        serv.login(username, password)
        serv.select("INBOX")
        typ, data = serv.search(None, "All") 
        print(data)
        count = 0
        numlist=data[0].split()      
        for num in numlist[-500:]:
            # print(num)
            type0, data0 = serv.fetch(num, '(RFC822)')            
            
            try:
                message=email.message_from_bytes(data0[0][1])
            except TypeError:
                continue           
            date11= message.get('date')
            try:         
                date1=parse(date11)            
            except ParserError:
                if date11.endswith('+0800 (GMT+08:00)'):
                    GMT_FORMAT='%a, %d %b %Y %H:%M:%S +0800 (GMT+08:00)'
                    date1=datetime.strptime(date11, GMT_FORMAT)
                else:
                    continue   
                # 此处可能有bug,忽略掉了无法解析日期的邮件
            if date1.date()<date0-timedelta(days=11):                
                continue
            date1 = date1.replace(tzinfo=None)
   
            #此处向前取12天的邮件，确保每次能取到至少两个交易日的邮件      
            subject = message.get('subject')
            # print(subject)
            subject = decode_str(subject) 
            if subject is None:
                continue         
            get_att(message)        
        print(count) 