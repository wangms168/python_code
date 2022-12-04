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

import email
import imaplib
import time


def decode_str(s): #字符编码转换
    
    try:
        subject = email.header.decode_header(s)
    except:       
        return None    
    sub_bytes = subject[0][0] 
    sub_charset = subject[0][1]
    if None == sub_charset:        
        try:        
             subject = str(sub_bytes,'utf8')
        except TypeError:
            try:
                subject = str(sub_bytes,'gbk')
            except TypeError:               
                subject=sub_bytes        
    elif 'unknown-8bit' == sub_charset:
        subject = str(sub_bytes, 'utf8')
    else:
        subject = str(sub_bytes, sub_charset)
    return subject

def get_att(message):       
    for part in message.walk():                      
        if not part.is_multipart():           
            name = part.get_filename()        # 获取附件名称 
            if name:                                               
                fname=decode_str(name)   # 对附件名称进行解码
                attach_data = part.get_payload(decode=True) # 下载附件
                att_file = open('E:\\调自有汇总邮件附件\\' + fname, 'wb') #指定目录下创建文件，注意二进制文件需要用wb模式打开
                att_file.write(attach_data) # 保存附件
                att_file.close()
      

def get_email(uname,pw):
    
    host = "mail.cgws.com" 
    username = uname
    password = pw
    str_name = "自有资金调拨申请|自有资金调拨申请书|自有资金申请|资金调拨申请表|自有资金调回|自有资金调拨|资金调拨|资金调拨申请书|自有资金调回申请|调自有资金|调回自有资金|调回自有|调回头寸|自有资金调回申请书|划拨自有资金|划拨资金"

    with imaplib.IMAP4(host, 143) as serv:
        serv.login(username, password)
        serv.select('INBOX')
        typ, data = serv.search(None,"UnSeen")       
        for num in data[0].split()[::-1]:        
            type0, data0 = serv.fetch(num, '(RFC822)')
            #print(type0)
            try:
                message=email.message_from_bytes(data0[0][1])
            except TypeError:
                continue
    
            subject = message.get('subject')        
            subject = decode_str(subject)         
            if subject is None:
                # serv.uid('STORE', num, '-FLAGS', '\SEEN')
                continue
            nameN=0    # 可以用来表示当前读取邮件的初始状态 0为未读
            t_sub = str_name.split('|')
            for sub in t_sub:
                if sub in subject:
                    nameN+=1  # 含有关键字，可以读取
                    # print(sub)
                    print('             '+subject)
                    get_att(message)
                    break
            if nameN==0: # 不含关键字，将状态退回未读
                serv.store(num, '-FLAGS', '\SEEN')
                continue       


            
    
if __name__ == '__main__':
    names = ['heyc@cgws.com','tanjing','fjie']
    passwords = ['Aa*940506','TJ445700112','Fjie168810@']
    start_time= time.time()
    for i in range(len(names)):
        get_email(names[i],passwords[i])
    print(time.time()-start_time)   
