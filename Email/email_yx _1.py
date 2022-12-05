import configparser
import time
from datetime import date, timedelta
import os
from imaplib import IMAP4
# from imaplib_imapobj import create_imapObj
from imaplib_folder_parse import folder_parse
from pprint import pprint

import email
import email.parser

config = configparser.ConfigParser()
config.read([os.path.expanduser('docs/config.cfg')] ,encoding='utf-8')

def header_decode(header):
    [(text, encoding)] = email.header.decode_header(header)
    # 在不转换字符集的情况下对消息标头值进行解码。 header 为标头值。这个函数返回一个 (decoded_string, charset) 对的列表，
    # 其中包含标头的每个已解码部分。 对于标头的未编码部分 charset 为 None，在其他情况下则为一个包含已编码字符串中所指定字符集名称的小写字符串。
    if isinstance(text, bytes):
        text = text.decode(encoding or "us-ascii")
    return text

def get_att(Obj, uid, msg):       
    """
    下载邮件中的附件
    """
    attachments = []
    for part in msg.walk():
        if part.get_content_maintype() == 'multipart':
            continue
        if part.get('Content-Disposition') is None:
            continue
        fileName = part.get_filename()
        fileName = header_decode(fileName)
        print('fileName', fileName)

        # 只获取指定拓展名的附件
        extension = os.path.splitext(os.path.split(fileName)[1])[1]
        if extension not in exts:
            continue

        # 如果获取到了文件，则将文件保存在指定的目录下
        if fileName:
            filePath = os.path.join(savedir, fileName)
            print('filePath:', filePath)
            if os.path.isfile(filePath):
                print("文件已存在，不下载啦！")
            else:
                if not os.path.exists(savedir):
                    os.makedirs(savedir)
                fp = open(filePath, 'wb')
                fp.write(part.get_payload(decode=True))
                fp.close()
                attachments.append(fileName)
                print("下载了uid:" + str(uid) + "的附件。")

                Obj.select('&bWZOHJT2iEw-', readonly=False)         # 以非只读方式select指定邮箱文件夹，此时，可改变标志flags     
                Obj.store(uid, '+FLAGS', '\Seen')                   # '+FLAGS', '\Seen' 添加已读标志。'\Seen'已读标志、'\UNSEEN'未读标志

    return attachments


def create_imapObj(hostname, port, username, password, verbose=False):
    if verbose:
        print('Connecting to', hostname, ':', port)
    imapObj = IMAP4(hostname, port)
    if verbose:
        print('Logging in as', username)
    try:
        imapObj.login(username, password)
        print('\nConnect to {0}:{1} successfully'.format(hostname, port))
        return imapObj
    except Exception as err:
        try:
            print('\nConnect to {0}:{1} failed'.format(hostname, port), err)
        finally:
            err = None
            del err

def get_email(hostname, port, username, password, verbose=False):

    with create_imapObj(hostname, port, username, password, verbose=False) as Obj:

# ----------------------------------- folder list ---------------------------------------
        
        (typ, mbox_list) = Obj.list()
        # print('select[typ]:', typ)
        # pprint(mbox_list)
        mbox_name_list = []
        for line in mbox_list:
            flags, delimiter, mbox_name = folder_parse(line)
            mbox_name_list.append(mbox_name)

# ----------------------------------- select folder ---------------------------------------
        
        (typ, msgsTotal_list) = Obj.select('&bWZOHJT2iEw-', readonly=True)          # if readonly="True" you can't change any flags. But,if it is false, you can do as follow,
        # (typ, msgs_total) = Obj.select('INBOX', readonly=True)                # if readonly="True" you can't change any flags. But,if it is false, you can do as follow,
        # IMAP4.select(mailbox='INBOX', readonly=False)
        # 选择一个邮箱。 返回的数据是 mailbox 中消息的数量 (EXISTS 响应)               
        # typ 同样是响应代码；响应数据 msgs_total是一个包含单个字节类型字符串的列表，该单个字符串包含邮箱中的邮件总数。

        print('select[typ]:', typ)
        print('select[msgsTotal_list]:', msgsTotal_list)
        if typ=='NO':
            print("邮箱文件夹不存在")
        elif typ=='OK':
            msgs_num = int(msgsTotal_list[0])
            print('There are {} messages in INBOX'.format(msgs_num))

# ----------------------------------- search mail ---------------------------------------
        
            Date = date.today() - timedelta(days=days)
            Date = Date.strftime("%d-%b-%Y")
            add = "ebank@eb.spdb.com.cn"
            criterion = f'(SINCE {Date} FROM {add})'
            (typ, msgsUids_list) = Obj.search(None, criterion)            # 'Seen'、'UnSeen'、'ALL'、'(BEFORE "01-Jan-2022")'
            # IMAP4.search(charset, criterion[, ...])，其第二个参数形状同status()第二个参数类似。
            # 在邮箱中搜索匹配的消息。 charset 字符集可以为 None    
            # 同c.status()一样，其响应数据也是一个包含单个字节类型(即字符串前面标有b的前缀)字符串的列表，
            # 该字符串是一个由空格分隔的连续消息(邮件)ID组成。

            print('search[typ]:', typ)
            print('search[msgsUids_list]:', msgsUids_list)

            if not msgsUids_list[0]:
                print("未搜索到符合条件的邮件！")
            else:
                uids = msgsUids_list[0].split()[::-1]
                num = len(uids)
                print('按搜索条件search到的邮件总数:', num, '\n')

# ----------------------------------- fetch messages ---------------------------------------
            # 这是按search返回的消息uid列表进行for循环，fetch以单个uid进行获取消息；另外一种是fetch以uid列表字符串str形式获取消息，
            # 其返回的元组的第二个项目是一个 num X 2 个元素的列表list，以msg_data[::2](正好跳过 b')' 这个元素)形式切片形成的以num个”两个元素的列表“组成的长度为num的列表list,
            i = 0
            for uid in uids:        
                (typ, [(msgID_bytes, msgData_bytes), rrb_bytes]) = Obj.fetch(uid, '(RFC822)')       # fetch()返回一个包含两个项目的tuple，第一个项目fetch()[0]是字符串'OK',是响应代码typ；
                # typ,              msg_data                     = Obj.fetch(','.join(uids), '(RFC822)')# fetch()返回一个包含两个项目的tuple，第一个项目fetch()[0]是字符串'OK',是响应代码typ；
                # OK  msg_data[0][0] msg_data[0][1]  msg_data[1]
                #   b'1 (RFC822 {39944}'                b')'
                                                                # 第二个项目fetch()[1]是一个含有两个元素的列表list,是响应数据msg_data。
                                                                    # 第二个项目的第一个元素msg_data[0]是含有两个项目的tuple:
                                                                        # 第一个项目msg_data[0][0]是一个字节类型字符串（b'1 (RFC822 {39944}'）
                                                                        # 第二个项目msg_data[0][1]是一个含有真正大量消息数据的字节类型字符串
                                                                            # email.message_from_bytes(msg_data[0][1])就是从一个 bytes-like object 中返回消息对象message_ojb。 这与 BytesParser().parsebytes(s) 等价。
                                                                            # 再从消息对象中获取get出各消息标头(<class 'str'>)
                                                                            # 再用email.header.decode_header()在不转换字符集的情况下对消息标头值进行解码，返回仅含有一个(decoded_string, charset)这样元素的列表。
                                                                            # 再对decoded_string进行str.decode(charset or "us-ascii")解码，至此解析解码完成。
                                                                    # 第二个项目的第二个元素是一个字节型字符串（ b')'）
                # IMAP4.fetch(message_set, message_parts)取回（部分）信息。“message_ids” 参数是逗号分隔的 ID（例如 “ 1”，“ 1,2”” 或 ID 范围（例如 “ 1：2”）列表。 message_parts应该是一串括在圆括号内的消息部分名。，例如: "(UID BODY[TEXT])"。 返回的数据是由消息部分信封和数据组成的元组。
                print('fetch:|', 'typ:', typ, '| msgID_bytes:', msgID_bytes, "| 右圆括号:", rrb_bytes, '|')

# ----------------------------------- 以上是imaplib的事，以下是email的事 ---------------------------------------
                
                if typ == 'NO':
                    print("获取uid="+ uid +"的消息失败！")
                elif typ == 'OK':
                    # for id, msgData_bytes in msg_data[::2]:             # 邮件id序列for循环
                    # 解析出邮件id以便回复邮件状态标志使用
                    # uid = id.split()[0]
                    # print('uid:', uid)

                    # 获取消息对象
                    msgOjb = email.message_from_bytes(msgData_bytes)
                    # email.message_from_bytes(s, _class=None, *, policy=policy.compat32)
                    # 从一个 bytes-like object 中返回消息对象。 这与 BytesParser().parsebytes(s) 等价。

                    # print('msgOjb.keys:', msgOjb.keys())           # msg.keys() https://stackoverflow.com/questions/703185/using-email-headerparser-with-imaplib-fetch-in-python

                    # 从消息对象中提取消息标头
                    Subject = msgOjb['Subject']

                    # 对消息标头进行解码
                    Subject = header_decode(Subject)

                    i+=1
                    print('uid_'+str(i)+':', uid, 'decoded Subject:', Subject)

                    if Subject is None:
                        # serv.uid('STORE', num, '+FLAGS', '\'UnSeen')
                        continue

                    titles = config['other']['titles']
                    titles = titles.split('|')
                    for title in titles:
                        if title in Subject:
                            attachments = get_att(Obj, uid, msgOjb)
                            print('attachments:', attachments, '\n')
                            break           # 含有关键字一次即可

if __name__ == '__main__':
    start_time= time.time()
    config = configparser.ConfigParser()
    config.read([os.path.expanduser('docs/config.cfg')], encoding='utf-8')
    hostname = config['server']['hostname']
    port = config['server']['port']
    usernames = config['account']['username'].split(',')
    passwords = config['account']['password'].split(',')
    exts = config['other']['exts'].split('|')
    savedir = config['other']['savedir']
    days = int(config['other']['days'])
    
    for i in range(len(usernames)):
        get_email(hostname, port, usernames[i], passwords[i], verbose=False)
    print("耗时：", time.time()-start_time)   
