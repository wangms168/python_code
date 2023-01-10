import configparser
import email
import os
import re
import time
from datetime import date, datetime, timedelta
import datetime
from imaplib import IMAP4
from pprint import pprint
from derow_xw import derow


def header_decode(header):
    [(text, encoding)] = email.header.decode_header(header)
    # 在不转换字符集的情况下对消息标头值进行解码。 header 为标头值。这个函数返回一个 (decoded_string, charset) 对的列表，
    # 其中包含标头的每个已解码部分。 对于标头的未编码部分 charset 为 None，在其他情况下则为一个包含已编码字符串中所指定字符集名称的小写字符串。
    if isinstance(text, bytes):
        text = text.decode(encoding or "us-ascii")
    return text


def get_att(imapObj, uid, msg, mbfolder, s_mbfolder, Subject, date_Ym_old):
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

        # 只获取指定拓展名的附件
        extension = os.path.splitext(os.path.split(fileName)[1])[1]
        if extension not in exts:
            continue

        # 如果获取到了文件，则将文件保存在指定的目录下
        if fileName:
            if mbfolder == '宝城期货':
                fileName = Subject + '_' + fileName[9:]

            date_Ym = re.search('(\d{4}\d{2})\d{2}', fileName).group(1)
            if not date_Ym == date_Ym_old:
                date_Ym_old = date_Ym

            savePath = os.path.join(savedir, '08-IRS客户日结单', date_Ym[:4]+'利率IRS互换日结单', date_Ym
                                    + '利率IRS互换日结单', date_Ym + '利率IRS互换日结单-' + mbfolder)
            if mbfolder == '宝城期货':
                savePath = os.path.join(savedir, '09-股指期货账单', date_Ym[:4]+'期货逐日逐笔单',  date_Ym + '期货逐日逐笔单')
            filePath = os.path.join(savePath, fileName)
            if not os.path.exists(savePath):
                os.makedirs(savePath)

            if os.path.isfile(filePath):
                pass
                # print("文件已存在，不下载啦！")
            else:
                fp = open(filePath, 'wb')
                fp.write(part.get_payload(decode=True))
                fp.close()
                attachments.append(fileName)
                # print("下载了uid:" + str(uid) + "的附件。")

                imapObj.select(s_mbfolder, readonly=False)
                # 以非只读方式select指定邮箱文件夹，此时fetch(search不改变标志flags)可改变标志flags
                imapObj.store(uid, '+FLAGS', '\\Seen')  # '+FLAGS', '\Seen' 添加已读标志。'\Seen'已读标志、'\UNSEEN'未读标志
                imapObj.select(s_mbfolder, readonly=True)   # 添加了标志后，马上恢复以只读方式select邮箱，以备顺序的下次使用。

                if mbfolder == '浦发银行':
                    filePath = os.path.abspath(filePath)  # win32com不认识相对路径，故需转换为绝对路径。
                    derow(filePath)

    return attachments


def do_msg(msgData_bytes, imapObj, uid, mbfolder, s_mbfolder):
    # 获取消息对象
    msgObj = email.message_from_bytes(msgData_bytes)
    # print('msgObj.keys:', msgObj.keys())
    # msg.keys() https://stackoverflow.com/questions/703185/using-email-headerparser-with-imaplib-fetch-in-python

    # 从消息对象中提取消息标头
    Subject = msgObj['Subject']

    # 对消息标头进行解码
    Subject = header_decode(Subject)

    date_Ym_old = None
    for title in titles:
        if title in Subject:
            get_att(imapObj, uid, msgObj, mbfolder, s_mbfolder, Subject, date_Ym_old)
            break                   # 含有关键字一次即可

    return Subject


def create_imapObj(hostname, port, username, password, verbose=False):
    if verbose:
        print('Connecting to', hostname, ':', port)
    imapObj = IMAP4(hostname, port)
    if verbose:
        print('Logging in as', username)
    try:
        imapObj.login(username, password)
        print('\nConnect to {0}:{1} successfully\n'.format(hostname, port))
        return imapObj
    except Exception as err:
        try:
            print('\nConnect to {0}:{1} failed\n'.format(hostname, port), err)
        finally:
            del err


def folder_parse(line):
    folder_pattern = re.compile(
        r'.(?P<flags>.*?). "(?P<delimiter>.*)" (?P<name>.*)'
    )
    match = folder_pattern.match(line.decode('utf-8'))
    flags, delimiter, mailbox_name = match.groups()
    mailbox_name = mailbox_name.strip('"')
    return flags, delimiter, mailbox_name


def get_email(hostname, port, username, password):
    with create_imapObj(hostname, port, username, password, verbose=False) as imapObj:

        # ----------------------------------- folder list ---------------------------------------
        (typ, mbfolder_list) = imapObj.list()
        # print('select[typ]:', typ)
        # pprint(mbfolder_list)
        """
        select[typ]: OK
        [b'() "/" "INBOX"',
         b'(\\Drafts) "/" "&g0l6P3ux-"',
         ...
         b'() "/" "&WSl5p08wUDyIaA-"']
        """
        mbfolder_name_list = []
        for line in mbfolder_list:
            flags, delimiter, mbfolder_name = folder_parse(line)
            mbfolder_name_list.append(mbfolder_name)

        # ----------------------------------- select folder ---------------------------------------
        for mbfolder in mbfolders:
            b_mbfolder = mbfolder.encode('utf-7')
            b_mbfolder = b_mbfolder.replace(b'+', b'&')
            s_mbfolder = b_mbfolder.decode('utf-8')
            (typ, msgsTotal_list) = imapObj.select(mailbox=s_mbfolder, readonly=True)
            # if readonly="True" you can't change any flags. But,if it is false, you can do as follow,
            # print('"' + mbfolder + '"' + 'select[typ]:', typ)
            # print('"' + mbfolder + '"' + 'select[msgsTotal_list]:', msgsTotal_list)
            """
            "浦发银行"select[typ]: OK
            "浦发银行"select[msgsTotal_list]: [b'1407']
            """
            if typ == 'NO':
                print("邮箱文件夹不存在\n\n")
            elif typ == 'OK':
                msgs_num = int(msgsTotal_list[0])
                print('There are {} messages in "{}"邮件文件夹'.format(msgs_num, mbfolder))

                # ----------------------------------- search mail ---------------------------------------
                Date = date.today() - timedelta(days=days)
                Date = Date.strftime("%d-%b-%Y")
                # criterion = f'(SINCE {Date} FROM {From})'
                criterion = f'(SINCE {Date})'
                (typ, msgsUids_list) = imapObj.search(None, criterion)
                # status, message = imapObj.search(None, 'OR FROM "ooxx@fuck.com"', 'SUBJECT "测试"'.encode('utf-8'))
                # IMAP4.search(charset, criterion[, ...])，其第二个参数形状同status()第二个参数类似。
                # 在邮箱中搜索匹配的消息。 charset 字符集可以为 None

                # print('"' + mbfolder + '"' + 'search[typ]:', typ)
                # print('"' + mbfolder + '"' + 'search[msgsUids_list]:', msgsUids_list, '\n')
                """
                "浦发银行"search[typ]: OK
                "浦发银行"search[msgsUids_list]: [b'1404 1405 1406 1407'] 
                """
                if not msgsUids_list[0]:
                    print("未搜索到符合条件的邮件！\n")
                else:
                    uids_list = msgsUids_list[0].split()[::-1]          # 原list是从早到近，[::-1]是顺序翻转。
                    # print("单个uid获取消息fetch:uids_list=:", uids_list)

                    # 未翻转的uids_bytes：
                    uids_bytes = msgsUids_list[0].replace(b' ', b',')   # 由多个uid组成的uids_bytes，一次性地批量fetch获取消息
                    # uids_bytes不管其顺序如何，fetch获取到的消息中的顺序都是从小到大的。

                    # print("批量fetch获取消息:uids_bytes=", uids_bytes, '\n')

                    print('按search条件搜索到的邮件总数:', len(uids_list), '\n')

                    # ----------------------------------- fetch messages ---------------------------------------
                    if sinflags:  # 单个模式fetch
                        i = 0
                        for uid_bytes in uids_list:
                            (typ, [(msgID_bytes, msgData_bytes), rrb_bytes]) = imapObj.fetch(uid_bytes, '(RFC822)')
                            # print('fetch:|', 'typ:', typ, '| msgID_bytes:', msgID_bytes, "| 右圆括号:", rrb_bytes, '|')
                            # fetch:| typ: OK | msgID_bytes: b'550 (RFC822 {439365}' | 右圆括号: b')' |
                            # ----------------------------------- 以上是imaplib的事，以下是email的事 -----------------------
                            if typ == 'NO':
                                print("单个获取uid_bytes=" + uid_bytes + "的消息失败！")
                            elif typ == 'OK':
                                Subject = do_msg(msgData_bytes, imapObj, uid_bytes, mbfolder, s_mbfolder)
                                i += 1
                                # print('uid_' + str(i) + ':', uid_bytes, 'decoded Subject:', Subject, '\n')

                    # ==================================================================================================
                    elif not sinflags:  # 批量模式fetch
                        (typ, msg_data) = imapObj.fetch(uids_bytes, '(RFC822)')
                        # msg_data[0]=(b'548 (RFC822 {330454}', b'Received: from ...')
                        # ----------------------------------- 以上是imaplib的事，以下是email的事 ---------------------------
                        if typ == 'NO':
                            print("批量获取uids_bytes=" + uids_bytes.decode('utf-8') + "的消息失败！")
                        elif typ == 'OK':
                            i = 0
                            for uid, msgData_bytes in msg_data[::2][::-1]:  # step=2 是隔一个取一个，不是指一次一次地取两个。
                                # uids_bytes不管其顺序翻不翻转msg_data中的顺序都是从小到大，故在这里对msg_data[::2]进行翻转。
                                uid_bytes = uid.split()[0]
                                Subject = do_msg(msgData_bytes, imapObj, uid_bytes, mbfolder, s_mbfolder)
                                i += 1
                                # print('uid_' + str(i) + ':', uid_bytes, 'decoded Subject:', Subject, '\n')


if __name__ == '__main__':
    start_time = time.time()

    config = configparser.ConfigParser()
    config.read('docs/config.cfg', encoding='utf-8')
    Hostname = config['server']['hostname']
    Port = config['server']['port']
    usernames = eval(config['account']['username'])
    passwords = eval(config['account']['password'])
    mbfolders = eval(config['other']['mbfolders'])
    titles = eval(config['other']['titles'])
    exts = eval(config['other']['exts'])
    savedir = config['other']['savedir']
    days = int(config['other']['days'])
    From = config['other']['From']
    sinflags = config['other']['sinflags']
    sinflags = eval(sinflags)
    mode = None
    
    if sinflags:
        mode = "单个模式"
    elif not sinflags:
        mode = "批量模式"
    else:
        print("未选定好模式")

    for n in range(len(usernames)):
        get_email(Hostname, Port, usernames[n], passwords[n])

    end_time = time.time()
    duration = (end_time-start_time)
    print(f'{mode}fetch，共耗时{duration}秒')
    # print(f'耗时：{time.time() - start_time}')
