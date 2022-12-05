import configparser, os
from imaplib_imapobj import create_imapObj

import email
import email.parser

def header_decode(header):
    hdr = ""
    for text, encoding in email.header.decode_header(header):
        if isinstance(text, bytes):
            text = text.decode(encoding or "us-ascii")
        hdr += text
    return hdr

config = configparser.ConfigParser()
config.read([os.path.expanduser('docs/config.cfg')], encoding='utf-8')

hostname = config['server']['hostname']
port = config['server']['port']
usernames = config['account']['username'].split(',')
passwords = config['account']['password'].split(',')

for i in range(len(usernames)):
    with create_imapObj(hostname, port, usernames[i], passwords[i], verbose=False) as (Obj):
        Obj.select('INBOX', readonly=True)

        (typ, [(msgXX_bytes, msgData_bytes), Rrb_bytes]) = Obj.fetch('1', '(RFC822)')    # fetch()返回一个包含两个项目的tuple，第一个项目fetch()[0]是字符串'OK',是响应代码typ；
        # typ,              msg_data                     = Obj.fetch('1', '(RFC822)')    # fetch()返回一个包含两个项目的tuple，第一个项目fetch()[0]是字符串'OK',是响应代码typ；
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
        # IMAP4.fetch(message_set, message_parts)取回（部分）信息。 message_parts应该是一串括在圆括号内的消息部分名。，例如: "(UID BODY[TEXT])"。 返回的数据是由消息部分信封和数据组成的元组。

        # ---------------------------------以上是imaplib的事，以下是email的事----------------------------------------------

        message_ojb = email.message_from_bytes(msgData_bytes)
        # email.message_from_bytes(s, _class=None, *, policy=policy.compat32)
        # 从一个 bytes-like object 中返回消息对象。 这与 BytesParser().parsebytes(s) 等价。

        # print('message_ojb:', message_ojb)
        print('message_ojb-type:', type(message_ojb))
        print('message_ojb.keys:',message_ojb.keys())           # msg.keys() https://stackoverflow.com/questions/703185/using-email-headerparser-with-imaplib-fetch-in-python

        Subject = message_ojb.get('Subject')
        print("message_ojb['Subject']:", Subject)
        print("message_ojb['Subject']-type:", type(Subject))


        Subject_decode_header = email.header.decode_header(Subject)
        print("decode_header_message_ojb['Subject']:", Subject_decode_header)
        print("decode_header_message_ojb['Subject']-type:", type(Subject_decode_header))
        print("decode_header_message_ojb['Subject']-len:", len(Subject_decode_header))


        print("message_ojb['Subject']:", Subject_decode_header[0][0])
        print("message_ojb['Subject']-type:", Subject_decode_header[0][0].decode(Subject_decode_header[0][1] or "us-ascii"))
        
        # for response_part in msg_data:
        #     if isinstance(response_part, tuple):
        #         email_parser = email.parser.BytesFeedParser()
        #         email_parser.feed(response_part[1])
        #         msg = email_parser.close()
        #         for header in ['subject', 'to', 'from']:
        #             print('{:^8}: {}'.format(
        #                 header.upper(), header_decode(msg[header])))