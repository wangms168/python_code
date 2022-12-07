import configparser
import os

from imaplib_folder_parse import folder_parse
from imaplib_imapobj import create_imapObj

# criterion = 'UnSeen'
# criterion = 'ALL'
# criterion = '(SUBJECT "Example message 2")'
criterion = '(FROM "Doug" SUBJECT "Example message 2")'

config = configparser.ConfigParser()
config.read([os.path.expanduser('docs/config.cfg')], encoding='utf-8')

hostname = config['server']['hostname']
port = config['server']['port']
usernames = config['account']['username'].split(',')
passwords = config['account']['password'].split(',')

for i in range(len(usernames)):
    with create_imapObj(hostname, port, usernames[i], passwords[i], verbose=False) as (imapObj):
        typ, mbox_data = imapObj.list()
        for line in mbox_data:
            flags, delimiter, mbox_name = folder_parse(line)
            imapObj.select('"{}"'.format(mbox_name), readonly=True)
            typ, msg_ids = imapObj.search(
                None,
                criterion,
            )
            # IMAP4.search(charset, criterion[, ...])，其第二个参数形状同status()第二个参数类似。
            # 在邮箱中搜索匹配的消息。 charset 字符集可以为 None    
            # 同c.status()一样，其响应数据也是一个包含单个字节类型(即字符串前面标有b的前缀)字符串的列表，
            # 该字符串是一个由空格分隔的连续消息(邮件)ID组成。
            print(mbox_name, typ, msg_ids)
