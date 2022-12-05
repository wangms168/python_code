import configparser, os
from imaplib_imapobj import create_imapObj

# [mailbox]
# 浦东银行 = &bWZOHJT2iEw-            附件有xlsx、zip两种，需下载xlsx、并须删行
# 国泰君安 = &Vv1s8FQbW4k-            附件是xlsx，只需下载即可     
# 兴业银行 = &UXROGpT2iEw-            附件是xlsx，只需下载即可
# 宝城期货 = &W51XzmcfjSc-            附件是txt文件，并须改名

config = configparser.ConfigParser()
config.read([os.path.expanduser('docs/config.cfg')], encoding='utf-8')

hostname = config['server']['hostname']
port = config['server']['port']
usernames = config['account']['username'].split(',')
passwords = config['account']['password'].split(',')

for i in range(len(usernames)):
    with create_imapObj(hostname, port, usernames[i], passwords[i], verbose=False) as (Obj):
        typ, data = Obj.select('INBOX')
        # IMAP4.select(mailbox='INBOX', readonly=False)
        # 选择一个邮箱。 返回的数据是 mailbox 中消息的数量 (EXISTS 响应)               
        # typ 同样是响应代码；响应数据 data 是一个包含单个字符串的列表，该单个字符串包含邮箱中的邮件总数。
        print(typ, data)
        num_msgs = int(data[0])
        print('There are {} messages in INBOX'.format(num_msgs))