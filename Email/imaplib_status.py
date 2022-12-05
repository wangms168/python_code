import configparser, os
from imaplib_imapobj import create_imapObj
from imaplib_folder_parse import folder_parse

config = configparser.ConfigParser()
config.read([os.path.expanduser('docs/config.cfg')], encoding='utf-8')

hostname = config['server']['hostname']
port = config['server']['port']
usernames = config['account']['username'].split(',')
passwords = config['account']['password'].split(',')

for i in range(len(usernames)):
    with create_imapObj(hostname, port, usernames[i], passwords[i], verbose=False) as (Obj):
        typ, data = Obj.list()
        for line in data:
            flags, delimiter, mailbox = folder_parse(line)
            print('Mailbox:', mailbox)
            status = Obj.status(                                      
                '"{}"'.format(mailbox),
                '(MESSAGES RECENT UIDNEXT UIDVALIDITY UNSEEN)',
            )
            # IMAP4.status(mailbox, names)， 其第二个参数形状同search()第二个参数类似。
            # 返回值status也是一个tuple，也可以写成 typ(响应代码), data(响应数据)，data(响应数据)也是一个列表，
            # 列表包含单个字符串，该字符串的格式为用引号引起来的邮箱名称，然后是括号中的状态条件和值。

            print(status)