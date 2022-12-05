import configparser, os
from pprint import pprint
from imaplib_imapobj import create_imapObj

config = configparser.ConfigParser()
config.read([os.path.expanduser('docs/config.cfg')], encoding='utf-8')

hostname = config['server']['hostname']
port = config['server']['port']
usernames = config['account']['username'].split(',')
passwords = config['account']['password'].split(',')

for i in range(len(usernames)):
    with create_imapObj(hostname, port, usernames[i], passwords[i], verbose=False) as (Obj):
        typ, data = Obj.list()                # IMAP4.list([directory[, pattern]]) 列出 directory 中与 pattern 相匹配的邮箱名称（邮箱文件夹）
        # c.list(directory='Example')
        # c.list(pattern='*Example*')
        # 返回值是一个 tuple，包含响应代码 typ 和响应数据 data 。除非出现错误，否则响应代码为 OK。
        # list() 的响应数据 data 是一个字符串序列，每个列表元素包含每个邮箱的标志，层次结构分隔符和邮箱名称 *。
        print('Response code:', typ)
        print('Response:')
        pprint(data)
