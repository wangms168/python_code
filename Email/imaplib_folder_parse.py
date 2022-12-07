# 对c.list()的响应数据data进行re(正则)解析(parse),得到flags(每个邮箱的标志), delimiter(层次结构分隔符), mailbox_name(邮箱名称)
import configparser
import os
import re

from imaplib_imapobj import create_imapObj

folder_pattern = re.compile(
    r'.(?P<flags>.*?). "(?P<delimiter>.*)" (?P<name>.*)'
)


def folder_parse(line):
    match = folder_pattern.match(line.decode('utf-8'))
    flags, delimiter, mailbox_name = match.groups()
    mailbox_name = mailbox_name.strip('"')
    return flags, delimiter, mailbox_name


def main():
    config = configparser.ConfigParser()
    config.read([os.path.expanduser('docs/config.cfg')], encoding='utf-8')

    hostname = config['server']['hostname']
    port = config['server']['port']
    usernames = config['account']['username'].split(',')
    passwords = config['account']['password'].split(',')

    for i in range(len(usernames)):
        with create_imapObj(hostname, port, usernames[i], passwords[i], verbose=False) as (Obj):
            typ, data = Obj.list()
        print('Response code:', typ)

        for line in data:
            print('Server response:', line)
            flags, delimiter, mailbox_name = folder_parse(line)
            print('Parsed response:', (flags, delimiter, mailbox_name))


if __name__ == '__main__':
    main()
