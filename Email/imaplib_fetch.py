import configparser
import email.parser
import os

from imaplib_imapobj import create_imapObj

config = configparser.ConfigParser()
config.read([os.path.expanduser('docs/config.cfg')], encoding='utf-8')

hostname = config['server']['hostname']
port = config['server']['port']
usernames = config['account']['username'].split(',')
passwords = config['account']['password'].split(',')

for i in range(len(usernames)):
    with create_imapObj(hostname, port, usernames[i], passwords[i], verbose=False) as (Obj):
        Obj.select('INBOX', readonly=True)

        typ, msg_data = Obj.fetch('1', '(RFC822)')
        for response_part in msg_data:
            if isinstance(response_part, tuple):
                email_parser = email.parser.BytesFeedParser()
                email_parser.feed(response_part[1])
                msg = email_parser.close()
                for header in ['subject', 'to', 'from']:
                    print('{:^8}: {}'.format(
                        header.upper(), msg[header]))