# uncompyle6 version 3.8.0
# Python bytecode 3.8.0 (3413)
# Decompiled from: Python 3.8.5 (tags/v3.8.5:580fbb0, Jul 20 2020, 15:57:54) [MSC v.1924 64 bit (AMD64)]
# Embedded file name: e:\python_code\Email\imaplib_imapObj.py
# Compiled at: 2022-12-04 12:15:40
# Size of source mod 2**32: 1238 bytes
import configparser, os
from imaplib import IMAP4

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


if __name__ == '__main__':
    config = configparser.ConfigParser()
    config.read([os.path.expanduser('docs/config.cfg')], encoding='utf-8')

    hostname = config['server']['hostname']
    port = config['server']['port']
    usernames = config['account']['username'].split(',')
    passwords = config['account']['password'].split(',')

    for i in range(len(usernames)):
        with create_imapObj(hostname, port, usernames[i], passwords[i], verbose=False) as (Obj):
             print(Obj)