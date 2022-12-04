from imaplib_imapobj import create_imapObj

import email
import email.parser

with create_imapObj() as Obj:
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