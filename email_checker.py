import re
import imaplib
import email
import random
import smtplib
import datetime
import os
import time
import json
import yaml

from email.mime.text import MIMEText
from util_func import securely_check_dir
from excel_handler import ExcelHandler


class EmailChecker:
    def __init__(self, config, config_location):
        self.keyword_pattern = re.compile(r'r\d{3}')

        self.config = config
        self.config_location = config_location

        self.imap_host = self.config['email']['imap_host']
        self.smtp_host = self.config['email']['smtp_host']
        self.user = self.config['email']['account']
        self.psw = self.config['email']['psw']

        self.keyword_list = []
        for key in self.config['responds'].keys():
            if not key.startswith('default') and not key.endswith('expired'):
                self.keyword_list.append(key)

        self.smtp_sender = smtplib.SMTP_SSL()
        self.smtp_sender.connect(self.smtp_host)
        self.smtp_sender.login(self.user, self.psw)
        self.imap = imaplib.IMAP4(self.imap_host)

    def check_imap_folder(self):
        self.imap.login(self.user, self.psw)
        self.imap.select("INBOX")

        while True:
            _, data = self.imap.search(None, 'UNSEEN')
            num = data[0].split()
            if len(num) == 0:
                break
            num = num[0]
            typ, mail = self.imap.fetch(num, '(RFC822)')
            content = mail[0][1]
            msg = email.message_from_bytes(content)
            subject = email.header.decode_header(msg["Subject"])[0]
            if subject[1] is not None:
                subject = subject[0].decode(subject[1])
            subject = self.keyword_pattern.findall(subject)
            if len(subject) == 0 or subject[0] not in self.keyword_list:
                continue
            subject = subject[0]
            from_time = self.config['responds'][subject]['from']
            if from_time == -1:
                from_time = datetime.datetime.fromtimestamp(time.time()).replace(tzinfo=None)
            to_time = self.config['responds'][subject]['to']
            if from_time <= datetime.datetime.fromtimestamp(time.time()).replace(tzinfo=None) <= to_time:
                expired = False
                if self.config['responds'][subject]['save']:
                    for part in msg.walk():
                        if part.is_multipart():
                            continue
                        if part.get('Content-Disposition') is None:
                            continue
                        file_data = part.get_payload(decode=True)
                        file_name = part.get_filename()
                        de_name = email.header.decode_header(file_name)[0]
                        if de_name[1] is not None:
                            file_name = de_name[0].decode(de_name[1])
                        att_folder_path = os.path.join('att', subject)
                        securely_check_dir(att_folder_path)
                        attachment_path = os.path.join(att_folder_path, file_name)
                        while os.path.isfile(attachment_path):
                            short_name, ext = os.path.splitext(attachment_path)
                            attachment_path = ''.join([short_name, str(random.randint(0, 999)), ext])
                        with open(attachment_path, 'wb') as fp:
                            fp.write(file_data)
                if self.config['responds'][subject]['handle']:
                    try:
                        securely_check_dir(os.path.join('config', '{}.yml'.format(subject)))
                        excel_handler = ExcelHandler(self.config)
                        excel_handler.handle()
                    except:
                        pass
            else:
                expired = True
                self.config['responds']['{}{}'.format(subject, '-expired')] = self.config['responds'].pop(subject)
                with open(self.config_location, 'w', encoding='utf-8') as fp:
                    fp.write(yaml.dump(self.config))
            # reply
            if not expired:
                if 'content' in self.config['responds'][subject].keys():
                    reply_content = self.config['responds'][subject]['content']
                else:
                    reply_content = self.config['responds']['default_content']
            else:
                if 'expire_content' in self.config['responds'][subject].keys():
                    reply_content = self.config['responds'][subject]['expire_content']
                else:
                    reply_content = self.config['responds']['default_expire_content']
            reply = MIMEText(reply_content, _subtype='plain', _charset='utf-8')
            if 'subject' in self.config['responds'][subject].keys():
                reply['Subject'] = self.config['responds'][subject]['subject']
            else:
                reply['Subject'] = self.config['responds']['default_subject']
            reply['From'] = '<{}>'.format(self.user)
            to = msg['From']
            reply['to'] = to

            self.smtp_sender.sendmail(self.user, to, reply.as_string())
            print('Message sent to {}.'.format(msg['From']))
            securely_check_dir('sent')
            sent = []
            try:
                with open(os.path.join('sent', '{}.json'.format(subject))) as fp:
                    sent = json.loads(fp.read())
            except:
                pass
            sent.append(to)
            with open(os.path.join('sent', '{}.json'.format(subject)), 'w') as fp:
                fp.write(json.dumps(sent, indent=4))

            # move this mail to folder
            result = self.imap.copy(num, self.config['email']['destination'])
            if result[0] == 'OK':
                self.imap.store(num, '+FLAGS', r'(\Deleted)')
                self.imap.expunge()
            print('Email moved to {}.'.format(self.config['email']['destination']))

        self.imap.close()
        self.imap.logout()
        self.smtp_sender.close()


if __name__ == '__main__':
    config_location = 'config/top.yml'
    if not os.path.exists(config_location):
        print('No top.yml file found!')
        exit(-1)
    with open(config_location, encoding='utf-8') as f:
        config_file = yaml.load(f.read())
    schedule_interval = config_file['schedule']['interval_minutes']
    email_checker = EmailChecker(config_file, config_location)
    email_checker.check_imap_folder()
