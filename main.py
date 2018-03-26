import yaml
import os
import imaplib
import email
import re
import random
import smtplib
from email.mime.text import MIMEText
import datetime
import time

keyword_pattern = re.compile(r'r\d{3}')


def check_imap_folder(config):
    host = config['email']['imap_host']
    user = config['email']['account']
    psw = config['email']['psw']
    keyword_list = [key for key in config['responds'].keys()]
    keyword_list.remove('default_content')
    keyword_list.remove('default_expire_content')
    keyword_list.remove('default_subject')

    smtp_sender = smtplib.SMTP_SSL()
    smtp_sender.connect(config['email']['smtp_host'])
    smtp_sender.login(user, psw)

    imap = imaplib.IMAP4(host)
    imap.login(user, psw)
    imap.select("INBOX")
    typ, data = imap.search(None, 'UNSEEN')
    for num in data[0].split():
        typ, data = imap.fetch(num, '(RFC822)')
        content = data[0][1]
        msg = email.message_from_bytes(content)
        subject = email.header.decode_header(msg["Subject"])[0]
        if subject[1] is not None:
            subject = subject[0].decode(subject[1])
        subject = keyword_pattern.findall(subject)
        if len(subject) == 0 or subject[0] not in keyword_list:
            continue
        subject = subject[0]
        from_time = config['responds'][subject]['from']
        if from_time == -1:
            from_time = datetime.datetime.fromtimestamp(time.time()).replace(tzinfo=None)
        to_time = config['responds'][subject]['to']
        expired = True
        if from_time <= datetime.datetime.fromtimestamp(time.time()).replace(tzinfo=None) <= to_time:
            expired = False
            if config['responds'][subject]['save']:
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
                    if not os.path.exists(att_folder_path):
                        os.makedirs(att_folder_path)
                    attachment_path = os.path.join(att_folder_path, file_name)
                    while os.path.isfile(attachment_path):
                        short_name, ext = os.path.splitext(attachment_path)
                        attachment_path = ''.join([short_name, str(random.randint(0, 999)), ext])
                    with open(attachment_path, 'wb') as fp:
                        fp.write(file_data)
            if config['responds'][subject]['handle']:
                if not os.path.exists(os.path.join('config', '{}.yml'.format(subject))):
                    pass
        # reply
        if not expired:
            if 'content' in config['responds'][subject].keys():
                reply_content = config['responds'][subject]['content']
            else:
                reply_content = config['responds']['default_content']
        else:
            if 'expire_content' in config['responds'][subject].keys():
                reply_content = config['responds'][subject]['expire_content']
            else:
                reply_content = config['responds']['default_expire_content']
        reply = MIMEText(reply_content, _subtype='plain', _charset='utf-8')
        if 'subject' in config['responds'][subject].keys():
            reply['Subject'] = config['responds'][subject]['subject']
        else:
            reply['Subject'] = config['responds']['default_subject']
        reply['From'] = '<{}>'.format(user)
        reply['to'] = msg['From']
        smtp_sender.sendmail(user, msg['From'], reply.as_string())
        print('Message sent to {}.'.format(msg['From']))

        # move this mail to folder
        result = imap.copy(num, config['email']['destination'])
        if result[0] == 'OK':
            imap.store(num, '+FLAGS', r'(\Deleted)')
            imap.expunge()
        print('Email moved to {}.'.format(config['email']['destination']))

    imap.close()
    imap.logout()
    smtp_sender.close()


if __name__ == '__main__':
    config_file = 'config/top.yml'
    if not os.path.exists(config_file):
        print('No top.yml file found!')
        exit(-1)
    with open(config_file, encoding='utf-8') as f:
        config_file = yaml.load(f.read())
    schedule_interval = config_file['schedule']['interval_minutes']
    while True:
        try:
            check_imap_folder(config_file)
            time.sleep(schedule_interval * 60)
        except (KeyboardInterrupt, SystemExit):
            print('Program shutdown!')
            break
