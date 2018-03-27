import os
import yaml
import logging

from apscheduler.schedulers.blocking import BlockingScheduler
from email_checker import EmailChecker


def job(config, config_location, login_first=False, logout_after=True):
    try:
        print('Checking...')
        net_gate = None
        if login_first:
            from thu_net_gate import THUNetGate
            import yaml
            net_gate = THUNetGate()
            with open('config/top.yml') as fp:
                account = yaml.load(fp.read())
                account = account['net_gate']
                net_gate.login(account['account'], account['psw'])
        email_checker = EmailChecker(config, config_location)
        email_checker.check_imap_folder()
        if logout_after:
            net_gate.logout()
        print('Check done...')
    except (KeyboardInterrupt, SystemExit):
        print('Program shutdown!')


if __name__ == '__main__':
    location = 'config/top.yml'
    if not os.path.exists(location):
        print('No top.yml file found!')
        exit(-1)
    with open(location, encoding='utf-8') as f:
        config_file = yaml.load(f.read())
    schedule_interval = config_file['schedule']['interval_minutes']
    scheduler = BlockingScheduler()
    scheduler.add_job(job, 'interval', seconds=schedule_interval * 60, args=[config_file, location, False, True])

    logging.basicConfig(level=logging.INFO,
                        format='%(asctime)s %(filename)s[line:%(lineno)d] %(levelname)s %(message)s',
                        datefmt='%Y-%m-%d %H:%M:%S',
                        filename='log.log',
                        filemode='a')

    scheduler._logger = logging
    scheduler.start()
