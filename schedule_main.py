import os
import yaml
import logging

from apscheduler.schedulers.blocking import BlockingScheduler
from email_checker import EmailChecker


def job(config_location, login=False, logout=True):
    try:
        print('Checking...')
        net_gate = None
        if login:
            from thu_net_gate import THUNetGate
            net_gate = THUNetGate()
            with open('config/top.yml', 'r', encoding='utf-8') as fp:
                account = yaml.load(fp.read())
                account = account['net_gate']
                net_gate.login(account['account'], account['psw'])
        with open(config_location, 'r', encoding='utf-8') as f:
            config = yaml.load(f.read())
        email_checker = EmailChecker(config, config_location)
        email_checker.check_imap_folder()
        if logout:
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
    login_first = config_file['net_gate']['login_first']
    logout_after = config_file['net_gate']['logout_after']
    job(location, login_first, logout_after)
    scheduler.add_job(job, 'interval', seconds=schedule_interval * 60, args=[location, login_first, logout_after])

    logging.basicConfig(level=logging.INFO,
                        format='%(asctime)s %(filename)s[line:%(lineno)d] %(levelname)s %(message)s',
                        datefmt='%Y-%m-%d %H:%M:%S',
                        filename='log.log',
                        filemode='a')

    scheduler._logger = logging
    scheduler.start()
