import yaml
import os
import time
from email_checker import EmailChecker

if __name__ == '__main__':
    config_location = 'config/top.yml'
    while True:
        try:
            if not os.path.exists(config_location):
                print('No top.yml file found!')
                exit(-1)
            with open(config_location, encoding='utf-8') as f:
                config_file = yaml.load(f.read())
            schedule_interval = config_file['schedule']['interval_minutes']
            email_checker = EmailChecker(config_file, config_location)
            email_checker.check_imap_folder()
            print('Check done...')
            print('Pending...')
            time.sleep(schedule_interval * 60)
        except (KeyboardInterrupt, SystemExit):
            print('Program shutdown!')
            break
