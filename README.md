# EMailResponder
Easily define respond messages of an email account

#### Prerequisites

* python3
  * pyyaml
  * xlrd
  * openpyxl
  * apscheduler

#### How to use?

1. Edit `config/top.yml`.
2. Simply run `python3 main.py` or `python3 schedule_main.py` (supposing you've met the prerequisites)


#### Note

1. If you need to login THU net gate before and/or logout after checking emails using `schedule_main.py`, you need to

   * edit

   ```
   scheduler.add_job(job, 'interval', seconds=schedule_interval * 60, args=[config_file, location, False, True])
   ```

   in `schedule_main.py`. The last 2 arguments in `args` fields denote whether need to login first and logout after checking emails.

   * edit `config/top.yml`. Replace `net_gate.account` and `net_gate.psw` fields with your login account and password for THU net gate.

2. ​

   ​