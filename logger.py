import os
from datetime import datetime as dt


def logger(action):
    if not os.path.exists('logs'):
        return

    log_file = open('logs/logs.txt', 'a')
    dt_now = dt.now()
    log_file.write(f'[{dt_now.strftime("%d-%m-%Y %H:%M")}]: \t{action}\n')
    log_file.close()
