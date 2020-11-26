import os
from datetime import datetime as dt


def logger(action):
    home_path = os.getenv('USERPROFILE')
    egrul_parser_path = home_path + '/egrul_parser'
    if not os.path.exists(egrul_parser_path):
        os.mkdir(egrul_parser_path)
    logs_dir = egrul_parser_path+'/logs'
    if not os.path.exists(logs_dir):
        os.mkdir(logs_dir)

    log_file = open(logs_dir+'/logs.txt', 'a')
    dt_now = dt.now()
    log_file.write(f'[{dt_now.strftime("%d-%m-%Y %H:%M")}]: \t{action}\n')
    log_file.close()
