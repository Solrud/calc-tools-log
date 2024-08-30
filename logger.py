from constants import LOGS_PATH
import os
import datetime

class Logger:
    def __init__(self):
        self.current_user = os.environ['USERNAME']

    def log_info(self, message):
        current_time = datetime.datetime.now().strftime("%d-%m-%Y %H:%M:%S")
        log_info = f'{current_time} | {self.current_user} | INFO | {message}\n'

        self.write_log(log_info)

    def log_error(self, message):
        current_time = datetime.datetime.now().strftime("%d-%m-%Y %H:%M:%S")
        log_info = f'{current_time} | {self.current_user} | ERROR | {message}\n'

        self.write_log(log_info)

    def write_log(self, log):
        try:
            with open(LOGS_PATH, 'a') as file_log:
                file_log.write(log)
        except Exception as ex:
            print('Ошибка в логгировании: ' + str(ex))

