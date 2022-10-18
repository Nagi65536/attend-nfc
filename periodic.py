import datetime
import os
import re
import sqlite3
import subprocess
import time

import gspread
from oauth2client.service_account import ServiceAccountCredentials

SPREADSHEET_APP = 'misc-list'
ROSTER_NAME = '名簿'
PATH = __file__

scope = ['https://spreadsheets.google.com/feeds',
         'https://www.googleapis.com/auth/drive']
creds = ServiceAccountCredentials.from_json_keyfile_name(
    'client_secret.json', scope)
client = gspread.authorize(creds)
sht = client.open_by_key("16QxcHhQBNo5RL1LgPiPMn4GKsOEd6FmAX4trDBPfN0Q")


def main():
    target_data = client.open(SPREADSHEET_APP)
    setting_sheet = target_data.worksheet("設定")

    t_delta = datetime.timedelta(hours=9)
    JST = datetime.timezone(t_delta, 'JST')

    isFirst = True
    isOn = True
    get_data = None
    set_time = None

    print('⚡️ Periodic is running!')
    while True:
        dt_now = datetime.datetime.now(JST)
        now_time = dt_now.strftime('%H:%M')
        now_min = int(dt_now.strftime('%M'))
        record_sheet_name = dt_now.strftime('%Y-%m')
        col = datetime.date.today().weekday() + 3

        if now_min % 10 == 0 or isFirst:
            get_data = setting_sheet.col_values(col)
            set_time = get_data[3]
            isOn = get_data[4].lower() == 'on'

        if isOn and now_time == set_time:
            print(' 【再生開始】蛍の光')
            subprocess.Popen(['mpg321', f'{PATH}sounds/fin.mp3', '-q'])
            time.sleep(210)

        time.sleep(50)


if __name__ == '__main__':
    main()
