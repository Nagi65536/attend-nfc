import datetime
import os
import sqlite3
import subprocess
import time

import gspread
from oauth2client.service_account import ServiceAccountCredentials

# TODO
MUSIC_FILE = 'hotarus_light.mp3'
SPREADSHEET_APP = 'misc-list'
SETTING_SHEET = '設定'
PATH = os.path.dirname(os.path.abspath(__file__)) + '/'

scope = ['https://spreadsheets.google.com/feeds',
         'https://www.googleapis.com/auth/drive']
creds = ServiceAccountCredentials.from_json_keyfile_name(
    f'{PATH}client_secret.json', scope)
client = gspread.authorize(creds)
sht = client.open_by_key("16QxcHhQBNo5RL1LgPiPMn4GKsOEd6FmAX4trDBPfN0Q")


def main():
    target_data = client.open(SPREADSHEET_APP)
    setting_sheet = target_data.worksheet(SETTING_SHEET)

    t_delta = datetime.timedelta(hours=9)
    JST = datetime.timezone(t_delta, 'JST')

    isFirst = True
    isOn = True
    get_data = None
    set_times = None

    print('⚡️ Periodic is running!')
    while True:
        dt_now = datetime.datetime.now(JST)
        now_time = dt_now.strftime('%H:%M')
        now_min = int(dt_now.strftime('%M'))
        record_sheet_name = dt_now.strftime('%Y-%m')
        col = datetime.date.today().weekday() + 3

        if now_min % 10 == 0 or isFirst:
            get_data = setting_sheet.col_values(col)
            set_times = [get_data[3], get_data[4]]

        print(f'now: {now_time}  set-1: {set_times[0]}  set-2: {set_times[1]}')
        if now_time in set_times:
            print('【再生開始】蛍の光')
            subprocess.Popen(['mpg321', f'{PATH}sounds/{MUSIC_FILE}', '-q'])
            time.sleep(210)

        time.sleep(30)


if __name__ == '__main__':
    main()
