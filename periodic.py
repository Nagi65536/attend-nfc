import datetime
import os
import re
import sqlite3
import subprocess
import time

import gspread
from oauth2client.service_account import ServiceAccountCredentials

# path = '/home/nagi/Documents/attend-misc/'
path = './'

scope = ['https://spreadsheets.google.com/feeds',
         'https://www.googleapis.com/auth/drive']
creds = ServiceAccountCredentials.from_json_keyfile_name(
    'client_secret.json', scope)
client = gspread.authorize(creds)
sht = client.open_by_key("16QxcHhQBNo5RL1LgPiPMn4GKsOEd6FmAX4trDBPfN0Q")


def init():
    try:
        cur.execute('DELETE FROM record')
        sht.duplicate_sheet(
            source_sheet_id=1120016017,
            new_sheet_name=record_sheet_name,
            insert_sheet_index=1
        )
        record_sheet = target_data.worksheet(record_sheet_name)
        all_data = main_sheet.get_all_values()

        for i, data in enumerate(all_data[1:]):
            record_sheet.append_row(
                [data[1], data[2], data[3], data[4], data[5], data[6]],
                table_range=f'B{i+1}'
            )
    except:
        pass


def main():
    t_delta = datetime.timedelta(hours=9)
    JST = datetime.timezone(t_delta, 'JST')

    all_data = setting_sheet.get_all_values()
    isOn = all_data[4][col].lower() == 'on'
    set_time = all_data[3][col]

    while True:
        dt_now = datetime.datetime.now(JST)
        now_time = dt_now.strftime('%H:%M')
        now_min = int(dt_now.strftime('%M'))
        record_sheet_name = dt_now.strftime('%Y-%m')
        col = datetime.date.today().weekday() + 3

        if now_min % 10 == 0:
            all_data = setting_sheet.get_all_values()
            isOn = all_data[4][col].lower() == 'on'
            set_time = all_data[3][col]

        if isOn and now_time == set_time:
            subprocess.Popen(['mpg321', f'{path}sounds/fin.mp3', '-q'])
            time.sleep(210)

        if now_time == "00:00":
            init()

        time.sleep(50)


if __name__ == '__main__':
    target_data = client.open("misc-list")
    main_sheet = target_data.worksheet("名簿")
    setting_sheet = target_data.worksheet("設定")

    conn = sqlite3.connect(f'{path}check.db', isolation_level=None)
    cur = conn.cursor()
    cur.execute('''
        CREATE TABLE IF NOT EXISTS record(
            idm            TEXT PRIMARY KEY,
            entry_time     TEXT DEFAULT 0,
            entry_time_int INTEGER DEFAULT 0,
            exit_time      TEXT DEFAULT 0
            exit_time_int  INTEGER DEFAULT 0
        )''')

    record_sheet_name = None
    print('⚡️ Periodic is running!')
    main()
