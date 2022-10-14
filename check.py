import binascii
import csv
import datetime
import os
import sqlite3
import subprocess
import time

import gspread
import nfc
from oauth2client.service_account import ServiceAccountCredentials

# path = '/home/nagi/Documents/attend-nfc-ver2/'
path = './'

scope = ['https://spreadsheets.google.com/feeds',
         'https://www.googleapis.com/auth/drive']
creds = ServiceAccountCredentials.from_json_keyfile_name(
    f'{path}client_secret.json', scope)
client = gspread.authorize(creds)
sht = client.open_by_key("16QxcHhQBNo5RL1LgPiPMn4GKsOEd6FmAX4trDBPfN0Q")


class MyCardReader(object):
    # カードがタッチされたら呼び出される
    def on_connect(self, tag):
        self.idm = binascii.hexlify(tag.identifier).upper().decode('utf-8')
        print("IDm : " + str(self.idm))
        record(self.idm)
        return

    # カードの検知
    def read_id(self):
        clf = nfc.ContactlessFrontend('usb')
        try:
            clf.connect(rdwr={'on-connect': self.on_connect})
        finally:
            clf.close()


# idm が未登録の場合
def unrecorded(idm_list):
    for i, value in enumerate(idm_list[1:]):
        if len(value) <= 10:
            print('【新規登録】\n', value)
            subprocess.Popen(['mpg321', f'{path}sounds/popi.mp3', '-q'])
            main_sheet.update_cell(i+1, 8, idm)
            return


# 入退室記録の登録
def record(idm):
    all_data = main_sheet.get_all_values()
    idm_list = [d[5] for d in all_data]

    if not idm in idm_list:
        unrecorded(idm_list)
        return

    t_delta = datetime.timedelta(hours=9)
    JST = datetime.timezone(t_delta, 'JST')
    dt_now = datetime.datetime.now(JST)
    time_str = dt_now.strftime('%H:%M:%S')
    time_int = time.time()

    record_row = int(dt_now.strftime('%-d')) + 10
    record_col = idm_list.index(idm) + 1
    record_txt = ''

    cur.execute(f'SELECT FROM entry_time_int WHERE idm="{self.idm}"')
    res = cur.fetchone()

    try:
        if not res:
            record_txt = '0'
            subprocess.Popen(['mpg321', f'{path}sounds/ppi.mp3', '-q'])

            cur.execute(f'''INSERT INTO record VALUES(
                "{idm}",
                "{time_str}",
                "{time_int}",
                0,
                0
            )''')

        elif res[0] == 0:
            record_txt = '0'
            subprocess.Popen(['mpg321', f'{path}sounds/ppi.mp3', '-q'])

            cur.execute(f'''UPDATE record SET
                entry_time="{time_str}",
                entry_time_int="{time_int}"
            WHERE idm="{idm}"
            ''')

        else:
            diff_time = time_int - res[0]
            if diff_time < 60*30:
                record_txt = str(diff_time)
            else:
                record_txt = '◯'
            subprocess.Popen(['mpg321', f'{path}sounds/teretere.mp3', '-q'])

            cur.execute(f'''UPDATE record SET
                exit_time="{time_str}",
                exit_time_int="{time_int}"
                WHERE idm="{idm}"
            ''')

        record_sheet.update_cell(record_col, record_row, record_txt)
        print(f'記録 {record_txt} ({record_row}, {record_col})')

    except:
        subprocess.Popen(['mpg321', f'{path}sounds/err.mp3', '-q'])


if __name__ == '__main__':
    target_data = client.open("misc-list")
    main_sheet = target_data.worksheet("名簿")

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

    cr = MyCardReader()
    print('⚡️ Attend System is running!')
    while True:
        cr.read_id()
