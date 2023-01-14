import binascii
import datetime
import os
import sqlite3
import subprocess
import time

import gspread
import nfc
from oauth2client.service_account import ServiceAccountCredentials

# TODO
SE_ENTRY = 'maou_system23.mp3'      # 入室時
SE_ALREADY = 'maou_system46.mp3'    # 退室時
SE_NEW = 'succeed.mp3'              # 新規登録時
SE_NONE = 'maou_system35.mp3'       # 未登録時
SE_ERROR = 'pickup02.mp3'           # システムエラー

# TODO
SPREADSHEET_APP = 'misc-list'
ROSTER_NAME = '名簿'


class MyCardReader(object):
    # カードがタッチされたら呼び出される
    def on_connect(self, tag):
        self.rfid = binascii.hexlify(tag.identifier).upper().decode('utf-8')
        print(f'rfid : {self.rfid}')
        record(self.rfid)
        return True

    # カードの検知
    def read_id(self):
        clf = nfc.ContactlessFrontend('usb')
        try:
            clf.connect(rdwr={'on-connect': self.on_connect})
        finally:
            clf.close()


# 入退室記録の登録
def record(rfid):
    dt_now = datetime.datetime.now(JST)
    date = int(dt_now.strftime('%Y%m'))
    month = int(dt_now.strftime('%m'))
    month_str = ["jan", "feb", "mar", "apr", "may",
                 "jun", "jul", "aug", "sep", "oct", "nov", "dec"](month-1)

    cur.execute(
        f'SELECT {month_str}, last_touch FROM user_data WHERE rfid="{rfid}"')
    res = cur.fetchone()

    # 登録済み and 未タッチ
    if res and res[1] != date:
        subprocess.Popen(['mpg321', f'{PATH}sounds/{SE_ENTRY}', '-q'])
        attend_num = res[0] + 1
        cur.execute(
            f'UPDATE user_data SET last_touch={date}, {month_str}={attend_num} WHERE rfid="{rfid}"')
        print('【記録】')

    # 登録済み and タッチ済
    elif res:
        subprocess.Popen(['mpg321', f'{PATH}sounds/{SE_ALREADY}', '-q'])
        print('【記録済み】')

    # 未登録
    else:
        all_data = main_sheet.get_all_values()
        idm_list = [d[7] for d in all_data]

        for i, value in enumerate(idm_list[1:]):
            # 新規登録(rfid が 10文字以下)があるか
            if len(value) > 0 and len(value) < 11:
                subprocess.Popen(['mpg321', f'{PATH}sounds/{SE_NEW}', '-q'])
                print(f'【新規登録】{i+1}, {value}')
                try:
                    cur.execute(f'INSERT INTO user_data(rfid, last_touch) VALUES ("{rfid}", {date})')
                    subprocess.Popen(['mpg321', f'{PATH}sounds/{SE_NEW}', '-q'])
                    print(f'【新規登録】{i+1}, {value}')
                except:
                    subprocess.Popen(['mpg321', f'{PATH}sounds/{SE_ERROR}', '-q'])
                    print(f'【登録失敗】')

                return

        subprocess.Popen(['mpg321', f'{PATH}sounds/{SE_NONE}', '-q'])
        print('【未登録】')


if __name__ == '__main__':
    t_delta=datetime.timedelta(hours = 9)
    JST=datetime.timezone(t_delta, 'JST')

    conn=sqlite3.connect(f'{PATH}check.db', isolation_level = None)
    cur=conn.cursor()

    # テーブルがなかった場合は作る
    cur.execute('''
        CREATE TABLE IF NOT EXISTS user_data(
            rfid       TEXT PRIMARY KEY,
            last_touch INTEGER,
            jan	       INTEGER DEFAULT 0,
            feb	       INTEGER DEFAULT 0,
            mar	       INTEGER DEFAULT 0,
            apr	       INTEGER DEFAULT 0,
            may	       INTEGER DEFAULT 0,
            jun	       INTEGER DEFAULT 0,
            jul	       INTEGER DEFAULT 0,
            aug	       INTEGER DEFAULT 0,
            sep	       INTEGER DEFAULT 0,
            oct	       INTEGER DEFAULT 0,
            nov	       INTEGER DEFAULT 0,
            dec	       INTEGER DEFAULT 0,
        )''')

    cr=MyCardReader()

    PATH=os.path.dirname(os.path.abspath(__file__)) + '/'
    scope=['https://spreadsheets.google.com/feeds',
             'https://www.googleapis.com/auth/drive']
    creds = ServiceAccountCredentials.from_json_keyfile_name(
        f'{PATH}client_secret.json', scope)
    client = gspread.authorize(creds)
    sht = client.open_by_key("16QxcHhQBNo5RL1LgPiPMn4GKsOEd6FmAX4trDBPfN0Q")

    target_data = client.open(SPREADSHEET_APP)
    main_sheet = target_data.worksheet(ROSTER_NAME)

    print('⚡️ Attend System is running!')
    while True:
        cr.read_id()
