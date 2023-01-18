import binascii
import datetime
import os
import sqlite3
import subprocess

import gspread
import nfc
from oauth2client.service_account import ServiceAccountCredentials

# TODO
SE_ENTRY = 'maou_system46.mp3'      # 記録時
SE_ALREADY = 'maou_system23.mp3'    # 記録済時
SE_NEW = 'succeed.mp3'              # 新規登録時
SE_NONE = 'maou_system35.mp3'       # 未登録時
SE_ERROR = 'pickup02.mp3'           # システムエラー
SE_BOOT = 'crrect_answer1.mp3'      # 起動音

SPREADSHEET_APP = 'misc-list'       # スプレッドシートのファイル名
ROSTER_NAME = '名簿'                 # 名簿シート名


class MyCardReader(object):
    # カードがタッチされたら呼び出される
    def on_connect(self, tag):
        try:
            self.rfid = binascii.hexlify(tag.identifier).upper().decode('utf-8')
            print(f'rfid : {self.rfid}')
            record(self.rfid)
        finally:
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
    date = dt_now.strftime('%m%d')
    month = int(dt_now.strftime('%m'))
    month_str = ["jan", "feb", "mar", "apr", "may",
                 "jun", "jul", "aug", "sep", "oct", "nov", "dec"][month-1]

    cur.execute(
        f'SELECT {month_str}, last_touch FROM user_data WHERE rfid="{rfid}"')
    res = cur.fetchone()

    # 登録済み and 未タッチ
    if res and res[1] != date:
        subprocess.Popen(['mpg321', f'{PATH}/sounds/{SE_ENTRY}', '-q'])
        attend_num = res[0] + 1
        cur.execute(
            f'UPDATE user_data SET last_touch="{date}", {month_str}={attend_num} WHERE rfid="{rfid}"')
        print('【記録】')

    # 登録済み and タッチ済
    elif res:
        subprocess.Popen(['mpg321', f'{PATH}/sounds/{SE_ALREADY}', '-q'])
        print('【記録済み】')

    # 未登録
    else:
        all_data = main_sheet.get_all_values()

        for i, data in enumerate(all_data[1:]):
            # 追加なのに既に RFID が入っている
            if data[8] == '追加' and len(data[7]) > 10:
                print(f'【追加失敗】')
                main_sheet.update(f'I{i+2}', [['ERROR']])
                subprocess.Popen(
                    ['mpg321', f'{PATH}/sounds/{SE_ERROR}', '-q'])
                return

            elif data[8] == '変更' and len(data[7]) <= 10:
                print(f'【変更失敗】')
                main_sheet.update(f'I{i+2}', [['ERROR']])
                subprocess.Popen(
                    ['mpg321', f'{PATH}/sounds/{SE_ERROR}', '-q'])
                return

            # 新規追加する
            elif data[8] == '追加':
                subprocess.Popen(['mpg321', f'{PATH}/sounds/{SE_NEW}', '-q'])
                try:
                    cur.execute(f'''INSERT INTO 
                        user_data (rfid, last_touch)
                        VALUES ("{rfid}", "0000"
                    )
                    ''')
                    main_sheet.update(f'H{i+2}', [[rfid, '']])
                    print(f'【新規登録】{i+2}, {rfid}')
                    subprocess.Popen(
                        ['mpg321', f'{PATH}/sounds/{SE_NEW}', '-q'])
                except:
                    print(f'【登録失敗】')
                    subprocess.Popen(
                        ['mpg321', f'{PATH}/sounds/{SE_ERROR}', '-q'])
                return

            # RFID を変更する
            elif data[8] == '変更':
                subprocess.Popen(['mpg321', f'{PATH}/sounds/{SE_NEW}', '-q'])
                try:
                    cur.execute(f'''UPDATE
                        user_data SET
                        rfid="{rfid}"
                        WHERE rfid="{data[7]}"
                    ''')
                    main_sheet.update(f'H{i+2}', [[rfid, '']])
                    print(f'【更新】{i+2}, {rfid}')
                    subprocess.Popen(
                        ['mpg321', f'{PATH}/sounds/{SE_NEW}', '-q'])
                except:
                    print(f'【更新失敗】')
                    subprocess.Popen(
                        ['mpg321', f'{PATH}/sounds/{SE_ERROR}', '-q'])
                return

        print('【未登録】')
        subprocess.Popen(['mpg321', f'{PATH}/sounds/{SE_NONE}', '-q'])
        return


if __name__ == '__main__':
    PATH = os.path.dirname(os.path.abspath(__file__))
    conn = sqlite3.connect(f'{PATH}/db/check.db', isolation_level=None)
    cur = conn.cursor()

    # テーブルがなかった場合は作る
    cur.execute('''
        CREATE TABLE IF NOT EXISTS user_data(
            rfid       TEXT PRIMARY KEY,
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
            last_touch TEXT
        )''')

    scope = ['https://spreadsheets.google.com/feeds',
             'https://www.googleapis.com/auth/drive']
    creds = ServiceAccountCredentials.from_json_keyfile_name(
        f'{PATH}/client_secret.json', scope)
    client = gspread.authorize(creds)
    sht = client.open_by_key("16QxcHhQBNo5RL1LgPiPMn4GKsOEd6FmAX4trDBPfN0Q")

    sht = client.open(SPREADSHEET_APP)
    main_sheet = sht.worksheet(ROSTER_NAME)

    t_delta = datetime.timedelta(hours=9)
    JST = datetime.timezone(t_delta, 'JST')

    cr = MyCardReader()

    print('⚡️ Attend System is running!')
    subprocess.Popen(['mpg321', f'{PATH}/sounds/{SE_BOOT}', '-q'])
    while True:
        cr.read_id()
