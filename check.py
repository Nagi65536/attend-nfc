import binascii
import csv
import datetime
import math
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


# 入退室記録の登録
def record(idm):
    input('aaa')
    global record_sheet
    global current_sheet_date
    global get_all_data

    # 現在日時•時刻設定
    JST = datetime.timezone(datetime.timedelta(hours=9), 'JST')
    dt_now = datetime.datetime.now(JST)

    # 今日の日付と今月のシート名
    todays_date = dt_now.strftime('%Y-%m-%d')
    todays_sheet_date = dt_now.strftime('%Y-%m')

    # 現在のシートが今月のではない場合
    if current_sheet_date != todays_sheet_date:
        record_sheet = target_data.worksheet(todays_sheet_date)
        current_sheet_date = todays_sheet_date

    # データがなければ取得
    if not get_all_data:
        get_all_data = main_sheet.get_all_values()

    idm_list = [d[7] for d in get_all_data]

    # idm が未登録の場合
    if idm not in idm_list:
        get_all_data = None

        for i, value in enumerate(idm_list[1:]):
            if len(value) > 0 and len(value) < 11:
                print(f'【新規登録】{i+1}, {value}')
                subprocess.Popen(['mpg321', f'{path}sounds/popi.mp3', '-q'])
                main_sheet.update_cell(i+2, 8, idm)
                return

        print('【新規追加なし】')

    try:
        record_txt = ''
        time_str = dt_now.strftime('%H:%M:%S')
        time_int = time.time()

        # 今日以外のデータを削除し、idmが一致するデータを取得
        cur.execute(
            f'DELETE FROM record WHERE last_entry_date!="{todays_date}"')
        cur.execute(
            f'SELECT entry_time_int, last_entry_date, current_data FROM record WHERE idm="{idm}"')
        res = cur.fetchone()

        # データベースにデータがあるか
        if res and not res[0] == 0:
            subprocess.Popen(['mpg321', f'{path}sounds/teretere.mp3', '-q'])
            diff_time = time_int - res[0]
            tmp = math.floor(diff_time/60)
            record_txt = tmp if tmp < 60*30 else '◯'
            cur.execute(
                f'UPDATE record SET exit_time="{time_str}", exit_time_int="{time_int}", current_data="{record_txt}" WHERE idm="{idm}"')

        else:
            subprocess.Popen(['mpg321', f'{path}sounds/ppi.mp3', '-q'])
            record_txt = '0'

            if not res:
                cur.execute(
                    f'''INSERT INTO record VALUES("{idm}", "{time_str}", "{time_int}", 0, 0,"{todays_date}", "{record_txt}")''')
            elif res[0] == 0:
                cur.execute(
                    f'''UPDATE record SET entry_time="{time_str}", entry_time_int="{time_int}", current_data="{record_txt}" WHERE idm="{idm}"''')

        # 値が更新されていない場合
        if res and str(res[2]) == str(record_txt):
            print('値は変更されていません')
            return

        # 生徒番号リストと生徒番号の取得
        student_num_list = record_sheet.col_values(7)
        student_num = get_all_data[idm_list.index(idm)][6]
        print("【生徒番号】", student_num)

        # 記録の追加または記録位置の取得
        if student_num not in student_num_list:
            record_col = len(student_num_list)
            data = get_all_data[idm_list.index(idm)]
            record_sheet.append_row(
                [None, data[1], data[2], data[3], data[4], data[5], data[6]], table_range=f'A{record_col}')
            print(f'【記録表新規追加】A{record_col} {data[1]} {data[6]}')
        else:
            record_col = student_num_list.index(student_num) + 1

        # スプレッドシートに書き込む
        record_row = int(dt_now.strftime('%-d')) + 10
        record_sheet.update_cell(record_col, record_row, record_txt)
        print(f'【記録】 "{record_txt}" {record_col}列目 {record_row - 10}日')

    except err:
        subprocess.Popen(['mpg321', f'{path}sounds/err.mp3', '-q'])
        print(f'【ERROR】{err}')

    else:
        # データを更新
        if get_all_data:
            get_all_data = main_sheet.get_all_values()


if __name__ == '__main__':
    target_data = client.open("misc-list")
    main_sheet = target_data.worksheet("名簿")
    record_sheet = None
    current_sheet_date = None

    conn = sqlite3.connect(f'{path}check.db', isolation_level=None)
    cur = conn.cursor()
    cur.execute('''
        CREATE TABLE IF NOT EXISTS record(
            idm             TEXT PRIMARY KEY,
            entry_time      TEXT DEFAULT 0,
            entry_time_int  INTEGER DEFAULT 0,
            exit_time       TEXT DEFAULT 0,
            exit_time_int   INTEGER DEFAULT 0,
            last_entry_date TEXT DEFAULT 0,
            current_data    TEXT DEFAULT "x"
        )''')

    cr = MyCardReader()
    get_all_data = None

    print('⚡️ Attend System is running!')
    while True:
        # cr.read_id()
        record('01010910DF16BF0E')