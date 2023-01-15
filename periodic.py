import datetime
import os
import sqlite3
import subprocess
import time

import gspread
from oauth2client.service_account import ServiceAccountCredentials

# TODO
MUSIC_FILE = 'hotarus_light.mp3'   # 帰宅時間の音楽
SE_REMOVE = 'succeed.mp3'          # 削除時
SE_NEW = 'succeed.mp3'             # 新規登録時

SPREADSHEET_APP = 'misc-list'      # スプレッドシートのファイル名
ROSTER_NAME = '名簿'                # 名簿シート名
SETTING_SHEET = '設定'              # 設定シート名


def str_to_int(value):
    if value.isdecimal():
        return int(value)
    else:
        return value


def gs_update(all_data=None):
    if all_data == None:
        all_data = main_sheet.get_all_values()

    datum = []
    for i, data in enumerate(all_data[1:]):
        if data[8] == '削除':
            continue

        cur.execute(f'''
            SELECT apr, may, jun, jul, aug, sep, oct, nov, dec, jan, feb, mar
            FROM user_data
            WHERE rfid="{'---' if len(data[7]) == 0 else data[7]}"
        ''')
        res = cur.fetchone()
        user_data = [str_to_int(d) for d in data[1:8]]

        if res:
            datum.append([*user_data, '', *res, sum(res)])
        else:
            datum.append([*user_data, '', *[0 for _ in range(12)], 0])

    courses = ('G', 'T', 'J')
    datum = sorted(datum, key=lambda x: (x[2], courses.index(x[1]), x[3], x[4]))
    main_sheet.batch_clear(['B2:V150'])
    main_sheet.update('B2', datum)
    print('【更新】')


def db_update():
    all_data = main_sheet.get_all_values()
    cur.execute('SELECT rfid FROM user_data')
    rfid_list_db = [r[0] for r in cur.fetchall()]
    rfid_list_gs = [data[7] for data in all_data]
    is_deleted = False

    for rfid in set(rfid_list_db) - set(rfid_list_gs):
        print(f'【削除】@DB {rfid}')
        cur.execute(f'DELETE FROM user_data WHERE rfid="{rfid}"')

    for i, data in enumerate(all_data[1:]):
        rfid = data[7] if data[7] else '---'

        if data[8] not in ['追加', '変更', '削除', '']:
            main_sheet.update_cell(i+2, 8, '')

        if data[8] == '削除':
            print(f'【削除】{rfid}')
            cur.execute(f'DELETE FROM user_data WHERE rfid="{rfid}"')
            is_deleted = True

        elif rfid not in rfid_list_db and rfid != '---':
            print(f'【新規登録】{rfid}')
            cur.execute(f'''INSERT INTO 
                user_data (rfid, last_touch)
                VALUES ("{rfid}", "0000")
            ''')

    if is_deleted:
        gs_update(all_data)


def main():
    dt_now = datetime.datetime.now(JST)
    now_time = dt_now.strftime('%H:%M')

    get_data = setting_sheet.col_values(3)
    set_times = get_data[2:]

    # 帰りの音楽を流す
    if now_time in set_times:
        print('【再生開始】蛍の光')
        subprocess.Popen(['mpg321', f'{PATH}/sounds/{MUSIC_FILE}', '-q'])

    # データを更新する
    elif now_time in ['00:00', '00:01']:
        gs_update()

    # 登録情報を更新
    else:
        db_update()


if __name__ == '__main__':
    PATH = os.path.dirname(os.path.abspath(__file__))
    conn = sqlite3.connect(f'{PATH}/db/check.db', isolation_level=None)
    cur = conn.cursor()
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

    sht = client.open(SPREADSHEET_APP)
    main_sheet = sht.worksheet(ROSTER_NAME)
    setting_sheet = sht.worksheet(SETTING_SHEET)

    t_delta = datetime.timedelta(hours=9)
    JST = datetime.timezone(t_delta, 'JST')

    print('⚡️ Run periodic.py')
    main()
