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


def main():
    dt_now = datetime.datetime.now(JST)
    now_time = dt_now.strftime('%H:%M')

    get_data = setting_sheet.col_values(3)
    set_times = [get_data[2], get_data[3]]

    # 帰りの音楽を流す
    if now_time in set_times:
        print('【再生開始】蛍の光')
        subprocess.Popen(['mpg321', f'{PATH}/sounds/{MUSIC_FILE}', '-q'])
        return

    # データを更新する
    elif now_time in ['15:21', '00:01']:
        all_data = main_sheet.get_all_values()
        main_sheet.batch_clear(['B2:V150'])

        datum = []
        for i, data in enumerate(all_data[1:]):
            cur.execute(f'''
                SELECT apr, may, jun, jul, aug, sep, oct, nov, dec, jan, feb, mar
                FROM user_data
                WHERE rfid="{'-' if len(data[7]) == 0 else data[7]}"
            ''')
            res = cur.fetchone()
            user_data = [str_to_int(d) for d in data[1:8]]
            if res:
                datum.append([*user_data, '', *res, sum(res)])
            else:
                datum.append([*user_data, '', *[0 for _ in range(12)], 0])

        main_sheet.update('B2', datum)
        print('【更新】')
        return

    # 登録情報を更新
    else:
        cur.execute('SELECT rfid FROM user_data')
        rfid_list_db = [r[0] for r in cur.fetchall()]
        rfid_list_gs = main_sheet.col_values(8)[1:]
        diff_rfid = list(set(rfid_list_db) ^ set(rfid_list_gs))

        for rfid in diff_rfid[1:]:
            if len(rfid) == 0:
                pass

            elif rfid in rfid_list_db:
                print(f'【削除】{rfid}')
                cur.execute(f'SELECT rfid FROM user_data rfid="{rfid}"')
                if cur.fetchone():
                    cur.execute(f'DELETE FROM user_data WHERE rfid="{rfid}"')

            else:
                print(f'【新規登録】{rfid}')
                cur.execute(f'''INSERT INTO 
                    user_data (rfid, last_touch)
                    VALUES ("{rfid}", "0000")
                ''')

        return


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
