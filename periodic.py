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

SPREADSHEET_APP = 'misc-list'      # スプレッドシートのファイル名
ROSTER_NAME = '名簿'                # 名簿シート名
SETTING_SHEET = '設定'              # 設定シート名


def main():
    dt_now = datetime.datetime.now(JST)
    now_time = dt_now.strftime('%H:%M')

    get_data = setting_sheet.col_values(3)
    set_times = [get_data[2], get_data[3]]

    if now_time in set_times:
        # 帰りの音楽を流す
        print('【再生開始】蛍の光')
        subprocess.Popen(['mpg321', f'{PATH}/sounds/{MUSIC_FILE}', '-q'])
        return

    elif now_time in ['00:00', '00:01']:
        # データを更新する
        all_data = main_sheet.get_all_values()
        main_sheet.clear()

        main_sheet.append_row([
            '', '氏名', '学科', '学年', 'クラス', '番号', '学籍番号', 'RFID', '',
            '4月', '5月', '6月', '7月', '8月', '9月', '10月', '11月', '12月', '1月', '2月', '3月'
        ])
        for i, data in enumerate(all_data[1:]):
            cur.execute(f'''SELECT
                apr, may, jun, jul, aug, sep,
                oct, nov, dec, jan, feb, mar
                FROM user_data
                WHERE rfid="{data[7]}"
            ''')
            res = cur.fetchone()
            main_sheet.append_row([i+1, *data[1:9], *res])
            time.sleep(1)
        return

    else:
        # 削除されたデータがないかの確認
        cur.execute('SELECT rfid FROM user_data')
        rfid_list_db = [r[0] for r in cur.fetchall()]
        rfid_list_gs = main_sheet.col_values(8)[1:]
        diff_rfid = set(rfid_list_db) ^ set(rfid_list_gs)

        for rfid in diff_rfid:
            if len(rfid) < 11:
                pass
            print(f'【削除】{rfid}')
            cur.execute(f'SELECT rfid FROM user_data rfid="{rfid}"')
            if cur.fetchone():
                cur.execute(f'DELETE FROM user_data WHERE rfid="{rfid}"')
                subprocess.Popen(
                    ['mpg321', f'{PATH}/sounds/{SE_REMOVE}', '-q'])
                time.sleep(1)
        return


if __name__ == '__main__':
    PATH = os.path.dirname(os.path.abspath(__file__))
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

    conn = sqlite3.connect(f'{PATH}/check.db', isolation_level=None)
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
            last_touch INTEGER,
        )''')

    print('⚡️ Run periodic.py')
    main()
