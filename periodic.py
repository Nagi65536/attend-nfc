import datetime
import time
import subprocess
import sqlite3
import os
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

target_data = client.open("misc-list")
main_sheet = target_data.worksheet("名簿")

conn = sqlite3.connect(f'{path}check.db')
cur = conn.cursor()

cur.execute('CREATE TABLE IF NOT EXISTS recorded(id INTEGER PRIMARY KEY AUTOINCREMENT, name TEXT, class TEXT, stunum TEXT)')
cur.execute('CREATE TABLE IF NOT EXISTS unrecorded(id INTEGER PRIMARY KEY AUTOINCREMENT, name TEXT, class TEXT, stunum TEXT, row INTEGER, record_sheet TEXT)')


print('⚡️ Periodic is running!')
settime = None
nowtime = None
t_delta = datetime.timedelta(hours=9)
JST = datetime.timezone(t_delta, 'JST')


def registered(stunum, name, class_, idm, record_cell_row, record_sheet_name):
    first = True

    try:
        record_sheet = sht.duplicate_sheet(
            source_sheet_id=1120016017, new_sheet_name=record_sheet_name, insert_sheet_index=1)
        users_data = main_sheet.get_all_values()
        i = 1
        for data in users_data[1:]:
            i += 1
            cl = f'{data[2]}{data[3]}{data[4]}'
            cl_sort = f'{data[3]}{data[2]}{data[4]}'
            record_sheet.append_row(
                [data[1], data[2], data[3], data[4], data[5], data[6]], table_range=f'B{i}')
    except:
        users_data = None
        record_sheet = target_data.worksheet(record_sheet_name)

    get_record_cell = record_sheet.find(stunum)
    if not get_record_cell:
        if not users_data:
            users_data = main_sheet.get_all_values()
        record_data = record_sheet.col_values(7)

        i = 1
        try:
            for data in users_data[1:]:
                if not data[6] in record_data:
                    input_row = len(record_data) + i
                    record_sheet.append_row(
                        [data[1], data[2], data[3], data[4], data[5], data[6]], table_range=f'B{input_row}')
                    i += 1
        except:
            pass
        
    get_record_cell = record_sheet.find(stunum)
    # 記録用シートに登録されているか
    if get_record_cell:
        record_cell_col = get_record_cell.row

        record_cell_row = int(now.strftime('%-d')) + 8
        record_sheet.update_cell(record_cell_col, record_cell_row, "◯")
        cur.execute(
            f'REPLACE INTO recorded(stunum, name, class, idm) VALUES("{stunum}", "{name}", "{class_}", "{idm}")')
        conn.commit()
        print('記録完了')
    else:
        record_cell_row = int(now.strftime('%-d')) + 8
        cur.execute(
            f'REPLACE INTO unrecorded(stunum, name, class, idm, row, record_sheet) VALUES("{stunum}", "{name}", "{class_}", "{idm}", "{record_cell_row}", "{record_sheet_name}")')
        conn.commit()
        print('記録できませんでした')
    print('\n---fin---\n')


def record():
    cur.execute('SELECT * FROM unrecorded')
    datas = cur.fetchall()
    cur.execute('DELETE FROM unrecorded')
    conn.commit()

    for data in datas:
        print(data)
        registered(data[0], data[1], data[2], data[3], data[4], data[5])


while True:
    if not os.path.isfile(f'{path}share/endtime.txt'):
        with open(f'{path}share/endtime.txt', 'w') as f2:
            f2.write(f'----\n18:00')

    now = datetime.datetime.now(JST)
    nowtime = now.strftime('%H:%M')
    nowmin = now.strftime('%M')

    with open(f'{path}share/endtime.txt', 'r') as f:
        times = f.read().split('\n')

        if times[0] != '----':
            with open(f'{path}share/endtime.txt', 'w') as f2:
                f2.write(f'----\n{times[1]}')
            settime = times[0]
        elif settime == None or len(times) > 2:
            settime = times[1]
            with open(f'{path}share/endtime.txt', 'w') as f2:
                f2.write(f'----\n{times[1]}')

    if nowtime == settime:
        print('【Playback "蛍の光"】')
        devnull = open('/dev/null', 'w')
        subprocess.Popen(['mpg321', f'{path}sounds/fin.mp3'],
                         stdout=devnull, stderr=devnull)

        time.sleep(220)

    if nowtime == '00:00':
        print('【リセット】')
        cur.execute('DELETE * FROM unrecorded')
        cur.execute('DELETE FROM recorded')
        conn.commit()
        settime = None
        today = now.strftime('%d')
        try:
            logfile = now.strftime(f'{path}log/%Y-%m.csv')
            record_sheet_name = now.strftime('%Y-%m')
            first = False

            sht.duplicate_sheet(source_sheet_id=1120016017,
                                new_sheet_name=record_sheet_name, insert_sheet_index=1)
            record_sheet = target_data.worksheet(record_sheet_name)
            users_data = main_sheet.get_all_values()

            i = 1
            datas = []
            for data in users_data[1:]:
                i += 1
                cl = f'{data[2]}{data[3]}{data[4]}'
                cl_sort = f'{data[3]}{data[2]}{data[4]}'
                record_sheet.append_row(
                    [data[1], data[2], data[3], data[4], data[5], data[6]], table_range=f'B{i}')
                time.sleep(1)
                print(i)
        except:
            pass
    print(nowmin)
    if nowmin == '00':
        print('【データ再登録】')
        record()

    time.sleep(20)
