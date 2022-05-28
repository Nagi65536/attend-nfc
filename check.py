import binascii
import csv
import datetime
import gspread
import nfc
import os
import subprocess
import sqlite3
import time
from oauth2client.service_account import ServiceAccountCredentials


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

cur.execute('CREATE TABLE IF NOT EXISTS recorded(stunum TEXT PRIMARY KEY, name TEXT, class TEXT, idm TEXT)')
cur.execute('CREATE TABLE IF NOT EXISTS unrecorded(stunum TEXT PRIMARY KEY, name TEXT, class TEXT, idm TEXT, row INTEGER, record_sheet TEXT)')


class MyCardReader(object):
    def on_connect(self, tag):
        self.idm = binascii.hexlify(tag.identifier).upper().decode('utf-8')
        print("IDm : " + str(self.idm))

        cur.execute(f'SELECT * FROM recorded WHERE idm="{self.idm}"')
        is_recorded = cur.fetchall()
        
        if is_recorded:
            subprocess.Popen(['mpg321', f'{path}sounds/piyopiyo.mp3', '-q'])
            print('【記録済】')
        else:
            cell = main_sheet.find(self.idm)
            if cell:
                registered(cell)
            else:
                unregistered(self.idm)
        
        print('\n---fin---\n')
        return True

    def read_id(self):
        clf = nfc.ContactlessFrontend('usb')

        try:
            clf.connect(rdwr={'on-connect': self.on_connect})
        finally:
            clf.close()


def registered(cell):
    subprocess.Popen(['mpg321', f'{path}sounds/ppi.mp3', '-q'])
    user_data = main_sheet.row_values(cell.row)
    print(user_data)
    print("【 登録済 】")

    t_delta = datetime.timedelta(hours=9)
    JST = datetime.timezone(t_delta, 'JST')
    now = datetime.datetime.now(JST)
    record_sheet_name = now.strftime('%Y-%m')
    first = True

    name = user_data[1]
    class_ = f'{user_data[2]}{user_data[3]}{user_data[4]}'
    stunum = str(user_data[6])

    try:
        record_sheet = sht.duplicate_sheet(
            source_sheet_id=1120016017, new_sheet_name=record_sheet_name, insert_sheet_index=1)
        users_data = main_sheet.get_all_values()
        i = 1
        for data in users_data[1:]:
            print(data)
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
                    print(data)
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
            f'REPLACE INTO recorded(stunum, name, class, idm) VALUES("{stunum}", "{name}", "{class_}", "{user_data[7]}")')
        conn.commit()
        print('記録完了')
    else:
        record_cell_row = int(now.strftime('%-d')) + 8
        cur.execute(
            f'REPLACE INTO unrecorded(stunum, name, class, idm, row, record_sheet) VALUES("{stunum}", "{name}", "{class_}", "{user_data[7]}", "{record_cell_row}", "{record_sheet_name}")')
        conn.commit()
        print('記録できませんでした')


def unregistered(idm):
    print('【未登録】')
    get_rfid = main_sheet.col_values(8)

    i = 2
    for value in get_rfid[1:]:
        if value and len(value) < 15:
            subprocess.Popen(['mpg321', f'{path}sounds/popi.mp3', '-q'])
            print('【新規登録】\n', value)
            main_sheet.update_cell(i, 8, idm)
            is_detection = True
            break
        else:
            is_detection = False
        i += 1

    if not is_detection:
        subprocess.Popen(['mpg321', f'{path}sounds/err.mp3', '-q'])


def conv_num_to_col(num):
    if num <= 26:
        return chr(64 + num)
    else:
        if num % 26 == 0:
            return conv_num_to_col(num//26-1) + 'Z'
        else:
            return conv_num_to_col(num//26) + chr(64+num % 26)


cr = MyCardReader()
print('⚡️ NFC checker is running!')
while True:
    cr.read_id()
