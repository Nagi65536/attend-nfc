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


path = '/home/nagi/Documents/attend-nfc-ver2/'

scope = ['https://spreadsheets.google.com/feeds',
         'https://www.googleapis.com/auth/drive']
creds = ServiceAccountCredentials.from_json_keyfile_name(
    f'{path}client_secret.json', scope)
client = gspread.authorize(creds)
sht = client.open_by_key("16QxcHhQBNo5RL1LgPiPMn4GKsOEd6FmAX4trDBPfN0Q")

target_data = client.open("misc-list")
main_sheet = target_data.worksheet("名簿")

conn = sqlite3.connect(f'{path}check.db')
cur = conn.cursor()

cur.execute('CREATE TABLE IF NOT EXISTS recorded(stunum TEXT PRIMARY KEY, name TEXT, class TEXT, idm TEXT, time INTEGER)')
cur.execute('CREATE TABLE IF NOT EXISTS unrecorded(stunum TEXT PRIMARY KEY, name TEXT, class TEXT, idm TEXT, row INTEGER, record_sheet TEXT)')


class MyCardReader(object):
    def on_connect(self, tag):
        self.idm = binascii.hexlify(tag.identifier).upper().decode('utf-8')
        print("IDm : " + str(self.idm))

        cur.execute(f'SELECT * FROM recorded WHERE idm="{self.idm}"')
        is_recorded = cur.fetchone()
        print('1-', is_recorded)

        if not is_recorded:
            print('in!')
            cur.execute(f'SELECT * FROM unrecorded WHERE idm="{self.idm}"')
            is_recorded = cur.fetchone()

        print('this', is_recorded)
        
        if is_recorded:
            diff_time = time.time() - int(is_recorded[4])
            print(diff_time)
            
            if diff_time < 60*30 or is_recorded[4] == 0:
                subprocess.Popen(['mpg321', f'{path}sounds/deliriteli.mp3', '-q'])
            else:
                leave_the_room(is_recorded[0])
            print('【記録済】')
        else:
            cell = main_sheet.find(self.idm)
            if cell:
                enter_the_room(cell)
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


def enter_the_room(cell):
    print('nyuusitu')
    subprocess.Popen(['mpg321', f'{path}sounds/ppi.mp3', '-q'])
    user_data = main_sheet.row_values(cell.row)
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
        
        time_now = int(time.time())
        record_cell_row = int(now.strftime('%-d')) + 10
        record_sheet.update_cell(record_cell_col, record_cell_row, "△")
        cur.execute(
            f'REPLACE INTO recorded(stunum, name, class, idm, time) VALUES("{stunum}", "{name}", "{class_}", "{user_data[7]}", "{time_now}")')
        conn.commit()
        print('記録完了')
    else:
        record_cell_row = int(now.strftime('%-d')) + 10
        cur.execute(
            f'REPLACE INTO unrecorded(stunum, name, class, idm, row, record_sheet) VALUES("{stunum}", "{name}", "{class_}", "{user_data[7]}", "{record_cell_row}", "{record_sheet_name}")')
        conn.commit()
        print('記録できませんでした')


def leave_the_room(stunum):
    print('taisitu')
    subprocess.Popen(['mpg321', f'{path}sounds/teretere.mp3', '-q'])
    t_delta = datetime.timedelta(hours=9)
    JST = datetime.timezone(t_delta, 'JST')
    now = datetime.datetime.now(JST)
    record_sheet_name = now.strftime('%Y-%m')
    record_sheet = target_data.worksheet(record_sheet_name)
    get_record_cell = record_sheet.find(stunum)
    record_cell_col = get_record_cell.row

    record_cell_row = int(now.strftime('%-d')) + 10
    record_sheet.update_cell(record_cell_col, record_cell_row, "◯")
    cur.execute(
        f'UPDATE recorded SET time=0 WHERE stunum="{stunum}"')
    conn.commit()
    print('更新完了')


def unregistered(idm):
    print('【未登録】')
    get_rfid = main_sheet.col_values(8)

    is_detection = False
    i = 2
    for value in get_rfid[1:]:
        if value and len(value) < 15:
            try:
                subprocess.Popen(['mpg321', f'{path}sounds/popi.mp3', '-q'])
                print('【新規登録】\n', value)
                main_sheet.update_cell(i, 8, idm)
                is_detection = True
            except:
                pass
            
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
