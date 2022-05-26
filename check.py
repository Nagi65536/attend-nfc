import binascii
import csv
import datetime
import gspread
import nfc
import os
import subprocess
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


class MyCardReader(object):
    def on_connect(self, tag):
        self.idm = binascii.hexlify(tag.identifier).upper().decode('utf-8')
        print("IDm : " + str(self.idm))

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
    logfile = now.strftime(f'{path}log/%Y-%m.csv')
    record_sheet_name = now.strftime('%Y-%m')
    first = False

    name = user_data[1]
    class_ = f'{user_data[2]}{user_data[3]}{user_data[4]}'
    stunum = str(user_data[6])

    if not os.path.isfile(logfile):
        first = True


    if first:
        worksheet = sht.add_worksheet(
            title=f'{record_sheet_name}', rows="100", cols="35")
    try:
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
        user_data = None
        record_sheet = target_data.worksheet(record_sheet_name)

    get_record_cell = record_sheet.find(stunum)
    if get_record_cell:
        record_cell_col = get_record_cell.row
    else:
        if not user_data:
            users_data = main_sheet.get_all_values()
        record_data = record_sheet.col_values(7)

        # 月の出席テーブルを更新
        i = 1
        for data in users_data[1:]:
            if not data[6] in record_data:
                print(data)
                print(record_data)
                input_row = len(record_data) + i
                record_sheet.append_row(
                    [data[1], data[2], data[3], data[4], data[5], data[6]], table_range=f'B{input_row}')
                i += 1

    record_cell_row = int(now.strftime('%-d')) + 8
    record_sheet.update_cell(record_cell_col, record_cell_row, "◯")
    print('fin')


def unregistered(idm):
    subprocess.Popen(['mpg321', f'{path}sounds/err.mp3', '-q'])
    get_rfid = main_sheet.col_values(8)

    i = 2
    for value in get_rfid[1:]:
        if value and len(value) < 5:
            print('発見', value)
            main_sheet.update_cell(i, 8, idm)
        i += 1


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
