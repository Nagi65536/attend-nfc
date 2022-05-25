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

        return True

    def read_id(self):
        clf = nfc.ContactlessFrontend('usb')

        try:
            clf.connect(rdwr={'on-connect': self.on_connect})
        finally:
            clf.close()


def registered(cell):
    user_data = main_sheet.row_values(cell.row)
    print(user_data)
    subprocess.Popen(['mpg321', f'{path}sounds/ppi.mp3', '-q'])
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

    with open(logfile, 'a') as f:
        writer = csv.writer(f)

        if first:
            worksheet = sht.add_worksheet(
                title=f'{record_sheet_name}', rows="100", cols="35")
            writer.writerow(['date', 'time', 'stunum', 'class', 'name'])
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
                record_sheet.append_row([data[6], cl, data[1], cl_sort],table_range=f'B{i}')
        except:
            pass

        # TODO: 来た日にチェック入れる処理
        # if user_data:
        #     write_cell_row = int(now.strftime('%-d')) - 6
        #     write_cell_col = int(res.col)

        #     record_sheet.update_acell(
        #         write_cell_col, write_cell_row, 'Hello World!')
        # date = now.strftime('%Y-%m-%d')
        # time = now.strftime('%H:%M:%S')

        # writer.writerow([date, time, stunum, class_, name])


def unregistered(idm):
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
