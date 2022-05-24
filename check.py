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
creds = ServiceAccountCredentials.from_json_keyfile_name('client_secret.json', scope)
client = gspread.authorize(creds)
sht = client.open_by_key("16QxcHhQBNo5RL1LgPiPMn4GKsOEd6FmAX4trDBPfN0Q")

target_data = client.open("misc-list")
main_sheet = target_data.worksheet("情報システム部")


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
    data = main_sheet.row_values(cell.row)
    subprocess.Popen(['mpg321', f'{path}sounds/ppi.mp3','-q'])
    print("【 登録済 】")

    t_delta = datetime.timedelta(hours=9)
    JST = datetime.timezone(t_delta, 'JST')
    now = datetime.datetime.now(JST)
    logfile = now.strftime(f'{path}log/%Y-%m.csv')
    record_sheet = now.strftime('%Y-%m')
    first = False

    name = data[1]
    class_ = f'{data[2]}{data[3]}{data[4]}'
    stunum = str(data[6])

    if not os.path.isfile(logfile):
        first = True

    with open(logfile, 'a') as f:
        print('okok')
        writer = csv.writer(f)

        if first:
            worksheet = sht.add_worksheet(title=f'{record_sheet}',rows="100",cols="35")
            writer.writerow(['date', 'time', 'stunum', 'class', 'name'])

        res = record_sheet.find(stunum)
        if res:
            # 記録する
            # record_sheet.update_acell('A1', 'Hello World!')
        date = now.strftime('%Y-%m-%d')
        time = now.strftime('%H:%M:%S')

        writer.writerow([date, time, stunum, class_, name])


def unregistered(idm):
    subprocess.Popen(['mpg321', f'{path}sounds/err.mp3'])
    
cr = MyCardReader()
print('⚡️ NFC checker is running!')
while True:
    cr.read_id()