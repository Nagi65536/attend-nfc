import datetime
import time
import subprocess
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


print('⚡️ End Music is running!')
settime = None
nowtime = None
t_delta = datetime.timedelta(hours=9)
JST = datetime.timezone(t_delta, 'JST')

while True:
    if not os.path.isfile(f'{path}share/endtime.txt'):
        with open(f'{path}share/endtime.txt', 'w') as f2:
            f2.write(f'----\n18:00')

    now = datetime.datetime.now(JST)
    nowtime = now.strftime('%H:%M')
    today = now.strftime('%d')

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
        print('Playback "蛍の光"')
        devnull = open('/dev/null', 'w')
        subprocess.Popen(['mpg321', f'{path}sounds/fin.mp3'],
                            stdout=devnull, stderr=devnull)
        
        time.sleep(220)

    if nowtime == "00:00":
        settime = None
        if today == '01':
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
    
    time.sleep(20)
