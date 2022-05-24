import binascii
import csv
import datetime
import nfc
import os
import sqlite3
import subprocess


class MyCardReader(object):
    def on_connect(self, tag):
        # タッチ時の処理
        print("【 Touched 】")

        # IDmのみ取得して表示
        self.idm = binascii.hexlify(tag.identifier).upper().decode('utf-8')
        print("IDm : " + str(self.idm))

        cur.execute(f'''
        SELECT * FROM {table_main} WHERE idm="{self.idm}"
        ''')
        res = cur.fetchone()

        # 特定のIDmだった場合のアクション
        if res:
            subprocess.Popen(['mpg321', './sounds/ppi.mp3'])
            print("【 登録されたIDです 】")
            cl = res[0] + res[1] + res[2]
            name = res[3]
            print(f"【登録済】- {name}")

            log_add(self.idm)

        else:
            subprocess.Popen(['mpg321', './sounds/err.mp3'])
            newidm(self.idm)

        return True

    def read_id(self):
        clf = nfc.ContactlessFrontend('usb')
        try:
            clf.connect(rdwr={'on-connect': self.on_connect})
        finally:
            clf.close()


def log_add(idm):
    t_delta = datetime.timedelta(hours=9)
    JST = datetime.timezone(t_delta, 'JST')
    now = datetime.datetime.now(JST)

    logfile = now.strftime('./logs/%Y-%m.csv')

    cur.execute(f'''
        SELECT * FROM {table_main} where idm="{idm}"
    ''')
    result = cur.fetchone()
    stunum = result[0]
    course = result[1]
    grade = result[2]
    class_ = result[3]
    name = result[4]
    cl2 = course + grade + class_
    first = False

    if not os.path.isfile(logfile):
        first = True

    with open(logfile, 'a') as f:
        writer = csv.writer(f)

        if first:
            writer.writerow(['date', 'time', 'stunum', 'class', 'name'])

        date = now.strftime('%Y-%m-%d')
        time = now.strftime('%H:%M:%S')

        writer.writerow([date, time, stunum, cl2, name])


def newidm(idm):
    print('【未登録】')
    with open('./share/lasttouch.txt', mode='w') as f:
        f.write(idm)


def nfcStart():
    print('⚡️ NFC checker is running!')
    dbname = 'attend'
    table_main = 'main'
    table_unreg = 'unreg'

    conn = sqlite3.connect(dbname)
    cur = conn.cursor()
    cr = MyCardReader()

    try:
        cur.execute(f'''
            CREATE TABLE {table_main}(
                stunum text,
                course text, 
                grade  text,
                class  text,
                name   text,
                idm    text)
            ''')

        while True:
            cr.read_id()

    except:
        pass


if __name__ == '__main__':
    nfcStart()
