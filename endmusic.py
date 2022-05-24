import datetime
import time
import subprocess
import os


def endmusicStart():
    print('⚡️ End Music is running!')
    settime = None
    nowtime = None
    t_delta = datetime.timedelta(hours=9)
    JST = datetime.timezone(t_delta, 'JST')

    while True:
        if not os.path.isfile('/home/nagi/Documents/attend-misc/share/endtime.txt'):
            with open('/home/nagi/Documents/attend-misc/share/endtime.txt', 'w') as f2:
                f2.write(f'----\n18:00')

        now = datetime.datetime.now(JST)
        nowtime = now.strftime('%H:%M')

        with open('/home/nagi/Documents/attend-misc/share/endtime.txt', 'r') as f:
            times = f.read().split('\n')

            if times[0] != '----':
                with open('/home/nagi/Documents/attend-misc/share/endtime.txt', 'w') as f2:
                    f2.write(f'----\n{times[1]}')
                settime = times[0]
            elif settime == None or len(times) > 2:
                settime = times[1]
                with open('/home/nagi/Documents/attend-misc/share/endtime.txt', 'w') as f2:
                    f2.write(f'----\n{times[1]}')

        if nowtime == settime:
            print('Playback "蛍の光"')
            devnull = open('/dev/null', 'w')
            subprocess.Popen(['mpg321', '/home/nagi/Documents/attend-misc/sounds/fin.mp3'],
                             stdout=devnull, stderr=devnull)
            
            time.sleep(220)

        if nowtime == "00:00":
            settime = None

        time.sleep(20)


if __name__ == '__main__':
    endmusicStart()
