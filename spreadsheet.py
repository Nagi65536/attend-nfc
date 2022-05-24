import cv2
import gspread
from oauth2client.service_account import ServiceAccountCredentials

scope =['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
creds = ServiceAccountCredentials.from_json_keyfile_name('client_secret.json', scope)
client = gspread.authorize(creds)

misc_list = client.open("misc-list").list
cap = cv2.VideoCapture(0)
detector = cv2.QRCodeDetector()
keep = None

while True:
    ret, frame = cap.read()
    data = detector.detectAndDecode(frame)

    if data[0] and data[0] != keep:
        print(data[0])
        keep = data[0]
    elif data[0] != keep:
        keep = None

cap.release()
cv2.destroyAllWindows()