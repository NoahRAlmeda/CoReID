# TODO
# - Comment code
# - Fix Bug: Token expires every hour
# - Fix Error: When a ID is not found it crashes
# - Create function that lets me add user info to new card Ids
# - Create Readme file
# - Come up with a solution in case if the student doesn't check out what will the program do
# -

import gspread
from oauth2client.service_account import ServiceAccountCredentials
import serial
import pandas as pd
from datetime import datetime
import time
import matplotlib.pyplot as plt
import seaborn as sns

scope = ['https://spreadsheets.google.com/feeds','https://www.googleapis.com/auth/spreadsheets','https://www.googleapis.com/auth/drive.file','https://www.googleapis.com/auth/drive']
creds = ServiceAccountCredentials.from_json_keyfile_name('creds.json', scope)
global client
client = gspread.authorize(creds)
global sheet
sheet = client.open('CoReLabCheckInList').sheet1
ser = serial.Serial('/dev/cu.usbmodem141401', 9600)  # open serial port

file = 'Book1.csv'
df = pd.read_csv(file)

def refreshToken():
    try:
        if creds.access_token_expired:
            print("Token Refreshed!")
            gs_client.login()
    except Exception:
        traceback.print_exc()

def scanCard():
    global ser
    while(True):
        if(ser.inWaiting() > 0):
            print("READING DATA")
            return ser.readline().decode('UTF-8').strip()

def findCell(n, i, t, d):
    try:
        return (sheet.find(n), False)
    except gspread.exceptions.CellNotFound:
        addEntry(n, i, t, d)
        return (sheet.find(n), True)

def addEntry(n, i, t, d):
    sheet.insert_row([n, i, t, d, '', 'False'], 2)
    print('Add Entry')

def getRowValue(row):
    return sheet.row_values(row)

def checkOut(cell):
    outTimeCell = 'E' + str(cell.row)
    switchToTrue = 'F' + str(cell.row)
    sheet.update_acell(outTimeCell, str(datetime.now().strftime('%H:%M:%S')))
    sheet.update_acell(switchToTrue, 'True')

def checkIn():
    print("WAITING...")
    cardId = scanCard()
    print("CARDID: ", cardId)
    name = df[df['Id'] == cardId]
    n = name['name'].values[0]
    i = name['studentId'].values[0]
    t = datetime.now().strftime('%H:%M:%S')
    d = datetime.now().strftime('%Y-%m-%d')
    cell, can_read = findCell(n, i, t, d)
    values_list = getRowValue(cell.row)
    print(values_list)

    if not can_read and (values_list[5] == 'False'):
        print('Cheking Out...')
        checkOut(cell)
    else:
        if can_read:
            print('Checked In Already')
        else:
            addEntry(n, i, t, d)

# def checkActivity(studentID):
#     listOflist = sheet.get_all_values()
#     allEntry = pd.DataFrame(listOflist, columns=['Student Name', 'ID', 'Time In', 'Date', 'Time Out', 'Checked Out']).drop(0)
#     studentEntries = allEntry[allEntry['ID'] == studentID]
#     # Create Graph showing the activity of a specific student
#     plt.plot(studentEntries['Date'], studentEntries['Time In'])
#     plt.show()


while True:
    refreshToken()
    checkIn()

# ======= This is en extra feature
# studID = input('Enter student number:')
# checkActivity(str(studID))
