# TODO
# - Fix Bug - If the person has never entered in the spreadsheet, app will crash because it can't find the cell of a non existed data entry
# - Comment code
# -

import gspread
from oauth2client.service_account import ServiceAccountCredentials
import serial
import pandas as pd
from datetime import datetime
import time

scope = ['https://spreadsheets.google.com/feeds','https://www.googleapis.com/auth/spreadsheets','https://www.googleapis.com/auth/drive.file','https://www.googleapis.com/auth/drive']

creds = ServiceAccountCredentials.from_json_keyfile_name('creds.json', scope)

client = gspread.authorize(creds)

sheet = client.open('CoReLabCheckInList').sheet1

ser = serial.Serial('/dev/cu.usbmodem141201', 9600)  # open serial port

file = 'Book1.csv'
df = pd.read_csv(file)


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
    print('Checked In')

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
        print('Not checked out')
        checkOut(cell)
    else:
        print('Checked In')
        addEntry(n, i, t, d)

while True:
    checkIn()
