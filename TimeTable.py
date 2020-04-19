
from __future__ import print_function
import httplib2
import os
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from apiclient import discovery
from oauth2client import client
from oauth2client import tools
from oauth2client.file import Storage
import time
import datetime

flags = None
SCOPES = 'https://www.googleapis.com/auth/calendar'
CLIENT_SECRET_FILE = 'client_secret.json'
APPLICATION_NAME = 'Time Table App'

def get_calendar_credentials():
    """Gets valid user credentials from storage.

    If nothing has been stored, or if the stored credentials are invalid,
    the OAuth2 flow is completed to obtain the new credentials.

    Returns:
        Credentials, the obtained credential.
    """
    home_dir = os.path.expanduser('~')
    credential_dir = os.path.join(home_dir, '.credentials')
    if not os.path.exists(credential_dir):
        os.makedirs(credential_dir)
    credential_path = os.path.join(credential_dir,
                                   'TimeTable-app.json')

    store = Storage(credential_path)
    credentials = store.get()
    if not credentials or credentials.invalid:
        flow = client.flow_from_clientsecrets(CLIENT_SECRET_FILE, SCOPES)
        flow.user_agent = APPLICATION_NAME
        if flags:
            credentials = tools.run_flow(flow, store, flags)
        print('Storing credentials to ' + credential_path)
    return credentials

def get_sheet(worksheet):
  scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
  creds = ServiceAccountCredentials.from_json_keyfile_name('sheets_secret.json', scope)
  client = gspread.authorize(creds)
  print("sheets creds collected")
  sheet = client.open('Time Table').get_worksheet(worksheet-1)
  return sheet
  
def main(worksheet):
    credentials = get_calendar_credentials()
    http = credentials.authorize(httplib2.Http())
    CAL = discovery.build('calendar', 'v3', http=http)

    GMT_OFF = '+03:00'
    print('calendar creds collected')
    sheet = get_sheet(worksheet)
    print("got the sheet")
    x = 0

    year = sheet.cell(13,2).value
    month = sheet.cell(14,2).value
    startDay = int(sheet.cell(15,2).value)
    '''
    if sheet.cell(14,5).value == 'y':
        recurring = "RRULE:FREQ=WEEKLY;UNTIL=20200701T170000Z;"
    else:
        recurring = ""
    '''
    for row in range(2,12):
        for col in range(2,8): 
            time.sleep(1) #google has a limit of 1 request per second.
            value = sheet.cell(row,col).value
            x+=1
            if value:
                start = sheet.cell(row,1).value[0:5]
                end = sheet.cell(row,1).value[8:13]
                day = str(startDay-2+col)
                if int(day)>30:
                    startDay = startDay-30
                    month = str(int(month)+1)
                    day = str(startDay-2+col) #recalculate day
                EVENT = {
                'summary': value,
                'start':   {'dateTime': '%s-%s-%sT%s:00%s' % (year, month, day, start, GMT_OFF),
                            },
                'end':     {'dateTime': '%s-%s-%sT%s:00%s' % (year, month, day, end, GMT_OFF),
                            },
                'colorId': '5',
               # "recurrence": [recurring,],
                }

                e = CAL.events().insert(calendarId='primary',
                    sendNotifications=True, body=EVENT).execute()
                print('''*** %r event added:
                Start: %s
                End:   %s''' % (e['summary'].encode('utf-8'),
                    e['start']['dateTime'], e['end']['dateTime']))
            print(x)

if __name__ == '__main__':
    main()