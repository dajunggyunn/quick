from __future__ import print_function

import datetime
import os.path
import pandas as pd

import json

from openpyxl import Workbook

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

# If modifying these scopes, delete the file token.json.
SCOPES = ['https://www.googleapis.com/auth/calendar']


def createEvent(summary, startDate, endDate, description):
    return {
        'summary': summary,
        'location': '800 Howard St., San Francisco, CA 94103',
        'description': description,
        'start': {
            'dateTime': startDate,
            'timeZone': 'Asia/Seoul',
        },
        'end': {
            'dateTime': endDate,
            'timeZone': 'Asia/Seoul',
        },
        # 'recurrence': [
        #     'RRULE:FREQ=DAILY;COUNT=2'
        # ],
        'attendees': [
            {'email': 'gyen530@bigvalue.co.kr'},
            {'email': 'showjihyun@bigvalue.co.kr'},
        ],
        'reminders': {
            'useDefault': False,
            'overrides': [
                {'method': 'email', 'minutes': 24 * 60},
                {'method': 'popup', 'minutes': 10},
            ],
        },
    }

def main():
    """Shows basic usage of the Google Calendar API.
    Prints the start and name of the next 10 events on the user's calendar.
    """
    # event = {
    #     'summary': 'test3',
    #     'location': '800 Howard St., San Francisco, CA 94103',
    #     'description': 'A chance to hear more about Google\'s developer products.',
    #     'start': {
    #         'dateTime': '2022-12-03T16:16:00',
    #         'timeZone': 'Asia/Seoul',
    #     },
    #     'end': {
    #         'dateTime': '2022-12-03T16:16:01',
    #         'timeZone': 'Asia/Seoul',
    #     },
    #     # 'recurrence': [
    #     #     'RRULE:FREQ=DAILY;COUNT=2'
    #     # ],
    #     'attendees': [
    #         {'email': 'gyen530@bigvalue.co.kr'},
    #         {'email': 'showjihyun@bigvalue.co.kr'},
    #     ],
    #     'reminders': {
    #         'useDefault': False,
    #         'overrides': [
    #             {'method': 'email', 'minutes': 24 * 60},
    #             {'method': 'popup', 'minutes': 10},
    #         ],
    #     },
    # }

    creds = None
    # The file token.json stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json', SCOPES)
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open('token.json', 'w') as token:
            token.write(creds.to_json())

    try:
        service = build('calendar', 'v3', credentials=creds)

        # Call the Calendar API
        now = datetime.datetime.utcnow().isoformat() + 'Z'  # 'Z' indicates UTC time
        print('Getting the upcoming 10 events')
        # events_result = service.events().list(calendarId='primary', timeMin=now,
        #                                       maxResults=10, singleEvents=True,
        #                                       orderBy='startTime').execute()
        # events = events_result.get('items', [])

        #엑셀 읽어서

        #wb = Workbook()

        directory = os.getcwd()

        wb = pd.read_excel(directory + '\\gongsi.xlsx', usecols = [0, 1, 3, 6])

        for idx, row in wb.iterrows():
            if idx > 0:
                startDate = row['Start Date']
                startDate = startDate.strftime("%Y-%m-%dT00:00:00")
                endDate = row['End Date']
                description = row['Description']
                subject = row['Subject'] + ' [' + description + ']'
                #date = '2022-12-02T00:00:00'
                if pd.isna(endDate):
                    #endDate = startDate
                    event2 = createEvent(subject, startDate, startDate, description)
                    event2 = service.events().insert(calendarId='primary', body=event2).execute()
                else:
                    endDate = endDate.strftime("%Y-%m-%dT00:00:00")
                    event2 = createEvent(subject + '시작일', startDate, startDate, description)
                    event2 = service.events().insert(calendarId='primary', body=event2).execute()
                    event2 = createEvent(subject + '종료일', endDate, endDate, description)
                    event2 = service.events().insert(calendarId='primary', body=event2).execute()

        #print(wb)

        #wb = load_workbook(filename='gongsi.xlsx')
        #summary = ws3['AA10'].value

        #event2 = createEvent(summary, '2022-12-03T16:16:00')


        #event = service.events().insert(calendarId='primary', body=event).execute()

        # if not events:
        #     print('No upcoming events found.')
        #     return

        # Prints the start and name of the next 10 events
        # for event in events:
        #     start = event['start'].get('dateTime', event['start'].get('date'))
        #     print(start, event['summary'])

        print("캘린더 추가 완료")



    except HttpError as error:
        print('An error occurred: %s' % error)


if __name__ == '__main__':
    main()