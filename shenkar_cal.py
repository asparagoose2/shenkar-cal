import camelot
import sys
import argparse
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from datetime import datetime
import pickle
import os.path
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from datetime import timedelta
import requests 
import pandas as pd

# If modifying these scopes, delete the file token.pickle.
SCOPES = ['https://www.googleapis.com/auth/calendar.events']

creds = None
service = None

def semester_schedule(table, parm):
    for index, row in  table.iterrows():
        title = row[4]
        teacher = row[1]
        start = row[7]
        end = row[6]
        day = row[8]
        splited_end = end.split(":")
        if not start or not end:
            print("Unable to add this event: " + title)
            continue
        startTime = datetime.strftime(parm.start, '%d/%m/%Y') + ' ' + str(start)
        firstDay = datetime.strptime(startTime, '%d/%m/%Y %H:%M')
        
        if day == 'ב':
            firstDay += timedelta(days=1)
        elif day == 'ג':
            firstDay += timedelta(days=2)
        elif day == 'ד':
            firstDay += timedelta(days=3)
        elif day == 'ה':
            firstDay += timedelta(days=4)
        elif day == 'ו':
            firstDay += timedelta(days=5)

        endTime = firstDay.replace(hour=int(splited_end[0]),minute=int(splited_end[1]))

        event = {
            'summary': title [::-1],
            "description": teacher [::-1],
            'start': {
                'dateTime': firstDay.strftime("%Y-%m-%d")+'T'+firstDay.strftime("%H:%M:%S"),
                'timeZone': 'Asia/Jerusalem'
            },
            'end': {
                'dateTime': endTime.strftime("%Y-%m-%d")+'T'+endTime.strftime("%H:%M:%S"),
                'timeZone': 'Asia/Jerusalem'
            },
            'recurrence': [
                'RRULE:FREQ=WEEKLY;UNTIL={end_day}T000000Z'.format(end_day = datetime.strftime(parm.end, '%Y%m%d')),
            ],
            }
        
        print(event)
        print('\n')
        recurring_event = service.events().insert(calendarId='0836ibjeeoiqh1ldb1lii447ig@group.calendar.google.com', body=event).execute()


def exams(table,moed):
    for index, row in  table.iterrows():
        title = row[3]
        start = row[6]
        end = row[5]
        day = row[7]
        splited_end = end.split(":")
        startTime = day + ' ' + str(start)
        firstDay = datetime.strptime(startTime, '%d/%m/%Y %H:%M')
        


        endTime = firstDay.replace(hour=int(splited_end[0]),minute=int(splited_end[1]))

        event = {
            'summary': title [::-1],
            'colorId': moed,
            'start': {
                'dateTime': firstDay.strftime("%Y-%m-%d")+'T'+firstDay.strftime("%H:%M:%S"),
                'timeZone': 'Asia/Jerusalem'
            },
            'end': {
                'dateTime': endTime.strftime("%Y-%m-%d")+'T'+endTime.strftime("%H:%M:%S"),
                'timeZone': 'Asia/Jerusalem'
            },
            }
        print(event)
        print('\n')
        event = service.events().insert(calendarId='0836ibjeeoiqh1ldb1lii447ig@group.calendar.google.com', body=event).execute()


def year_pdf(table):
    for index, row in  table.iterrows():
        if not row[1]:
            continue
        date = row[1]
        start = None
        end = None
        if '-' in date:
            start = date.split('-')[0] + '.' + date.split('.',1)[1]
            start = datetime.strptime(start, '%d.%m.%y')
            end = date.split('-')[1]
            end = datetime.strptime(end, '%d.%m.%y')
        else:
            date = datetime.strptime(date, '%d.%m.%y')
        title = row[3]
        note = row[0]
        title = title.replace('(cid:26)', 'ם')
        title = title.replace('(cid:23)', 'ם')
        title = title.replace('(cid:16)', 'ן')
        title = title.replace('(cid:31)', 'ן')
        title = title.replace('(cid:11)', 'ס')
        title = title.replace('(cid:12)', 'ס')
        title = title.replace('(cid:4)', 'ץ')
        title = title.replace('(cid:10)', 'ץ')
        title = title.replace('(cid:13)', "'")
        title = title.replace('(cid:30)', '"')
        title = title.replace('(cid:9)', '"')
        

        event = {
            'summary': title [::-1],
            "description": note [::-1],
            'start': {
                'date': start.strftime("%y-%m-%d") if start else date.strftime("%Y-%m-%d"),
                'timeZone': 'Asia/Jerusalem'
            },
            'end': {
                'date': end.strftime("%y-%m-%d") if end else date.strftime("%Y-%m-%d"),
                'timeZone': 'Asia/Jerusalem'
            },
            }
        print(event)
        print('\n')    
        event = service.events().insert(calendarId='0836ibjeeoiqh1ldb1lii447ig@group.calendar.google.com', body=event).execute()
            
def year_events():
    url = "https://www.shenkar.ac.il/he/pages/academic-secretariat-studies-calander"
    html = requests.get(url, verify=False).content
    df_list = pd.read_html(html)
    rows = df_list[0].values.tolist()
    for row in rows:
        date = row[1].split(' ')[-1]
        start = None
        end = None
        if '-' in date:
            start = date.split('-')[0] + '.' + date.split('.',1)[1]
            start = datetime.strptime(start, '%d.%m.%y')
            end = date.split('-')[1]
            end = datetime.strptime(end, '%d.%m.%y')
        else:
            try:
                date = datetime.strptime(date, '%d.%m.%y')
            except:
                print('\033[93m' + "Could not add this event:\n", row , '\033[0m', '\n')
                continue
        title = row[0]
        note = row[2]
        event = {
            'summary': title [::-1],
            "description": note [::-1] if not pd.isna(note) else "",
            'start': {
                'date': start.strftime("%y-%m-%d") if start else date.strftime("%Y-%m-%d"),
                'timeZone': 'Asia/Jerusalem'
            },
            'end': {
                'date': end.strftime("%y-%m-%d") if end else date.strftime("%Y-%m-%d"),
                'timeZone': 'Asia/Jerusalem'
            },
            }
        print(event)
        print('\n') 
        event = service.events().insert(calendarId='0836ibjeeoiqh1ldb1lii447ig@group.calendar.google.com', body=event).execute()


def year_events_old():
    driver = webdriver.Firefox('.')
    driver.get("https://www.shenkar.ac.il/he/pages/academic-secretariat-studies-calander")
    table = driver.find_element_by_tag_name("table")

    rows = table.find_elements_by_tag_name("tr")

    rows = rows[1:]

    for row in rows:
            row_data = row.find_elements_by_tag_name("td")
            title = row_data[0]
            date = row_data[1]
            info = row_data[2]
            date = date.text.split(" ")[-1].strip()
            if len(date) < 3:
                continue
            if '-' in date:
                splited_date = date.split("-")
                end_time = datetime.strptime(splited_date[1],'%d.%m.%y')
                start_time = end_time.replace(day=int(splited_date[0]))
                start_time = start_time.strftime("%Y-%m-%d")
                end_time = end_time.strftime("%Y-%m-%d")
                event = {
                    'summary': title.text.encode('utf-8'),
                    'description': info.text.encode('utf-8'),
                    'start': {
                    'date': start_time,
                    'timeZone' : 'Asia/Jerusalem'
                    },
                    'end': {
                    'date': end_time,
                    'timeZone': 'Asia/Jerusalem',
                },
                }

                event = service.events().insert(calendarId='primary', body=event).execute()
    
            else:
                date =  datetime.strptime(date,'%d.%m.%y')
                date = date.strftime("%Y-%m-%d")
                event = {
                    'summary': title.text.encode('utf-8'),
                    'description': info.text.encode('utf-8'),
                    'start': {
                    'date': date,
                    'timeZone' : 'Asia/Jerusalem'
                    },
                    'end': {
                    'date': date,
                    'timeZone': 'Asia/Jerusalem',
                },
                }

                event = service.events().insert(calendarId='primary', body=event).execute()
                
    driver.close()


def main():
    global creds
    global service
    parser = argparse.ArgumentParser()
    parser.add_argument('--start', type=lambda s: datetime.strptime(s, '%d/%m/%Y'), help="start date")
    parser.add_argument('--end', type=lambda s: datetime.strptime(s, '%d/%m/%Y'), help="end date")
    parser.add_argument('--pdf', type=str, help="file path")
    parser.add_argument('--exams', default=False, action='store_true')
    parser.add_argument('--schedule', default=False, action='store_true')
    parser.add_argument('--year', default=False, action='store_true')
    parameters = parser.parse_args()

    if os.path.exists('token.pickle'):
        with open('token.pickle', 'rb') as token:
            creds = pickle.load(token)
        # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open('token.pickle', 'wb') as token:
            pickle.dump(creds, token)

    service = build('calendar', 'v3', credentials=creds)


    if parameters.schedule:
        if not parameters.pdf:
                print("Error: file path missing\n USAGE: --pdf <path to file>")
                return
        if not parameters.start or not parameters.end:
                print("Error: start/end missing\n USAGE: --end / --start <dd/mm/yyyy>")
                return

        Table = camelot.read_pdf(parameters.pdf)[1].df
        Table = Table[1:]
        semester_schedule(Table, parameters)

    elif parameters.exams:
         # moed A = 9, moed B = 7
        exams(camelot.read_pdf(parameters.pdf)[1].df,9)
        exams(camelot.read_pdf(parameters.pdf)[2].df,7)

    elif parameters.year:
        if parameters.pdf:
            Table = camelot.read_pdf(parameters.pdf)[0].df
            Table = Table[1:]
            year_pdf(Table)
        else:
            year_events()

    else:
        print("USAGE: choose mode: --exams | --schedule | --year")


if __name__ == "__main__":
    main()