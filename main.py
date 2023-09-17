import os.path
import dictdiffer as dd

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from pymongo import MongoClient, DESCENDING
from dotenv import load_dotenv

# If modifying these scopes, delete the file token.json.
SCOPES = [
    'https://www.googleapis.com/auth/calendar',
    'https://www.googleapis.com/auth/calendar.calendarlist',
    'https://www.googleapis.com/auth/calendar.events'
]

load_dotenv()

INCOMING_UNIVERSITY_NAME = os.environ.get("INCOMING_UNIVERSITY_NAME")
NEW_UNIVERSITY_NAME = os.environ.get("NEW_UNIVERSITY_NAME")
DB_PORT = int(os.environ.get("DB_PORT"))
DB_NAME = os.environ.get("DB_NAME")

EXAM_COLOR = os.environ.get("EXAM_COLOR")
LESSON_COLOR = os.environ.get("LESSON_COLOR")
LAB_COLOR = os.environ.get("LAB_COLOR")
WORK_GROUP_COLOR = os.environ.get("WORK_GROUP_COLOR")
OTHER_COLOR = os.environ.get("OTHER_COLOR")


def check_credentials():
    creds = None
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
    return creds


def get_uni_calendar_ids(service):
    # Get all the calendars information
    user_calendars = service.calendarList().list().execute().get("items", [])
    incoming_calendar_uni = None
    new_university_calendar = None
    for calendarList in user_calendars:
        # find the id's based on the summary or summaryOverride that should be equal to the calendar name provided in
        # the .env file
        if "summaryOverride" in calendarList:
            if calendarList['summaryOverride'] == INCOMING_UNIVERSITY_NAME:
                incoming_calendar_uni = calendarList["id"]
            elif calendarList['summaryOverride'] == NEW_UNIVERSITY_NAME:
                new_university_calendar = calendarList["id"]
        elif "summary" in calendarList:
            if calendarList['summary'] == INCOMING_UNIVERSITY_NAME:
                incoming_calendar_uni = calendarList["id"]
            elif calendarList['summary'] == NEW_UNIVERSITY_NAME:
                new_university_calendar = calendarList["id"]
    if incoming_calendar_uni is None:
        raise Exception(
            f"Please make sure you changed your incoming university calendar to {INCOMING_UNIVERSITY_NAME}")
    if new_university_calendar is None:
        raise Exception(
            f"Please make sure you changed your new university calendar to {NEW_UNIVERSITY_NAME}")
    return incoming_calendar_uni, new_university_calendar


def create_db_connection():
    client = MongoClient("localhost", DB_PORT)
    return client[DB_NAME]


def set_color_id(summary):
    if "Tentamen" in summary or "Hertentamen" in summary or "Deeltoets" in summary:
        return EXAM_COLOR
    elif "Hoorcollege" in summary:
        return LESSON_COLOR
    elif "Practicum" in summary:
        return LAB_COLOR
    elif "Werkgroep" in summary:
        return WORK_GROUP_COLOR
    else:
        return OTHER_COLOR


def make_new_event(event):
    new_event = dict()
    list_of_keys = ["kind", "status", "created", "updated", "summary", "description", "location", "creator",
                    "organizer", "start", "end", "eventType"]
    for i in list_of_keys:
        new_event[i] = event[i]

    new_event["colorId"] = set_color_id(event["summary"])
    return new_event


def check_update_diff_event(new_event, event, old_event):
    diff = dd.diff(old_event, event)
    new_event = dd.patch(diff, new_event)
    return new_event


def insert_unknown_event(db, event, service, cal_id):
    inserted_event = service.events().insert(calendarId=cal_id, body=make_new_event(event)).execute()
    new_event_json = dict()
    new_event_json["eventIcalUID"] = event["iCalUID"]
    new_event_json["eventInfo"] = inserted_event
    db.events.insert_one(event)
    db.new_events.insert_one(new_event_json)


def insert_or_update_event(db, event, service, cal_id):
    old_event = db.events.find_one({"iCalUID": event["iCalUID"]})
    if old_event is None:
        # Insert entry, when there isn't one in the database or calendar
        insert_unknown_event(db, event, service, cal_id)
    else:
        difference = check_update_diff_event(dict(), old_event, event)
        if event["status"] == "cancelled":
            if difference == {}:
                # Delete entry from calendar, keep in database in case re-confirmed
                corresponding_new_event = db.new_events.find_one({"eventIcalUID": event["iCalUID"]})
                try:
                    inserted_event = service.events().delete(
                        calendarId=cal_id, eventId=corresponding_new_event["eventInfo"]["id"]
                    ).execute()
                    db.events.replace_one({"iCalUID": event["iCalUID"]}, event)
                    db.new_events.update_one({"eventIcalUID": event["iCalUID"]}, {"$set": inserted_event})
                except HttpError:
                    pass
        else:
            if difference == {}:
                # Update or re-insert event, if there are changes
                new_event = db.new_events.find_one({"eventIcalUID": event["iCalUID"]})["eventInfo"]
                if new_event["status"] == "cancelled":
                    # if the event exist, but calendar entry has been deleted, re-insert based on changes and
                    # older_new_event
                    inserted_event = service.events().insert(
                        calendarId=cal_id,
                        body=check_update_diff_event(new_event, event, old_event)
                    ).execute()
                else:
                    # if the event exist and calendar entry exists
                    inserted_event = service.events().update(
                        calendarId=cal_id,
                        eventId =new_event["id"],
                        body=check_update_diff_event(new_event, event, old_event)
                    ).execute()
                db.events.replace_one({"iCalUID": event["iCalUID"]}, event)
                db.new_events.update_one({"eventIcalUID": event["iCalUID"]}, {"$set": inserted_event})


def get_latest_updated_min(db):
    sorted_events = list(db.events.find().sort("updated", DESCENDING))
    if sorted_events:
        return sorted_events[0]["updated"]
    else:
        return None


def main():
    creds = check_credentials()

    try:
        service = build('calendar', 'v3', credentials=creds)
        university_id, new_university_id = get_uni_calendar_ids(service)
        db = create_db_connection()

        min_updated = get_latest_updated_min(db)
        events_executable = None
        while True:
            if events_executable is None:
                if min_updated is None:
                    events_executable = service.events().list(
                        calendarId=university_id,
                        singleEvents=True,
                        showDeleted=True
                    )
                else:
                    events_executable = service.events().list(
                        calendarId=university_id,
                        updatedMin=get_latest_updated_min(db),
                        singleEvents=True,
                        orderBy='updated',
                        showDeleted=True
                    )
            else:
                events_executable = service.events().list_next(
                    previous_request=events_executable, previous_response=events_result
                )

            try:
                events_result = events_executable.execute()
            except:
                break

            events = events_result.get('items', [])
            if not events:
                print('No upcoming events found.')
                return

            for event in events:
                print(event)
                insert_or_update_event(db, event, service, new_university_id)

    except HttpError as error:
        print('An error occurred: %s' % error)


if __name__ == '__main__':
    main()
