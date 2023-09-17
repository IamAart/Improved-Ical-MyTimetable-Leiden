import os.path
import dictdiffer as dd

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from pymongo import MongoClient, DESCENDING

# If modifying these scopes, delete the file token.json.
SCOPES = [
    'https://www.googleapis.com/auth/calendar',
    'https://www.googleapis.com/auth/calendar.calendarlist',
    'https://www.googleapis.com/auth/calendar.events'
]

UNIVERSITY_CALENDAR_NAME = "Oud-Universiteit"
BEAUTIFUL_UNIVERSITY = "Universiteit"
DB_PORT = 27017
DB_NAME = "AdjustIcal"


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
    user_calendars = service.calendarList().list().execute().get("items", [])
    calendar_id_from_uni = None
    beautiful_uni_calendar_id = None
    for calendarList in user_calendars:
        if "summaryOverride" in calendarList:
            if calendarList['summaryOverride'] == UNIVERSITY_CALENDAR_NAME:
                calendar_id_from_uni = calendarList["id"]
            elif calendarList['summaryOverride'] == BEAUTIFUL_UNIVERSITY:
                beautiful_uni_calendar_id = calendarList["id"]
        elif "summary" in calendarList:
            if calendarList['summary'] == UNIVERSITY_CALENDAR_NAME:
                calendar_id_from_uni = calendarList["id"]
            elif calendarList['summary'] == BEAUTIFUL_UNIVERSITY:
                beautiful_uni_calendar_id = calendarList["id"]
    if calendar_id_from_uni is None:
        raise Exception(
            f"Please make sure you changed your university calendar to {UNIVERSITY_CALENDAR_NAME}")
    elif beautiful_uni_calendar_id is None:
        raise Exception(
            f"Please make sure you changed your new university calendar to {BEAUTIFUL_UNIVERSITY}")
    return calendar_id_from_uni, beautiful_uni_calendar_id


def create_db_connection():
    client = MongoClient("localhost", DB_PORT)
    return client[DB_NAME]


def set_color_id(summary):
    # colorIds with their corresponding color are:
    # 1 = light purple, 2 = light green, 3 = dark purple, 4 = light orange, 5 = yellow, 6 = bright orange,
    # 7 = light blue, 8 = dark grey, 9 = dark blue, 10 = green, 11 = red

    if "Tentamen" in summary or "Hertentamen" in summary or "Deeltoets" in summary:
        return "11"
    elif "Hoorcollege" in summary:
        return "1"
    elif "Practicum" in summary:
        return "2"
    elif "Werkgroep" in summary:
        return "4"
    else:
        return "7"


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
        insert_unknown_event(db, event, service, cal_id)
    elif event["status"] == "cancelled":
        if check_update_diff_event(dict(), old_event, event) == {}:
            corresponding_new_event = db.new_events.find_one({"eventIcalUID": event["iCalUID"]})
            try:
                inserted_event = service.events().delete(calendarId=cal_id, eventId=corresponding_new_event["eventInfo"]["id"]).execute()
                db.events.replace_one({"iCalUID": event["iCalUID"]}, event)
                db.new_events.update_one({"eventIcalUID": event["iCalUID"]}, {"$set": inserted_event})
            except HttpError:
                pass
    else:
        if check_update_diff_event(dict(), old_event, event) == {}:
            # Change or insert event here
            new_event = db.new_events.find_one({"eventIcalUID": event["iCalUID"]})["eventInfo"]
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

        while True:
            if "events_executable" not in locals():
                min_updated = get_latest_updated_min(db)
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

            if events_executable is None:
                break

            events_result = events_executable.execute()
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
