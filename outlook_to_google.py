import datetime as dt
import json
import pickle
import pytz
import time

from bs4 import BeautifulSoup
from googleapiclient.discovery import build
from google.auth.transport.requests import Request
from O365 import Account
from O365 import Connection
from O365 import FileSystemTokenBackend

import config


def authenticate_outlook():
    # authenticate microsoft graph api credentials

    credentials = (config.outlook_client_id, config.outlook_client_secret)
    token_backend = FileSystemTokenBackend(
        token_path=config.outlook_token_path, token_filename=config.outlook_token_filename
    )
    account = Account(credentials, token_backend=token_backend)
    if not account.is_authenticated:
        # not authenticated, throw error
        account.authenticate(scopes=config.outlook_scopes)

    connection = Connection(credentials, token_backend=token_backend, scopes=config.outlook_scopes)
    connection.refresh_token()

    print("Authenticated Outlook.")
    return account


def authenticate_google():
    # authenticate google api credentials

    with open(config.google_token_path, "rb") as token:
        creds = pickle.load(token)
    if creds.expired:
        creds.refresh(Request())
    with open(config.google_token_path, "wb") as token:
        pickle.dump(creds, token)

    service = build("calendar", "v3", credentials=creds)
    se = service.events()

    print("Authenticated Google.")
    return se


def get_outlook_events(cal):
    # get all events from an outlook calendar
    start_time = time.time()

    start = dt.datetime.today() - dt.timedelta(days=config.previous_days)
    end = dt.datetime.today() + dt.timedelta(days=config.future_days)
    query = (
        cal.new_query("start").greater_equal(start).chain("and").on_attribute("end").less_equal(end)
    )
    events = cal.get_events(query=query, limit=None, include_recurring=True)
    events = list(events)

    elapsed_time = time.time() - start_time
    print("Retrieved {} events from Outlook in {:.1f} secs.".format(len(events), elapsed_time))
    return events


def clean_subject(subject):
    # remove prefix clutter from an outlook event subject
    remove = ["Fwd: ", "Invitation: ", "Updated invitation: ", "Updated invitation with note: "]
    for s in remove:
        subject = subject.replace(s, "")
    return subject


def clean_body(body):
    # strip out html and excess line returns from outlook event body
    text = BeautifulSoup(body, "html.parser").get_text()
    return text.replace("\n", " ").replace("\r", "\n")


def build_gcal_event(event):
    # construct a google calendar event from an outlook event

    e = {
        "summary": clean_subject(event.subject),
        "location": event.location["displayName"],
        "description": clean_body(event.body),
    }

    if event.is_all_day:
        # all day events just get a start/end date
        # use UTC start date to get correct day
        date = str(event.start.astimezone(pytz.utc).date())
        start_end = {"start": {"date": date}, "end": {"date": date}}
    else:
        # normal events have start/end datetime/timezone
        start_end = {
            "start": {
                "dateTime": str(event.start).replace(" ", "T"),
                "timeZone": str(event.start.tzinfo),
            },
            "end": {
                "dateTime": str(event.end).replace(" ", "T"),
                "timeZone": str(event.end.tzinfo),
            },
        }

    e.update(start_end)
    return e


def delete_google_events(se):
    # delete all events from google calendar
    start_time = time.time()
    gcid = config.google_calendar_id
    mr = 2500

    # retrieve a list of all events
    result = se.list(calendarId=gcid, maxResults=mr).execute()
    gcal_events = result.get("items", [])

    # if nextPageToken exists, we need to paginate: sometimes a few items are
    # spread across several pages of results for whatever reason
    i = 1
    while "nextPageToken" in result:
        npt = result["nextPageToken"]
        result = se.list(calendarId=gcid, maxResults=mr, pageToken=npt).execute()
        gcal_events.extend(result.get("items", []))
        i += 1

    print("Retrieved {} events across {} pages from Google.".format(len(gcal_events), i))

    # delete each event retrieved
    for gcal_event in gcal_events:
        request = se.delete(calendarId=config.google_calendar_id, eventId=gcal_event["id"])
        result = request.execute()
        assert result == ""
        time.sleep(config.pause)

    elapsed_time = time.time() - start_time
    print("Deleted {} events from Google in {:.1f} secs.".format(len(gcal_events), elapsed_time))


def add_google_events(se, events):
    # add all events to google calendar
    start_time = time.time()

    for event in events:
        e = build_gcal_event(event)
        result = se.insert(calendarId=config.google_calendar_id, body=e).execute()
        assert isinstance(result, dict)
        time.sleep(config.pause)

    elapsed_time = time.time() - start_time
    print("Added {} events to Google in {:.1f} secs.".format(len(events), elapsed_time))


def get_event_timestamps(outlook_events):
    # ids and timestamps of new events retrieved during current run
    ts = {}
    for e in outlook_events:
        ts[e.ical_uid] = {
            "created_ts": int(e.created.timestamp()),
            "modified_ts": int(e.modified.timestamp()),
        }
    return ts


def check_ts_match(new_events):
    # compare old event ids/timestamps to new ones retrieved during current run

    try:
        # load the old events' ids/timestamps saved to disk during previous run
        with open(config.events_ts_json_path, "r") as f:
            old_events = json.load(f)

        # make sure all ids and timestamps match between old and new
        assert new_events.keys() == old_events.keys()
        for k, new_event in new_events.items():
            old_event = old_events[k]
            assert new_event["created_ts"] == old_event["created_ts"]
            assert new_event["modified_ts"] == old_event["modified_ts"]

    except Exception:
        # if json file doesn't exist or if any id or timestamp is different
        print("Changes found.")
        return False

    return True


current_time = "{:%Y-%m-%d %H:%M:%S}".format(dt.datetime.now())
print("Started at {}.".format(current_time))
start_time = time.time()

# authenticate outlook and google credentials
outlook_acct = authenticate_outlook()
se = authenticate_google()

# get all events from outlook
outlook_cal = outlook_acct.schedule().get_default_calendar()
outlook_events = get_outlook_events(outlook_cal)
outlook_events_ts = get_event_timestamps(outlook_events)

# check if all the current event ids/timestamps match the previous run
# only update google calendar if they don't all match (means there are changes)
if config.force or not check_ts_match(outlook_events_ts):
    # delete all existing google events then add all outlook events
    delete_google_events(se)
    add_google_events(se, outlook_events)

    # save event ids/timestamps json to disk for the next run
    with open(config.events_ts_json_path, "w") as f:
        json.dump(outlook_events_ts, f)
else:
    print("No changes found.")

# all done
elapsed_time = time.time() - start_time
print("Finished in {:.1f} secs.\n".format(elapsed_time))
