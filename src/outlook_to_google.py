"""Sync Outlook calendar events to a dedicated Google calendar."""

import datetime as dt
import json
import pickle
import time
from pathlib import Path
from typing import Any

from bs4 import BeautifulSoup
from google.auth.transport.requests import Request
from googleapiclient.discovery import build
from O365 import Account, Connection, FileSystemTokenBackend

import config


def authenticate_outlook() -> Any:
    """Authenticate Microsoft Graph API credentials."""
    credentials = (config.outlook_client_id, config.outlook_client_secret)
    token_backend = FileSystemTokenBackend(
        token_path=config.outlook_token_path,
        token_filename=config.outlook_token_filename,
    )
    account = Account(credentials, token_backend=token_backend)
    if not account.is_authenticated:
        # not authenticated, throw error
        account.authenticate(scopes=config.outlook_scopes)

    connection = Connection(
        credentials,
        token_backend=token_backend,
        scopes=config.outlook_scopes,
    )
    connection.refresh_token()

    print(f"{timestamp()} Authenticated Outlook.")
    return account


def authenticate_google() -> Any:
    """Authenticate Google API credentials."""
    google_token_path = Path(config.google_token_path)
    with google_token_path.open("rb") as token:
        creds = pickle.load(token)  # noqa: S301
    if creds.expired:
        creds.refresh(Request())
    with google_token_path.open("wb") as token:
        pickle.dump(creds, token)

    service = build("calendar", "v3", credentials=creds)
    se = service.events()

    print(f"{timestamp()} Authenticated Google.")
    return se


def get_outlook_events(cal: Any) -> list[Any]:
    """Get all events from an Outlook calendar."""
    now = dt.datetime.now(tz=cal.protocol.timezone)
    start = now - dt.timedelta(days=config.previous_days)
    end = now + dt.timedelta(days=config.future_days)
    events = cal.get_events(
        limit=None,
        include_recurring=True,
        start_recurring=start.isoformat(),
        end_recurring=end.isoformat(),
    )
    events = [event for event in events if event.start >= start and event.end <= end]

    print(f"{timestamp()} Retrieved {len(events)} events from Outlook.")
    return events


def clean_subject(subject: str) -> str:
    """Remove prefix clutter from an Outlook event subject."""
    remove = [
        "Fwd: ",
        "Invitation: ",
        "Updated invitation: ",
        "Updated invitation with note: ",
    ]
    for s in remove:
        subject = subject.replace(s, "")
    return subject


def clean_body(body: str) -> str:
    """Strip out HTML and excess line returns from an Outlook event body."""
    text = BeautifulSoup(body, "html.parser").get_text()
    return text.replace("\n", " ").replace("\r", "\n")


def build_gcal_event(event: Any) -> dict[str, Any]:
    """Construct a Google Calendar event from an Outlook event."""
    e = {
        "summary": clean_subject(event.subject),
        "location": event.location["displayName"],
        "description": clean_body(event.body),
    }

    if event.is_all_day:
        # all day events just get a start/end date
        start_date = event.start.date()
        end_date = event.end.date()
        if end_date <= start_date:
            end_date = start_date + dt.timedelta(days=1)
        start_end = {"start": {"date": str(start_date)}, "end": {"date": str(end_date)}}
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


def delete_google_events(se: Any) -> None:
    """Delete all events from the Google calendar."""
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

    print(
        f"{timestamp()} Retrieved {len(gcal_events)} events across {i} pages from Google.",
    )

    # delete each event retrieved
    for gcal_event in gcal_events:
        request = se.delete(
            calendarId=config.google_calendar_id,
            eventId=gcal_event["id"],
        )
        result = request.execute()
        if result != "":
            msg = f"Unexpected Google delete response: {result!r}"
            raise RuntimeError(msg)
        time.sleep(config.pause)
    print(f"{timestamp()} Deleted {len(gcal_events)} events from Google.")


def add_google_events(se: Any, google_events: list[dict[str, Any]]) -> None:
    """Add all events to the Google calendar."""
    for google_event in google_events:
        result = se.insert(calendarId=config.google_calendar_id, body=google_event).execute()
        if not isinstance(result, dict):
            msg = f"Unexpected Google insert response: {result!r}"
            raise TypeError(msg)
        time.sleep(config.pause)

    print(f"{timestamp()} Added {len(google_events)} events to Google.")


def get_event_timestamps(outlook_events: list[Any]) -> dict[str, dict[str, int]]:
    """Get IDs and timestamps of events retrieved during the current run."""
    ts = {}
    for e in outlook_events:
        ts[e.ical_uid] = {
            "created_ts": int(e.created.timestamp()),
            "modified_ts": int(e.modified.timestamp()),
        }
    return ts


def check_ts_match(new_events: dict[str, dict[str, int]]) -> bool:
    """Compare old event IDs and timestamps to newly retrieved events."""
    # load the old events' ids/timestamps saved to disk during previous run
    try:
        with Path(config.events_ts_json_path).open() as f:
            old_events = json.load(f)

        # make sure all ids and timestamps match between old and new
        if new_events.keys() != old_events.keys():
            print(f"{timestamp()} Changes found.")
            return False
        for k, new_event in new_events.items():
            old_event = old_events[k]
            if new_event["created_ts"] != old_event["created_ts"]:
                print(f"{timestamp()} Changes found.")
                return False
            if new_event["modified_ts"] != old_event["modified_ts"]:
                print(f"{timestamp()} Changes found.")
                return False

    except (FileNotFoundError, KeyError, TypeError, json.JSONDecodeError):
        # if json file doesn't exist or if any id or timestamp is different
        print(f"{timestamp()} Changes found.")
        return False

    return True


def timestamp() -> str:
    """Return the current timestamp for log messages."""
    return f"{dt.datetime.now(tz=dt.UTC).astimezone():%Y-%m-%d %H:%M:%S}"


if __name__ == "__main__":
    print(f"{timestamp()} Started process.")

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
        google_events = [build_gcal_event(event) for event in outlook_events]
        delete_google_events(se)
        add_google_events(se, google_events)

        # save event ids/timestamps json to disk for the next run
        with Path(config.events_ts_json_path).open("w") as f:
            json.dump(outlook_events_ts, f)
    else:
        print(f"{timestamp()} No changes found.")

    # all done
    print(f"{timestamp()} Finished process.")
