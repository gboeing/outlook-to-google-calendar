#!/usr/bin/env python
"""Tests for event conversion helpers."""

import datetime as dt

import outlook_to_google as otg

ROME = dt.timezone(dt.timedelta(hours=2), name="Europe/Rome")


class Event:
    """Small test event with the Outlook attributes this script uses."""

    def __init__(self, start: dt.datetime, end: dt.datetime, *, is_all_day: bool) -> None:
        self.subject = "Invitation: Planning"
        self.location = {"displayName": "Conference Room"}
        self.body = "<p>Hello<br>World</p>"
        self.is_all_day = is_all_day
        self.start = start
        self.end = end
        self.ical_uid = "event-1"
        self.created = start
        self.modified = end


def test_clean_subject() -> None:
    """Remove Outlook subject prefixes."""
    assert otg.clean_subject("Fwd: Meeting") == "Meeting"
    assert otg.clean_subject("Invitation: Meeting") == "Meeting"
    assert otg.clean_subject("Updated invitation: Meeting") == "Meeting"
    assert otg.clean_subject("Meeting") == "Meeting"


def test_clean_body() -> None:
    """Strip HTML markup from event bodies."""
    assert otg.clean_body("<p>Hello</p>\n<p>World</p>") == "Hello World"


def test_timed_event_payload() -> None:
    """Build Google payload for a timed event."""
    event = Event(
        dt.datetime(2026, 5, 27, 9, 30, tzinfo=ROME),
        dt.datetime(2026, 5, 27, 10, 30, tzinfo=ROME),
        is_all_day=False,
    )

    payload = otg.build_gcal_event(event)

    assert payload["summary"] == "Planning"
    assert payload["location"] == "Conference Room"
    assert payload["description"] == "HelloWorld"
    assert payload["start"] == {
        "dateTime": "2026-05-27T09:30:00+02:00",
        "timeZone": "Europe/Rome",
    }
    assert payload["end"] == {
        "dateTime": "2026-05-27T10:30:00+02:00",
        "timeZone": "Europe/Rome",
    }


def test_one_day_all_day_event_payload() -> None:
    """Build Google payload for a one-day all-day event."""
    event = Event(
        dt.datetime(2026, 5, 27, tzinfo=ROME),
        dt.datetime(2026, 5, 28, tzinfo=ROME),
        is_all_day=True,
    )

    payload = otg.build_gcal_event(event)

    assert payload["start"] == {"date": "2026-05-27"}
    assert payload["end"] == {"date": "2026-05-28"}


def test_multi_day_all_day_event_payload() -> None:
    """Build Google payload for a multi-day all-day event."""
    event = Event(
        dt.datetime(2026, 5, 29, tzinfo=ROME),
        dt.datetime(2026, 5, 31, tzinfo=ROME),
        is_all_day=True,
    )

    payload = otg.build_gcal_event(event)

    assert payload["start"] == {"date": "2026-05-29"}
    assert payload["end"] == {"date": "2026-05-31"}


def test_same_day_all_day_event_gets_exclusive_end() -> None:
    """Give same-day all-day events an exclusive end date."""
    event = Event(
        dt.datetime(2026, 5, 27, tzinfo=ROME),
        dt.datetime(2026, 5, 27, tzinfo=ROME),
        is_all_day=True,
    )

    payload = otg.build_gcal_event(event)

    assert payload["start"] == {"date": "2026-05-27"}
    assert payload["end"] == {"date": "2026-05-28"}


def test_get_event_timestamps() -> None:
    """Build timestamp state from Outlook events."""
    event = Event(
        dt.datetime(2026, 5, 27, 9, tzinfo=ROME),
        dt.datetime(2026, 5, 27, 10, tzinfo=ROME),
        is_all_day=False,
    )
    event.created = dt.datetime(2026, 5, 27, 9, tzinfo=ROME)
    event.modified = dt.datetime(2026, 5, 27, 10, tzinfo=ROME)

    timestamps = otg.get_event_timestamps([event])

    assert timestamps == {
        "event-1": {
            "created_ts": int(event.created.timestamp()),
            "modified_ts": int(event.modified.timestamp()),
        },
    }
