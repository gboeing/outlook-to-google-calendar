#!/usr/bin/env python
"""Tests for event conversion helpers."""

import datetime as dt
from pathlib import Path
from tempfile import TemporaryDirectory
from unittest.mock import Mock

import pytest
from O365 import Account

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


def test_create_outlook_account_does_not_sign_in() -> None:
    """Construct the account without requesting browser sign-in or calendar data."""
    assert isinstance(otg.create_outlook_account(), Account)


def test_unattended_outlook_authentication_requires_saved_login() -> None:
    """Reject interactive login in Lambda before attempting token refresh."""
    account = Mock(is_authenticated=False)
    with pytest.raises(RuntimeError, match="Sign in locally"):
        otg.authenticate_outlook(account, interactive=False)
    account.authenticate.assert_not_called()
    account.con.refresh_token.assert_not_called()


def test_outlook_authentication_uses_the_account_connection() -> None:
    """Refresh the same connection that later retrieves calendar events."""
    account = Mock(is_authenticated=True)
    assert otg.authenticate_outlook(account, interactive=False) is account
    account.authenticate.assert_not_called()
    account.con.refresh_token.assert_called_once_with()


def test_local_outlook_authentication_can_request_login() -> None:
    """Keep local browser sign-in and report incomplete login."""
    account = Mock(is_authenticated=False)
    account.authenticate.return_value = True
    assert otg.authenticate_outlook(account) is account
    account.authenticate.assert_called_once_with(scopes=otg.config.outlook_scopes)
    account.con.refresh_token.assert_called_once_with()
    account.authenticate.return_value = False
    with pytest.raises(RuntimeError, match="not completed"):
        otg.authenticate_outlook(account)


def test_state_comparison_handles_missing_invalid_and_changed_state() -> None:
    """Compare state without using the user's configured state file."""
    current = {"event": {"created_ts": 1, "modified_ts": 2}}
    with TemporaryDirectory() as temporary:
        path = Path(temporary) / "state.json"
        assert otg.check_ts_match(current, path) is False
        for content in ("invalid", "{}", '{"event": {}}'):
            path.write_text(content)
            assert otg.check_ts_match(current, path) is False
        path.write_text('{"event": {"created_ts": 1, "modified_ts": 2}}')
        assert otg.check_ts_match(current, path) is True
        path.write_text('{"event": {"created_ts": 0, "modified_ts": 2}}')
        assert otg.check_ts_match(current, path) is False
        path.write_text('{"event": {"created_ts": 1, "modified_ts": 0}}')
        assert otg.check_ts_match(current, path) is False


def test_sync_returns_rebuild_status_and_reuses_saved_state() -> None:
    """Save successful state and skip writes when the source is unchanged."""
    calendar = Mock()
    calendar.protocol.timezone = dt.UTC
    calendar.get_events.return_value = []
    service = Mock()
    service.list.return_value.execute.return_value = {"items": []}
    with TemporaryDirectory() as temporary:
        path = Path(temporary) / "state.json"
        assert otg.sync_calendar(calendar, service, path, force=False) is True
        assert path.read_text() == "{}"
        assert otg.sync_calendar(calendar, service, path, force=False) is False
        assert otg.sync_calendar(calendar, service, path, force=True) is True
    expected_rebuilds = 2
    assert service.list.call_count == expected_rebuilds


def test_failed_rebuild_does_not_replace_state() -> None:
    """Keep the previous state when Google Calendar requests fail."""
    calendar = Mock()
    calendar.protocol.timezone = dt.UTC
    calendar.get_events.return_value = []
    service = Mock()
    service.list.return_value.execute.side_effect = RuntimeError("calendar unavailable")
    with TemporaryDirectory() as temporary:
        path = Path(temporary) / "state.json"
        path.write_text('{"previous": {}}')
        with pytest.raises(RuntimeError, match="calendar unavailable"):
            otg.sync_calendar(calendar, service, path, force=False)
        assert path.read_text() == '{"previous": {}}'
