#!/usr/bin/env python
"""Test AWS file handling with local files and in-memory service doubles."""

from io import BytesIO
from pathlib import Path
from tempfile import TemporaryDirectory
from types import ModuleType
from unittest.mock import Mock, call

import pytest

import aws_handler as handler


def test_download_files_and_close_responses() -> None:
    """Download all runtime objects and close their response streams."""
    streams = [BytesIO(content) for content in (b"force = False", b"outlook", b"google", b"{}")]
    s3 = Mock()
    s3.exceptions.NoSuchKey = KeyError
    s3.get_object.side_effect = [{"Body": stream} for stream in streams]
    with TemporaryDirectory() as temporary:
        directory = Path(temporary)
        handler.download_files(s3, "test-bucket", directory)
        assert (directory / "events_ts.json").read_bytes() == b"{}"
        assert (directory / "outlook_token.txt").read_bytes() == b"outlook"
    assert all(stream.closed for stream in streams)
    assert s3.get_object.call_args_list == [
        call(Bucket="test-bucket", Key=f"runtime/{filename}")
        for filename in ("config.py", "outlook_token.txt", "google_token.pickle", "events_ts.json")
    ]


def test_only_missing_state_is_optional() -> None:
    """Keep first-run behavior without treating missing credentials as empty state."""
    s3 = Mock()
    s3.exceptions.NoSuchKey = KeyError
    s3.get_object.side_effect = [
        {"Body": BytesIO(b"force = False")},
        {"Body": BytesIO(b"outlook")},
        {"Body": BytesIO(b"google")},
        KeyError("events_ts.json"),
    ]
    with TemporaryDirectory() as temporary:
        directory = Path(temporary)
        handler.download_files(s3, "test-bucket", directory)
        assert not (directory / "events_ts.json").exists()
        s3.get_object.side_effect = KeyError("config.py")
        with pytest.raises(KeyError, match=r"config\.py"):
            handler.download_files(s3, "test-bucket", directory)


def test_download_denial_and_read_failure_propagate() -> None:
    """Do not convert failed reads into a first-run sync."""
    s3 = Mock()
    s3.exceptions.NoSuchKey = KeyError
    with TemporaryDirectory() as temporary:
        s3.get_object.side_effect = PermissionError("denied")
        with pytest.raises(PermissionError, match="denied"):
            handler.download_files(s3, "test-bucket", Path(temporary))
        body = Mock()
        body.read.side_effect = OSError("read failed")
        s3.get_object.side_effect = None
        s3.get_object.return_value = {"Body": body}
        with pytest.raises(OSError, match="read failed"):
            handler.download_files(s3, "test-bucket", Path(temporary))
        body.close.assert_called_once_with()


def test_load_config_is_fresh_and_only_changes_file_locations() -> None:
    """Use the latest downloaded values without changing calendar settings."""
    with TemporaryDirectory() as temporary:
        directory = Path(temporary)
        source = directory / "config.py"
        source.write_text('google_calendar_id = "first"\nforce = False\n')
        first = handler.load_config(directory)
        # Change the length too so filesystem bytecode caching cannot hide this edit.
        source.write_text('google_calendar_id = "second-calendar"\nforce = True\n')
        second = handler.load_config(directory)
        assert first.google_calendar_id == "first"
        assert second.google_calendar_id == "second-calendar"
        assert second.force is True
        assert second.outlook_token_path == str(directory)
        assert second.outlook_token_filename == "outlook_token.txt"  # noqa: S105
        assert second.google_token_path == str(directory / "google_token.pickle")
        assert second.events_ts_json_path == str(directory / "events_ts.json")


def test_run_sync_saves_state_and_both_token_uploads() -> None:
    """Persist successful state and save tokens before and after calendar writes."""
    s3 = Mock()
    sync = Mock(spec=ModuleType)
    sync.config = Mock(force=False)
    sync.create_outlook_account = Mock()
    sync.authenticate_outlook = Mock()
    sync.authenticate_google = Mock()
    sync.sync_calendar = Mock(return_value=True)
    with TemporaryDirectory() as temporary:
        directory = Path(temporary)
        (directory / "outlook_token.txt").write_bytes(b"outlook")
        (directory / "google_token.pickle").write_bytes(b"google")
        (directory / "events_ts.json").write_bytes(b'{"event": {}}')
        assert handler.run_sync(s3, "test-bucket", directory, sync) is True
    assert [request.kwargs["Key"] for request in s3.put_object.call_args_list] == [
        "runtime/outlook_token.txt",
        "runtime/google_token.pickle",
        "runtime/events_ts.json",
        "runtime/outlook_token.txt",
        "runtime/google_token.pickle",
    ]
    sync.authenticate_outlook.assert_called_once_with(
        sync.create_outlook_account.return_value,
        interactive=False,
    )


def test_unchanged_calendar_does_not_upload_state() -> None:
    """Keep the successful state object untouched when no rebuild is needed."""
    s3 = Mock()
    sync = Mock(spec=ModuleType)
    sync.config = Mock(force=False)
    sync.create_outlook_account = Mock()
    sync.authenticate_outlook = Mock()
    sync.authenticate_google = Mock()
    sync.sync_calendar = Mock(return_value=False)
    with TemporaryDirectory() as temporary:
        directory = Path(temporary)
        (directory / "outlook_token.txt").write_bytes(b"outlook")
        (directory / "google_token.pickle").write_bytes(b"google")
        assert handler.run_sync(s3, "test-bucket", directory, sync) is False
    assert all(
        request.kwargs["Key"] != "runtime/events_ts.json"
        for request in s3.put_object.call_args_list
    )


def test_authentication_and_sync_failures_still_save_tokens() -> None:
    """Retain token changes and propagate failures without uploading event state."""
    for failed_step in ("authenticate_outlook", "authenticate_google", "sync_calendar"):
        s3 = Mock()
        sync = Mock(spec=ModuleType)
        sync.config = Mock(force=False)
        sync.create_outlook_account = Mock()
        sync.authenticate_outlook = Mock()
        sync.authenticate_google = Mock()
        sync.sync_calendar = Mock()
        getattr(sync, failed_step).side_effect = RuntimeError("failed step")
        with TemporaryDirectory() as temporary:
            directory = Path(temporary)
            (directory / "outlook_token.txt").write_bytes(b"refreshed-outlook")
            (directory / "google_token.pickle").write_bytes(b"refreshed-google")
            with pytest.raises(RuntimeError, match="failed step"):
                handler.run_sync(s3, "test-bucket", directory, sync)
        assert s3.put_object.call_args_list[-2:] == [
            call(Bucket="test-bucket", Key="runtime/outlook_token.txt", Body=b"refreshed-outlook"),
            call(Bucket="test-bucket", Key="runtime/google_token.pickle", Body=b"refreshed-google"),
        ]
        assert all(
            request.kwargs["Key"] != "runtime/events_ts.json"
            for request in s3.put_object.call_args_list
        )


def test_upload_failure_is_not_reported_as_success() -> None:
    """Surface S3 write failures to the caller for Lambda error reporting."""
    s3 = Mock()
    s3.put_object.side_effect = PermissionError("write denied")
    with TemporaryDirectory() as temporary:
        directory = Path(temporary)
        (directory / "outlook_token.txt").write_bytes(b"outlook")
        (directory / "google_token.pickle").write_bytes(b"google")
        with pytest.raises(PermissionError, match="write denied"):
            handler.upload_tokens(s3, "test-bucket", directory)
