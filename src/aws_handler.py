"""Run the calendar sync with files stored in one private S3 bucket."""

import importlib
import importlib.util
import os
import sys
from pathlib import Path
from tempfile import TemporaryDirectory
from types import ModuleType
from typing import Any

import boto3


def download_files(s3: Any, bucket: str, directory: Path) -> None:  # noqa: ANN401
    """Download the runtime files, allowing only the state file to be absent.

    Parameters
    ----------
    s3 : Any
        S3 client using execution-role credentials.
    bucket : str
        Bucket containing the fixed runtime object keys.
    directory : Path
        Empty working directory for this invocation.

    Returns
    -------
    None
        Downloaded files are written into ``directory``.
    """
    for filename in ("config.py", "outlook_token.txt", "google_token.pickle", "events_ts.json"):
        try:
            response = s3.get_object(Bucket=bucket, Key=f"runtime/{filename}")
        except s3.exceptions.NoSuchKey:
            if filename != "events_ts.json":
                raise
        else:
            # Close each HTTP response even when reading or writing fails.
            body = response["Body"]
            try:
                (directory / filename).write_bytes(body.read())
            finally:
                body.close()


def load_config(directory: Path) -> ModuleType:
    """Load a fresh configuration and replace only its local file locations.

    Parameters
    ----------
    directory : Path
        Invocation directory containing the downloaded configuration and tokens.

    Returns
    -------
    ModuleType
        Fresh Python configuration with paths pointing into ``directory``.
    """
    spec = importlib.util.spec_from_file_location("config", directory / "config.py")
    if spec is None or spec.loader is None:
        msg = "Cannot load the downloaded config.py."
        raise ImportError(msg)
    settings = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(settings)
    # Module attributes are dynamic because this is the user's Python configuration.
    settings.__dict__.update(
        outlook_token_path=str(directory),
        outlook_token_filename=(directory / "outlook_token.txt").name,
        google_token_path=str(directory / "google_token.pickle"),
        events_ts_json_path=str(directory / "events_ts.json"),
    )
    return settings


def upload_tokens(s3: Any, bucket: str, directory: Path) -> None:  # noqa: ANN401
    """Save both token files without adding token-change tracking.

    Parameters
    ----------
    s3 : Any
        S3 client using execution-role credentials.
    bucket : str
        Runtime bucket.
    directory : Path
        Invocation directory containing both token files.

    Returns
    -------
    None
        Both S3 token objects are replaced with their current local contents.
    """
    for filename in ("outlook_token.txt", "google_token.pickle"):
        s3.put_object(
            Bucket=bucket,
            Key=f"runtime/{filename}",
            Body=(directory / filename).read_bytes(),
        )


def run_sync(s3: Any, bucket: str, directory: Path, sync: ModuleType) -> bool:  # noqa: ANN401
    """Authenticate and sync, retaining token changes when later work fails.

    Parameters
    ----------
    s3 : Any
        S3 client using execution-role credentials.
    bucket : str
        Runtime bucket.
    directory : Path
        Invocation directory with downloaded files and configured paths.
    sync : ModuleType
        Calendar module bound to this invocation's configuration.

    Returns
    -------
    bool
        True when the calendar was rebuilt and its state uploaded successfully.
    """
    try:
        account = sync.authenticate_outlook(sync.create_outlook_account(), interactive=False)
        service = sync.authenticate_google()
        upload_tokens(s3, bucket, directory)
        calendar = account.schedule().get_default_calendar()
        state_path = directory / "events_ts.json"
        updated = sync.sync_calendar(calendar, service, state_path, force=sync.config.force)
        if updated:
            s3.put_object(
                Bucket=bucket,
                Key="runtime/events_ts.json",
                Body=state_path.read_bytes(),
            )
        return bool(updated)
    finally:
        # A Python exception still reaches this block. Forced termination may not.
        upload_tokens(s3, bucket, directory)


def lambda_handler(event: object, context: object) -> dict[str, bool]:  # noqa: ARG001
    """Run one scheduled invocation using fresh runtime files from S3.

    Parameters
    ----------
    event : object
        Scheduler event, which is not used to select configuration or calendars.
    context : object
        Lambda invocation context.

    Returns
    -------
    dict[str, bool]
        Whether this invocation rebuilt the destination calendar.
    """
    bucket = os.environ["SYNC_BUCKET"]
    s3 = boto3.client("s3")
    # No account, S3 client, or configuration is created when importing this file.
    with TemporaryDirectory(prefix="calendar-sync-", dir="/tmp") as temporary:
        directory = Path(temporary)
        download_files(s3, bucket, directory)
        settings = load_config(directory)
        # The first import needs config.py; warm invocations also need rebinding.
        sys.modules["config"] = settings
        sync = importlib.import_module("outlook_to_google")
        sync.__dict__["config"] = settings
        return {"updated": run_sync(s3, bucket, directory, sync)}
