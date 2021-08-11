# Microsoft Office 365 calendar -> Google calendar

## Overview

One-way sync from a Microsoft Office 365 Outlook calendar to a Google calendar, handling new, updated, and deleted events.

The script connects to the Microsoft API using the [O365 package](https://github.com/O365/python-o365#calendar) and connects to the Google API using its [Python client](https://developers.google.com/docs/api/quickstart/python). See also the Google Calendar API [reference](https://developers.google.com/calendar/v3/reference/events).

## Setup

  - Create `config.py` (you can adapt [`config_sample.py`](config_sample.py)) to hold your personal configuration details, include your Microsoft `client_id` and `client_secret` and Google calendar ID. **Create a new Google calendar just for this, or else your existing events will be deleted!**
      - Get the Microsoft `client_id` and `client_secret` by following the O365 [instructions](https://github.com/O365/python-o365#authentication) on how to authenticate on behalf of a user.
  - Run `pip install --upgrade -r requirements.txt` to install the [required Python dependencies](requirements.txt).
  - In the [credentials folder](credentials), run [`python quickstart.py`](credentials/quickstart.py) to create and pickle permanent Google API access token.
  - Microsoft API access token is created interactively via URL on first run, then permanently stored. It expires in 90 days *if* you don't run the script within that time.
  - On your server, set up a cron job to run [`outlook_to_google.py`](outlook_to_google.py) (using [run.sh](run.sh)) every 15 minutes (or however often you need).
  - The script will check Microsoft for calendar events and compare them to the calendar events it saved (in events_ts.json) during the previous run. **If they differ (in IDs or timestamps), it will delete all Google calendar events and then add all Microsoft calendar events to the Google calendar.**
