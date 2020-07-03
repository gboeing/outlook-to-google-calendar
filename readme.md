# Microsoft Office 365 calendar -> Google calendar

## Overview

One-way sync from a Microsoft Office365 Outlook calendar to a Google calendar, handling new, updated, and deleted events. 

Connect to Microsoft API using the [O365](https://github.com/O365/python-o365#calendar) package. Connect to Google API using its Python [client](https://developers.google.com/docs/api/quickstart/python). Google API [reference]( https://developers.google.com/calendar/v3/reference/events).

## Setup

  - Create config.py to hold configuration details
  - In credentials folder, run quickstart.py to create and pickle permanent Google API access token.
  - Microsoft API access token is created interactively via URL on first run, then permanently stored. It expires in 90 days if you don't run the script within that time.
  - On server, set up cronjob to run script every 15 minutes. The script will check Microsoft for calendar events and compare them to the calendar events it saved during the previous run. If they differ (in IDs or timestamps), it will delete all Google calendar events and then add all the Microsoft calendar events to the Google calendar.
