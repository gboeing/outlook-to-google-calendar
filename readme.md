Connect to Outlook using O365: https://github.com/O365/python-o365#calendar

Connect to Google API using Python client: https://developers.google.com/docs/api/quickstart/python

Google API ref: https://developers.google.com/calendar/v3/reference/events

In credentials folder, run quickstart.py to create and pickle permanent google access token.

Outlook access token is created interactively via URL on first run, then permanently stored. It expires in 90 days if you don't run the script within that time.