# outlook
outlook_client_id = "your-client-id-here"
outlook_client_secret = "your-client-secret-here"
outlook_scopes = ["basic", "calendar"]
outlook_token_path = "./credentials/"
outlook_token_filename = "outlook_token.txt"
previous_days = 40  # retrieve this many past days of events
future_days = 365  # retrieve this many future days of events

# google
google_token_path = "./credentials/google_token.pickle"
google_calendar_id = "your-calendar-id-here@group.calendar.google.com"

# misc
events_ts_json_path = "./events_ts.json"
pause = 0.1
force = False  # force full run even if no changes
only_title = False  # sync only title of event
no_all_day_events = False  # do not sync all-day events
incremental_sync = False  # use incremental sync - add only missing appointments, remove only additional
