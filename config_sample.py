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

# calendars
calendars = [
	["You Outlook Calendar Name Here", "your-calendar-id-here@group.calendar.google.com"]
]


# misc
events_ts_json_path = "./events_ts_{0}.json"
pause = 0.1
force = False  # force full run even if no changes
