# Outlook to Google calendar

One-way sync from Outlook to a dedicated Google calendar, including new, changed,
and deleted events. If event IDs or timestamps change, the tool deletes all events
in the destination calendar and copies the Outlook events again. Use this calendar
only for the sync.

## AWS setup

EventBridge Scheduler runs a Python 3.13 Lambda function at :10, :25, :40, and :55
each hour. The deployment is in Oregon (`us-west-2`), under stack `calendar-sync`.
One private S3 bucket, `outlook-to-google-calendar-[ID]-us-west-2`, stores
the code package and the runtime configuration, calendar tokens, and sync state.

Each run downloads the runtime files, refreshes credentials, checks Outlook, and
rebuilds the Google calendar only when needed. It saves refreshed tokens and
successful sync state to S3. Logs are kept for 7 days. Failed executions retry once.

## Updates and maintenance

1. To update code, sign in and deploy from the repository root. Deployment
   preserves the runtime files in S3.

   ```bash
   aws login --profile calendar-sync
   bash aws/deploy.sh calendar-sync outlook-to-google-calendar-[ID]-us-west-2 calendar-sync ENABLED
   ```

2. To update settings, upload the revised `src/config.py` to `runtime/config.py`
   in the bucket. Lambda uses the new settings on its next run.
3. If sync stops, check the function's CloudWatch logs. If calendar sign-in is
   required, disable the schedule and let pending runs finish, sign in locally,
   upload the replacement tokens to `runtime/`, then enable the schedule.
4. Before restarting the server job, disable the AWS schedule, let pending runs
   finish, and download the current tokens and state from S3. Run only one scheduler.

A failed rebuild can leave the destination incomplete until a later successful run.
