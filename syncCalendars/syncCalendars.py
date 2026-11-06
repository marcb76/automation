import argparse
import datetime
import os.path
import sys
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
import win32com.client

SCOPES = ['https://www.googleapis.com/auth/calendar']

# Default synchronization limits (in days)
BACK_LOG_SPAN_IN_DAYS = 30
FRONT_LOG_SPAN_IN_DAYS = 60


def print_header():
  """Prints the script header with metadata, description, and objective."""
  print('==================================================')
  print(' Script Name : syncCalendars.py')
  print(' Version     : 1.1.6')
  print(' Created     : November 2026')
  print(' Author      : Marc Bonet')
  print(' Description : Automated synchronization utility between Windows')
  print('               Outlook Desktop Calendar and Google Calendar.')
  print(' Objective   : Keep local appointments mirrored in Google Calendar')
  print('               safely, maintaining idempotency via unique IDs.')
  print('==================================================\n')


def get_google_service():
  creds = None
  if os.path.exists('token.json'):
    creds = Credentials.from_authorized_user_file('token.json', SCOPES)
  if not creds or not creds.valid:
    if creds and creds.expired and creds.refresh_token:
      creds.refresh(Request())
    else:
      flow = InstalledAppFlow.from_client_secrets_file(
          'credentials.json', SCOPES
      )
      creds = flow.run_local_server(port=0)
    with open('token.json', 'w') as token:
      token.write(creds.to_json())
  return build('calendar', 'v3', credentials=creds)


def format_datetime_with_local_offset(dt):
  """Safely converts an Outlook pywintypes/datetime object to an ISO 8601 string 
  preserving local wall-clock time and attaching the system time zone offset.
  """
  if not dt:
    return None
  
  # Extract components explicitly to handle pywintypes compatibility
  naive_dt = datetime.datetime(
      dt.year, dt.month, dt.day, dt.hour, dt.minute, dt.second
  )
  # Attach local system time zone offset
  local_dt = naive_dt.astimezone()
  return local_dt.isoformat()


def sync_calendar(
    dry_run=True,
    backlog_days=BACK_LOG_SPAN_IN_DAYS,
    frontlog_days=FRONT_LOG_SPAN_IN_DAYS,
):
  mode_str = 'SIMULATION (Dry Run)' if dry_run else 'REAL UPDATE'

  # Use timezone-aware local datetime for window calculations
  now = datetime.datetime.now().astimezone()
  start_date_limit = now - datetime.timedelta(days=backlog_days)
  end_date_limit = now + datetime.timedelta(days=frontlog_days)

  print('--- STARTING SYNCHRONIZATION ---')
  print(f'Execution Mode: {mode_str}')
  print(f'Backlog Span (Days): {backlog_days}')
  print(f'Frontlog Span (Days): {frontlog_days}')
  print(
      'Initial Synchronization Date:'
      f' {start_date_limit.strftime("%Y-%m-%d %H:%M:%S %z")}'
  )
  print(
      'Final Synchronization Date:'
      f' {end_date_limit.strftime("%Y-%m-%d %H:%M:%S %z")}'
  )
  print('----------------------------------')

  service = get_google_service()

  # MAPI Connection to Windows Outlook
  outlook = win32com.client.Dispatch('Outlook.Application')
  namespace = outlook.GetNamespace('MAPI')
  calendar = namespace.GetDefaultFolder(9)  # 9 = Calendar folder
  items = calendar.Items
  items.IncludeRecurrences = True

  # Apply MAPI date restriction bounded by backlog and frontlog spans
  restriction = (
      f"[Start] >= '{start_date_limit.strftime('%m/%d/%Y %H:%M')}' "
      f"AND [Start] <= '{end_date_limit.strftime('%m/%d/%Y %H:%M')}'"
  )
  restricted_items = items.Restrict(restriction)
  restricted_items.Sort('[Start]')

  # 1. Fetch existing events from Google Calendar within the defined window
  g_events_result = (
      service.events()
      .list(
          calendarId='primary',
          timeMin=start_date_limit.isoformat(),
          timeMax=end_date_limit.isoformat(),
          singleEvents=True,
      )
      .execute()
  )
  g_events = g_events_result.get('items', [])

  g_map = {}
  for ge in g_events:
    desc = ge.get('description', '')
    if '[OutlookID:' in desc:
      start_i = desc.find('[OutlookID:') + len('[OutlookID:')
      end_i = desc.find(']', start_i)
      out_id = desc[start_i:end_i]
      g_map[out_id] = ge

  print(f'Events found in Google with linked Outlook ID: {len(g_map)}')

  # 2. Process Outlook appointments within the restricted time frame
  created_count = 0
  updated_count = 0
  skipped_count = 0

  for item in restricted_items:
    if item.Class != 26:  # 26 = olAppointment
      continue

    item_start = item.Start
    naive_item_start = datetime.datetime(
        item_start.year, item_start.month, item_start.day,
        item_start.hour, item_start.minute, item_start.second
    )

    start_limit_naive = start_date_limit.replace(tzinfo=None)
    end_limit_naive = end_date_limit.replace(tzinfo=None)

    # Double-check Python-side boundary enforcement
    if naive_item_start < start_limit_naive or naive_item_start > end_limit_naive:
      skipped_count += 1
      continue

    out_id = item.EntryID
    subject = item.Subject or 'Untitled'
    
    # Format datetimes preserving exact local wall-clock time with local offset
    start_iso = format_datetime_with_local_offset(item.Start)
    end_iso = format_datetime_with_local_offset(item.End)
    
    body = f"{item.Body or ''}\n\n[OutlookID:{out_id}]"

    body_payload = {
        'summary': subject,
        'description': body,
        'start': {'dateTime': start_iso},
        'end': {'dateTime': end_iso},
    }

    if out_id in g_map:
      updated_count += 1
      action_label = 'Would update' if dry_run else 'Updating'
      print(
          f'[{ "SIMULATION" if dry_run else "REAL" }] {action_label} in Google'
          f' -> Subject: "{subject}" (Start: {start_iso})'
      )
      if not dry_run:
        service.events().update(
            calendarId='primary',
            eventId=g_map[out_id]['id'],
            body=body_payload,
        ).execute()
    else:
      created_count += 1
      action_label = 'Would create' if dry_run else 'Creating'
      print(
          f'[{ "SIMULATION" if dry_run else "REAL" }] {action_label} in Google'
          f' -> Subject: "{subject}" (Start: {start_iso})'
      )
      if not dry_run:
        service.events().insert(
            calendarId='primary', body=body_payload
        ).execute()

  print('--- EXECUTION SUMMARY ---')
  print(f'Total events processed for creation: {created_count}')
  print(f'Total events processed for update: {updated_count}')
  print(f'Total events skipped outside window: {skipped_count}')
  print('-----------------------------------------')


def valid_span_days(value):
  """Custom validator to ensure day span is an integer greater than 15."""
  try:
    ivalue = int(value)
  except ValueError:
    raise argparse.ArgumentTypeError(f"'{value}' is not a valid integer.")

  if ivalue <= 15:
    raise argparse.ArgumentTypeError(
        f'Value ({ivalue}) must be an integer greater than 15.'
    )
  return ivalue


def main():
  """Main execution flow."""
  print_header()

  parser = argparse.ArgumentParser(
      description=(
          'Synchronizes Outlook desktop calendar with Google Calendar'
          ' bidirectionally.'
      ),
      usage=(
          'python syncCalendars.py [--help] [--dryRun | --now]'
          ' [--backLogSpanInDays N] [--frontLogSpanInDays N]'
      ),
      add_help=True,
  )

  group = parser.add_mutually_exclusive_group(required=True)
  group.add_argument(
      '--dryRun',
      action='store_true',
      help=(
          'Simulation mode: processes events and prints planned changes'
          ' to console without modifying Google Calendar.'
      ),
  )
  group.add_argument(
      '--now',
      action='store_true',
      help=(
          'Production mode: executes the real synchronization by creating'
          ' or updating events directly in Google Calendar.'
      ),
  )

  parser.add_argument(
      '--backLogSpanInDays',
      type=valid_span_days,
      default=BACK_LOG_SPAN_IN_DAYS,
      help=(
          'Look-back window in days (integer > 15). Defines how many days'
          ' in the past from today to search for Outlook appointments.'
          f' Default: {BACK_LOG_SPAN_IN_DAYS}.'
      ),
  )
  parser.add_argument(
      '--frontLogSpanInDays',
      type=valid_span_days,
      default=FRONT_LOG_SPAN_IN_DAYS,
      help=(
          'Look-ahead window in days (integer > 15). Defines how many days'
          ' into the future from today to include for synchronization.'
          f' Default: {FRONT_LOG_SPAN_IN_DAYS}.'
      ),
  )

  args = parser.parse_args()

  if args.dryRun:
    sync_calendar(
        dry_run=True,
        backlog_days=args.backLogSpanInDays,
        frontlog_days=args.frontLogSpanInDays,
    )
  elif args.now:
    sync_calendar(
        dry_run=False,
        backlog_days=args.backLogSpanInDays,
        frontlog_days=args.frontLogSpanInDays,
    )


if __name__ == '__main__':
  main()