from googleapiclient.discovery import build
from httplib2 import Http
from oauth2client import file, client, tools
from datetime import datetime, timedelta
from os import chdir

# If modifying these scopes, delete the file token.json.
SCOPES = 'https://www.googleapis.com/auth/spreadsheets'

# The ID and range of a sample spreadsheet.
SPREADSHEET_ID = '1AQTMkpMejOoGDy5BZu35y2NDbENkzIHWxdOV3a-AFFk'

WEEKDAY_MAP = {
    0: 'Mon',
    1: 'Tue',
    2: 'Wed',
    3: 'Thu',
    4: 'Fri',
    5: 'Sat',
    6: 'Sun',
}

LAST_X_DAYS = 7


def parse_row(row):
    """:returns row_datetime, work_mode"""
    try:
        return datetime.strptime(row[0], '%B %d, %Y at %I:%M%p'), row[1]
    except ValueError:
        return None, row[1]


def main():
    """Shows basic usage of the Sheets API.
    Prints values from a sample spreadsheet.
    """
    chdir('/home/dan/development/worktime')
    store = file.Storage('token.json')
    creds = store.get()
    if not creds or creds.invalid:
        flow = client.flow_from_clientsecrets('credentials.json', SCOPES)
        creds = tools.run_flow(flow, store)
    service = build('sheets', 'v4', http=creds.authorize(Http()))

    # Call the Sheets API
    result = service.spreadsheets().values().get(
        spreadsheetId=SPREADSHEET_ID, range='Data!A2:B'
    ).execute()
    values = result.get('values', [])
    data_dict = {}
    today = datetime.today().date()
    for num_days in range(LAST_X_DAYS-1, -1, -1):
        data_dict[today - timedelta(days=num_days)] = {'dev': 0, 'mtg': 0, 'oth': 0}

    for index, row in enumerate(values):
        row_datetime, work_mode = parse_row(row)
        if work_mode == "end":
            continue

        if not data_dict.get(row_datetime.date()):
            continue

        try:
            next_row_datetime, _ = parse_row(values[index + 1])
        except IndexError:
            if row_datetime.date() == today:
                dur_seconds = (datetime.now() - row_datetime).total_seconds()
            else:
                continue
        else:
            dur_seconds = (next_row_datetime - row_datetime).total_seconds()

        dur_hours = dur_seconds / 3600
        data_dict[row_datetime.date()][work_mode] += dur_hours

    headers = [""]
    dev = ["dev"]
    mtg = ["mtg"]
    oth = ["oth"]

    for row_date, modes in data_dict.items():
        headers += [f'{WEEKDAY_MAP[row_date.weekday()]} {row_date.month}/{row_date.day}']
        dev += [modes['dev']]
        mtg += [modes['mtg']]
        oth += [modes['oth']]

    data = [headers, dev, mtg, oth]
    result = service.spreadsheets().values().update(
        spreadsheetId=SPREADSHEET_ID, range='Computed!A1:H4',
        valueInputOption="RAW", body={'values': data}
    ).execute()


if __name__ == '__main__':
    main()
