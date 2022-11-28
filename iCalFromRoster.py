import csv
import sys
from datetime import datetime, timedelta
import uuid
import argparse
from os.path import exists
from collections import deque

calPreamble = [
    'BEGIN:VCALENDAR',
    'VERSION:2.0',
    'PRODID:-//JwF//Roster to Calendar//EN',
    'X-APPLE-CALENDAR-COLOR:#44A703',
    'X-WR-TIMEZONE:Europe/London',
    'CALSCALE:GREGORIAN',
]

calPostamble = [
    'END:VCALENDAR'
]


def main():
    print("Roster to Calendar Script...")
    # OPEN AND PARSE ROSTER SHEET
    parser = argparse.ArgumentParser()
    parser.add_argument('-i', '--input', help='Roster file to parse')
    parser.add_argument('-o', '--output', help='Calendar output file', default=None)
    parser.add_argument('-l', '--line', help='Starting line in the link', default=None)
    parser.add_argument('-v', '--verbose', help='Displays extra processing information', default=False, action='store_true')
    args = parser.parse_args()

    # WAS A FILE SUPPLIED
    if args.input is None:
        print("Please Supply Roster CSV File with the -i option")
        sys.exit(2)

    # DOES THE FILE EXIST
    if exists(args.input):
        print("Found Roster...")
        rosterCSV = args.input
    else:
        rosterCSV = input("File NOT found... please supply another: ")
        if not exists(rosterCSV):
            print("File NOT found... Re-run the script.")
            sys.exit(2)
        else:
            print("Found Roster...")

    # FIND OUT WHEN THE ROSTER STARTS
    startDate = input("When does this roster start (YYYYMMDD) ? : ")
    # Validate Roster Start
    rosterStart = datetime(int(startDate[0:4]), int(startDate[4:6]), int(startDate[6:8]))
    if not rosterStart.weekday() == 6:
        print("Roster MUST start on a Sunday... Re-start the script")
        sys.exit(2)
    if args.verbose:
        print(rosterStart.strftime('Roster Starts %a %d %b'))

    # FIND OUT STARTING POSITION IN THE LINK
    if args.line is None:
        startingLine = input("What line do you START in this link? : ")
    else:
        startingLine = args.line
    # Validate Starting Line
    try:
        startingLine = int(startingLine)
    except ValueError as verr:
        print(f'Invalid Starting Line, `{startingLine}`, Aborting!')
        sys.exit(2)
    if not startingLine >= 1:
        print(f'`{startingLine}` is not a valid option for Starting Line')
        sys.exit(2)

    # print(f'`{startingLine}` == Starting Line')

    # Setting required variables
    requiredFieldnames = ['Link Name', 'Line Number', 'Id', 'On', 'Off', 'Duration', 'Day', 'Rest Day', 'Spare Turn']
    shiftsDeque = deque()
    desiredLink = None

    # OPEN ROSTER AND CHECK FOR REQUIRED FIELDS / PARSE INTO A DEQUE OBJECT
    with open(rosterCSV, newline='') as rosterFile:
        reader = csv.DictReader(rosterFile)
        if not all(item in reader.fieldnames for item in requiredFieldnames):
            print("The roster file does not have the correct headings.\nExpecting;\n\t-", end='')
            print('\n\t-'.join(requiredFieldnames))
            print("\nCorrect the roster file or alter this script and re-run.")
            sys.exit(2)

        print("Reading Roster File...")
        for row in reader:
            eventDict = {
                'link':     row['Link Name'],
                'line':     row['Line Number'],
                'job':      row['Id'],
                'start':    row['On'].replace(':', ''),
                'finish':   row['Off'].replace(':', ''),
                'duration': row['Duration'].replace(':', ''),
                'dayName':  row['Day'],
                'rest':     bool(int(row['Rest Day'])),
                'spare':    bool(int(row['Spare Turn']))
            }

            # ONLY CONSIDER THE FIRST LINK ENCOUNTERED
            if not desiredLink:
                desiredLink = eventDict['link']
            if eventDict['link'] == desiredLink:
                shiftsDeque.append(eventDict)
            elif args.verbose:
                print(f"Ignoring entry - {eventDict['link']}")

    if args.verbose:
        print("Finished reading roster...")

    # DETERMINE HOW WHEN TO START / STOP OUTPUTTING CALENDAR ENTRIES
    outputStartRequest = input("What date do you want calendar entries START (YYYYMMDD) ? : ")
    outputStart = datetime(int(outputStartRequest[0:4]), int(outputStartRequest[4:6]), int(outputStartRequest[6:8]))
    if outputStart < rosterStart:
        print("You can't ask for calendar entries before the roster starts. Try again.")
        sys.exit(2)

    outputEndRequest = input("What date do you want calendar entries to END (YYYYMMDD) ? : ")
    outputEnd = datetime(int(outputEndRequest[0:4]), int(outputEndRequest[4:6]), int(outputEndRequest[6:8]), 23, 59, 59)
    if outputEnd < outputStart:
        print("You can't ask for calendar entries to end before they have started. Try again.")
        sys.exit(2)

    # ROTATE DEQUE TO ALIGN WITH START REQUEST
    deltaDays = (outputStart - rosterStart).days
    print(f'Skipping {deltaDays} days to get to {outputStart.strftime("%a %d %b")} ...')
    deltaDays += ((startingLine - 1) * 7)
    print(f'Skipping {startingLine - 1} weeks to get to correct position in the link ...')
    shiftsDeque.rotate(-deltaDays)

    currentDay = outputStart        # Counter
    calendarList = calPreamble      # Container for calendar events as they are processed

    while currentDay < outputEnd:
        currentShift = shiftsDeque.popleft()
        # CHECK FOR DAY OF WEEK ALIGNMENT ERROR
        if not currentDay.strftime('%A') == currentShift['dayName']:
            print(f'{currentDay.strftime("%Y-%m-%d")}; DAY MISMATCH ERROR -- Expecting {currentDay.strftime("%A")} GOT {currentShift["dayName"]} -- Aborting!')
            sys.exit(2)
        if args.verbose:
            print(f'Processing {currentDay.strftime("%a %d %b")}')
            print(currentShift)
        calendarList += makeEvent(title=currentShift['job'], date=currentDay.strftime('%Y%m%d'),
                                  allDay=currentShift['rest'], start=currentShift['start'],
                                  duration=currentShift['duration'], line=currentShift['line'],
                                  spare=currentShift['spare'], verbose=args.verbose)
        shiftsDeque.append(currentShift)
        currentDay = currentDay + timedelta(days=1)

    # calendarList += makeEvent('MP405', '20221117', allDay=False, start='2158', duration='0658')
    # calendarList += makeEvent('MP2403', '20221120', allDay=False, start='0905', duration='0818')
    # calendarList += makeEvent('RD', '20221121', allDay=True, start='2158', duration='0658')
    calendarList += calPostamble
    calendarOutputText = '\n'.join(text for text in calendarList)
    outputFile = None

    if args.output is not None:
        # if exists(args.output):
        #     outputFile = args.output
        # else:
        #     outputFile = None
        outputFile = args.output

    if outputFile is None:
        outputFile = f'Roster_{outputStart.strftime("%b%d")}-{outputEnd.strftime("%b%d")}_Line{startingLine}.ics'

    with open(outputFile, 'w+') as f:
        f.seek(0)
        f.write(calendarOutputText)
        f.truncate()


def makeEvent(title, date, line=1, allDay=False, start=False, duration=False, spare=False, verbose=False):
    nowString = datetime.now().strftime('%Y%m%dT%H%M%S')
    # print(int(date[0:4]), int(date[4:6]), int(date[6:8]))
    if allDay:
        startTime = datetime(int(date[0:4]), int(date[4:6]), int(date[6:8]))
        endTime = startTime
    else:
        startTime = datetime(int(date[0:4]), int(date[4:6]), int(date[6:8]), int(start[0:2]), int(start[2:4]))
        endTime = startTime + timedelta(hours=int(duration[0:2]), minutes=int(duration[2:4]))
    if verbose:
        print(f'Shift Starts {startTime.strftime("%a %d %b - %H.%M")}')
        print(f'Shift Ends {endTime.strftime("%a %d %b - %H.%M")}')
    if allDay:
        dstart = f';VALUE=DATE:{startTime.strftime("%Y%m%d")}'
        dend = f';VALUE=DATE:{endTime.strftime("%Y%m%d")}'
        desc = f'Line {line}'
        title = 'Rest Day'
        locationString = ''
    else:
        dstart = f':{startTime.strftime("%Y%m%dT%H%M00")}'
        dend = f':{endTime.strftime("%Y%m%dT%H%M00")}'
        desc = f'Shift Length, {duration[0:2]}:{duration[2:4]}\\nLine {line}'
        title = title if not spare else 'Spare'
        locationString = '5 Boad St\\nManchester\\n M1 2DW\\n, England\\n M1 2DW'

    eventBuffer = [
        'BEGIN:VEVENT',
        'TRANSP: OPAQUE',
        f'DTSTAMP:{nowString}',
        f'DTSTART{dstart}',
        f'DTEND{dend}',
        f'UID:{uuid.uuid4()}',
        f'LOCATION:{locationString}',
        'SEQUENCE:1',
        'X-APPLE-TRAVEL-ADVISORY-BEHAVIOR:AUTOMATIC',
        f'SUMMARY:{title}',
        f'DESCRIPTION:{desc}',
        f'LAST-MODIFIED:{nowString}',
        f'CREATED:{nowString}',
        'BEGIN:VALARM',
        f'UID:{uuid.uuid4()}',
        'TRIGGER;VALUE=DATE-TIME:19760401T005545Z',
        'ACTION:NONE',
        'END:VALARM',
        'END:VEVENT'
    ]

    # eventText = '\n'.join(text for text in eventBuffer)
    return eventBuffer


if __name__ == "__main__":
    main()