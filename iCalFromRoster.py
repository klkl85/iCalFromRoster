import csv
import sys
from datetime import datetime, timedelta
import uuid
import argparse
from os.path import exists
from collections import deque
import random


class col:
    red = '\033[0m\033[91m'
    Red = '\033[0m\033[1m\033[91m'
    RED = '\033[0m\033[1m\033[30m\033[41m'
    Ured = '\033[0m\033[1m\033[4m\033[91m'
    grn = '\033[0m\033[92m'
    Grn = '\033[0m\033[1m\033[92m'
    GRN = '\033[0m\033[1m\033[30m\033[42m'
    Ugrn = '\033[0m\033[1m\033[4m\033[92m'
    blu = '\033[0m\033[94m'
    Blu = '\033[0m\033[1m\033[94m'
    BLU = '\033[0m\033[1m\033[30m\033[44m'
    Ublu = '\033[0m\033[1m\033[4m\033[94m'
    cyn = '\033[0m\033[96m'
    Cyn = '\033[0m\033[1m\033[96m'
    CYN = '\033[0m\033[1m\033[30m\033[46m'
    Ucyn = '\033[0m\033[1m\033[4m\033[96m'
    yel = '\033[0m\033[93m'
    Yel = '\033[0m\033[1m\033[93m'
    YEL = '\033[0m\033[1m\033[30m\033[43m'
    Uyel = '\033[0m\033[1m\033[4m\033[93m'
    mag = '\033[0m\033[95m'
    Mag = '\033[0m\033[1m\033[95m'
    MAG = '\033[0m\033[1m\033[30m\033[45m'
    Umag = '\033[0m\033[1m\033[4m\033[95m'
    k = '\033[0m\033[47m\033[30m'
    K = '\033[0m\033[1m\033[47m\033[30m'
    Uk = '\033[0m\033[1m\033[4m\033[47m\033[30m'
    UK = '\033[0m\033[1m\033[4m\033[7m\033[47m\033[30m'
    end = '\033[0m'
    bold = '\033[1m'
    italic = '\033[3m'
    underline = '\033[4m'
    blink = '\033[5m'
    selected = '\033[7m'


def randomColour():
    r = lambda: random.randint(0, 255)
    return '#%02X%02X%02X' % (r(), r(), r())


calPreamble = [
    'BEGIN:VCALENDAR',
    'VERSION:2.0',
    'PRODID:-//JwF//Roster to Calendar//EN',
    f'X-APPLE-CALENDAR-COLOR:{randomColour()}',
    'X-WR-TIMEZONE:Europe/London',
    'CALSCALE:GREGORIAN',
]

calPostamble = [
    'END:VCALENDAR'
]


def main():
    print(f"{col.Yel}Roster to Calendar Script...")
    # OPEN AND PARSE ROSTER SHEET
    parser = argparse.ArgumentParser()
    parser.add_argument('-i', '--input', help='Roster file to parse')
    parser.add_argument('-o', '--output', help='Calendar output file', default=None)
    parser.add_argument('-l', '--line', help='Starting line in the link', default=None)
    parser.add_argument('-c', '--commence', help='Date the link starts formatted as YYYYMMDD', default=False)
    parser.add_argument('-s', '--startCalendar', help='Date (YYYYMMDD) you want to start the calendar entries from. Leave blank to start from commencement', default=False)
    parser.add_argument('-S', '--startOnCommence', help='Calendar entries will start from the link commencement date.', dest='startCalendar', action='store_true')
    parser.add_argument('-f', '--finishCalendar', help='Date (YYYYMMDD) you want to the calendar entries to finish on.')
    parser.add_argument('-r', '--restOnly', help='Only Output Rest Days', default=False, action='store_true')
    parser.add_argument('-v', '--verbose', help='Displays extra processing information', default=False, action='store_true')
    args = parser.parse_args()

    # WAS A FILE SUPPLIED
    if args.input is None:
        print(f"{col.Red}Please Supply Roster CSV File with the -i option")
        sys.exit(2)

    # DOES THE FILE EXIST
    if exists(args.input):
        print(f"{col.Grn}Found Roster...{col.end}")
        rosterCSV = args.input
    else:
        rosterCSV = input(f"{col.Yel}File NOT found... {col.Cyn}please supply another: {col.end}")
        if not exists(rosterCSV):
            print("File NOT found... Re-run the script.")
            sys.exit(2)
        else:
            print(f"{col.grn}Found Roster...{col.end}")

    # FIND OUT WHEN THE ROSTER STARTS
    invalidCommence = False
    if args.commence:
        valid, outcome = validateDate(args.commence)
        if valid:
            rosterStart = outcome
            print(f"{col.Mag}Link {col.Grn}commences on {col.Umag}{rosterStart.strftime('%a %d %b %Y')}{col.end}")
        else:
            invalidCommence = True
    else:
        startDate = input(f"What {col.Cyn}date{col.end} does this roster {col.Cyn}START{col.end} {col.blu}(YYYYMMDD){col.end} ? : ")
        valid, outcome = validateDate(startDate)
        if valid:
            rosterStart = outcome
            if args.verbose:
                print(rosterStart.strftime('Roster Starts %a %d %b'))
        else:
            invalidCommence = True
    if invalidCommence:
        print(f"{col.Red}Invalid {col.Yel}Commence Date{col.Red} supplied... must use {col.Uyel}YYYYMMDD{col.Red} format.{col.end}")
        sys.exit(2)

    # FIND OUT STARTING POSITION IN THE LINK
    if args.line is None:
        startingLine = input(f"What {col.Cyn}line{col.end} do you {col.Cyn}START{col.end} in this link? : ")
    else:
        startingLine = args.line
    # Validate Starting Line
    try:
        startingLine = int(startingLine)
    except ValueError as verr:
        print(f'{col.Red}Invalid Starting Line, {col.Yel}`{startingLine}`{col.Red}, Aborting!')
        sys.exit(2)
    if not startingLine >= 1:
        print(f'{col.Yel}`{startingLine}`{col.Red} is not a valid option for Starting Line')
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
            print(f"{col.Red}The roster file does not have the correct headings.\n{col.Yel}Expecting;", end='')
            for item in requiredFieldnames:
                color = col.Red if item not in reader.fieldnames else col.Grn
                print(f"\n\t{color}* {item}", end="")
            if args.verbose:
                print(f"\n{col.Yel}Received;", end="")
                for item in reader.fieldnames:
                    print(f"\n\t{col.Mag}# {item}")
            print(f"\n{col.Cyn}Correct the roster file or alter this script and re-run.")
            sys.exit(2)

        print(f"{col.Grn}Reading Roster File...{col.end}")
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


    # DETERMINE WHEN TO START OUTPUTTING CALENDAR ENTRIES
    startDateError = False
    if args.startCalendar:
        if isinstance(args.startCalendar, bool):
            outputStart = rosterStart
            print(f"{col.Grn}Output {col.Mag}starting {col.Grn}from {col.Umag}{outputStart.strftime('%a %d %b %Y')}{col.end}")
        else:
            valid, outcome = validateDate(args.startCalendar)
            if valid:
                outputStart = outcome
                print(f"{col.Grn}Output {col.Mag}starting {col.Grn}from {col.Umag}{outputStart.strftime('%a %d %b %Y')}{col.end}")
            else:
                startDateError = True
    else:
        outputStartRequest = input(f"What {col.Cyn}date{col.end} do you want {col.Cyn}calendar{col.end} entries {col.Cyn}START {col.blu}(YYYYMMDD){col.end} ? : ")
        valid, outcome = validateDate(outputStartRequest)
        if valid:
            outputStart = outcome
        else:
            startDateError = True
    if startDateError:
        print(f"{col.Red}Invalid {col.Yel}Start Date{col.Red} supplied... must use {col.Uyel}YYYYMMDD{col.Red} format.")
        sys.exit(2)
    # Handle Earlier than roster start
    if outputStart < rosterStart:
        print(f"{col.Red}You can't ask for calendar entries before the roster starts. {col.Yel}Try again.")
        sys.exit(2)


    # DETERMINE WHEN TO STOP OUTPUTTING CALENDAR ENTRIES
    finishDateError = False
    if args.finishCalendar:
        valid, outcome = validateDate(args.finishCalendar)
        if valid:
            outputEnd = outcome
            print(f"{col.Grn}Output {col.Mag}finishing {col.Grn}on {col.Umag}{outputEnd.strftime('%a %d %b %Y')}{col.end}")
        else:
            finishDateError = True
    else:
        outputEndRequest = input(f"What {col.Cyn}date{col.end} do you want {col.Cyn}calendar{col.end} entries {col.Cyn}END {col.blu}(YYYYMMDD){col.end} ? : ")
        valid, outcome = validateDate(outputEndRequest)
        if valid:
            outputEnd = outcome
        else:
            finishDateError = True
    if finishDateError:
        print(f"{col.Red}Invalid {col.Yel}Finish Date{col.Red} supplied... must use {col.Uyel}YYYYMMDD{col.Red} format.")
        sys.exit(2)
    if outputEnd < outputStart:
        print(f"{col.Red}You can't ask for calendar entries to end before they have started. {col.Yel}Try again.")
        sys.exit(2)


    # ROTATE DEQUE TO ALIGN WITH START REQUEST
    deltaDays = (outputStart - rosterStart).days
    print(f'{col.grn}Skipping {col.cyn}{deltaDays} days {col.grn}to get to {col.cyn}{outputStart.strftime("%a %d %b")}{col.grn} ...')
    deltaDays += ((startingLine - 1) * 7)
    print(f'Skipping {col.cyn}{startingLine - 1} weeks {col.grn}to get to correct position in the link ...{col.end}')
    shiftsDeque.rotate(-deltaDays)

    currentDay = outputStart        # Counter
    calendarList = calPreamble      # Container for calendar events as they are processed

    while currentDay < outputEnd:
    while currentDay <= outputEnd:
        currentShift = shiftsDeque.popleft()
        # CHECK FOR DAY OF WEEK ALIGNMENT ERROR
        if not currentDay.strftime('%A') == currentShift['dayName']:
            print(f'{col.Red}{currentDay.strftime("%Y-%m-%d")}; DAY MISMATCH ERROR -- {col.Yel}Expecting {currentDay.strftime("%A")} {col.Cyn}GOT {currentShift["dayName"]} {col.Red}-- Aborting!')
            sys.exit(2)
        if args.verbose:
            print(f'{col.Yel}Processing{col.end} {col.mag}{currentDay.strftime("%a %d %b")}{col.end}')
            print(currentShift)
            if args.restOnly and currentShift['rest'] is False:
                print(f'`{col.Cyn}Rest Days Only: {col.mag}`{currentShift["job"]}` on `{currentDay.strftime("%Y%m%d")}` has been OMITTED{col.end}')
            # else:
            #     print(f'`{args.restOnly}` = Rest Only DECLINED')

        if not args.restOnly or currentShift['rest'] is True:
            calendarList += makeEvent(title=currentShift['job'], date=currentDay.strftime('%Y%m%d'),
                                      allDay=currentShift['rest'], start=currentShift['start'],
                                      duration=currentShift['duration'], line=currentShift['line'],
                                      spare=currentShift['spare'], verbose=args.verbose)
        else:
            print(f'{col.Ured}{col.blink}ABORT!!!{col.end}')
            sys.exit(2)
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
    print(f'{col.mag}Calendar File Saved as ... {col.Yel}{outputFile}')


def makeEvent(title, date, line=1, allDay=False, start=False, duration=False, spare=False, verbose=False):
    nowString = datetime.now().strftime('%Y%m%dT%H%M%S')
    # print(f'{col.Cyn}Making Event{col.end} for {nowString}')
    # print(int(date[0:4]), int(date[4:6]), int(date[6:8]))
    if allDay:
        startTime = datetime(int(date[0:4]), int(date[4:6]), int(date[6:8]))
        endTime = startTime
    else:
        startTime = datetime(int(date[0:4]), int(date[4:6]), int(date[6:8]), int(start[0:2]), int(start[2:4]))
        endTime = startTime + timedelta(hours=int(duration[0:2]), minutes=int(duration[2:4]))
    if verbose:
        print(f'Shift {col.Blu}Starts{col.end} {startTime.strftime("%a %d %b - %H.%M")}')
        print(f'Shift {col.Blu}Ends{col.end} {endTime.strftime("%a %d %b - %H.%M")}')
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


def validateDate(suspectData):
    error = False
    # if len(startDate) == 8:
    try:
        dateObject = datetime(int(suspectData[0:4]), int(suspectData[4:6]), int(suspectData[6:8]))
    except ValueError:
        error = "Invalid date format supplied"
        return False, error
    return True, dateObject


if __name__ == "__main__":
    main()