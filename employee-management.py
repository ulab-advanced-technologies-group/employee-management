import sys
SID = int(sys.argv[1])

from GoogleCalendar import quickstartCal
from GoogleSheets import quickstartSheets

def get_groups(SID):
  	return quickstartSheets.main(SID)

def get_events(my_groups):
    events = quickstartCal.main()
    sortedevents = []
    if not events:
        print('No upcoming events found.')
    else:
        for event in events:
            start = event['start'].get('dateTime', event['start'].get('date'))
            try:
                description = event['description']
                groups = [description]
                splitgroups = groups[0].split(":")
                for group in splitgroups:
                    if group in my_groups:
                        if event not in sortedevents:
                            sortedevents.append(event)
            except:
                pass
    print('------These are your events------')
    for event in sortedevents:
        start = event['start'].get('dateTime', event['start'].get('date'))
        end = event['end'].get('dateTime', event['end'].get('date)'))
        print(event["summary"])
        print('Date & Time:', start[0:10], start[11:25], '-', end[0:10], end[11:25])
        print('Description:', event["description"])
        print()

if __name__ == "__main__":
    SID = int(sys.argv[1])

groups = get_groups(SID)
