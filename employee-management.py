import sys
SID = int(sys.argv[1])

# HOME_DIR ==

from GoogleCalendar import quickstartCal
from GoogleSheets import quickstartSheets

def get_groups(SID):
  	return quickstartSheets.py.main(SID)

#groups = get_groups(SID)

def get_events(my_groups):
    events = quickstartCal.py.main()
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
            except KeyError:
                None
    return sortedevents
