import sys
# HOME_DIR ==
sys.path.append('/Users/khangnguyen/Desktop/Google API/Google Sheets')
sys.path.append('/Users/khangnguyen/Desktop/Google API/Google Calendar')

import quickstartCal.py
import quickstartSheets.py

SID = int(sys.argv[0])

def get_groups(SID):
  	return quickstartSheets.py.main(SID)

groups = get_groups(SID)

def get_events(my_groups):
		events = quickstartCal.py.main()
    sortedevents = []
    if not events:
        print('No upcoming events found.')
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
