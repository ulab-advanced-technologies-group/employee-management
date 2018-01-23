import sys
SID = int(sys.argv[1])

from GoogleCalendar import quickstartCal
from GoogleSheets import employee_management

def get_groups(SID):
  	return employee_management.groups(SID)

def get_events(my_groups):
    events = quickstartCal.main()
    sortedevents = []
    if not events:
        print('No upcoming events found.')
    else:
        for event in events:
            start = event['start'].get('dateTime', event['start'].get('date'))
            try:
                attendees = event['attendees']
                for attendee in attendees:
                    try:
                        if attendee['displayName'] in my_groups: # we can look through attendees for group instead of description
                            if event not in sortedevents:
                                sortedevents.append(event)
                    except:
                        pass
            except:
                pass
    print('------These are your events------')
    for event in sortedevents:
        allday = False
        if event['start'].get('dateTime') is None:
            allday = True

        print(event["summary"])

        try:
            if allday:
                start = event['start'].get('date')
                end = event['end'].get('date')
                print('Date & Time:', start[0:10], '-', end[0:9] + str(int(end[9]) - 1))
            else:
                start = event['start'].get('dateTime', event['start'].get('date'))
                end = event['end'].get('dateTime', event['end'].get('date)'))
                # if first two numbers > 12, subtract by 12
                # 24 = 12 AM

                if int(start[11:13]) >= 20 and int(start[11:13]) != 24:
                    start = start[0:11] + str(int(start[11:13])-12) + start[13:16] + ' PM'
                elif int(start[11:13]) > 12:
                    start = start[0:11] + str(int(start[11:13])-12) + start[13:16] + ' PM'
                elif int(start[11:13]) == 12:
                    start = start[0:16] + ' PM'
                elif int(start[11:13]) == 24:
                    start = start[0:11] + str(int(start[11:13])-12) + start[13:16] + ' AM'
                else:
                    start = start[0:16] + ' AM'

                if int(end[11:13]) >= 20 and int(end[11:13]) != 24:
                    end = end[0:11] + str(int(end[11:13])-12) + end[13:16] + ' PM'
                elif int(end[11:13]) > 12:
                    end = end[0:11] + str(int(end[11:13])-12) + end[13:16] + ' PM'
                elif int(end[11:13]) == 12:
                    end = end[0:16] + ' PM'
                elif int(end[11:13]) == 24:
                    end = end[0:11] + str(int(end[11:13])-12) + end[13:16] + ' AM'
                else:
                    end = end[0:16] + ' AM'

                print('Date & Time:', start[0:10], start[11:20], '-', end[0:10], end[11:20])
        except KeyError:
            pass

        try:
            print('Location:', event['location'])
        except KeyError:
            pass

        try:
            print('Description:', event["description"])
        except KeyError:
            pass

        try:
            print('Hangout Link:', event['hangoutLink'])
        except KeyError:
            pass

        print()

        # start = event['start'].get('dateTime', event['start'].get('date'))
        # end = event['end'].get('dateTime', event['end'].get('date)'))
        # print(event["summary"])
        # print('Date & Time:', start[0:10], start[11:25], '-', end[0:10], end[11:25])
        # print('Description:', event["description"])
        # print()

if __name__ == "__main__":
    SID = int(sys.argv[1])

groups = get_groups(SID)
