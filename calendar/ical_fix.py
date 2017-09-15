# icalendar is required (python -m pip install icalendar)
import icalendar

def fix_ical_missing_summary(filename):
    ''' Copy the summary line to events missing them if there is an event
    with a matching UID.'''
    ics_string = open(filename, 'rb').read()
    cal = icalendar.Calendar().from_ical(ics_string)

    # Walk the calendar and save a summary for the UIDs, applying it to the
    # events without one. This should be OK as the first occurance of a UID
    # seems to always have a summary. Following events sometimes do not.
    summary_dict = dict()
    for event in cal.walk('vevent'):
        summary = event.get('summary')
        uid = event.get('uid')
        if summary is None:
            event.add('summary', summary_dict[uid])
        else:
            summary_dict[uid] = summary

    # Write the edited calendar back out to the same file
    with open(filename, 'wb') as ofile:
        ofile.write(cal.to_ical())
        ofile.close()
