#!/usr/bin/env python3

from flask import Flask, render_template

app = Flask(__name__, static_url_path='', static_folder='AdminLTE')

import os
import sys
import poplib
import email
from icalendar import Calendar, Event
from datetime import datetime
import json

def read_mbox():
    Mailbox = poplib.POP3_SSL('zimbra.dmmp.me')
    Mailbox.user('room303@ntel.ru')
    Mailbox.pass_('TraClaDra303')
    events = []
    numMessages = len(Mailbox.list()[1])
    for i in range(numMessages):
        is_calendar = False
        raw_email = b"\n".join(Mailbox.retr(i+1)[1])
        parsed_email = email.message_from_bytes(raw_email)
        for part in parsed_email.walk():
            if part.get_content_type() == 'text/calendar' and '.ics' in part.get_filename():
                is_calendar = True
                attachment = part.get_payload()
                cal = Calendar()
                events.append(cal.from_ical(attachment))
        if not is_calendar:
            Mailbox.dele(1)
    Mailbox.quit()
    return events


def ical_dumps(events):
    js_events = []
    for meeting in events:
        for component in meeting.walk():
            if component.name == "VEVENT":
                js_event = {}
                js_event['uid'] = component.decoded('uid').decode('utf-8')
                js_event['LAST-MODIFIED'] = component.decoded(
                    'LAST-MODIFIED').strftime("%s")
                # js_event['title'] = "%s %s" % (component.decoded("summary"), component.decoded("location"))
                js_event['title'] = '%s %s' % (component.decoded("summary").decode('utf-8'), '')
                js_event['dtstart'] = component.decoded("dtstart").strftime("%Y-%m-%d %H:%M")
                js_event['dtend'] = component.decoded("dtend").strftime("%Y-%m-%d %H:%M")
                js_event['backgroundColor'] = '#3c8dbc'
                js_event['borderColor'] = '#3c8dbc'
                js_event['allday'] = 'false'

                # New event in the calendar
                if js_event['uid'] not in [element['uid'] for element in js_events]:
                    js_events.append(js_event)
                else:
                    for i in range(len(js_events)):
                        # Changed event. Replace it with new version.
                        if js_events[i]['uid'] == js_event['uid'] and js_events[i]['LAST-MODIFIED'] < js_event['LAST-MODIFIED']:
                            del js_events[i]
                            js_events.append(js_event)

    return js_events



@app.route('/')
def hello():
    return render_template('index.html')

@app.route('/calendar/')
def calendar():
    ical_events = read_mbox()
    return render_template('calendar.html', js_var=ical_dumps(ical_events))
