from __future__ import print_function
import httplib2
import os

from apiclient import discovery
from oauth2client import client
from oauth2client import tools
from oauth2client.file import Storage
from ConfigParser import SafeConfigParser
from dateutil import parser

import pytz

import datetime

try:
    import argparse

    flags = argparse.ArgumentParser(parents=[tools.argparser]).parse_args()
except ImportError:
    flags = None


class FreeTime:
    def __init__(self):
        safe_parser = SafeConfigParser()
        cwd = os.path.dirname(os.path.realpath(__file__))
        safe_parser.read(os.path.join(cwd, 'config.ini'))
        self.scopes = safe_parser.get('settings', 'SCOPES')
        self.client_secret_file = safe_parser.get('settings', 'CLIENT_SECRET_FILE')
        self.application_name = safe_parser.get('settings', 'APPLICATION_NAME')
        self.time_zone = safe_parser.get('settings', 'TIME_ZONE')
        self.calendar_id = safe_parser.get('settings', 'CALENDAR_ID')
        """Gets valid user credentials from storage.

        If nothing has been stored, or if the stored credentials are invalid,
        the OAuth2 flow is completed to obtain the new credentials.

        Returns:
            Credentials, the obtained credential.
        """
        home_dir = os.path.expanduser('~')
        credential_dir = os.path.join(home_dir, '.credentials')
        if not os.path.exists(credential_dir):
            os.makedirs(credential_dir)
        credential_path = os.path.join(credential_dir,
                                       'calendar-credentials.json')

        store = Storage(credential_path)
        credentials = store.get()
        if not credentials or credentials.invalid:
            flow = client.flow_from_clientsecrets(self.client_secret_file, self.scopes)
            flow.user_agent = self.application_name
            if flags:
                credentials = tools.run_flow(flow, store, flags)
            else:  # Needed only for compatibility with Python 2.6
                credentials = tools.run(flow, store)
            print('Storing credentials to ' + credential_path)
        http = credentials.authorize(httplib2.Http())
        self.service = discovery.build('calendar', 'v3', http=http)

    def minutes_between_two_dates(self, start, end):
        c = end - start
        return c.total_seconds() / 60.0

    def get_days_till_next_weekend(self):
        # next_weekend = datetime.date.today()
        next_weekend = datetime.datetime.now(pytz.timezone(self.time_zone))
        now = datetime.datetime.now(pytz.timezone(self.time_zone))

        while next_weekend.weekday() != 4:
            next_weekend += datetime.timedelta(days=1)
        next_weekend += datetime.timedelta(days=3)
        while next_weekend.weekday() != 4:
            next_weekend += datetime.timedelta(days=1)
        next_weekend += datetime.timedelta(days=1)

        return (next_weekend - now).days

    def get_free_time(self, start_time, body, min_free_slot=1):
        """
        Prints the free time available from your google calendar.
        :param start_time: When your availability starts as a datetime object
        :param body freebusy api body
        :param min_free_slot: The minimum minutes to show a free slot. Default is 1 minute
        """
        times = {}
        events_result = self.service.freebusy().query(body=body).execute()
        cal_dict = events_result[u'calendars']
        date_fmt = "%m/%d %a"
        time_fmt = "%I:%M%p"
        for cal_name in cal_dict:
            start_availability = start_time
            for item in cal_dict[cal_name]['busy']:
                start_obj = parser.parse(item['start'])
                end_obj = parser.parse(item['end'])
                time_difference = self.minutes_between_two_dates(start=start_availability, end=start_obj)
                if time_difference >= min_free_slot:
                    date_str = start_availability.strftime(date_fmt).lstrip('0')
                    if date_str not in times:
                        times[date_str] = []
                    time_str = "{start}-{end}".format(start=start_availability.strftime(time_fmt).lstrip('0'),
                                                      end=start_obj.strftime(time_fmt).lstrip('0'))
                    times[date_str].append(time_str)
                start_availability = end_obj
        # last time for the day
        last_time = parser.parse(body['timeMax'])
        time_difference = self.minutes_between_two_dates(start=start_availability, end=last_time)
        if time_difference >= min_free_slot:
            date_str = start_availability.strftime(date_fmt).lstrip('0')
            if date_str not in times:
                times[date_str] = []
            time_str = "{start}-{end}".format(start=start_availability.strftime(time_fmt).lstrip('0'),
                                              end=last_time.strftime(time_fmt).lstrip('0'))
            times[date_str].append(time_str)
        return times

    def run(self):
        tz = pytz.timezone(self.time_zone)
        now = datetime.datetime.now(pytz.timezone(self.time_zone))
        current_day = tz.localize(datetime.datetime(year=now.year, month=now.month, day=now.day, hour=17, minute=0))

        friday = 4
        saturday = 5
        sunday = 6

        print("Availability this week:")
        days_till_next_weekend = ft.get_days_till_next_weekend()
        for i in range(days_till_next_weekend + 1):
            if current_day.weekday() == friday:
                min_day = tz.localize(
                    datetime.datetime(year=current_day.year, month=current_day.month, day=current_day.day, hour=10,
                                      minute=30))
                max_day = tz.localize(
                    datetime.datetime(year=current_day.year, month=current_day.month, day=current_day.day, hour=20,
                                      minute=0))
            elif current_day.weekday() == saturday:
                min_day = tz.localize(
                    datetime.datetime(year=current_day.year, month=current_day.month, day=current_day.day, hour=10,
                                      minute=30))
                max_day = tz.localize(
                    datetime.datetime(year=current_day.year, month=current_day.month, day=current_day.day, hour=18,
                                      minute=30))
            else:  # Weekday
                min_day = tz.localize(
                    datetime.datetime(year=current_day.year, month=current_day.month, day=current_day.day, hour=17,
                                      minute=0))
                max_day = tz.localize(
                    datetime.datetime(year=current_day.year, month=current_day.month, day=current_day.day, hour=23,
                                      minute=30))
            if current_day.weekday() != sunday:
                request_body = {
                    "timeMin": min_day.isoformat(),
                    "timeMax": max_day.isoformat(),
                    "timeZone": self.time_zone,
                    "items": [{"id": self.calendar_id}]
                }
                if current_day.weekday() == saturday:
                    temp = current_day.replace(hour=10, minute=30)
                    availability_dict = ft.get_free_time(temp, request_body, min_free_slot=60)
                else:
                    availability_dict = ft.get_free_time(current_day, request_body, min_free_slot=60)
                keys = availability_dict.keys()
                if len(keys) > 0:
                    key = availability_dict.keys()[0]
                    print("{date}: {time}".format(date=key, time=', '.join(availability_dict[key])))
            else:
                print("\nAvailability next week:")
            current_day += datetime.timedelta(days=1)


if __name__ == '__main__':
    ft = FreeTime()
    ft.run()
