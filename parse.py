import pandas as pd
from datetime import datetime
from dotenv import load_dotenv
from typing import Hashable
import os
from icalendar import Calendar, Event, vCalAddress
from pytz import timezone
import os

load_dotenv()

file_name = "./data.xlsx"

target_tab_name = 'END'

feature_row_index = 1

name_column_index = 2

cur_timezone = timezone("Asia/Singapore")

TARTET_NAME =  os.environ.get("TARTET_NAME")
SUMMARY_PREFIX = os.environ.get("SUMMARY_PREFIX")
ATTENDEES = os.environ.get("ATTENDEES").split(",")

calendar = Calendar()

def main():
    
  xl_file = pd.read_excel(file_name, sheet_name=None)
  
  tab_data = xl_file[target_tab_name]

  parse_target_name_row(
    tab_data=tab_data, 
    date_time_dict= get_column_key_to_datetime_dict(tab_data=tab_data),
  )
  print(calendar)
  f = open('./result.ics', 'wb')
  f.write(calendar.to_ical())
  f.close()

def get_column_key_to_datetime_dict(tab_data: pd.DataFrame):
  column_key_to_datetime_dict = {}
  for index, row_series in tab_data.iterrows():
    if index != feature_row_index:
       continue  
    for key, item in row_series.items():
        if type(item) == datetime:
            column_key_to_datetime_dict[key] = item
  return column_key_to_datetime_dict


def parse_target_name_row(tab_data: pd.DataFrame, date_time_dict: dict[Hashable, datetime]):
   for _, row_series in tab_data.iterrows():
      name = row_series.tolist()[name_column_index]
      if name == TARTET_NAME:
        for key in date_time_dict.keys():
          item = row_series.get(key=key)
          if type(item) != str:
             continue
          if 'HO' in item:
            add_ho_event(event_name=item, cur_datetime=date_time_dict[key])

def add_ho_event(event_name,cur_datetime):
  event = Event()
  event.add('summary', SUMMARY_PREFIX + ' ' + event_name)
  event.add('dtstart', datetime(cur_datetime.year, cur_datetime.month,cur_datetime.day,6,0,0,tzinfo=cur_timezone))
  event.add('dtend',  datetime(cur_datetime.year, cur_datetime.month,cur_datetime.day + 1,8,0,0,tzinfo=cur_timezone))

  calendar.add_component(event)

if __name__ == "__main__":
    main()