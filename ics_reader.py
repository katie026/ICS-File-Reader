import os
import re
import datetime

class Calendar:
    version = None
    prod_id = None
    name = "" # X-WR-CALNAME
    events = []

class CalEvent(object):
    start_date = None
    end_date = None
    creation_date = None
    summary = None
    uid = None

def print_event_info(event):
    print(f"Event Name: {event.summary}")
    print(f"Start Date: {event.start_date}")
    print(f"End Date: {event.end_date}")
    print(f"Created: {event.creation_date}")
    print(f"UID: {event.uid}")

def create_event(start_date, end_date, summary):
    event = CalEvent()
    event.start_date = start_date
    event.end_date = end_date
    event.summary = summary
    return event

def get_download_path():
    # Returns the default downloads path for linux or windows
    if os.name == 'nt':
        import winreg
        sub_key = r'SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders'
        downloads_guid = '{374DE290-123F-4565-9164-39C4925E467B}'
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, sub_key) as key:
            location = winreg.QueryValueEx(key, downloads_guid)[0]
        return location
    else:
        return os.path.join(os.path.expanduser('~'), 'downloads')

def clean_filename(filename):
    # \W character class matches any non-word character (i.e., any character that is not a letter, digit, or underscore)
    clean_filename = re.sub(r'[\W]', '', filename.lower())
    return clean_filename

def get_ics_filename(filename):
    if filename.endswith(".ics"):
        # file already ends with .ics
        return filename
    else:
        # add .ics extension
        filename += ".ics"
    
    # return value
    return filename

def get_filepath():
    download_path = get_download_path()

    print("Enter the filepath for the .ics file. This program will autoamtically look in the downloads folder if not otherwise defined.")

    filepath = ""
    valid_file = False
    # start input loop
    while valid_file == False:
        # get user input
        user_input = input("File Name: ")

        # check if filepath or not
        if user_input.find("/") == -1:
            # there is no "/" in the input, assume file is in the downloads folder
            filename = user_input
            # ensure file is in ics format
            filename = get_ics_filename(filename)
            filepath = f"{download_path}/{filename}"
        else:
            # assume user entered a filepath
            filepath = user_input
        
        # check if filepath is valid
        if os.path.isfile(filepath):
            # exit input loop
            valid_file = True
            break

        # ask user again
        print(f"{filepath} is not a valid .ics file.")

    return(filepath)

def string_to_date(string, time = True):
    # The “T” character separates the date from the time portion of the string. The “Z” character is the UTC offset representation. UTC is Coordinated Universal Time, which is the primary time standard used to regulate times worldwide. 
    # YYYYMMDDTHHmmssZ
    # 0123456789012345
    # string[inclusive:exclusive]
    try:
        year = int(string[0:4])
        month = int(string[4:6])
        day = int(string[6:8])
        
        date = datetime.datetime(year, month, day)
        
        if time == True:
            hour = int(string[9:11])
            minute = int(string[11:13])
            second = int(string[13:15])
            date = datetime.datetime(year, month, day, hour, minute, second)

        return date
    except ValueError:
        print("Invalid datetime input. Could not convert to datetime object.")

def create_event_dictionaries(lines):
    # define index where events start
    start_indexes = []
    # define index where events end
    end_indexes = []

    i = 0
    for line in lines:
        if "BEGIN:VEVENT" in line:
            start_indexes.append(i)
        if "END:VEVENT" in line:
            end_indexes.append(i)
        i+=1

    # create list of dictionaries to store start and end indices
    event_dictionaries = []

    # for each event
    i = 0
    for x in range(len(start_indexes)):
        # create a dictionary to store start and end date indexes
        event_dict = {
            "start_date_index" : start_indexes[i],
            "end_date_index" : end_indexes[i]
        }
        # add to dictionary list
        event_dictionaries.append(event_dict)
        i+=1
    
    # print(f"lines: {len(lines)}, starts: {len(start_indexes)}, ends: {len(end_indexes)}, dictionaries: {len(event_dictionaries)}")
    return event_dictionaries

def create_events(lines, event_index_dictionaries):
    # create list of events
    events = []

    # for each event dictionary
    for event_dict in event_index_dictionaries:
        # start date index
        start_date_index = event_dict["start_date_index"]
        # end date index
        end_date_index = event_dict["end_date_index"]

        # create an event object
        event = CalEvent()

        # read each line of the event
        for i in range(start_date_index,end_date_index+1):
            line_value = lines[i]
            line_values = line_value.split(":",1)

            # assign start date
            if "DTSTART" in line_values[0]:
                if line_values[0] == "DTSTART":
                    date = string_to_date(line_values[1])
                if line_values[0] == "DTSTART;VALUE=DATE":
                    date = string_to_date(line_values[1], False)
                event.start_date = date
            # assign end date
            if "DTEND" in line_values[0]:
                if line_values[0] == "DTEND":
                    date = string_to_date(line_values[1])
                if line_values[0] == "DTEND;VALUE=DATE":
                    date = string_to_date(line_values[1], False)
                event.end_date = date
            # assign creation date
            if line_values[0] == "CREATED":
                date = string_to_date(line_values[1])
                event.creation_date = date
            # assign summary
            if line_values[0] == "SUMMARY":
                event.summary = line_values[1].strip('\n')
            # assign uid
            if line_values[0] == "UID":
                event.uid = line_values[1]

        # add event to events list
        events.append(event)
    
    return events

def get_event_by_max_min_attribute(events, attribute, extreme = "max"):
    if not events:
        return None  # or raise an error, depending on the use case
    
    if extreme == "min":
        event = min(events, key=lambda x: getattr(x, attribute))
    elif extreme == "max":
        event = max(events, key=lambda x: getattr(x, attribute))
    else:
        raise ValueError('extreme must be either "min" or "max"')
    
    return event

def main():
    # Introduction
    print("\nHello! I will help you find information on the Earliest Event in your .ics file.\n")

    # get file
    filepath = get_filepath()
    
    with open(filepath, 'r' ) as ics_file:
        # read lines from file
        lines = ics_file.readlines()

        # create dictionaries to store start and end date indices
        event_indices = create_event_dictionaries(lines)

        # create list of events using indices
        events = create_events(lines, event_indices)

        # find event with the oldest start_date
        event_with_oldest_start_date = get_event_by_max_min_attribute(events, "start_date", "min")

        # print the event info
        print("")
        print_event_info(event_with_oldest_start_date)

main()