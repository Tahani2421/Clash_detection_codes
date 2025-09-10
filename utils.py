import pandas as pd
import pyexcel
import ipdb
import json
from rich import print
import os
import csv
#Import all the necessary meeting

#Read all excel files and convert into list of dictionaries
def read_excel_to_dict(file_dir):
    records = pyexcel.get_records(file_name=file_dir)
    products = [dict(row) for row in records]

    return products

# Function to detect instructor and room clashes from timetable data
def find_instructor_and_room_clash_data(data_dict):
    instructor = {} # Dictionary to store instructor clashes grouped by staff+day
    class_room = {} # Dictionary to store room clashes grouped by room+day
    time = {}
    ignore = [0, 'TBA - Bus']
    # Iterate over each row in dataset
    for indx, item in enumerate(data_dict):
        day = item['Scheduled Days']
        staff = item['Allocated Staff Name']
        room = item['Allocated Location Name']
        
        try:
            # Skip rows with ignored values
            if staff in ignore or room in ignore or day in ignore:
                continue

            else:
                # Create unique keys for instructor/day and room/day
                ins_key = staff + '__' + day
                cls_key = room + '__' + day

                # Extract session start and end times
                start = item['Scheduled Start Time']
                end = item['Scheduled End Time']

                # Check if instructor and room has overlapping clashes
                instructor_clashes_check(ins_key, item, instructor)
                room_clashes_check(cls_key, item, class_room)
                
        except Exception as e:
            # Print error details if something goes wrong in processing
            print(e)
            print(item['Allocated Staff Name'], day, item['Allocated Location Name'])

    # Return only the cleaned clash data (with conflicts)
    return clean_clash_data(instructor, class_room)


# Helper function to check instructor clashes (overlapping times)
def instructor_clashes_check(ins_key, item, instructor):
    start = item['Scheduled Start Time']
    end = item['Scheduled End Time']

# If instructor already has entries for the day, check for overlaps
    if ins_key in instructor:
        new_instructor = instructor[ins_key]
        # Compare the new session time against existing sessions
        for indx, data in enumerate(instructor[ins_key]):
            if start >= data['Scheduled Start Time'] and start < data['Scheduled End Time']:
                new_instructor.append(item)
                break

            elif end > data['Scheduled Start Time'] and end <= data['Scheduled End Time']:
                new_instructor.append(item)
                break
        instructor[ins_key] = new_instructor
    else:
        instructor[ins_key] = [item]
    


# Helper function to check room clashes (overlapping times)
def room_clashes_check(cls_key, item, class_room):
    start = item['Scheduled Start Time']
    end = item['Scheduled End Time']

# If room already has entries for the day, check for overlaps
    if cls_key in class_room:
        new_class = class_room[cls_key]
        for indx, data in enumerate(class_room[cls_key]):
            if start >= data['Scheduled Start Time'] and start < data['Scheduled End Time']:
                new_class.append(item)
                break

            elif end > data['Scheduled Start Time'] and end <= data['Scheduled End Time']:
                new_class.append(item)
                break
        class_room[cls_key] = new_class
    else:
        class_room[cls_key] = [item]


# Function to clean up clash data and keep only actual conflicts
def clean_clash_data(instructor, class_room):
    clash_instructor = []
    clash_room = []
    # Keep only entries where instructor has more than one session (clash)
    for item in instructor:
        if len(instructor[item]) > 1:
            clash_instructor.append(instructor[item])
    
    # Keep only entries where room has more than one session (clash)
    for item in class_room:
        if len(class_room[item]) > 1:
            clash_room.append(class_room[item])
    
    return clash_instructor, clash_room

# Function to write clash data to a CSV file
def write_excel_file(data, file_dir_name):
    with open(f'{file_dir_name}.csv', 'w', newline='') as output:
        writer = csv.writer(output)
        writer.writerow(['Name', 'Module Name', 'Activity Type Name', 'Scheduled Days', 'Scheduled Start Time', 'Scheduled End Time',
        'Duration', 'Scheduled Weeks', 'Allocated Location Name', 'Planned Size', 'Real Size', 'Allocated Staff Name'])
        for item in data:
            for spec in item:
                row = []
                row.append(spec['Name'])
                row.append(spec['Module Name'])
                row.append(spec['Activity Type Name'])
                row.append(spec['Scheduled Days'])
                row.append(spec['Scheduled Start Time'])
                row.append(spec['Scheduled End Time'])
                row.append(spec['Duration'])
                row.append(spec['Scheduled Weeks'])
                row.append(spec['Allocated Location Name'])
                row.append(spec['Planned Size'])
                row.append(spec['Real Size'])
                row.append(spec['Allocated Staff Name'])
                
                writer.writerow(row)    
            writer.writerow([])