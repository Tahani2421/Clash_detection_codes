from utils import *
# Import all helper functions from utils.py
import ipdb
# Import ipdb for debugging
import json
 # For potential JSON handling
from rich import print
# Rich library for proper console output

def main():
    excel_file_dir = 'BOS Live TT 24-25_for TT project.xlsx' 
    # Define the Excel file path that contains timetable data
    data = read_excel_to_dict(excel_file_dir) # parse the file and get data according to our needed format
    instructor, place = find_instructor_and_room_clash_data(data)
    # Find clashes for instructors and rooms using the dataset
    write_excel_file(instructor, '../instructor_clash_output')
    write_excel_file(place, '../room_clash_output')
    print("Here are the Instructors that clash time")
    print('##############################################')
    # print(instructor)

    print("\n\nHere are the rooms that clash time")
    print('##############################################')
    # print(place)
    
    

# Entry point of the script
if __name__ == '__main__':
    main()