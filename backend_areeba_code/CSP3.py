import pandas as pd
import os
import ast
import re
import time
import random
from tqdm import tqdm
from collections import defaultdict
import numpy as np

# Function to extract section number for sorting
def extract_section_number(section):
    match = re.search(r'V(\d+)', section)
    return int(match.group(1)) if match else float('inf')

# Load CSV files from 'validated' folder
data_folder = "validated"
output_folder = "output"

# Create output directory if it doesn't exist
os.makedirs(output_folder, exist_ok=True)

print("Loading data files...")
faculty_df = pd.read_csv(os.path.join(data_folder, "Faculty.csv"))
rooms_df = pd.read_csv(os.path.join(data_folder, "Rooms.csv"))
schedule_df = pd.read_csv(os.path.join(data_folder, "Time Slots.csv"))
courses_df = pd.read_csv(os.path.join(data_folder, "Courses.csv"))

print("Processing faculty assignments...")
# Convert assigned courses from string to dict
faculty_df["Courses_Assigned"] = faculty_df["Courses_Assigned"].apply(ast.literal_eval)

# Function to get 3 courses for faculty based on ratings
def get_three_courses_for_faculty(faculty_courses, all_courses):
    # Step 1: Sort courses by rating
    courses_with_ratings = [(course, rating) for course, rating in faculty_courses.items()]
    courses_with_ratings.sort(key=lambda x: x[1], reverse=True)
    
    assigned_courses_with_ratings = {}
    
    # Step 1: If more than 3 courses, take top 3 highest rated
    if len(courses_with_ratings) >= 3:
        for course, rating in courses_with_ratings[:3]:
            assigned_courses_with_ratings[course] = rating
        return assigned_courses_with_ratings
    
    # Step 2: If less than 3, first take all assigned courses
    for course, rating in courses_with_ratings:
        assigned_courses_with_ratings[course] = rating
    
    # Step 2: Try to fill with additional sections of existing courses
    if len(assigned_courses_with_ratings) < 3:
        original_courses = list(assigned_courses_with_ratings.keys())
        for course in original_courses:
            if len(assigned_courses_with_ratings) >= 3:
                break
            # Get the base course name and rating
            base_rating = assigned_courses_with_ratings[course]
            # Find all sections of this course in courses_df
            course_sections = courses_df[courses_df['Course_ID'] == course]['Section'].iloc[0].split(', ')
            # Add sections as new courses with same rating
            for section in course_sections:
                if len(assigned_courses_with_ratings) >= 3:
                    break
                section_course = course
                if section_course not in assigned_courses_with_ratings:
                    assigned_courses_with_ratings[section_course] = base_rating
    
    # Step 3: If still less than 3, add random courses with rating 1
    if len(assigned_courses_with_ratings) < 3:
        remaining_slots = 3 - len(assigned_courses_with_ratings)
        available_courses = [c for c in all_courses if c not in assigned_courses_with_ratings]
        if available_courses:
            random_courses = random.sample(available_courses, min(remaining_slots, len(available_courses)))
            for course in random_courses:
                assigned_courses_with_ratings[course] = 1  # Assign lowest rating to random courses
    
    return assigned_courses_with_ratings

# Initialize empty list for processed courses
processed_courses_list = []
processed_ratings_list = []

# Process faculty assignments
print("\nProcessing faculty assignments...")
for _, faculty_row in faculty_df.iterrows():
    processed_with_ratings = get_three_courses_for_faculty(
        faculty_row['Courses_Assigned'],
        courses_df['Course_ID'].unique().tolist()
    )
    processed_courses_list.append(list(processed_with_ratings.keys()))
    processed_ratings_list.append(processed_with_ratings)

# Add processed courses and their ratings as new columns
faculty_df['Processed_Courses'] = processed_courses_list
faculty_df['Processed_Ratings'] = processed_ratings_list

# Print summary of assignments
print("\nFinal Faculty Course Assignments:")
for _, faculty_row in faculty_df.iterrows():
    print(f"\nFaculty: {faculty_row['Faculty_Name']} (ID: {faculty_row['Faculty_ID']})")
    print("Original Courses with Ratings:")
    for course, rating in faculty_row['Courses_Assigned'].items():
        print(f"  {course}: {rating}")
    print("Processed Courses with Ratings:")
    for course, rating in faculty_row['Processed_Ratings'].items():
        print(f"  {course}: {rating}")

# Expand courses to include separate entries for each section
expanded_courses = []
for _, row in courses_df.iterrows():
    sections = row["Section"].split(", ")
    for section in sections:
        new_row = row.copy()
        new_row["Course_Name"] = f"{row['Course_Name']}-{section}"
        new_row["Section"] = section
        expanded_courses.append(new_row)

# Sort expanded_courses by Section
def natural_sort_key(course):
    section = course["Section"]
    match = re.search(r'V(\d+)', section)
    return int(match.group(1)) if match else float('inf')

expanded_courses.sort(key=natural_sort_key)

# Convert data into required lists
faculty = faculty_df.to_dict("records")
rooms = rooms_df.sort_values("Room_Capacity").to_dict("records")
schedule = schedule_df.to_dict("records")
courses = expanded_courses

# Define fixed standardized time slots instead of generating them dynamically
timeslots = []

# Define all days from the schedule
days = [day["Day"] for day in schedule]

# Define standard time blocks for different durations
standard_slots = {
    "1h15m": [  # 1 hour 15 minutes blocks
        {"start": "8:00 AM", "end": "9:15 AM"},
        {"start": "9:30 AM", "end": "10:45 AM"},
        {"start": "11:00 AM", "end": "12:15 PM"},
        {"start": "12:30 PM", "end": "1:45 PM"},
        {"start": "2:00 PM", "end": "3:15 PM"},
        {"start": "3:30 PM", "end": "4:45 PM"},
        {"start": "5:00 PM", "end": "6:15 PM"}
    ],
    "2h30m": [  # 2 hours 30 minutes blocks
        {"start": "8:00 AM", "end": "10:30 AM"},
        {"start": "11:00 AM", "end": "1:30 PM"},
        {"start": "2:00 PM", "end": "4:30 PM"},
        {"start": "5:00 PM", "end": "7:30 PM"}
    ]
}

# Add 1 hour blocks and 1.5 hour blocks for completeness
standard_slots["1h"] = [
    {"start": "8:00 AM", "end": "9:00 AM"},
    {"start": "9:00 AM", "end": "10:00 AM"},
    {"start": "10:00 AM", "end": "11:00 AM"},
    {"start": "11:00 AM", "end": "12:00 PM"},
    {"start": "12:00 PM", "end": "1:00 PM"},
    {"start": "1:00 PM", "end": "2:00 PM"},
    {"start": "2:00 PM", "end": "3:00 PM"},
    {"start": "3:00 PM", "end": "4:00 PM"},
    {"start": "4:00 PM", "end": "5:00 PM"},
    {"start": "5:00 PM", "end": "6:00 PM"}
]

standard_slots["1h30m"] = [
    {"start": "8:00 AM", "end": "9:30 AM"},
    {"start": "9:30 AM", "end": "11:00 AM"},
    {"start": "11:00 AM", "end": "12:30 PM"},
    {"start": "12:30 PM", "end": "2:00 PM"},
    {"start": "2:00 PM", "end": "3:30 PM"},
    {"start": "3:30 PM", "end": "5:00 PM"},
    {"start": "5:00 PM", "end": "6:30 PM"}
]

# Create timeslots for each day in the schedule using the standard blocks
for day in days:
    for duration_type, time_blocks in standard_slots.items():
        for block in time_blocks:
            timeslots.append({
                "Day": day,
                "Start_Time": block["start"],
                "End_Time": block["end"],
                "Duration_Type": duration_type
            })

# Add this function to get faculty ratings for a course
def get_faculty_ratings_for_course(course_id, relaxed=False):
    if not relaxed:
        faculty_ratings = [(f, f['Courses_Assigned'].get(course_id, 0)) 
                         for f in faculty 
                         if course_id in f['Courses_Assigned']]
        faculty_ratings.sort(key=lambda x: x[1], reverse=True)
    else:
        faculty_ratings = [(f, f['Courses_Assigned'].get(course_id, 0)) 
                         for f in faculty]
    return faculty_ratings

# Pre-compute time slot combinations
time_slot_combinations = {}
for day in schedule:
    day_name = day["Day"]
    time_slot_combinations[day_name] = {}
    
    # Calculate the maximum possible duration in 15-minute blocks from minutes
    if 'Duration_Minutes' in courses_df.columns:
        # Use vectorized operations instead of conditional
        # For each duration value, calculate how many 15-minute blocks needed
        # Ceiling division to round up to nearest 15-minute block
        blocks_needed = np.ceil(courses_df['Duration_Minutes'] / 15).astype(int)
        max_duration_blocks = blocks_needed.max()
    else:
        # Fallback to the Duration column if Duration_Minutes doesn't exist
        # Assuming Duration is in hours, convert to 15-minute blocks
        max_duration_blocks = max(int(course['Duration']) * 4 for course in courses)
    
    # Create combinations for each possible duration (measured in 15-minute blocks)
    for duration_blocks in range(1, max_duration_blocks + 1):
        time_slot_combinations[day_name][duration_blocks] = []
        for i in range(len(timeslots) - duration_blocks + 1):
            # Only combine consecutive slots from the same day
            if all(t['Day'] == day_name for t in timeslots[i:i + duration_blocks]):
                time_slot_combinations[day_name][duration_blocks].append(timeslots[i:i + duration_blocks])

# Pre-compute all possible time slots for each day and room type
time_slot_cache = {}
for day in schedule:
    day_name = day["Day"]
    time_slot_cache[day_name] = {
        "Lab": [],
        "Lecture": []
    }
    
    # For each timeslot in the day
    for timeslot in [t for t in timeslots if t['Day'] == day_name]:
        # Pre-compute which rooms can be used for this time slot
        for room in rooms:
            if room["Room_Capacity"] >= min(course["Capacity"] for course in courses):
                time_slot_cache[day_name][room["Room_Type"]].append((timeslot, room))

# Pre-compute room assignments by type and capacity
room_assignments = {
    'Lab': [],
    'Lecture': []
}
for room in rooms:
    room_assignments[room['Room_Type']].append(room)

# Sort rooms by capacity for faster filtering
for room_type in room_assignments:
    room_assignments[room_type].sort(key=lambda x: x['Room_Capacity'])

# Utility function to count faculty assignments (no cache)
def count_faculty_assignments(timetable, faculty_id):
    unique_courses = set()
    for (course_name, section, _), details in timetable.items():
        if details['FacultyID'] == faculty_id:
            unique_courses.add((course_name, section))
    return len(unique_courses)

# Restore is_consistent function (no cache, original logic)
def is_consistent(timetable, selected_times, room, faculty_id, section, course_type, conflicts, course_name, day):
    # Check room type constraints first (fastest check)
    if course_type == "Lab" and room["Room_Type"] != "Lab":
        conflicts.append({"CourseName": course_name, "Day": day, "Conflict": "Lab Course Assigned to Non-Lab Room"})
        return False
    if course_type != "Lab" and room["Room_Type"] != "Lecture":
        conflicts.append({"CourseName": course_name, "Day": day, "Conflict": "Lecture Course Assigned to Lab Room"})
        return False

    # Check faculty's processed courses (using direct lookup)
    faculty_member = faculty_df[faculty_df['Faculty_ID'] == faculty_id].iloc[0]
    course_id = courses_df[courses_df['Course_Name'] == course_name.rsplit('-', 1)[0]]['Course_ID'].iloc[0]
    if course_id not in faculty_member['Processed_Courses']:
        conflicts.append({"Course_Name": course_name, "Day": day, "Conflict": "Course not in faculty's processed courses"})
        return False

    # Check faculty course limits using direct count
    current_course_count = count_faculty_assignments(timetable, faculty_id)
    if current_course_count >= 3:
        conflicts.append({"Course_Name": course_name, "Day": day, "Conflict": "Faculty already has 3 courses"})
        return False

    # Create time keys for checking conflicts
    selected_time_keys = [(t['Day'], t['Start_Time']) for t in selected_times]

    # Check for conflicting times in existing timetable
    for (_, _, _), details in timetable.items():
        if details['Room'] == room['Room_ID']:
            for t in details['Times']:
                if (t['Day'], t['Start_Time']) in selected_time_keys:
                    conflicts.append({"CourseName": course_name, "Day": day, "Conflict": f"Room Conflict at {t['Start_Time']}"})
                    return False
        if details['FacultyID'] == faculty_id:
            for t in details['Times']:
                if (t['Day'], t['Start_Time']) in selected_time_keys:
                    conflicts.append({"CourseName": course_name, "Day": day, "Conflict": f"Faculty Conflict at {t['Start_Time']}"})
                    return False
        if details['Section'] == section:
            for t in details['Times']:
                if (t['Day'], t['Start_Time']) in selected_time_keys:
                    conflicts.append({"CourseName": course_name, "Day": day, "Conflict": f"Section Conflict at {t['Start_Time']}"})
                    return False
    return True

def backtrack(timetable, course_index, conflicts, relaxed=False, pbar=None):
    if pbar:
        pbar.update(1)
        pbar.set_description(f"Processing course {course_index + 1}/{len(courses)}")
    if course_index == len(courses):
        return True
    course = courses[course_index]
    course_id = course['Course_ID']
    section = course["Section"]
    # Get requirements directly from course
    def get_course_requirements(course):
        return {
            'capacity': course['Capacity'],
            'duration': course['Duration'],
            'weekdays': course['Weekdays'],
            'type': course['Course_Type']
        }
    requirements = get_course_requirements(course)
    # Determine the appropriate time slot duration type based on course duration
    duration_type = None
    if 'Duration_Minutes' in courses_df.columns:
        course_df_row = courses_df[courses_df['Course_ID'] == course_id].iloc[0]
        duration_minutes = course_df_row['Duration_Minutes']
        if duration_minutes <= 60:
            duration_type = "1h"
        elif duration_minutes <= 75:
            duration_type = "1h15m"
        elif duration_minutes <= 90:
            duration_type = "1h30m"
        else:
            duration_type = "2h30m"
    else:
        if int(requirements['duration']) <= 1:
            duration_type = "1h"
        elif int(requirements['duration']) <= 1.5:
            duration_type = "1h30m"
        else:
            duration_type = "2h30m"
    print(f"\nCourse {course_id}-{section} using {duration_type} time blocks")
    faculty_ratings = get_faculty_ratings_for_course(course_id, relaxed)
    if not faculty_ratings and not relaxed:
        return backtrack(timetable, course_index, conflicts, relaxed=True, pbar=pbar)
    available_faculty = []
    if not relaxed:
        available_faculty = [
            f for f, rating in faculty_ratings 
            if course_id in f['Processed_Courses'] and 
            count_faculty_assignments(timetable, f['Faculty_ID']) < 3
        ]
    else:
        available_faculty = [
            f for f in faculty 
            if count_faculty_assignments(timetable, f['Faculty_ID']) < 3
        ]
    available_faculty.sort(key=lambda f: f['Courses_Assigned'].get(course_id, 0), reverse=True)
    for faculty_member in available_faculty:
        if time.time() - start_time > MAX_TIME_PER_COURSE * (course_index + 1):
            print(f"\nSkipping complex course {course_id}-{section} and moving on.")
            conflicts.append({"CourseName": course['Course_Name'], "Section": section, 
                            "Conflict": "Scheduling timeout - too complex"})
            return backtrack(timetable, course_index + 1, conflicts, relaxed, pbar)
        assignments = []
        assigned_days = 0
        last_assigned_day = None
        for day in schedule:
            day_name = day["Day"]
            if assigned_days >= requirements['weekdays']:
                break
            if last_assigned_day and schedule.index(day) - schedule.index(last_assigned_day) < 3:
                continue
            available_timeslots = [
                ts for ts in timeslots 
                if ts['Day'] == day_name and ts['Duration_Type'] == duration_type
            ]
            if not available_timeslots:
                print(f"No available {duration_type} slots for day {day_name}")
                continue
            room_type = "Lab" if requirements['type'] == "Lab" else "Lecture"
            suitable_rooms = [
                r for r in rooms 
                if r['Room_Capacity'] >= requirements['capacity'] and r['Room_Type'] == room_type
            ]
            if not suitable_rooms:
                continue
            for slot in available_timeslots:
                for room in suitable_rooms:
                    if is_consistent(
                        timetable, [slot], room, faculty_member['Faculty_ID'], section,
                        requirements['type'], conflicts, course['Course_Name'], day_name
                    ):
                        assignment = {
                            "CourseID": course_id,
                            "CourseType": requirements['type'],
                            "FacultyID": faculty_member['Faculty_ID'],
                            "FacultyName": faculty_member['Faculty_Name'],
                            "Times": [slot],
                            "Room": room['Room_ID'],
                            "RoomType": room['Room_Type'],
                            "Section": section,
                            "Rating": faculty_member['Courses_Assigned'].get(course_id, 1)
                        }
                        assignments.append(((course['Course_Name'], section, day_name), assignment))
                        assigned_days += 1
                        last_assigned_day = day
                        break
                if last_assigned_day == day:
                    break
            if assigned_days == requirements['weekdays']:
                for key, value in assignments:
                    timetable[key] = value
                # No cache update needed
                if backtrack(timetable, course_index + 1, conflicts, relaxed, pbar):
                    return True
                for key, _ in assignments:
                    del timetable[key]
                # No cache update needed
    if not relaxed:
        return backtrack(timetable, course_index, conflicts, relaxed=True, pbar=pbar)
    return False

# Measure Execution Time
print("\nStarting timetable generation...")
start_time = time.time()

# Initialize timetable and conflicts
timetable = {}
conflicts = []

# Add timeout and set maximum time for scheduling
MAX_SCHEDULING_TIME = 300  # 5 minutes timeout
MAX_TIME_PER_COURSE = 5    # 5 seconds per course maximum

# Create progress bar
with tqdm(total=len(courses), desc="Generating timetable", unit="course") as pbar:
    try:
        # Add timeout mechanism
        def backtrack_with_timeout(timetable, course_index, conflicts, relaxed=False, pbar=None):
            # Check if we've been running too long overall
            if time.time() - start_time > MAX_SCHEDULING_TIME:
                print("\n⚠️ Scheduling timeout reached. Returning partial schedule.")
                return True
                
            # Check if we're spending too long on this course
            course_start_time = time.time()
            
            result = backtrack(timetable, course_index, conflicts, relaxed, pbar)
            
            # If this course took too long, print a warning
            course_time = time.time() - course_start_time
            if course_time > MAX_TIME_PER_COURSE:
                print(f"\n⚠️ Course {course_index} took {course_time:.2f}s to schedule.")
                
            return result
            
        if backtrack_with_timeout(timetable, 0, conflicts, relaxed=False, pbar=pbar):
            end_time = time.time()
            execution_time = end_time - start_time

            print("\nTimetable generation completed!")
            print(f"Execution Time: {execution_time:.2f} seconds")

            # Calculate Accuracy Metrics
            total_course_sections = len(expanded_courses)
            unique_scheduled_course_sections = len(set((course, section) for (course, section, _) in timetable.keys()))
            unscheduled_course_sections = total_course_sections - unique_scheduled_course_sections
            accuracy = (unique_scheduled_course_sections / total_course_sections) * 100 if total_course_sections > 0 else 0
            
            print(f"Scheduled Course-Sections: {unique_scheduled_course_sections}/{total_course_sections}")
            print(f"Unscheduled Course-Sections: {unscheduled_course_sections}")
            print(f"Accuracy: {accuracy:.2f}%")

            print("\nPreparing output files...")
            # Convert timetable into a DataFrame grouped by sections
            timetable_data = []
            for (course, section, day), details in timetable.items():
                timetable_data.append({
                    "CourseID": details["CourseID"],
                    "CourseName": course,
                    "CourseType": details["CourseType"],
                    "Section": section,
                    "FacultyID": details["FacultyID"],
                    "FacultyName": details["FacultyName"],
                    "Day": day,
                    "Start Time": details["Times"][0]["Start_Time"],
                    "End Time": details["Times"][-1]["End_Time"],
                    "Room": details["Room"],
                    "RoomType": details["RoomType"],
                    "Rating": details.get("Rating", "N/A")  # Include rating in output
                })

            # Sort and save to Excel
            timetable_data = sorted(timetable_data, key=lambda x: extract_section_number(x["Section"]))
            df = pd.DataFrame(timetable_data)

            # Add summary of faculty course assignments with ratings
            print("\nFaculty Course Assignment Summary:")
            faculty_summary = defaultdict(list)
            faculty_ratings = defaultdict(dict)
            for data in timetable_data:
                faculty_id = data["FacultyID"]
                course_name = data["CourseName"].rsplit('-', 1)[0]  # Remove section
                if course_name not in faculty_summary[faculty_id]:
                    faculty_summary[faculty_id].append(course_name)
                    faculty_ratings[faculty_id][course_name] = data["Rating"]
            
            for faculty_id, courses in faculty_summary.items():
                faculty_name = faculty_df[faculty_df['Faculty_ID'] == faculty_id]['Faculty_Name'].iloc[0]
                print(f"\nFaculty: {faculty_name} (ID: {faculty_id})")
                for course in courses:
                    rating = faculty_ratings[faculty_id][course]
                    print(f"Course: {course} (Rating: {rating})")

            # Update Excel output to include ratings
            print("\nSaving timetable to Excel...")
            timetable_output_path = os.path.join(output_folder, "timetable-section.xlsx")
            with pd.ExcelWriter(timetable_output_path, engine="xlsxwriter") as writer:
                workbook = writer.book
                worksheet = workbook.add_worksheet("Timetable")
                writer.sheets["Timetable"] = worksheet

                # Define formats
                bold_format = workbook.add_format({"bold": True, "font_size": 14, "align": "center"})
                header_format = workbook.add_format({"bold": True, "bg_color": "#D9E1F2", "border": 1})

                row = 0
                for section in sorted(df["Section"].unique(), key=extract_section_number):
                    section_df = df[df["Section"] == section]

                    worksheet.merge_range(row, 0, row, len(section_df.columns) - 1, f"Section-{section} Courses", bold_format)
                    row += 1

                    # Write Column Headers
                    for col_num, value in enumerate(section_df.columns):
                        worksheet.write(row, col_num, value, header_format)
                    row += 1
                    
                    # Write Data
                    for record in section_df.itertuples(index=False):
                        for col_num, value in enumerate(record):
                            worksheet.write(row, col_num, value)
                        row += 1
                    
                    row += 2  # Add space before next section

                # Update Faculty Summary Sheet to include ratings
                summary_sheet = workbook.add_worksheet("Faculty Summary")
                summary_sheet.write(0, 0, "Faculty ID", header_format)
                summary_sheet.write(0, 1, "Faculty Name", header_format)
                summary_sheet.write(0, 2, "Course", header_format)
                summary_sheet.write(0, 3, "Rating", header_format)
                
                row = 1
                for faculty_id, courses in faculty_summary.items():
                    faculty_name = faculty_df[faculty_df['Faculty_ID'] == faculty_id]['Faculty_Name'].iloc[0]
                    for course in courses:
                        summary_sheet.write(row, 0, faculty_id)
                        summary_sheet.write(row, 1, faculty_name)
                        summary_sheet.write(row, 2, course)
                        summary_sheet.write(row, 3, faculty_ratings[faculty_id][course])
                        row += 1

            print(f"Timetable saved to {timetable_output_path}")

            if conflicts:
                print("Saving conflicts to CSV...")
                conflicts_df = pd.DataFrame(conflicts)
                conflicts_output_path = os.path.join(output_folder, "conflicts.csv")
                conflicts_df.to_csv(conflicts_output_path, index=False)
                print(f"Conflicts saved to {conflicts_output_path}")
    except Exception as e:
        print(f"Error during timetable generation: {e}")


