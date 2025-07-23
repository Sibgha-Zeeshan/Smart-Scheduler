import pandas as pd
import numpy as np
import os
import sys  # To exit script if errors are found
import ast  # For safely evaluating string literals
import re  # For parsing duration strings

# Parse duration string to minutes
def parse_duration(duration_str):
    if isinstance(duration_str, (int, float)):
        return int(duration_str) * 60  # Convert to minutes
    
    if not isinstance(duration_str, str):
        return None
        
    # Clean the input string
    duration_str = duration_str.lower().strip()
    
    # Try to match patterns like "1 hour 15 min", "2 hours", "45 minutes", etc.
    hours = 0
    minutes = 0
    
    # Extract hours
    hour_match = re.search(r'(\d+)\s*(?:hour|hr|h)s?', duration_str)
    if hour_match:
        hours = int(hour_match.group(1))
    
    # Extract minutes
    minute_match = re.search(r'(\d+)\s*(?:minute|min|m)s?', duration_str)
    if minute_match:
        minutes = int(minute_match.group(1))
    
    # If no pattern matched but it's just a number, assume it's hours
    if hours == 0 and minutes == 0:
        try:
            hours = int(duration_str)
        except ValueError:
            return None
    
    return hours * 60 + minutes  # Return total minutes

# Load Excel file
def load_excel(file_path):
    try:
        return pd.ExcelFile(file_path)
    except FileNotFoundError:
        print("❌ File not found! Please upload a valid Excel file.")
        sys.exit(1)  # Stop execution immediately

# Standardize column names
def standardize_columns(df, column_map, sheet_name):
    standardized_columns = {}
    for standard_name, variations in column_map.items():
        for col in df.columns:
            if col.lower().strip() in [v.lower().strip() for v in variations]:
                standardized_columns[col] = standard_name

    df.rename(columns=standardized_columns, inplace=True)
    missing_columns = set(column_map.keys()) - set(standardized_columns.values())

    if missing_columns:
        return None, f"🚨 Missing columns in {sheet_name}: {missing_columns}"

    return df, None

# Validate data types
def check_data_types(df, expected_types, sheet_name):
    errors = []
    for col, expected_type in expected_types.items():
        if col in df.columns:
            # Special handling for Duration
            if col == "Duration":
                for idx, value in df[col].items():
                    duration_minutes = parse_duration(value)
                    if duration_minutes is None:
                        errors.append(f"❌ Invalid duration format in {sheet_name}, row {idx}: '{value}'")
            elif not df[col].apply(lambda x: isinstance(x, expected_type) or pd.isnull(x)).all():
                errors.append(f"❌ Data type mismatch in {sheet_name}: '{col}' should be {expected_type.__name__}.")
    return errors

# Validate and standardize data
def validate_data(excel_data):
    column_mappings = {
        "Courses": {
            "Course_ID": ["Course_ID", "course id", "COURSE_ID"],
            "Course_Name": ["Course_Name", "course name", "COURSE_NAME"],
            "Duration": ["Duration", "DURATION", "duration (hours)"],
            "Course_Type": ["Course_Type", "course type", "COURSE_TYPE"],
            "Capacity": ["Capacity", "capacity", "CAPACITY"],
            "Weekdays": ["Weekdays", "weekdays", "WEEKDAYS"]
        },
        "Rooms": {
            "Room_ID": ["Room_ID", "room id", "ROOM_ID"],
            "Room_Capacity": ["Room_Capacity", "room capacity", "ROOM_CAPACITY"],
            "Room_Type": ["Room_Type", "room type", "ROOM_TYPE"],
        },
        "Faculty": {
            "Faculty_ID": ["Faculty_ID", "faculty id", "FACULTY_ID"],
            "Faculty_Name": ["Faculty_Name", "faculty name", "FACULTY_NAME"],
            "Courses_Assigned": ["Courses_Assigned", "courses assigned", "COURSES_ASSIGNED"],
            "Availability": ["Availability", "availability", "AVAILABILITY"],
            "Start_Time": ["Start_Time", "start time", "START_TIME"],
            "End_Time": ["End_Time", "end time", "END_TIME"],
        },
        "Time Slots": {
            "Day": ["Day", "day", "DAY"],
            "Start_Time": ["Start_Time", "start time", "START_TIME"],
            "End_Time": ["End_Time", "end time", "END_TIME"],
        },
        "Students": {
            "Total_Students": ["Total_Students", "total students", "TOTAL_STUDENTS"]
        }
    }
    expected_dtypes = {
        "Courses": {
            "Course_ID": str,
            "Course_Name": str,
            "Duration": object,  # Changed from int to object to handle string durations
            "Course_Type": str,
            "Capacity": int,
            "Weekdays": int
        },
        "Rooms": {
            "Room_ID": str,
            "Room_Capacity": int,
            "Room_Type": str
        },
        "Faculty": {
            "Faculty_ID": int,
            "Faculty_Name": str,
            "Courses_Assigned": str,
            "Availability": str,
            "Start_Time": int,
            "End_Time": int
        },
        "Time Slots": {
            "Day": str,
            "Start_Time": int,
            "End_Time": int
        },
        "Students": {
            "Total_Students": int
        }
    }

    validated_sheets = {}
    errors = []  # Store all errors at once

    for sheet, column_map in column_mappings.items():
        try:
            df = pd.read_excel(excel_data, sheet_name=sheet, index_col=False)
            df.index += 1  # Shift index to start from 1

            # Remove duplicate rows and log removed values
            duplicate_rows = df[df.duplicated()]
            if not duplicate_rows.empty:
                print(f"⚠️ Found {len(duplicate_rows)} duplicate rows in {sheet}. Removing them...\n")
                df.drop_duplicates(inplace=True)

            df, error = standardize_columns(df, column_map, sheet)
            if error:
                errors.append(error)
                continue
            
            # Pre-process Duration column in Courses sheet
            if sheet == "Courses" and "Duration" in df.columns:
                # Convert the Duration values to minutes and store as a new column
                df["Duration_Minutes"] = df["Duration"].apply(parse_duration)
                invalid_durations = df[df["Duration_Minutes"].isna()]["Duration"].tolist()
                if invalid_durations:
                    errors.append(f"❌ Invalid duration formats in {sheet}: {invalid_durations}")
            
            # Validate data types
            dtype_errors = check_data_types(df, expected_dtypes[sheet], sheet)
            if dtype_errors:
                errors.extend(dtype_errors)

            # Special handling for Courses_Assigned column in Faculty sheet
            if sheet == "Faculty" and "Courses_Assigned" in df.columns:
                # Convert the Courses_Assigned data to a standardized string format
                df["Courses_Assigned"] = df["Courses_Assigned"].apply(lambda x: str(x) if pd.notnull(x) else "")

            # Check for missing values
            missing_values = df.isnull().sum()
            missing_cols = missing_values[missing_values > 0]
            if not missing_cols.empty:
                missing_info = f"🚨 Missing values in {sheet}:"
                for col in missing_cols.index:
                    rows_with_missing = df[df[col].isnull()].index.tolist()
                    missing_info += f"\n  - Column '{col}' has missing values at rows: {rows_with_missing}"
                errors.append(missing_info)

            validated_sheets[sheet] = df
        except Exception as e:
            errors.append(f"❌ Error in {sheet}: {str(e)}")

    # If any errors exist, show all at once and stop execution
    if errors:
        print("\n".join(errors))
        print("\n❗ Please fix all the above errors and re-upload the Excel file.")
        sys.exit(1)  # Stop execution immediately

    print("✅ All sheets validated successfully!")
    return validated_sheets


def allocate_sections(validated_sheets):
    courses = validated_sheets["Courses"]
    students = validated_sheets["Students"]

    total_students = students["Total_Students"].sum()

    for index, course in courses.iterrows():
        num_sections = int(np.ceil(total_students / course["Capacity"]))

        if course["Capacity"] <= 50:
            sections = [f"V{i+1}" for i in range(num_sections)]
        else:  # 50 < Capacity ≤ 100
            sections = [f"C{i+1}" for i in range(num_sections)]

        courses.at[index, "Section"] = ", ".join(sections)

    return courses

# Format duration for display
def format_duration(minutes):
    hours = minutes // 60
    mins = minutes % 60
    
    if mins == 0:
        return f"{hours} hour{'s' if hours > 1 else ''}"
    else:
        return f"{hours} hour{'s' if hours > 1 else ''} {mins} minute{'s' if mins > 1 else ''}"

# Save validated data
def save_to_csv(validated_sheets, folder_path):
    if not os.path.exists(folder_path):
        os.makedirs(folder_path)

    validated_sheets["Courses"] = allocate_sections(validated_sheets)

    for sheet, df in validated_sheets.items():
        # Special handling for Faculty sheet's Courses_Assigned column
        if sheet == "Faculty" and "Courses_Assigned" in df.columns:
            # Ensure the data is stored as is without any additional processing
            df["Courses_Assigned"] = df["Courses_Assigned"].astype(str)
            
        # Format Duration_Minutes back to human-readable format for Courses sheet
        if sheet == "Courses" and "Duration_Minutes" in df.columns:
            df["Duration"] = df["Duration_Minutes"].apply(format_duration)
            # Keep the minutes value as well for calculations
            # df.drop("Duration_Minutes", axis=1, inplace=True)  # Uncomment if you don't want to keep this column

        file_path = os.path.join(folder_path, f"{sheet}.csv")
        df.to_csv(file_path, index=False)
        print(f"📂 {sheet} data saved to {file_path}")


# Main function
def main():
    file_path = "Data.xlsx"
    excel_data = load_excel(file_path)
    validated_sheets = validate_data(excel_data)  # Validation must pass
    save_to_csv(validated_sheets, "validated")
    print("✅ All sheets processed successfully!")

if __name__ == "__main__":
    main()
