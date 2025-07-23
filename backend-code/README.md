# Timetable Generation API

This FastAPI application provides endpoints for validating, generating, and downloading timetables based on faculty, rooms, time slots, courses, and student data.

## Setup and Installation

1. Install the required dependencies:
   ```
   pip install -r requirements.txt
   ```

2. Run the API server using the run script:
   ```
   python run.py
   ```
   
   To enable auto-reload during development:
   ```
   python run.py --reload
   ```

3. Access the application:
   - Web interface: http://localhost:8000/
   - API documentation: http://localhost:8000/docs

## Using the Application

The application provides a simple step-by-step workflow:

1. **Upload Excel File**: Upload a single Excel file with the following sheets:
   - Faculty
   - Rooms
   - Time Slots
   - Courses
   - Students
2. **Validate Excel Data**: Process the Excel file using v.py to generate CSV files
3. **Generate Timetable**: Start the timetable generation process using CSP3.py
4. **Download Files**: Check the status and download the generated timetable and conflicts file

## API Endpoints

### 1. Upload Excel File
- **Endpoint**: `/upload-excel/`
- **Method**: POST
- **Description**: Upload a single Excel file containing all required data

### 2. Validate Excel Data
- **Endpoint**: `/validate-excel/`
- **Method**: POST
- **Description**: Process the Excel file using v.py to generate CSV files

### 3. Generate Timetable
- **Endpoint**: `/generate-timetable/`
- **Method**: POST
- **Description**: Generate a timetable using the validated data (runs as a background task)

### 4. Check Status
- **Endpoint**: `/status/`
- **Method**: GET
- **Description**: Check the status of timetable generation and available files

### 5. Download Files
- **Endpoint**: `/download/{file_name}`
- **Method**: GET
- **Description**: Download the generated timetable or conflicts file

## Excel File Requirements

The Excel file should have the following sheets with these columns:

### Faculty Sheet
- Required columns: Faculty_ID, Faculty_Name, Courses_Assigned
- Courses_Assigned should be a dictionary in string format (e.g., '{"CS101": 5, "CS102": 3}')

### Rooms Sheet
- Required columns: Room_ID, Room_Capacity, Room_Type
- Room_Type should be either "Lab" or "Lecture"

### Time Slots Sheet
- Required columns: Day, Start_Time, End_Time

### Courses Sheet
- Required columns: Course_ID, Course_Name, Duration, Course_Type, Capacity, Weekdays
- Course_Type should be either "Lab" or "Lecture"
- Duration can be specified as "1 hour", "1 hour 30 minutes", etc.

### Students Sheet
- Required columns: Total_Students

## Output Files

- **timetable-section.xlsx**: The generated timetable organized by sections
- **conflicts.csv**: Any conflicts that occurred during timetable generation

## How It Works

The application uses two main components:

1. **v.py**: Handles Excel file validation and processing, generating CSV files with standardized data
2. **CSP3.py**: Implements a Constraint Satisfaction Problem approach to generate a timetable based on the validated data

## Project Structure

```
timetable-generation-api/
├── app.py               # FastAPI application
├── v.py                 # Excel validation and processing
├── CSP3.py              # Timetable generation algorithm
├── requirements.txt     # Dependencies
├── static/              # Static files
│   └── index.html       # Web interface
├── uploads/             # Directory for uploaded Excel files
├── validated/           # Directory for validated CSV files
└── output/              # Directory for generated output files
``` 