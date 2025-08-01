# type: ignore 
import pandas as pd
import numpy as np
import random
import time
import os
import re
import ast
from typing import List, Dict, Tuple, Any
from dataclasses import dataclass
from collections import defaultdict
from tqdm import tqdm

@dataclass
class Gene:
    """Represents a single session in the timetable"""
    CourseID: str
    Section: str
    FacultyID: int
    RoomID: str
    Day: str
    StartTime: float
    EndTime: float
    
    def __hash__(self):
        return hash((self.CourseID, self.Section, self.FacultyID, self.RoomID, self.Day, self.StartTime))

@dataclass
class Chromosome:
    """Represents a complete timetable solution"""
    genes: List[Gene]
    fitness: float = 0.0
    
    def __len__(self):
        return len(self.genes)

class GeneticTimetableGenerator:
    def __init__(self, input_file: str, population_size: int = 75, generations: int = 150, 
                 mutation_rate: float = 0.05, tournament_size: int = 3, timeout_minutes: int = 2,
                 skip_soft_constraints: bool = False):
        """
        Initialize the genetic algorithm timetable generator
        
        Args:
            input_file: Path to the Excel file with input data
            population_size: Number of chromosomes in population (reduced for speed)
            generations: Maximum number of generations (reduced for speed)
            mutation_rate: Probability of mutation per gene (increased for faster convergence)
            tournament_size: Size of tournament for selection (reduced for speed)
            timeout_minutes: Maximum runtime in minutes (reduced for speed)
            skip_soft_constraints: Skip soft constraint checking for faster execution
        """
        self.input_file = input_file
        self.population_size = population_size
        self.generations = generations
        self.mutation_rate = mutation_rate
        self.tournament_size = tournament_size
        self.timeout_seconds = timeout_minutes * 60
        self.skip_soft_constraints = skip_soft_constraints
        
        # Data storage
        self.courses_df = None
        self.faculty_df = None
        self.rooms_df = None
        self.timeslots_df = None
        self.students_df = None
        
        # Processed data
        self.courses = []
        self.faculty = []
        self.rooms = []
        self.time_slots = []
        self.total_students = 0
        
        # Genetic algorithm state
        self.population = []
        self.best_chromosome = None
        self.generation_history = []
        
        # Conflict tracking (matching CSP3.py)
        self.all_conflict_reasons = defaultdict(set)
        
        # Constraint weights (reduced for faster computation)
        self.hard_constraint_weight = 50  # Reduced from 100
        self.soft_constraint_weight = 5   # Reduced from 10
        
    def step1_preprocess_inputs(self):
        """Step 1: Preprocessing Inputs - Read and parse validated CSV files"""
        print("Step 1: Preprocessing Inputs...")
        
        # Read from validated CSV files (matching CSP3.py approach)
        data_folder = "validated"
        
        # Check if validated folder exists
        if not os.path.exists(data_folder):
            raise FileNotFoundError(f"Validated data folder '{data_folder}' not found. "
                                  f"Please ensure your input file has been processed through the validation step first.")
        
        # Check if required CSV files exist
        required_files = ["Courses.csv", "Faculty.csv", "Rooms.csv", "Time Slots.csv", "Students.csv"]
        missing_files = [f for f in required_files if not os.path.exists(os.path.join(data_folder, f))]
        if missing_files:
            raise FileNotFoundError(f"Missing required CSV files in '{data_folder}': {missing_files}. "
                                  f"Please ensure your input file has been properly validated.")
        self.courses_df = pd.read_csv(os.path.join(data_folder, "Courses.csv"))
        self.faculty_df = pd.read_csv(os.path.join(data_folder, "Faculty.csv"))
        self.rooms_df = pd.read_csv(os.path.join(data_folder, "Rooms.csv"))
        self.timeslots_df = pd.read_csv(os.path.join(data_folder, "Time Slots.csv"))
        self.students_df = pd.read_csv(os.path.join(data_folder, "Students.csv"))
        
        # Convert assigned courses from string to dict (matching CSP3.py)
        self.faculty_df["Courses_Assigned"] = self.faculty_df["Courses_Assigned"].apply(ast.literal_eval)
        
        # Verify required columns exist
        required_columns = {
            'courses_df': ['Course_ID', 'Course_Name', 'Duration', 'Course_Type', 'Capacity', 'Weekdays', 'Section'],
            'faculty_df': ['Faculty_ID', 'Faculty_Name', 'Courses_Assigned'],
            'rooms_df': ['Room_ID', 'Room_Capacity', 'Room_Type'],
            'timeslots_df': ['Day', 'Start_Time', 'End_Time'],
            'students_df': ['Total_Students']
        }
        
        for df_name, columns in required_columns.items():
            df = getattr(self, df_name)
            missing_columns = [col for col in columns if col not in df.columns]
            if missing_columns:
                raise ValueError(f"Missing required columns in {df_name}: {missing_columns}. "
                               f"Please ensure your input file has been properly validated and contains all required columns.")
        
        # Parse duration to minutes
        self.courses_df['Duration_Minutes'] = self.courses_df['Duration'].apply(self._parse_duration)
        
        # Parse faculty availability
        self.faculty_df['Available_Days'] = self.faculty_df['Availability'].apply(self._parse_availability)
        
        # Parse courses assigned (convert string to dict)
        self.faculty_df['Courses_Assigned'] = self.faculty_df['Courses_Assigned'].apply(self._parse_courses_assigned)
        
        # Process faculty assignments like CSP3.py
        print("Processing faculty assignments...")
        
        # Function to get 3 courses for faculty based on ratings (matching CSP3.py)
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
                    course_sections = self.courses_df[self.courses_df['Course_ID'] == course]['Section'].iloc[0].split(', ')
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
        for _, faculty_row in self.faculty_df.iterrows():
            processed_with_ratings = get_three_courses_for_faculty(
                faculty_row['Courses_Assigned'],
                self.courses_df['Course_ID'].unique().tolist()
            )
            processed_courses_list.append(list(processed_with_ratings.keys()))
            processed_ratings_list.append(processed_with_ratings)
        
        # Add processed courses and their ratings as new columns
        self.faculty_df['Processed_Courses'] = processed_courses_list
        self.faculty_df['Processed_Ratings'] = processed_ratings_list
        
        # Print summary of assignments
        print("\nFinal Faculty Course Assignments:")
        for _, faculty_row in self.faculty_df.iterrows():
            print(f"\nFaculty: {faculty_row['Faculty_Name']} (ID: {faculty_row['Faculty_ID']})")
            print("Original Courses with Ratings:")
            for course, rating in faculty_row['Courses_Assigned'].items():
                print(f"  {course}: {rating}")
            print("Processed Courses with Ratings:")
            for course, rating in faculty_row['Processed_Ratings'].items():
                print(f"  {course}: {rating}")
        
        # Generate sections from student count
        self.total_students = self.students_df['Total_Students'].iloc[0]
        self._generate_sections()
        
        # Convert to internal data structures
        self._convert_to_internal_structures()
        
        print(f"✓ Processed {len(self.courses)} course sections")
        print(f"✓ Processed {len(self.faculty)} faculty members")
        print(f"✓ Processed {len(self.rooms)} rooms")
        print(f"✓ Generated {len(self.time_slots)} time slots")
        
    def _parse_duration(self, duration_str: str) -> int:
        """Parse duration string to minutes"""
        if pd.isna(duration_str):
            return 60  # Default 1 hour
            
        duration_str = str(duration_str).lower().strip()
        
        # Extract hours and minutes
        hours = 0
        minutes = 0
        
        hour_match = re.search(r'(\d+)\s*(?:hour|hr|h)s?', duration_str)
        if hour_match:
            hours = int(hour_match.group(1))
        
        minute_match = re.search(r'(\d+)\s*(?:minute|min|m)s?', duration_str)
        if minute_match:
            minutes = int(minute_match.group(1))
        
        # If no pattern matched but it's just a number, assume it's hours
        if hours == 0 and minutes == 0:
            try:
                hours = int(duration_str)
            except ValueError:
                return 60  # Default 1 hour
        
        return hours * 60 + minutes
    
    def _parse_availability(self, availability_str: str) -> List[str]:
        """Parse faculty availability to list of days"""
        if pd.isna(availability_str):
            return ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday"]
            
        availability_str = str(availability_str).strip()
        
        if "Monday-Friday" in availability_str:
            return ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday"]
        elif "Monday-Thursday" in availability_str:
            return ["Monday", "Tuesday", "Wednesday", "Thursday"]
        elif "Tuesday-Friday" in availability_str:
            return ["Tuesday", "Wednesday", "Thursday", "Friday"]
        elif "Monday-Wednesday" in availability_str:
            return ["Monday", "Tuesday", "Wednesday"]
        elif "Tuesday-Thursday" in availability_str:
            return ["Tuesday", "Wednesday", "Thursday"]
        elif "Wednesday-Friday" in availability_str:
            return ["Wednesday", "Thursday", "Friday"]
        else:
            # Parse individual days
            days = []
            day_mapping = {
                "monday": "Monday", "tuesday": "Tuesday", "wednesday": "Wednesday",
                "thursday": "Thursday", "friday": "Friday", "saturday": "Saturday"
            }
            
            for day in availability_str.split('-'):
                day = day.strip().lower()
                if day in day_mapping:
                    days.append(day_mapping[day])
            
            return days if days else ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday"]
    
    def _parse_courses_assigned(self, courses_str: str) -> Dict[str, int]:
        """Parse courses assigned string to dictionary"""
        if pd.isna(courses_str):
            return {}
            
        try:
            # Handle both string and dict formats
            if isinstance(courses_str, dict):
                return courses_str
            elif isinstance(courses_str, str):
                # Remove any extra quotes and evaluate
                courses_str = courses_str.strip()
                if courses_str.startswith("'") and courses_str.endswith("'"):
                    courses_str = courses_str[1:-1]
                return eval(courses_str)
            else:
                return {}
        except:
            return {}
    
    def _generate_sections(self):
        """Generate sections based on student count and course capacity"""
        expanded_courses = []
        
        for _, course in self.courses_df.iterrows():
            capacity = course['Capacity']
            num_sections = int(np.ceil(self.total_students / capacity))
            
            # Generate section names
            if capacity <= 50:
                sections = [f"V{i+1}" for i in range(num_sections)]
            else:
                sections = [f"C{i+1}" for i in range(num_sections)]
            
            # Create separate entries for each section
            for section in sections:
                new_course = course.copy()
                new_course['Section'] = section
                new_course['Course_Name'] = f"{course['Course_Name']}-{section}"
                expanded_courses.append(new_course)
        
        self.courses_df = pd.DataFrame(expanded_courses)
    
    def _convert_to_internal_structures(self):
        """Convert DataFrames to internal data structures"""
        # Convert courses
        self.courses = self.courses_df.to_dict('records')
        
        # Convert faculty
        self.faculty = self.faculty_df.to_dict('records')
        
        # Convert rooms
        self.rooms = self.rooms_df.to_dict('records')
        
        # Generate time slots
        self._generate_time_slots()
    
    def _generate_time_slots(self):
        """Generate time slots matching CSP3.py approach"""
        print("Generating time slots...")
        
        # Define standard time blocks like CSP3.py
        standard_slots = {
            "1h": [  # 1 hour blocks
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
            ],
            "1h15m": [  # 1 hour 15 minutes blocks
                {"start": "8:00 AM", "end": "9:15 AM"},
                {"start": "9:30 AM", "end": "10:45 AM"},
                {"start": "11:00 AM", "end": "12:15 PM"},
                {"start": "12:30 PM", "end": "1:45 PM"},
                {"start": "2:00 PM", "end": "3:15 PM"},
                {"start": "3:30 PM", "end": "4:45 PM"},
                {"start": "5:00 PM", "end": "6:15 PM"}
            ],
            "1h30m": [  # 1.5 hour blocks
                {"start": "8:00 AM", "end": "9:30 AM"},
                {"start": "9:30 AM", "end": "11:00 AM"},
                {"start": "11:00 AM", "end": "12:30 PM"},
                {"start": "12:30 PM", "end": "2:00 PM"},
                {"start": "2:00 PM", "end": "3:30 PM"},
                {"start": "3:30 PM", "end": "5:00 PM"},
                {"start": "5:00 PM", "end": "6:30 PM"}
            ],
            "2h30m": [  # 2.5 hour blocks
                {"start": "8:00 AM", "end": "10:30 AM"},
                {"start": "11:00 AM", "end": "1:30 PM"},
                {"start": "2:00 PM", "end": "4:30 PM"},
                {"start": "5:00 PM", "end": "7:30 PM"}
            ]
        }
        
        # Get days from timeslots_df
        days = self.timeslots_df['Day'].unique().tolist()
        
        # Create timeslots for each day using standard blocks
        self.time_slots = []
        for day in days:
            for duration_type, time_blocks in standard_slots.items():
                for block in time_blocks:
                    # Convert time strings to float hours
                    start_time = self._time_str_to_float(block["start"])
                    end_time = self._time_str_to_float(block["end"])
                    
                    self.time_slots.append({
                        "Day": day,
                        "Start_Time": start_time,
                        "End_Time": end_time,
                        "Duration_Type": duration_type,
                        "Start_Time_Str": block["start"],
                        "End_Time_Str": block["end"]
                    })
        
        print(f"Generated {len(self.time_slots)} time slots")
    
    def _time_str_to_float(self, time_str: str) -> float:
        """Convert time string to float hours (e.g., '8:00 AM' -> 8.0)"""
        time_str = time_str.strip().upper()
        match = re.match(r'(\d+):(\d+)\s*([AP]M)', time_str)
        if not match:
            return 0.0
        
        hour, minute, ampm = int(match.group(1)), int(match.group(2)), match.group(3)
        if ampm == 'PM' and hour != 12:
            hour += 12
        if ampm == 'AM' and hour == 12:
            hour = 0
        
        return hour + minute / 60.0
    
    def _get_duration_type(self, duration_minutes: int) -> str:
        """Get duration type based on minutes (matching CSP3.py logic)"""
        if duration_minutes <= 60:
            return "1h"
        elif duration_minutes <= 75:
            return "1h15m"
        elif duration_minutes <= 90:
            return "1h30m"
        else:
            return "2h30m"
    
    def _get_faculty_ratings_for_course(self, course_id: str, relaxed: bool = False):
        """Get faculty ratings for a course (matching CSP3.py logic)"""
        if not relaxed:
            faculty_ratings = [(f, f['Courses_Assigned'].get(course_id, 0)) 
                             for f in self.faculty 
                             if course_id in f['Courses_Assigned']]
            faculty_ratings.sort(key=lambda x: x[1], reverse=True)
        else:
            faculty_ratings = [(f, f['Courses_Assigned'].get(course_id, 0)) 
                             for f in self.faculty]
        return faculty_ratings
    
    def step2_create_chromosome_structure(self) -> List[Gene]:
        """Step 2: Create chromosome structure - Generate genes for all course sections"""
        print("Step 2: Creating Chromosome Structure...")
        
        genes = []
        
        for course in self.courses:
            course_id = course['Course_ID']
            section = course['Section']
            duration_minutes = course['Duration_Minutes']
            weekdays = course['Weekdays']
            
            # Create genes for each required session
            for session in range(weekdays):
                gene = Gene(
                    CourseID=course_id,
                    Section=section,
                    FacultyID=0,  # Will be assigned randomly
                    RoomID="",    # Will be assigned randomly
                    Day="",       # Will be assigned randomly
                    StartTime=0.0, # Will be assigned randomly
                    EndTime=0.0    # Will be calculated
                )
                genes.append(gene)
        
        print(f"✓ Created {len(genes)} genes for chromosome structure")
        return genes
    
    def step3_generate_initial_population(self):
        """Step 3: Generate initial population"""
        print("Step 3: Generating Initial Population...")
        
        base_genes = self.step2_create_chromosome_structure()
        
        for i in tqdm(range(self.population_size), desc="Generating population"):
            chromosome = self._create_random_chromosome(base_genes)
            self.population.append(chromosome)
        
        print(f"✓ Generated {len(self.population)} chromosomes")
    
    def _create_random_chromosome(self, base_genes: List[Gene]) -> Chromosome:
        """Create a random chromosome with proper duration matching"""
        genes = []
        room_bookings = defaultdict(set)  # (day, start_time) -> set of room_ids
        faculty_bookings = defaultdict(set)  # (day, start_time) -> set of faculty_ids
        section_bookings = defaultdict(set)  # (day, start_time) -> set of (course_id, section)
        faculty_section_count = defaultdict(int)  # faculty_id -> number of sections assigned
        
        # Track faculty assignments for Core courses to ensure same faculty for both sessions
        core_course_faculty_assignments = {}  # (course_id, section) -> faculty_id
        
        for base_gene in base_genes:
            try:
                # Get course details
                course = next((c for c in self.courses if c['Course_ID'] == base_gene.CourseID), None)
                if not course:
                    continue
                
                # Get duration type for this course
                duration_type = self._get_duration_type(course['Duration_Minutes'])
                
                # Check if this is a Core course that already has a faculty assigned
                course_section_key = (base_gene.CourseID, base_gene.Section)
                if course['Course_Type'] == "Core" and course_section_key in core_course_faculty_assignments:
                    # Use the same faculty as the first session
                    assigned_faculty_id = core_course_faculty_assignments[course_section_key]
                    faculty = next((f for f in self.faculty if f['Faculty_ID'] == assigned_faculty_id), None)
                    if not faculty:
                        continue
                else:
                    # Get available faculty for this course
                    faculty_ratings = self._get_faculty_ratings_for_course(base_gene.CourseID, relaxed=False)
                    available_faculty = [
                        f for f, rating in faculty_ratings 
                        if base_gene.CourseID in f.get('Processed_Courses', []) and 
                        faculty_section_count.get(f['Faculty_ID'], 0) < 3
                    ]
                    
                    if not available_faculty:
                        # Try relaxed mode
                        available_faculty = [
                            f for f in self.faculty 
                            if faculty_section_count.get(f['Faculty_ID'], 0) < 3
                        ]
                    
                    if not available_faculty:
                        continue
                    
                    # Select faculty (prefer higher rated)
                    faculty = random.choice(available_faculty)
                    
                    # For Core courses, store the faculty assignment for the second session
                    if course['Course_Type'] == "Core":
                        core_course_faculty_assignments[course_section_key] = faculty['Faculty_ID']
                
                faculty_section_count[faculty['Faculty_ID']] += 1
                
                # Get suitable rooms
                room_type = "Lab" if course['Course_Type'] == "Lab" else "Lecture"
                suitable_rooms = [
                    r for r in self.rooms 
                    if r['Room_Capacity'] >= course['Capacity'] and r['Room_Type'] == room_type
                ]
                
                if not suitable_rooms:
                    continue
                
                # Get time slots matching duration type
                available_timeslots = [
                    ts for ts in self.time_slots 
                    if ts['Duration_Type'] == duration_type
                ]
                
                if not available_timeslots:
                    continue
                
                # Try to find a valid assignment
                assigned = False
                for _ in range(50):  # Try up to 50 times
                    time_slot = random.choice(available_timeslots)
                    room = random.choice(suitable_rooms)
                    
                    time_key = (time_slot['Day'], time_slot['Start_Time'])
                    
                    # Check conflicts
                    if (room['Room_ID'] in room_bookings[time_key] or
                        faculty['Faculty_ID'] in faculty_bookings[time_key] or
                        (base_gene.CourseID, base_gene.Section) in section_bookings[time_key]):
                        continue
                    
                    # Create gene
                    gene = Gene(
                        CourseID=base_gene.CourseID,
                        Section=base_gene.Section,
                        FacultyID=faculty['Faculty_ID'],
                        RoomID=room['Room_ID'],
                        Day=time_slot['Day'],
                        StartTime=time_slot['Start_Time'],
                        EndTime=time_slot['End_Time']
                    )
                    
                    # Update bookings
                    room_bookings[time_key].add(room['Room_ID'])
                    faculty_bookings[time_key].add(faculty['Faculty_ID'])
                    section_bookings[time_key].add((base_gene.CourseID, base_gene.Section))
                    
                    genes.append(gene)
                    assigned = True
                    break
                
                if not assigned:
                    # Create gene with random assignment (will be penalized)
                    time_slot = random.choice(available_timeslots)
                    room = random.choice(suitable_rooms)
                    
                    gene = Gene(
                        CourseID=base_gene.CourseID,
                        Section=base_gene.Section,
                        FacultyID=faculty['Faculty_ID'],
                        RoomID=room['Room_ID'],
                        Day=time_slot['Day'],
                        StartTime=time_slot['Start_Time'],
                        EndTime=time_slot['End_Time']
                    )
                    genes.append(gene)
                    
                    # Update bookings even for conflicted assignments
                    time_key = (time_slot['Day'], time_slot['Start_Time'])
                    room_bookings[time_key].add(room['Room_ID'])
                    faculty_bookings[time_key].add(faculty['Faculty_ID'])
                    section_bookings[time_key].add((base_gene.CourseID, base_gene.Section))
                    
            except Exception as e:
                # Create a placeholder gene to maintain chromosome length consistency
                # This will be heavily penalized in fitness calculation
                gene = Gene(
                    CourseID=base_gene.CourseID,
                    Section=base_gene.Section,
                    FacultyID=1,  # Default faculty ID
                    RoomID="R1",  # Default room ID
                    Day="Monday", # Default day
                    StartTime=8.0, # Default start time
                    EndTime=9.0    # Default end time
                )
                genes.append(gene)
        
        return Chromosome(genes=genes)
    
    def step4_calculate_fitness(self, chromosome: Chromosome) -> float:
        """Step 4: Calculate fitness function"""
        total_penalty = 0
        
        # Hard constraint penalties
        hard_penalties = self._calculate_hard_constraint_penalties(chromosome)
        total_penalty += hard_penalties * self.hard_constraint_weight
        
        # Soft constraint penalties
        if not self.skip_soft_constraints:
            soft_penalties = self._calculate_soft_constraint_penalties(chromosome)
            total_penalty += soft_penalties * self.soft_constraint_weight
        
        # Calculate fitness (higher is better)
        fitness = 1000 - total_penalty
        
        # Reject chromosomes with too many hard constraint violations
        if hard_penalties > 50:  # Threshold for rejection
            fitness = 0
        
        chromosome.fitness = max(0, fitness)
        return chromosome.fitness
    
    def _calculate_hard_constraint_penalties(self, chromosome: Chromosome) -> int:
        """Calculate hard constraint violations (matching CSP3.py constraints)"""
        penalties = 0
        # Track bookings for conflict detection
        room_bookings = defaultdict(set)  # (day, start_time) -> set of room_ids
        faculty_bookings = defaultdict(set)  # (day, start_time) -> set of faculty_ids
        section_bookings = defaultdict(set)  # (day, start_time) -> set of (course_id, section)
        faculty_section_count = defaultdict(int)  # faculty_id -> number of sections assigned
        
        # Track faculty day times for continuous window constraint
        faculty_day_times = defaultdict(list)  # (faculty_id, day) -> list of (start_time, end_time)
        
        # Track lecture start times for fixed time constraint
        lecture_start_times = defaultdict(set)  # (faculty_id, section) -> set of start_times
        
        # Track faculty assignments for Core courses to check same faculty constraint
        core_course_faculty_assignments = defaultdict(set)  # (course_id, section) -> set of faculty_ids
        
        for gene in chromosome.genes:
            try:
                # Get course details
                course = next((c for c in self.courses if c['Course_ID'] == gene.CourseID), None)
                room = next((r for r in self.rooms if r['Room_ID'] == gene.RoomID), None)
                faculty = next((f for f in self.faculty if f['Faculty_ID'] == gene.FacultyID), None)
                if not course or not room or not faculty:
                    penalties += 100  # Heavy penalty for missing data
                    continue
                
                # Faculty can only be assigned to their Processed_Courses
                if gene.CourseID not in faculty.get('Processed_Courses', []):
                    penalties += 100
                
                # Track number of sections assigned to this faculty
                faculty_section_count[gene.FacultyID] += 1
                
                # Check room type mismatch
                if course['Course_Type'] == "Lab" and room['Room_Type'] != "Lab":
                    penalties += 100
                elif course['Course_Type'] != "Lab" and room['Room_Type'] != "Lecture":
                    penalties += 100
                
                # Check room capacity
                if room['Room_Capacity'] < course['Capacity']:
                    penalties += 100
                
                # Check faculty assignment
                if gene.CourseID not in faculty['Courses_Assigned']:
                    penalties += 100
                
                # Check time slot bounds
                if gene.StartTime < 8.0 or gene.EndTime > 18.0:
                    penalties += 100
                
                # Check course duration fit
                expected_duration = course['Duration_Minutes'] / 60.0
                actual_duration = gene.EndTime - gene.StartTime
                if abs(actual_duration - expected_duration) > 0.1:  # Allow small tolerance
                    penalties += 100
                
                # Check double bookings
                time_key = (gene.Day, gene.StartTime)
                
                # Room double booking
                if gene.RoomID in room_bookings[time_key]:
                    penalties += 100
                room_bookings[time_key].add(gene.RoomID)
                
                # Faculty double booking
                if gene.FacultyID in faculty_bookings[time_key]:
                    penalties += 100
                faculty_bookings[time_key].add(gene.FacultyID)
                
                # Section double booking
                section_key = (gene.CourseID, gene.Section)
                if section_key in section_bookings[time_key]:
                    penalties += 100
                section_bookings[time_key].add(section_key)
                
                # Track faculty day times for continuous window constraint
                faculty_day_times[(gene.FacultyID, gene.Day)].append((gene.StartTime, gene.EndTime))
                
                # Track lecture start times for fixed time constraint
                if course['Course_Type'] != "Lab":  # Only for lectures
                    lecture_start_times[(gene.FacultyID, gene.Section)].add(gene.StartTime)
                
                # Track faculty assignments for Core courses
                if course['Course_Type'] == "Core":
                    core_course_faculty_assignments[(gene.CourseID, gene.Section)].add(gene.FacultyID)
                
            except Exception as e:
                penalties += 50  # Penalty for any errors
                continue
        
        # After all genes, check faculty section count
        for faculty_id, count in faculty_section_count.items():
            if count > 3:
                penalties += (count - 3) * 100  # Heavy penalty for each extra section
        
        # Check course frequency violation
        course_sessions = defaultdict(int)
        for gene in chromosome.genes:
            course_sessions[gene.CourseID] += 1
        for course in self.courses:
            expected_sessions = course['Weekdays']
            actual_sessions = course_sessions[course['Course_ID']]
            if actual_sessions != expected_sessions:
                penalties += abs(actual_sessions - expected_sessions) * 100
        
        # Check faculty continuous window constraint (CSP3.py style)
        for (faculty_id, day), times in faculty_day_times.items():
            if len(times) > 1:
                start_minutes = [s * 60 for s, _ in times]
                end_minutes = [e * 60 for _, e in times]
                earliest_start = min(start_minutes)
                latest_end = max(end_minutes)
                
                for start_time, end_time in times:
                    s_min = start_time * 60
                    e_min = end_time * 60
                    if s_min < earliest_start or e_min > latest_end:
                        penalties += 50  # Soft constraint penalty
        
        # Check lecture fixed time across week constraint (CSP3.py style)
        for (faculty_id, section), start_times in lecture_start_times.items():
            if len(start_times) > 1:
                penalties += (len(start_times) - 1) * 50  # Soft constraint penalty
        
        # Check same faculty constraint for Core courses (both lecture sessions must have same faculty)
        for (course_id, section), faculty_ids in core_course_faculty_assignments.items():
            if len(faculty_ids) > 1:
                penalties += (len(faculty_ids) - 1) * 100  # Heavy penalty for different faculty assignments
        
        return penalties
    
    def _track_conflicts(self, gene: Gene, reason: str):
        """Track conflicts for reporting (matching CSP3.py approach)"""
        course_name = next((c['Course_Name'] for c in self.courses if c['Course_ID'] == gene.CourseID), gene.CourseID)
        self.all_conflict_reasons[(course_name, gene.Section)].add(reason)
    
    def _get_conflict_summary(self) -> List[Dict]:
        """Get conflict summary for unassigned courses (matching CSP3.py)"""
        if self.best_chromosome is None:
            return []
        
        assigned_keys = set((gene.CourseID, gene.Section) for gene in self.best_chromosome.genes)
        final_conflicts = []
        
        for (cname, section), reasons in self.all_conflict_reasons.items():
            if (cname, section) not in assigned_keys:
                final_conflicts.append({
                    "CourseName": cname,
                    "Section": section,
                    "Reason": "; ".join(sorted(reasons))
                })
        
        return final_conflicts
    
    def _calculate_soft_constraint_penalties(self, chromosome: Chromosome) -> int:
        """Calculate soft constraint violations"""
        penalties = 0
        
        # Check faculty ratings
        for gene in chromosome.genes:
            try:
                faculty = next((f for f in self.faculty if f['Faculty_ID'] == gene.FacultyID), None)
                if faculty:
                    rating = faculty['Courses_Assigned'].get(gene.CourseID, 1)
                    if rating < 3:  # Penalize low-rated faculty
                        penalties += (4 - rating)
            except:
                continue
        
        # Check faculty time window spread
        faculty_day_times = defaultdict(list)
        for gene in chromosome.genes:
            faculty_day_times[(gene.FacultyID, gene.Day)].append(gene.StartTime)
        
        for (faculty_id, day), times in faculty_day_times.items():
            if len(times) > 1:
                time_spread = max(times) - min(times)
                if time_spread > 4:  # Penalize spread-out schedules
                    penalties += int(time_spread - 4)
        
        # Check consistent start times for lectures
        course_start_times = defaultdict(set)
        for gene in chromosome.genes:
            try:
                course = next((c for c in self.courses if c['Course_ID'] == gene.CourseID), None)
                if course and course['Course_Type'] != "Lab":  # Only for lectures
                    course_start_times[gene.CourseID].add(gene.StartTime)
            except:
                continue
        
        for course_id, start_times in course_start_times.items():
            if len(start_times) > 1:
                penalties += len(start_times) - 1
        
        return penalties
    
    def step5_selection(self) -> List[Chromosome]:
        """Step 5: Selection using tournament selection"""
        parents = []
        
        for _ in range(self.population_size // 2):
            # Tournament selection
            tournament = random.sample(self.population, self.tournament_size)
            winner = max(tournament, key=lambda c: c.fitness)
            parents.append(winner)
        
        return parents
    
    def step6_crossover(self, parents: List[Chromosome]) -> List[Chromosome]:
        """Step 6: Crossover - gene-level crossover"""
        children = []
        
        for i in range(0, len(parents), 2):
            if i + 1 < len(parents):
                parent1 = parents[i]
                parent2 = parents[i + 1]
                
                # Create two children
                child1 = self._perform_crossover(parent1, parent2)
                child2 = self._perform_crossover(parent2, parent1)
                
                children.extend([child1, child2])
        
        return children
    
    def _perform_crossover(self, parent1: Chromosome, parent2: Chromosome) -> Chromosome:
        """Perform crossover between two parents"""
        # Alternate genes from parents
        child_genes = []
        
        # Use the minimum length to avoid index out of range
        min_length = min(len(parent1.genes), len(parent2.genes))
        
        for i in range(min_length):
            if i % 2 == 0:
                child_genes.append(parent1.genes[i])
            else:
                child_genes.append(parent2.genes[i])
        
        # If one parent has more genes, add them from the longer parent
        if len(parent1.genes) > min_length:
            child_genes.extend(parent1.genes[min_length:])
        elif len(parent2.genes) > min_length:
            child_genes.extend(parent2.genes[min_length:])
        
        # Validate and fix conflicts
        child = Chromosome(genes=child_genes)
        self._fix_conflicts(child)
        
        return child
    
    def _fix_conflicts(self, chromosome: Chromosome):
        """Fix conflicts in a chromosome"""
        # Track bookings
        room_bookings = defaultdict(set)
        faculty_bookings = defaultdict(set)
        section_bookings = defaultdict(set)
        
        # Track faculty assignments for Core courses to maintain same faculty constraint
        core_course_faculty_assignments = {}  # (course_id, section) -> faculty_id
        
        for gene in chromosome.genes:
            time_key = (gene.Day, gene.StartTime)
            
            # Get course details
            course = next((c for c in self.courses if c['Course_ID'] == gene.CourseID), None)
            if not course:
                continue
            
            # Check if this is a Core course that should have same faculty as previous session
            course_section_key = (gene.CourseID, gene.Section)
            if course['Course_Type'] == "Core" and course_section_key in core_course_faculty_assignments:
                expected_faculty_id = core_course_faculty_assignments[course_section_key]
                if gene.FacultyID != expected_faculty_id:
                    # Fix: assign the same faculty as the first session
                    gene.FacultyID = expected_faculty_id
            
            # Check for conflicts
            has_conflict = False
            
            if gene.RoomID in room_bookings[time_key]:
                has_conflict = True
            if gene.FacultyID in faculty_bookings[time_key]:
                has_conflict = True
            if (gene.CourseID, gene.Section) in section_bookings[time_key]:
                has_conflict = True
            
            if has_conflict:
                # Try to find alternative time slot
                self._find_alternative_slot(gene, room_bookings, faculty_bookings, section_bookings)
            
            # Update bookings
            room_bookings[time_key].add(gene.RoomID)
            faculty_bookings[time_key].add(gene.FacultyID)
            section_bookings[time_key].add((gene.CourseID, gene.Section))
            
            # Store faculty assignment for Core courses
            if course['Course_Type'] == "Core":
                core_course_faculty_assignments[course_section_key] = gene.FacultyID
    
    def _find_alternative_slot(self, gene: Gene, room_bookings: dict, faculty_bookings: dict, section_bookings: dict):
        """Find alternative time slot for a gene"""
        course = next(c for c in self.courses if c['Course_ID'] == gene.CourseID)
        faculty = next(f for f in self.faculty if f['Faculty_ID'] == gene.FacultyID)
        
        # Try different time slots
        for time_slot in self.time_slots:
            if time_slot['Day'] not in faculty['Available_Days']:
                continue
            
            time_key = (time_slot['Day'], time_slot['Start_Time'])
            
            # Check if this slot is available
            if (gene.RoomID not in room_bookings[time_key] and
                gene.FacultyID not in faculty_bookings[time_key] and
                (gene.CourseID, gene.Section) not in section_bookings[time_key]):
                
                # Update gene
                gene.Day = time_slot['Day']
                gene.StartTime = time_slot['Start_Time']
                gene.EndTime = gene.StartTime + (course['Duration_Minutes'] / 60.0)
                return
        
        # If no alternative found, keep original (will be penalized in fitness)
        pass
    
    def step7_mutation(self, chromosomes: List[Chromosome]):
        """Step 7: Mutation - randomly mutate genes"""
        for chromosome in chromosomes:
            for gene in chromosome.genes:
                if random.random() < self.mutation_rate:
                    self._mutate_gene(gene)
    
    def _mutate_gene(self, gene: Gene):
        """Mutate a single gene"""
        mutation_type = random.choice(['time', 'room', 'faculty'])
        
        course = next(c for c in self.courses if c['Course_ID'] == gene.CourseID)
        
        if mutation_type == 'time':
            # Assign new time slot
            time_slot = random.choice(self.time_slots)
            gene.Day = time_slot['Day']
            gene.StartTime = time_slot['Start_Time']
            gene.EndTime = gene.StartTime + (course['Duration_Minutes'] / 60.0)
        
        elif mutation_type == 'room':
            # Assign new room
            room_type = "Lab" if course['Course_Type'] == "Lab" else "Lecture"
            eligible_rooms = [
                r for r in self.rooms 
                if r['Room_Capacity'] >= course['Capacity'] and 
                r['Room_Type'] == room_type
            ]
            if eligible_rooms:
                room = random.choice(eligible_rooms)
                gene.RoomID = room['Room_ID']
        
        elif mutation_type == 'faculty':
            # For Core courses, we need to be careful about faculty mutation
            # to maintain the same faculty constraint for both sessions
            course = next((c for c in self.courses if c['Course_ID'] == gene.CourseID), None)
            if course and course['Course_Type'] == "Core":
                # For Core courses, we should avoid faculty mutation to maintain constraint
                # Instead, mutate time or room
                if random.random() < 0.5:
                    # Assign new time slot
                    time_slot = random.choice(self.time_slots)
                    gene.Day = time_slot['Day']
                    gene.StartTime = time_slot['Start_Time']
                    gene.EndTime = gene.StartTime + (course['Duration_Minutes'] / 60.0)
                else:
                    # Assign new room
                    room_type = "Lab" if course['Course_Type'] == "Lab" else "Lecture"
                    eligible_rooms = [
                        r for r in self.rooms 
                        if r['Room_Capacity'] >= course['Capacity'] and 
                        r['Room_Type'] == room_type
                    ]
                    if eligible_rooms:
                        room = random.choice(eligible_rooms)
                        gene.RoomID = room['Room_ID']
            else:
                # For Lab courses, faculty mutation is allowed
                eligible_faculty = [
                    f for f in self.faculty 
                    if gene.CourseID in f.get('Processed_Courses', [])
                ]
                if eligible_faculty:
                    faculty = random.choice(eligible_faculty)
                    gene.FacultyID = faculty['Faculty_ID']
    
    def step8_local_search(self, chromosomes: List[Chromosome]):
        """Step 8: Local search - improve chromosomes"""
        for chromosome in chromosomes:
            self._apply_local_search(chromosome)
    
    def _apply_local_search(self, chromosome: Chromosome):
        """Apply local search to improve a chromosome"""
        improved = True
        max_iterations = 10
        
        while improved and max_iterations > 0:
            improved = False
            max_iterations -= 1
            
            # Try swapping time slots
            for i in range(len(chromosome.genes)):
                for j in range(i + 1, len(chromosome.genes)):
                    gene1 = chromosome.genes[i]
                    gene2 = chromosome.genes[j]
                    
                    # Check if swap is valid
                    if self._is_swap_valid(gene1, gene2):
                        # Calculate fitness before swap
                        old_fitness = chromosome.fitness
                        
                        # Perform swap
                        gene1.Day, gene2.Day = gene2.Day, gene1.Day
                        gene1.StartTime, gene2.StartTime = gene2.StartTime, gene1.StartTime
                        gene1.EndTime, gene2.EndTime = gene2.EndTime, gene1.EndTime
                        
                        # Recalculate fitness
                        new_fitness = self.step4_calculate_fitness(chromosome)
                        
                        # Keep swap if it improves fitness
                        if new_fitness > old_fitness:
                            improved = True
                        else:
                            # Revert swap
                            gene1.Day, gene2.Day = gene2.Day, gene1.Day
                            gene1.StartTime, gene2.StartTime = gene2.StartTime, gene1.StartTime
                            gene1.EndTime, gene2.EndTime = gene2.EndTime, gene1.EndTime
                            chromosome.fitness = old_fitness
    
    def _is_swap_valid(self, gene1: Gene, gene2: Gene) -> bool:
        """Check if swapping two genes is valid"""
        # Check if both genes are from the same faculty
        if gene1.FacultyID == gene2.FacultyID:
            return False
        
        # Check if both genes use the same room
        if gene1.RoomID == gene2.RoomID:
            return False
        
        # Check if both genes are from the same section
        if gene1.CourseID == gene2.CourseID and gene1.Section == gene2.Section:
            return False
        
        return True
    
    def _validate_chromosome_structure(self, chromosome: Chromosome, expected_length: int) -> bool:
        """Validate that chromosome has the expected structure"""
        if len(chromosome.genes) != expected_length:
            print(f"⚠️ Chromosome length mismatch: expected {expected_length}, got {len(chromosome.genes)}")
            return False
        return True
    
    def _print_chromosome_stats(self, population: List[Chromosome], generation: int = 0):
        """Print statistics about chromosome lengths in population"""
        lengths = [len(chromosome.genes) for chromosome in population]
        unique_lengths = set(lengths)
        print(f"Generation {generation}: Chromosome lengths - {unique_lengths}")
        if len(unique_lengths) > 1:
            print(f"⚠️ Multiple chromosome lengths detected: {unique_lengths}")
            for length in unique_lengths:
                count = lengths.count(length)
                print(f"  Length {length}: {count} chromosomes")
    
    def step9_evolution_loop(self):
        """Step 9: Main evolution loop"""
        print("Step 9: Starting Evolution Loop...")
        
        start_time = time.time()
        best_fitness_history = []
        
        # Get expected chromosome length from base structure
        base_genes = self.step2_create_chromosome_structure()
        expected_length = len(base_genes)
        
        # Validate and fix initial population structure
        print("Validating initial population structure...")
        self._print_chromosome_stats(self.population, 0)
        for i, chromosome in enumerate(self.population):
            if not self._validate_chromosome_structure(chromosome, expected_length):
                print(f"Fixing chromosome {i} structure...")
                chromosome.genes = self._create_random_chromosome(base_genes).genes
        
        # Calculate initial fitness for all chromosomes
        print("Calculating initial fitness...")
        for i, chromosome in enumerate(self.population):
            try:
                self.step4_calculate_fitness(chromosome)
                if i % 10 == 0:  # Progress indicator
                    print(f"  Processed {i+1}/{len(self.population)} chromosomes")
            except Exception as e:
                print(f"Error calculating fitness for chromosome {i}: {e}")
                chromosome.fitness = 0
        
        print("Initial fitness calculation complete")
        
        for generation in tqdm(range(self.generations), desc="Evolution"):
            # Check timeout
            if time.time() - start_time > self.timeout_seconds:
                print(f"⚠️ Timeout reached after {generation} generations")
                break
            
            # Sort population by fitness
            self.population.sort(key=lambda c: c.fitness, reverse=True)
            
            # Update best chromosome
            if self.best_chromosome is None or self.population[0].fitness > self.best_chromosome.fitness:
                self.best_chromosome = Chromosome(
                    genes=[Gene(**vars(gene)) for gene in self.population[0].genes],
                    fitness=self.population[0].fitness
                )
            
            # Track best fitness
            best_fitness_history.append(self.best_chromosome.fitness)
            
            # Print progress every 10 generations
            if generation % 10 == 0:
                print(f"Generation {generation}: Best fitness = {self.best_chromosome.fitness:.2f}")
            
            # Check for convergence
            if len(best_fitness_history) > 20:
                recent_improvement = max(best_fitness_history[-20:]) - min(best_fitness_history[-20:])
                if recent_improvement < 1.0:
                    print(f"✓ Convergence reached after {generation} generations")
                    break
            
            # Selection
            parents = self.step5_selection()
            
            # Crossover
            try:
                children = self.step6_crossover(parents)
                
                # Validate children structure
                for child in children:
                    if not self._validate_chromosome_structure(child, expected_length):
                        # Fix chromosome by recreating it
                        child.genes = self._create_random_chromosome(base_genes).genes
            except Exception as e:
                print(f"Error in crossover: {e}")
                # Create new random children as fallback
                children = [self._create_random_chromosome(base_genes) for _ in range(len(parents))]
            
            # Print chromosome stats after crossover
            if generation % 10 == 0:
                self._print_chromosome_stats(children, generation)
            
            # Mutation
            self.step7_mutation(children)
            
            # Local search (skip for now to avoid performance issues)
            # self.step8_local_search(children)
            
            # Calculate fitness for children
            for chromosome in children:
                try:
                    self.step4_calculate_fitness(chromosome)
                except Exception as e:
                    print(f"Error calculating fitness for child: {e}")
                    chromosome.fitness = 0
            
            # Elitism: keep best 10% of parents
            elite_count = max(1, self.population_size // 10)
            elite = self.population[:elite_count]
            
            # Create new population
            new_population = elite + children[:self.population_size - elite_count]
            
            # Ensure population size
            while len(new_population) < self.population_size:
                new_population.append(random.choice(self.population))
            
            # Final validation of new population
            for i, chromosome in enumerate(new_population):
                if not self._validate_chromosome_structure(chromosome, expected_length):
                    print(f"Final fix for chromosome {i} in generation {generation}")
                    chromosome.genes = self._create_random_chromosome(base_genes).genes
            
            self.population = new_population[:self.population_size]
        
        print(f"✓ Evolution completed in {time.time() - start_time:.2f} seconds")
        if self.best_chromosome:
            print(f"✓ Best fitness achieved: {self.best_chromosome.fitness:.2f}")
        else:
            print("✓ No valid solution found")
    
    def step10_export_output(self, output_file: str):
        """Step 10: Export output to Excel (matching CSP3.py formatting)"""
        print("Step 10: Exporting Output...")
        
        if self.best_chromosome is None:
            print("❌ No valid solution found")
            return
        
        # Handle file permission issues by adding timestamp if file exists
        import os
        import time
        base_name, ext = os.path.splitext(output_file)
        if os.path.exists(output_file):
            timestamp = time.strftime("%Y%m%d_%H%M%S")
            output_file = f"{base_name}_{timestamp}{ext}"
            print(f"⚠️  Original file exists, using: {output_file}")
        
        # Create timetable data
        timetable_data = []
        for gene in self.best_chromosome.genes:
            try:
                course = next(c for c in self.courses if c['Course_ID'] == gene.CourseID)
                faculty = next(f for f in self.faculty if f['Faculty_ID'] == gene.FacultyID)
                room = next(r for r in self.rooms if r['Room_ID'] == gene.RoomID)
                
                timetable_data.append({
                    'CourseID': gene.CourseID,
                    'CourseName': course['Course_Name'],
                    'CourseType': course['Course_Type'],
                    'Section': gene.Section,
                    'FacultyID': gene.FacultyID,
                    'FacultyName': faculty['Faculty_Name'],
                    'Day': gene.Day,
                    'StartTime': f"{int(gene.StartTime):02d}:{int((gene.StartTime % 1) * 60):02d}",
                    'EndTime': f"{int(gene.EndTime):02d}:{int((gene.EndTime % 1) * 60):02d}",
                    'Room': gene.RoomID,
                    'RoomType': room['Room_Type'],
                    'Rating': faculty['Courses_Assigned'].get(gene.CourseID, 1)
                })
            except Exception as e:
                print(f"Error processing gene: {e}")
                continue
        
        # Create unassigned courses data
        unassigned_data = []
        assigned_courses = set((gene.CourseID, gene.Section) for gene in self.best_chromosome.genes)
        
        # Get conflict summary for better reporting
        conflict_summary = self._get_conflict_summary()
        
        for course in self.courses:
            course_key = (course['Course_ID'], course['Section'])
            if course_key not in assigned_courses:
                # Find specific reason from conflict tracking
                course_name = course['Course_Name']
                section = course['Section']
                reason = "Could not be scheduled due to constraints"
                
                # Look for specific conflict reason
                for conflict in conflict_summary:
                    if conflict['CourseName'] == course_name and conflict['Section'] == section:
                        reason = conflict['Reason']
                        break
                
                unassigned_data.append({
                    'CourseName': course_name,
                    'Section': section,
                    'Reason': reason
                })
        
        # Create Excel file with CSP3.py style formatting
        print(f"\nSaving timetable to Excel: {output_file}")
        try:
            with pd.ExcelWriter(output_file, engine='xlsxwriter') as writer:
                workbook = writer.book
                worksheet = workbook.add_worksheet("Timetable")
                writer.sheets["Timetable"] = worksheet
                
                # Define formats (matching CSP3.py)
                bold_format = workbook.add_format({"bold": True, "font_size": 14, "align": "center"})
                header_format = workbook.add_format({"bold": True, "bg_color": "#D9E1F2", "border": 1})
                
                if timetable_data:
                    timetable_df = pd.DataFrame(timetable_data)
                    
                    # Sort by section number (matching CSP3.py)
                    def extract_section_number(section):
                        match = re.search(r'V(\d+)', section)
                        return int(match.group(1)) if match else float('inf')
                    
                    timetable_df = timetable_df.sort_values('Section', key=lambda x: x.apply(extract_section_number))
                    
                    # Write data grouped by sections (matching CSP3.py)
                    row = 0
                    for section in sorted(timetable_df["Section"].unique(), key=extract_section_number):
                        section_df = timetable_df[timetable_df["Section"] == section]
                        
                        # Write section header
                        worksheet.merge_range(row, 0, row, len(section_df.columns) - 1, f"Section-{section} Courses", bold_format)
                        row += 1
                        
                        # Write column headers
                        for col_num, value in enumerate(section_df.columns):
                            worksheet.write(row, col_num, value, header_format)
                        row += 1
                        
                        # Write data
                        for record in section_df.itertuples(index=False):
                            for col_num, value in enumerate(record):
                                worksheet.write(row, col_num, value)
                            row += 1
                        
                        row += 2  # Add space before next section
                    
                    # Set column widths
                    for col_num, column in enumerate(timetable_df.columns):
                        max_length = max(
                            len(str(column)),
                            timetable_df[column].astype(str).str.len().max()
                        )
                        worksheet.set_column(col_num, col_num, max_length + 2)
                
                # Unassigned courses sheet
                if unassigned_data:
                    unassigned_df = pd.DataFrame(unassigned_data)
                    unassigned_df.to_excel(writer, sheet_name='Unassigned Courses', index=False)
        
        except PermissionError as e:
            print(f"❌ Permission Error: {e}")
            print("💡 Solutions:")
            print("   1. Close the Excel file if it's open in another application")
            print("   2. Check if you have write permissions in the current directory")
            print("   3. Try running the program as administrator")
            print("   4. The program will try to create a new file with timestamp")
            
            # Try with a different filename
            timestamp = time.strftime("%Y%m%d_%H%M%S")
            new_output_file = f"genetic_timetable_output_{timestamp}.xlsx"
            print(f"🔄 Trying alternative filename: {new_output_file}")
            
            try:
                with pd.ExcelWriter(new_output_file, engine='xlsxwriter') as writer:
                    # Recreate the Excel content
                    workbook = writer.book
                    worksheet = workbook.add_worksheet("Timetable")
                    writer.sheets["Timetable"] = worksheet
                    
                    # Define formats
                    bold_format = workbook.add_format({"bold": True, "font_size": 14, "align": "center"})
                    header_format = workbook.add_format({"bold": True, "bg_color": "#D9E1F2", "border": 1})
                    
                    if timetable_data:
                        timetable_df = pd.DataFrame(timetable_data)
                        
                        # Sort by section number
                        def extract_section_number(section):
                            match = re.search(r'V(\d+)', section)
                            return int(match.group(1)) if match else float('inf')
                        
                        timetable_df = timetable_df.sort_values('Section', key=lambda x: x.apply(extract_section_number))
                        
                        # Write data grouped by sections
                        row = 0
                        for section in sorted(timetable_df["Section"].unique(), key=extract_section_number):
                            section_df = timetable_df[timetable_df["Section"] == section]
                            
                            # Write section header
                            worksheet.merge_range(row, 0, row, len(section_df.columns) - 1, f"Section-{section} Courses", bold_format)
                            row += 1
                            
                            # Write column headers
                            for col_num, value in enumerate(section_df.columns):
                                worksheet.write(row, col_num, value, header_format)
                            row += 1
                            
                            # Write data
                            for record in section_df.itertuples(index=False):
                                for col_num, value in enumerate(record):
                                    worksheet.write(row, col_num, value)
                                row += 1
                            
                            row += 2  # Add space before next section
                        
                        # Set column widths
                        for col_num, column in enumerate(timetable_df.columns):
                            max_length = max(
                                len(str(column)),
                                timetable_df[column].astype(str).str.len().max()
                            )
                            worksheet.set_column(col_num, col_num, max_length + 2)
                    
                    # Unassigned courses sheet
                    if unassigned_data:
                        unassigned_df = pd.DataFrame(unassigned_data)
                        unassigned_df.to_excel(writer, sheet_name='Unassigned Courses', index=False)
                
                print(f"✓ Output exported to {new_output_file}")
                output_file = new_output_file
                
            except Exception as e2:
                print(f"❌ Failed to create alternative file: {e2}")
                print("💡 Please close any open Excel files and try again")
                return
        except Exception as e: # Catch any other unexpected errors during export
            print(f"❌ An unexpected error occurred during export: {e}")
            return

        print(f"✓ Output exported to {output_file}")
        print(f"✓ Scheduled {len(timetable_data)} sessions")
        print(f"✓ {len(unassigned_data)} courses could not be scheduled")
        
        # Calculate and display accuracy metrics (matching CSP3.py)
        total_course_sections = len(self.courses)
        unique_scheduled_course_sections = len(set((gene.CourseID, gene.Section) for gene in self.best_chromosome.genes))
        unscheduled_course_sections = total_course_sections - unique_scheduled_course_sections
        accuracy = (unique_scheduled_course_sections / total_course_sections) * 100 if total_course_sections > 0 else 0
        
        print(f"\n📊 Accuracy Metrics:")
        print(f"Scheduled Course-Sections: {unique_scheduled_course_sections}/{total_course_sections}")
        print(f"Unscheduled Course-Sections: {unscheduled_course_sections}")
        print(f"Accuracy: {accuracy:.2f}%")
    
    def run(self, output_file: str = "genetic_timetable_output_new.xlsx"):
        """Run the complete genetic algorithm"""
        print("🧬 Starting Genetic Algorithm Timetable Generator")
        print("=" * 50)
        
        import time
        start_time = time.time()
        
        # Step 1: Preprocess inputs
        self.step1_preprocess_inputs()
        
        # Step 3: Generate initial population
        self.step3_generate_initial_population()
        
        # Step 9: Evolution loop
        self.step9_evolution_loop()
        
        # Step 10: Export output
        self.step10_export_output(output_file)
        
        total_time = time.time() - start_time
        print(f"⏱️  Total time taken: {total_time:.2f} seconds")
        print("=" * 50)
        print("🎉 Genetic Algorithm Timetable Generation Complete!")

def main():
    """Main function to run the genetic algorithm"""
    # Example usage
    input_file = "f2025.xlsx"  # Change to your input file
    
    if not os.path.exists(input_file):
        print(f"❌ Input file {input_file} not found")
        return
    
    # Create generator with optimized parameters for better time complexity
    generator = GeneticTimetableGenerator(
        input_file=input_file,
        population_size=75,      # Reduced from 200 for speed
        generations=150,         # Reduced from 300 for speed
        mutation_rate=0.05,      # Increased from 0.02 for faster convergence
        tournament_size=3,       # Reduced from 5 for speed
        timeout_minutes=2,       # Reduced from 5 for speed
        skip_soft_constraints=False  # Can be set to True for even faster execution
    )
    
    # Run the algorithm
    generator.run("genetic_timetable_output_new.xlsx")

if __name__ == "__main__":
    main() 