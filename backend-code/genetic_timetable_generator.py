# type: ignore 
import pandas as pd
import numpy as np
import random
import time
import os
from typing import List, Dict, Tuple, Any
from dataclasses import dataclass
from collections import defaultdict
import re
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
        
        # Constraint weights (reduced for faster computation)
        self.hard_constraint_weight = 50  # Reduced from 100
        self.soft_constraint_weight = 5   # Reduced from 10
        
    def step1_preprocess_inputs(self):
        """Step 1: Preprocessing Inputs - Read and parse Excel sheets"""
        print("Step 1: Preprocessing Inputs...")
        
        # Read Excel sheets
        self.courses_df = pd.read_excel(self.input_file, sheet_name="Courses")
        self.faculty_df = pd.read_excel(self.input_file, sheet_name="Faculty")
        self.rooms_df = pd.read_excel(self.input_file, sheet_name="Rooms")
        self.timeslots_df = pd.read_excel(self.input_file, sheet_name="Time Slots")
        self.students_df = pd.read_excel(self.input_file, sheet_name="Students")
        
        # Parse duration to minutes
        self.courses_df['Duration_Minutes'] = self.courses_df['Duration'].apply(self._parse_duration)
        
        # Parse faculty availability
        self.faculty_df['Available_Days'] = self.faculty_df['Availability'].apply(self._parse_availability)
        
        # Parse courses assigned (convert string to dict)
        self.faculty_df['Courses_Assigned'] = self.faculty_df['Courses_Assigned'].apply(self._parse_courses_assigned)
        
        # Compute Processed_Courses and Processed_Ratings (top 3 by rating)
        def get_top3_courses(courses_assigned):
            if not courses_assigned:
                return [], {}
            sorted_courses = sorted(courses_assigned.items(), key=lambda x: x[1], reverse=True)
            top3 = sorted_courses[:3]
            return [c for c, _ in top3], {c: r for c, r in top3}
        self.faculty_df['Processed_Courses'] = self.faculty_df['Courses_Assigned'].apply(lambda d: get_top3_courses(d)[0])
        self.faculty_df['Processed_Ratings'] = self.faculty_df['Courses_Assigned'].apply(lambda d: get_top3_courses(d)[1])
        
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
        """Generate time slots based on available days and times"""
        self.time_slots = []
        
        # Get unique days from timeslots
        days = self.timeslots_df['Day'].unique()
        
        # Standard time blocks
        time_blocks = [
            {"start": 8.0, "end": 9.0},    # 8:00 AM - 9:00 AM
            {"start": 9.0, "end": 10.0},   # 9:00 AM - 10:00 AM
            {"start": 10.0, "end": 11.0},  # 10:00 AM - 11:00 AM
            {"start": 11.0, "end": 12.0},  # 11:00 AM - 12:00 PM
            {"start": 12.0, "end": 13.0},  # 12:00 PM - 1:00 PM
            {"start": 13.0, "end": 14.0},  # 1:00 PM - 2:00 PM
            {"start": 14.0, "end": 15.0},  # 2:00 PM - 3:00 PM
            {"start": 15.0, "end": 16.0},  # 3:00 PM - 4:00 PM
            {"start": 16.0, "end": 17.0},  # 4:00 PM - 5:00 PM
            {"start": 17.0, "end": 18.0},  # 5:00 PM - 6:00 PM
        ]
        
        for day in days:
            for block in time_blocks:
                self.time_slots.append({
                    'Day': day,
                    'StartTime': block['start'],
                    'EndTime': block['end']
                })
    
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
        """Create a random chromosome with valid assignments"""
        genes = []
        faculty_section_count = {}
        for base_gene in base_genes:
            try:
                # Find course details
                course = next((c for c in self.courses if c['Course_ID'] == base_gene.CourseID), None)
                if not course:
                    continue
                # Find eligible faculty (only those with this course in Processed_Courses and <3 sections assigned)
                eligible_faculty = [
                    f for f in self.faculty
                    if base_gene.CourseID in f.get('Processed_Courses', [])
                    and faculty_section_count.get(f['Faculty_ID'], 0) < 3
                ]
                if not eligible_faculty:
                    continue
                faculty = random.choice(eligible_faculty)
                # Track number of sections assigned to this faculty
                faculty_section_count[faculty['Faculty_ID']] = faculty_section_count.get(faculty['Faculty_ID'], 0) + 1
                # Find eligible rooms
                eligible_rooms = [
                    r for r in self.rooms
                    if r['Room_Capacity'] >= course['Capacity'] and
                    r['Room_Type'] == course['Course_Type']
                ]
                if not eligible_rooms:
                    eligible_rooms = self.rooms
                if not eligible_rooms:
                    continue
                room = random.choice(eligible_rooms)
                # Find eligible time slots
                eligible_slots = [
                    ts for ts in self.time_slots
                    if ts['Day'] in faculty['Available_Days']
                ]
                if not eligible_slots:
                    eligible_slots = self.time_slots
                if not eligible_slots:
                    continue
                time_slot = random.choice(eligible_slots)
                # Calculate end time based on duration
                duration_hours = course['Duration_Minutes'] / 60.0
                end_time = time_slot['StartTime'] + duration_hours
                # Create gene
                gene = Gene(
                    CourseID=base_gene.CourseID,
                    Section=base_gene.Section,
                    FacultyID=faculty['Faculty_ID'],
                    RoomID=room['Room_ID'],
                    Day=time_slot['Day'],
                    StartTime=time_slot['StartTime'],
                    EndTime=end_time
                )
                genes.append(gene)
            except Exception as e:
                continue
        chromosome = Chromosome(genes=genes)
        return chromosome
    
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
        """Calculate hard constraint violations"""
        penalties = 0
        # Track bookings for conflict detection
        room_bookings = defaultdict(set)  # (day, start_time) -> set of room_ids
        faculty_bookings = defaultdict(set)  # (day, start_time) -> set of faculty_ids
        section_bookings = defaultdict(set)  # (day, start_time) -> set of (course_id, section)
        faculty_section_count = defaultdict(int)  # faculty_id -> number of sections assigned
        for gene in chromosome.genes:
            try:
                # Get course details
                course = next((c for c in self.courses if c['Course_ID'] == gene.CourseID), None)
                room = next((r for r in self.rooms if r['Room_ID'] == gene.RoomID), None)
                faculty = next((f for f in self.faculty if f['Faculty_ID'] == gene.FacultyID), None)
                if not course or not room or not faculty:
                    penalties += 10  # Heavy penalty for missing data
                    continue
                # Faculty can only be assigned to their Processed_Courses
                if gene.CourseID not in faculty.get('Processed_Courses', []):
                    penalties += 10
                # Track number of sections assigned to this faculty
                faculty_section_count[gene.FacultyID] += 1
                # Check room type mismatch
                if course['Course_Type'] == "Lab" and room['Room_Type'] != "Lab":
                    penalties += 1
                elif course['Course_Type'] != "Lab" and room['Room_Type'] != "Lecture":
                    penalties += 1
                # Check room capacity
                if room['Room_Capacity'] < course['Capacity']:
                    penalties += 1
                # Check faculty assignment
                if gene.CourseID not in faculty['Courses_Assigned']:
                    penalties += 1
                # Check time slot bounds
                if gene.StartTime < 8.0 or gene.EndTime > 18.0:
                    penalties += 1
                # Check course duration fit
                expected_duration = course['Duration_Minutes'] / 60.0
                actual_duration = gene.EndTime - gene.StartTime
                if abs(actual_duration - expected_duration) > 0.1:  # Allow small tolerance
                    penalties += 1
                # Check double bookings
                time_key = (gene.Day, gene.StartTime)
                # Room double booking
                if gene.RoomID in room_bookings[time_key]:
                    penalties += 1
                room_bookings[time_key].add(gene.RoomID)
                # Faculty double booking
                if gene.FacultyID in faculty_bookings[time_key]:
                    penalties += 1
                faculty_bookings[time_key].add(gene.FacultyID)
                # Section double booking
                section_key = (gene.CourseID, gene.Section)
                if section_key in section_bookings[time_key]:
                    penalties += 1
                section_bookings[time_key].add(section_key)
            except Exception as e:
                penalties += 5  # Penalty for any errors
                continue
        # After all genes, check faculty section count
        for faculty_id, count in faculty_section_count.items():
            if count > 3:
                penalties += (count - 3) * 10  # Heavy penalty for each extra section
        # Check course frequency violation
        course_sessions = defaultdict(int)
        for gene in chromosome.genes:
            course_sessions[gene.CourseID] += 1
        for course in self.courses:
            expected_sessions = course['Weekdays']
            actual_sessions = course_sessions[course['Course_ID']]
            if actual_sessions != expected_sessions:
                penalties += abs(actual_sessions - expected_sessions)
        return penalties
    
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
        
        for i in range(len(parent1.genes)):
            if i % 2 == 0:
                child_genes.append(parent1.genes[i])
            else:
                child_genes.append(parent2.genes[i])
        
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
        
        for gene in chromosome.genes:
            time_key = (gene.Day, gene.StartTime)
            
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
    
    def _find_alternative_slot(self, gene: Gene, room_bookings: dict, faculty_bookings: dict, section_bookings: dict):
        """Find alternative time slot for a gene"""
        course = next(c for c in self.courses if c['Course_ID'] == gene.CourseID)
        faculty = next(f for f in self.faculty if f['Faculty_ID'] == gene.FacultyID)
        
        # Try different time slots
        for time_slot in self.time_slots:
            if time_slot['Day'] not in faculty['Available_Days']:
                continue
            
            time_key = (time_slot['Day'], time_slot['StartTime'])
            
            # Check if this slot is available
            if (gene.RoomID not in room_bookings[time_key] and
                gene.FacultyID not in faculty_bookings[time_key] and
                (gene.CourseID, gene.Section) not in section_bookings[time_key]):
                
                # Update gene
                gene.Day = time_slot['Day']
                gene.StartTime = time_slot['StartTime']
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
            gene.StartTime = time_slot['StartTime']
            gene.EndTime = gene.StartTime + (course['Duration_Minutes'] / 60.0)
        
        elif mutation_type == 'room':
            # Assign new room
            eligible_rooms = [
                r for r in self.rooms 
                if r['Room_Capacity'] >= course['Capacity'] and 
                r['Room_Type'] == course['Course_Type']
            ]
            if eligible_rooms:
                room = random.choice(eligible_rooms)
                gene.RoomID = room['Room_ID']
        
        elif mutation_type == 'faculty':
            # Assign new faculty
            eligible_faculty = [
                f for f in self.faculty 
                if gene.CourseID in f['Courses_Assigned']
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
    
    def step9_evolution_loop(self):
        """Step 9: Main evolution loop"""
        print("Step 9: Starting Evolution Loop...")
        
        start_time = time.time()
        best_fitness_history = []
        
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
            children = self.step6_crossover(parents)
            
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
            
            self.population = new_population[:self.population_size]
        
        print(f"✓ Evolution completed in {time.time() - start_time:.2f} seconds")
        if self.best_chromosome:
            print(f"✓ Best fitness achieved: {self.best_chromosome.fitness:.2f}")
        else:
            print("✓ No valid solution found")
    
    def step10_export_output(self, output_file: str):
        """Step 10: Export output to Excel"""
        print("Step 10: Exporting Output...")
        
        if self.best_chromosome is None:
            print("❌ No valid solution found")
            return
        
        # Create timetable data
        timetable_data = []
        for gene in self.best_chromosome.genes:
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
        
        # Create unassigned courses data
        unassigned_data = []
        assigned_courses = set((gene.CourseID, gene.Section) for gene in self.best_chromosome.genes)
        
        for course in self.courses:
            course_key = (course['Course_ID'], course['Section'])
            if course_key not in assigned_courses:
                unassigned_data.append({
                    'CourseName': course['Course_Name'],
                    'Section': course['Section'],
                    'Reason': 'Could not be scheduled due to constraints'
                })
        
        # Create Excel file
        with pd.ExcelWriter(output_file, engine='xlsxwriter') as writer:
            # Timetable sheet
            if timetable_data:
                timetable_df = pd.DataFrame(timetable_data)
                timetable_df.to_excel(writer, sheet_name='Timetable', index=False)
            
            # Unassigned courses sheet
            if unassigned_data:
                unassigned_df = pd.DataFrame(unassigned_data)
                unassigned_df.to_excel(writer, sheet_name='Unassigned Courses', index=False)
        
        print(f"✓ Output exported to {output_file}")
        print(f"✓ Scheduled {len(timetable_data)} sessions")
        print(f"✓ {len(unassigned_data)} courses could not be scheduled")
    
    def run(self, output_file: str = "genetic_timetable.xlsx"):
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
    generator.run("genetic_timetable_output.xlsx")

if __name__ == "__main__":
    main() 