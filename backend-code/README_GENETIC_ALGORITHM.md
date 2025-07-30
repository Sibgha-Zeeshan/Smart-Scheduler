# Genetic Algorithm Timetable Generator

This is a genetic algorithm-based implementation for generating university timetables, designed as an alternative to the existing CSP (Constraint Satisfaction Problem) approach.

## 🧬 Overview

The genetic algorithm approach uses evolutionary computation to find optimal timetable solutions by:

- Representing timetables as chromosomes with genes for each session
- Using fitness functions to evaluate solution quality
- Applying selection, crossover, and mutation operators
- Incorporating local search for improvement

## 📁 Files

- `genetic_timetable_generator.py` - Main genetic algorithm implementation
- `test_genetic_algorithm.py` - Test script to demonstrate usage
- `README_GENETIC_ALGORITHM.md` - This documentation file

## 🚀 Quick Start

### Prerequisites

Ensure you have the required dependencies:

```bash
pip install pandas numpy tqdm xlsxwriter
```

### Basic Usage

```python
from genetic_timetable_generator import GeneticTimetableGenerator

# Create generator
generator = GeneticTimetableGenerator(
    input_file="f2025.xlsx",
    population_size=100,
    generations=300,
    mutation_rate=0.02,
    tournament_size=5,
    timeout_minutes=5
)

# Run the algorithm
generator.run("genetic_timetable_output.xlsx")
```

### Running Tests

```bash
python test_genetic_algorithm.py
```

## 🔧 Algorithm Steps

### Step 1: Preprocessing Inputs

- Reads 5 Excel sheets (Courses, Faculty, Rooms, Time Slots, Students)
- Parses duration strings to minutes
- Converts faculty availability to day lists
- Parses course assignments to dictionaries
- Generates sections based on student count

### Step 2: Chromosome Structure

Each gene represents one session:

```python
Gene = {
    "CourseID": "CC141",
    "Section": "V1",
    "FacultyID": 5,
    "RoomID": "Room2",
    "Day": "Tuesday",
    "StartTime": 9.0,
    "EndTime": 10.25,
}
```

### Step 3: Initial Population

- Generates P chromosomes (default: 100)
- Randomly assigns rooms, faculty, and time slots
- Avoids obvious violations during creation

### Step 4: Fitness Function

Calculates penalties for constraint violations:

**Hard Constraints (Weight: 100)**

- Room type mismatch
- Room capacity exceeded
- Faculty not assigned course
- Double booking (room, faculty, section)
- Time slot out of bounds
- Course duration misfit
- Course frequency violation

**Soft Constraints (Weight: 10)**

- Low-rated faculty assigned
- Spread-out faculty time window
- Inconsistent start time per course
- Unbalanced section allocation

### Step 5: Selection

Uses tournament selection to choose parents for reproduction.

### Step 6: Crossover

Performs gene-level crossover with conflict resolution.

### Step 7: Mutation

Randomly mutates 1-2% of genes with constraint validation.

### Step 8: Local Search

Applies local improvements through time slot swaps.

### Step 9: Evolution Loop

Repeats for G generations or until convergence/timeout.

### Step 10: Export Output

Creates Excel file with timetable and unassigned courses.

## Parameters

The genetic algorithm can be configured with the following parameters:

| Parameter               | Default | Description                                         | Impact on Performance                    |
| ----------------------- | ------- | --------------------------------------------------- | ---------------------------------------- |
| `population_size`       | 75      | Number of chromosomes in population                 | **High** - Smaller = faster              |
| `generations`           | 150     | Maximum number of generations                       | **High** - Fewer = faster                |
| `mutation_rate`         | 0.05    | Probability of mutation per gene                    | **Medium** - Higher = faster convergence |
| `tournament_size`       | 3       | Size of tournament for selection                    | **Medium** - Smaller = faster            |
| `timeout_minutes`       | 2       | Maximum runtime in minutes                          | **High** - Shorter = faster termination  |
| `skip_soft_constraints` | False   | (Must remain False) Always include soft constraints | **Required** - Always include            |

### Parameter Recommendations

**For Faster Execution:**

```python
generator = GeneticTimetableGenerator(
    input_file="f2025.xlsx",
    population_size=30,      # Very small population
    generations=50,          # Few generations
    mutation_rate=0.08,      # High mutation rate
    tournament_size=2,       # Small tournament
    timeout_minutes=1        # Short timeout
    # skip_soft_constraints is always False (soft constraints always included)
)
```

**For Better Solutions:**

```python
generator = GeneticTimetableGenerator(
    input_file="f2025.xlsx",
    population_size=100,     # Larger population
    generations=200,         # More generations
    mutation_rate=0.03,      # Lower mutation rate
    tournament_size=4,       # Larger tournament
    timeout_minutes=3        # Longer timeout
    # skip_soft_constraints is always False (soft constraints always included)
)
```

**For Balanced Performance:**

```python
generator = GeneticTimetableGenerator(
    input_file="f2025.xlsx",
    population_size=50,      # Medium population
    generations=100,         # Medium generations
    mutation_rate=0.05,      # Balanced mutation
    tournament_size=3,       # Medium tournament
    timeout_minutes=2        # Medium timeout
    # skip_soft_constraints is always False (soft constraints always included)
)
```

**Note:** Soft constraints are always included in the fitness calculation. Do not set `skip_soft_constraints=True`.

## 📊 Output Format

The algorithm generates an Excel file with two sheets:

### Timetable Sheet

- CourseID, CourseName, CourseType, Section
- FacultyID, FacultyName
- Day, StartTime, EndTime
- Room, RoomType, Rating

### Unassigned Courses Sheet

- CourseName, Section, Reason

## 🔄 Comparison with CSP Approach

| Aspect               | Genetic Algorithm          | CSP Approach               |
| -------------------- | -------------------------- | -------------------------- |
| **Method**           | Evolutionary computation   | Backtracking search        |
| **Solution Quality** | Near-optimal               | Exact (if found)           |
| **Runtime**          | Predictable (timeout)      | Variable (may timeout)     |
| **Scalability**      | Good for large problems    | Limited by search space    |
| **Flexibility**      | Easy to modify constraints | Requires algorithm changes |
| **Convergence**      | May not find optimal       | Guaranteed if exists       |

## 🎯 Advantages of Genetic Algorithm

1. **Predictable Performance**: Timeout ensures completion
2. **Scalability**: Handles large problem instances well
3. **Flexibility**: Easy to add new constraints
4. **Robustness**: Less likely to get stuck in local optima
5. **Parallelization**: Can be easily parallelized

## 🎯 Advantages of CSP Approach

1. **Optimality**: Finds exact solutions when they exist
2. **Completeness**: Guaranteed to find solution if one exists
3. **Deterministic**: Same input always produces same output
4. **Memory Efficient**: Uses backtracking instead of population

## 📈 Performance Tuning

### For Better Solutions

```python
generator = GeneticTimetableGenerator(
    population_size=300,    # Larger population
    generations=500,        # More generations
    mutation_rate=0.01,     # Lower mutation rate
    timeout_minutes=10      # More time
)
```

### For Faster Execution

```python
generator = GeneticTimetableGenerator(
    population_size=100,    # Smaller population
    generations=100,        # Fewer generations
    mutation_rate=0.05,     # Higher mutation rate
    timeout_minutes=2       # Less time
)
```

## 🧪 Testing

Run the test suite to verify functionality:

```bash
python test_genetic_algorithm.py
```

This will:

1. Test basic functionality with existing data
2. Test different parameter combinations
3. Provide comparison metrics

## 🔍 Debugging

Enable verbose output by modifying the print statements in the code. Key areas to monitor:

- Fitness calculation
- Constraint violations
- Population diversity
- Convergence patterns

## 📝 Example Output

```
🧬 Starting Genetic Algorithm Timetable Generator
==================================================
Step 1: Preprocessing Inputs...
✓ Processed 260 course sections
✓ Processed 98 faculty members
✓ Processed 15 rooms
✓ Generated 60 time slots

Step 3: Generating Initial Population...
✓ Generated 100 chromosomes

Step 9: Starting Evolution Loop...
✓ Evolution completed in 45.23 seconds
✓ Best fitness achieved: 856.78

Step 10: Exporting Output...
✓ Output exported to genetic_timetable_output.xlsx
✓ Scheduled 245 sessions
✓ 15 courses could not be scheduled

==================================================
🎉 Genetic Algorithm Timetable Generation Complete!
```

## 🤝 Integration

The genetic algorithm can be integrated into the existing FastAPI application by:

1. Adding a new endpoint for genetic algorithm generation
2. Modifying the app.py to include genetic algorithm option
3. Providing parameter configuration through the API

## 📚 References

- Genetic Algorithms in Search, Optimization, and Machine Learning (Goldberg)
- Timetabling using Genetic Algorithms (Burke & Newall)
- Constraint Satisfaction Problems (Russell & Norvig)

## 🐛 Known Issues

1. **Memory Usage**: Large populations may consume significant memory
2. **Local Optima**: May get stuck in local optima for complex constraints
3. **Parameter Sensitivity**: Performance depends on parameter tuning

## 🔮 Future Improvements

1. **Multi-objective Optimization**: Handle conflicting objectives
2. **Adaptive Parameters**: Self-tuning mutation and crossover rates
3. **Parallel Processing**: Multi-threaded evolution
4. **Hybrid Approaches**: Combine with other metaheuristics
5. **Real-time Constraints**: Handle dynamic scheduling requirements
