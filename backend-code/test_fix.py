#!/usr/bin/env python3
"""
Test script to verify the genetic algorithm fixes
"""

import os
import sys
from genetic_timetable_generator import GeneticTimetableGenerator

def test_genetic_algorithm():
    """Test the genetic algorithm with the fixes"""
    print("🧬 Testing Genetic Algorithm with Fixes")
    print("=" * 50)
    
    input_file = "f2025.xlsx"
    
    if not os.path.exists(input_file):
        print(f"❌ Input file {input_file} not found")
        return False
    
    try:
        # Create generator with conservative parameters for testing
        generator = GeneticTimetableGenerator(
            input_file=input_file,
            population_size=20,      # Small population for testing
            generations=10,          # Few generations for testing
            mutation_rate=0.05,
            tournament_size=3,
            timeout_minutes=1,       # Short timeout for testing
            skip_soft_constraints=False
        )
        
        # Run the algorithm
        output_file = "test_genetic_fix.xlsx"
        generator.run(output_file)
        
        print(f"\n✅ Test completed successfully! Check {output_file} for results")
        return True
        
    except Exception as e:
        print(f"\n❌ Test failed with error: {e}")
        import traceback
        traceback.print_exc()
        return False

if __name__ == "__main__":
    success = test_genetic_algorithm()
    if success:
        print("\n🎉 All tests passed!")
    else:
        print("\n💥 Tests failed!")
        sys.exit(1) 