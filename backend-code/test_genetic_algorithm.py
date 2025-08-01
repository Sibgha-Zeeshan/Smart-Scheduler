#!/usr/bin/env python3
"""
Test script for the Genetic Algorithm Timetable Generator

This script demonstrates how to use the genetic algorithm with the existing data
from the backend-code folder.
"""

import os
import sys
from genetic_timetable_generator import GeneticTimetableGenerator
import time # Added for timing tests

def test_with_existing_data():
    """Test the genetic algorithm with existing data"""
    print("🧬 Testing Genetic Algorithm with Existing Data")
    print("=" * 50)
    
    input_file = "f2025.xlsx"
    
    if not os.path.exists(input_file):
        print(f"❌ Input file {input_file} not found")
        return
    
    # Create generator with optimized parameters
    generator = GeneticTimetableGenerator(
        input_file=input_file,
        population_size=75,      # Optimized for speed
        generations=150,         # Optimized for speed
        mutation_rate=0.05,      # Increased for faster convergence
        tournament_size=3,       # Reduced for speed
        timeout_minutes=2,       # Reduced for speed
        skip_soft_constraints=False  # Always include soft constraints
    )
    
    # Run the algorithm
    output_file = "genetic_timetable_test.xlsx"
    generator.run(output_file)
    
    print(f"\n✅ Test completed! Check {output_file} for results")

def test_with_different_parameters():
    """Test with different parameter combinations"""
    print("🧬 Testing Genetic Algorithm with Different Parameters")
    print("=" * 50)
    
    input_file = "f2025.xlsx"
    
    if not os.path.exists(input_file):
        print(f"❌ Input file {input_file} not found")
        return
    
    # Test configurations
    configs = [
        {
            "name": "Fast",
            "population_size": 50,
            "generations": 100,
            "mutation_rate": 0.08,
            "tournament_size": 2,
            "timeout_minutes": 1
        },
        {
            "name": "Balanced",
            "population_size": 75,
            "generations": 150,
            "mutation_rate": 0.05,
            "tournament_size": 3,
            "timeout_minutes": 2
        },
        {
            "name": "Quality",
            "population_size": 100,
            "generations": 200,
            "mutation_rate": 0.03,
            "tournament_size": 4,
            "timeout_minutes": 3
        }
    ]
    
    for config in configs:
        print(f"\n🔧 Testing {config['name']} Configuration:")
        print(f"Population: {config['population_size']}, Generations: {config['generations']}")
        
        generator = GeneticTimetableGenerator(
            input_file=input_file,
            population_size=config['population_size'],
            generations=config['generations'],
            mutation_rate=config['mutation_rate'],
            tournament_size=config['tournament_size'],
            timeout_minutes=config['timeout_minutes'],
            skip_soft_constraints=False  # Always include soft constraints
        )
        
        output_file = f"genetic_timetable_{config['name'].lower()}.xlsx"
        generator.run(output_file)
        
        print(f"✅ {config['name']} test completed!")

def compare_with_csp():
    """Compare genetic algorithm with CSP approach"""
    print("🧬 Comparing Genetic Algorithm with CSP")
    print("=" * 50)
    
    input_file = "f2025.xlsx"
    
    if not os.path.exists(input_file):
        print(f"❌ Input file {input_file} not found")
        return
    
    # Run genetic algorithm
    print("\n🔬 Running Genetic Algorithm...")
    start_time = time.time()
    
    generator = GeneticTimetableGenerator(
        input_file=input_file,
        population_size=75,
        generations=150,
        mutation_rate=0.05,
        tournament_size=3,
        timeout_minutes=2,
        skip_soft_constraints=False  # Always include soft constraints
    )
    
    generator.run("genetic_comparison.xlsx")
    ga_time = time.time() - start_time
    
    print(f"\n📊 Genetic Algorithm Results:")
    print(f"Time: {ga_time:.2f} seconds")
    
    # Note: CSP comparison would require running CSP3.py separately
    print("\n💡 To compare with CSP:")
    print("1. Run: python CSP3.py f2025.xlsx")
    print("2. Compare the output files")
    print("3. Check accuracy and constraint satisfaction")

def main():
    """Main test function"""
    print("🧬 Genetic Algorithm Timetable Generator - Test Suite")
    print("=" * 60)
    
    # Test 1: Basic functionality
    print("\n1️⃣ Testing Basic Functionality")
    success = test_with_existing_data()
    
    if success:
        # Test 2: Different parameters
        print("\n2️⃣ Testing Different Parameters")
        test_with_different_parameters()
        
        # Test 3: Comparison
        print("\n3️⃣ Comparison Test")
        compare_with_csp()
    
    print("\n" + "=" * 60)
    print("🎉 Test Suite Complete!")
    
    # Instructions for running CSP comparison
    print("\n📝 To compare with CSP approach:")
    print("   1. Run: python CSP3.py f2025.xlsx")
    print("   2. Compare the output files")
    print("   3. Analyze fitness vs scheduling success rate")

if __name__ == "__main__":
    main() 