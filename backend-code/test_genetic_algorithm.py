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
    """Test the genetic algorithm with existing validated data"""
    print("Testing Genetic Algorithm with existing data...")
    
    generator = GeneticTimetableGenerator(
        input_file="f2025.xlsx",
        population_size=50,
        generations=100,
        mutation_rate=0.05,
        tournament_size=3,
        timeout_minutes=2
        # skip_soft_constraints is omitted (default False)
    )
    
    try:
        generator.run("test_genetic_output.xlsx")
        print("✓ Test completed successfully")
    except Exception as e:
        print(f"❌ Test failed: {e}")

def test_with_different_parameters():
    """Test with different parameter combinations"""
    print("Testing different parameter combinations...")
    
    # All configs use skip_soft_constraints=False
    configs = [
        ("Fast", {
            'population_size': 30,
            'generations': 50,
            'mutation_rate': 0.08,
            'tournament_size': 2,
            'timeout_minutes': 1
        }),
        ("Balanced", {
            'population_size': 50,
            'generations': 100,
            'mutation_rate': 0.05,
            'tournament_size': 3,
            'timeout_minutes': 2
        }),
        ("Quality", {
            'population_size': 75,
            'generations': 150,
            'mutation_rate': 0.03,
            'tournament_size': 4,
            'timeout_minutes': 3
        })
    ]
    
    for config_name, params in configs:
        print(f"\nTesting {config_name} configuration...")
        try:
            generator = GeneticTimetableGenerator("f2025.xlsx", **params)
            start_time = time.time()
            generator.run(f"test_{config_name.lower()}_output.xlsx")
            end_time = time.time()
            print(f"✓ {config_name} test completed in {end_time - start_time:.2f} seconds")
        except Exception as e:
            print(f"❌ {config_name} test failed: {e}")

def compare_with_csp():
    """Compare genetic algorithm with CSP approach"""
    print("Comparing Genetic Algorithm with CSP...")
    
    generator = GeneticTimetableGenerator(
        input_file="f2025.xlsx",
        population_size=50,
        generations=100,
        mutation_rate=0.05,
        tournament_size=3,
        timeout_minutes=2
    )
    
    try:
        start_time = time.time()
        generator.run("genetic_comparison_output.xlsx")
        genetic_time = time.time() - start_time
        print(f"✓ Genetic Algorithm completed in {genetic_time:.2f} seconds")
        print("Note: CSP comparison not yet implemented")
    except Exception as e:
        print(f"❌ Comparison failed: {e}")

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