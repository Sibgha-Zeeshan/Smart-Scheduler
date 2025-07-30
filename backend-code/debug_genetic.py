#!/usr/bin/env python3
"""
Debug script for the Genetic Algorithm Timetable Generator

This script helps identify issues that might cause the algorithm to hang.
"""

import os
import sys
import time
from genetic_timetable_generator import GeneticTimetableGenerator

def debug_data_loading():
    """Debug data loading and preprocessing"""
    print("🔍 Debugging Data Loading...")
    
    try:
        generator = GeneticTimetableGenerator(
            input_file="f2025.xlsx",
            population_size=10,
            generations=5,
            mutation_rate=0.05,
            tournament_size=2,
            timeout_minutes=1
            # skip_soft_constraints is omitted (default False)
        )
        generator.step1_preprocess_inputs()
        print("✓ Data preprocessing completed")
        base_genes = generator.step2_create_chromosome_structure()
        print(f"✓ Created {len(base_genes)} base genes")
        generator.step3_generate_initial_population()
        print(f"✓ Generated {len(generator.population)} chromosomes")
        if generator.population:
            fitness = generator.step4_calculate_fitness(generator.population[0])
            print(f"✓ First chromosome fitness: {fitness:.2f}")
        return True
    except Exception as e:
        print(f"❌ Data loading debug failed: {e}")
        import traceback
        traceback.print_exc()
        return False

def debug_single_chromosome():
    """Debug a single chromosome in detail"""
    print("\n🔍 Debugging Single Chromosome...")
    
    input_file = "f2025.xlsx"
    if not os.path.exists(input_file):
        print(f"❌ Input file {input_file} not found")
        return
    
    try:
        generator = GeneticTimetableGenerator(
            input_file=input_file,
            population_size=1,   # Just one chromosome
            generations=1,       # Just one generation
            mutation_rate=0.0,   # No mutation
            tournament_size=1,
            timeout_minutes=1
        )
        
        # Load data
        generator.step1_preprocess_inputs()
        generator.step3_generate_initial_population()
        
        if not generator.population:
            print("❌ No population generated")
            return
        
        chromosome = generator.population[0]
        print(f"Chromosome has {len(chromosome.genes)} genes")
        
        # Analyze genes
        print("\n📋 Gene Analysis:")
        for i, gene in enumerate(chromosome.genes[:5]):  # Show first 5 genes
            print(f"Gene {i+1}: {gene.CourseID}-{gene.Section} | "
                  f"Faculty: {gene.FacultyID} | Room: {gene.RoomID} | "
                  f"Day: {gene.Day} | Time: {gene.StartTime}-{gene.EndTime}")
        
        # Test fitness calculation step by step
        print("\n⚖️ Testing Fitness Calculation Step by Step...")
        
        # Test hard constraints
        print("Testing hard constraints...")
        hard_penalties = generator._calculate_hard_constraint_penalties(chromosome)
        print(f"Hard penalties: {hard_penalties}")
        
        # Test soft constraints
        print("Testing soft constraints...")
        soft_penalties = generator._calculate_soft_constraint_penalties(chromosome)
        print(f"Soft penalties: {soft_penalties}")
        
        # Calculate total fitness
        total_penalty = hard_penalties * generator.hard_constraint_weight + soft_penalties * generator.soft_constraint_weight
        fitness = 1000 - total_penalty
        print(f"Total penalty: {total_penalty}")
        print(f"Fitness: {fitness:.2f}")
        
    except Exception as e:
        print(f"❌ Error during single chromosome debugging: {str(e)}")
        import traceback
        traceback.print_exc()

def debug_evolution_loop():
    """Debug the evolution loop step by step"""
    print("🔍 Debugging Evolution Loop...")
    
    try:
        generator = GeneticTimetableGenerator(
            input_file="f2025.xlsx",
            population_size=10,
            generations=5,
            mutation_rate=0.05,
            tournament_size=2,
            timeout_minutes=1
        )
        generator.step1_preprocess_inputs()
        generator.step2_create_chromosome_structure()
        generator.step3_generate_initial_population()
        print(f"✓ Initial population size: {len(generator.population)}")
        parents = generator.step5_selection()
        print(f"✓ Selected {len(parents)} parents")
        children = generator.step6_crossover(parents)
        print(f"✓ Generated {len(children)} children")
        generator.step7_mutation(children)
        print("✓ Applied mutations")
        for i, child in enumerate(children[:3]):
            fitness = generator.step4_calculate_fitness(child)
            print(f"✓ Child {i+1} fitness: {fitness:.2f}")
        return True
    except Exception as e:
        print(f"❌ Evolution loop debug failed: {e}")
        import traceback
        traceback.print_exc()
        return False

def main():
    """Main debug function"""
    print("🔍 Genetic Algorithm Debug Suite")
    print("=" * 50)
    
    # Test 1: Data loading
    print("\n1️⃣ Testing Data Loading")
    if not debug_data_loading():
        print("❌ Data loading failed, stopping debug")
        return
    
    # Test 2: Single chromosome analysis
    print("\n2️⃣ Testing Single Chromosome")
    debug_single_chromosome()
    
    # Test 3: Evolution loop
    print("\n3️⃣ Testing Evolution Loop")
    debug_evolution_loop()
    
    print("\n" + "=" * 50)
    print("🎉 Debug Suite Complete!")

if __name__ == "__main__":
    main() 