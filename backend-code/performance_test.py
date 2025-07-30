#!/usr/bin/env python3
"""
Performance Test for Genetic Algorithm Timetable Generator
Tests different parameter configurations to demonstrate time complexity improvements
"""

import time
import os
from genetic_timetable_generator import GeneticTimetableGenerator

def test_performance_configurations():
    """Test different parameter configurations and measure performance"""
    print("🚀 Performance Test - Genetic Algorithm Timetable Generator")
    print("=" * 60)
    
    # Check if input file exists
    input_file = "f2025.xlsx"
    if not os.path.exists(input_file):
        print(f"❌ Input file {input_file} not found")
        print("Please ensure f2025.xlsx is in the current directory")
        return
    
    # Define test configurations
    configurations = {
        "Fast": {
            'population_size': 30,
            'generations': 50,
            'mutation_rate': 0.08,
            'tournament_size': 2,
            'timeout_minutes': 1
        },
        "Balanced": {
            'population_size': 50,
            'generations': 100,
            'mutation_rate': 0.05,
            'tournament_size': 3,
            'timeout_minutes': 2
        },
        "Quality": {
            'population_size': 75,
            'generations': 150,
            'mutation_rate': 0.03,
            'tournament_size': 4,
            'timeout_minutes': 3
        }
    }
    
    results = {}
    
    for config_name, params in configurations.items():
        print(f"\n📊 Testing {config_name} Configuration...")
        print(f"   Parameters: {params}")
        
        try:
            # Create generator
            generator = GeneticTimetableGenerator(input_file, **params)
            
            # Measure execution time
            start_time = time.time()
            generator.run(f"performance_test_{config_name.lower()}.xlsx")
            end_time = time.time()
            
            execution_time = end_time - start_time
            
            # Get results
            best_fitness = generator.best_chromosome.fitness if generator.best_chromosome else 0
            sessions_scheduled = len(generator.best_chromosome.genes) if generator.best_chromosome else 0
            
            results[config_name] = {
                'execution_time': execution_time,
                'best_fitness': best_fitness,
                'sessions_scheduled': sessions_scheduled,
                'generations_completed': len(generator.generation_history) if hasattr(generator, 'generation_history') else 0
            }
            
            print(f"   ⏱️  Execution Time: {execution_time:.2f} seconds")
            print(f"   🏆 Best Fitness: {best_fitness:.2f}")
            print(f"   📅 Sessions Scheduled: {sessions_scheduled}")
            print(f"   ✅ Test completed successfully")
            
        except Exception as e:
            print(f"   ❌ Test failed: {e}")
            results[config_name] = {
                'execution_time': float('inf'),
                'best_fitness': 0,
                'sessions_scheduled': 0,
                'generations_completed': 0,
                'error': str(e)
            }
    
    # Print summary
    print("\n" + "=" * 60)
    print("📈 PERFORMANCE SUMMARY")
    print("=" * 60)
    
    print(f"{'Configuration':<15} {'Time (s)':<10} {'Fitness':<10} {'Sessions':<10} {'Speedup':<10}")
    print("-" * 60)
    
    baseline_time = None
    for config_name, result in results.items():
        if result['execution_time'] != float('inf'):
            if baseline_time is None:
                baseline_time = result['execution_time']
                speedup = 1.0
            else:
                speedup = baseline_time / result['execution_time']
            
            print(f"{config_name:<15} {result['execution_time']:<10.2f} {result['best_fitness']:<10.2f} {result['sessions_scheduled']:<10} {speedup:<10.2f}x")
        else:
            print(f"{config_name:<15} {'FAILED':<10} {'N/A':<10} {'N/A':<10} {'N/A':<10}")
    
    # Recommendations
    print("\n💡 RECOMMENDATIONS:")
    print("-" * 30)
    
    fastest_config = min(results.items(), key=lambda x: x[1]['execution_time'] if x[1]['execution_time'] != float('inf') else float('inf'))
    best_fitness_config = max(results.items(), key=lambda x: x[1]['best_fitness'])
    
    if fastest_config[1]['execution_time'] != float('inf'):
        print(f"🚀 Fastest: {fastest_config[0]} ({fastest_config[1]['execution_time']:.2f}s)")
    
    if best_fitness_config[1]['best_fitness'] > 0:
        print(f"🏆 Best Quality: {best_fitness_config[0]} (Fitness: {best_fitness_config[1]['best_fitness']:.2f})")
    
    print("\n📝 Time Complexity Analysis:")
    print("   • Population size has the highest impact on performance")
    print("   • Soft constraints are always included in fitness calculation")
    print("   • Higher mutation rates can lead to faster convergence")
    print("   • Smaller tournament sizes reduce selection overhead")

def main():
    """Main function to run performance tests"""
    test_performance_configurations()

if __name__ == "__main__":
    main() 