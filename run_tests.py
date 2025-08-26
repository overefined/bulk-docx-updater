#!/usr/bin/env python3
"""
Test runner script for the DOCX bulk updater.

Provides convenient commands for running tests with different configurations:
- Run all tests
- Run tests with coverage reporting  
- Run specific test categories (unit, integration)
- Run tests with detailed output

Usage:
    python run_tests.py                    # Run all tests
    python run_tests.py --coverage         # Run with coverage
    python run_tests.py --unit             # Run unit tests only
    python run_tests.py --integration      # Run integration tests only
    python run_tests.py --verbose          # Run with verbose output
"""

import sys
import subprocess
import argparse
from pathlib import Path


def run_command(cmd, description=""):
    """Run a command and return the result."""
    if description:
        print(f"\n{description}")
        print("=" * len(description))
    
    print(f"Running: {' '.join(cmd)}")
    result = subprocess.run(cmd, capture_output=True, text=True)
    
    if result.stdout:
        print(result.stdout)
    if result.stderr:
        print(result.stderr, file=sys.stderr)
    
    return result.returncode == 0


def main():
    """Main test runner function."""
    parser = argparse.ArgumentParser(description="Test runner for DOCX bulk updater")
    parser.add_argument("--coverage", "-c", action="store_true", help="Run tests with coverage reporting")
    parser.add_argument("--unit", "-u", action="store_true", help="Run unit tests only")
    parser.add_argument("--integration", "-i", action="store_true", help="Run integration tests only")
    parser.add_argument("--verbose", "-v", action="store_true", help="Run tests with verbose output")
    parser.add_argument("--fast", "-f", action="store_true", help="Run tests without slow tests")
    parser.add_argument("--file", help="Run specific test file")
    parser.add_argument("--function", help="Run specific test function")
    
    args = parser.parse_args()
    
    # Ensure we're in the correct directory
    script_dir = Path(__file__).parent
    print(f"Running tests from: {script_dir}")
    
    # Build pytest command
    cmd = ["python", "-m", "pytest"]
    
    # Add test selection options
    if args.unit:
        cmd.extend(["-m", "unit"])
    elif args.integration:
        cmd.extend(["-m", "integration"])
    
    # Add coverage options
    if args.coverage:
        cmd.extend(["--cov=.", "--cov-report=term-missing", "--cov-report=html"])
    
    # Add verbosity options
    if args.verbose:
        cmd.extend(["-v", "-s"])
    
    # Add fast option (exclude slow tests)
    if args.fast:
        cmd.extend(["-m", "not slow"])
    
    # Add specific file or function
    if args.file:
        if args.function:
            cmd.append(f"{args.file}::{args.function}")
        else:
            cmd.append(args.file)
    elif args.function:
        cmd.extend(["-k", args.function])
    
    # Run the tests
    success = run_command(cmd, "Running DOCX Bulk Updater Tests")
    
    if success:
        print("\n✓ All tests passed!")
        
        # Show coverage report location if coverage was run
        if args.coverage:
            coverage_dir = script_dir / "htmlcov"
            if coverage_dir.exists():
                print(f"\nCoverage report available at: {coverage_dir / 'index.html'}")
    else:
        print("\n✗ Some tests failed!")
        sys.exit(1)
    
    # Additional helpful information
    print("\n" + "=" * 60)
    print("Test Commands Reference:")
    print("  python run_tests.py                    # Run all tests")
    print("  python run_tests.py --coverage         # Run with coverage")
    print("  python run_tests.py --unit             # Run unit tests only")
    print("  python run_tests.py --integration      # Run integration tests")
    print("  python run_tests.py --verbose          # Detailed output")
    print("  python run_tests.py --fast             # Skip slow tests")
    print("  python run_tests.py --file tests/test_formatting.py")
    print("  python run_tests.py --function test_process_formatting_tokens")
    print("=" * 60)


if __name__ == "__main__":
    main()