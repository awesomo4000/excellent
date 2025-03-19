#!/usr/bin/env python3
"""
Automated checks for Excel files to identify potential issues before manual verification.
This tool examines Excel files for common problems like formula warnings, missing data,
or encoding issues.

Common usage:
  python3 utils/autocheck.py example_name --build --run        # Build, run and check
  python3 utils/autocheck.py example_name --ignore-styles      # Ignore style differences
  python3 utils/autocheck.py --all                             # Check all examples
"""

import os
import sys
import argparse
from pathlib import Path
import openpyxl
import traceback

# Add the project root to sys.path to allow imports to work both when run as a module and directly
current_dir = Path(__file__).parent
project_root = current_dir.parent
sys.path.insert(0, str(project_root))

# Import local modules
try:
    # When run as module (python -m utils.autocheck)
    from utils.excel_checks import (
        check_formulas,
        check_string_null_termination,
        check_xml_content,
        check_binary_compatibility
    )
    from utils.file_comparison import compare_with_reference, get_relative_path
    from utils.example_runner import build_example, run_example
except ModuleNotFoundError:
    # When run directly (python utils/autocheck.py)
    from excel_checks import (
        check_formulas,
        check_string_null_termination,
        check_xml_content,
        check_binary_compatibility
    )
    from file_comparison import compare_with_reference, get_relative_path
    from example_runner import build_example, run_example

# Set up paths
PROJECT_ROOT = Path(__file__).parent.parent
EXAMPLES_DIR = PROJECT_ROOT / "examples"
REFERENCE_DIR = PROJECT_ROOT / "testing" / "reference-xls"
RESULTS_DIR = PROJECT_ROOT / "testing" / "results"


def check_single_example(example_name, args):
    """Run checks on a single example"""
    # Check if example exists, unless in file-only mode
    if not args.file_only:
        example_file = EXAMPLES_DIR / f"{example_name}.zig"
        if not example_file.exists():
            print(f"❌ Example file not found: {get_relative_path(example_file, PROJECT_ROOT)}")
            return False
    
    # Build if requested
    if args.build:
        if not build_example(example_name, PROJECT_ROOT):
            return False
    
    # Run if requested
    if args.run:
        if not run_example(example_name, PROJECT_ROOT):
            return False
    
    # Look for the Excel file
    excel_file = Path(f"{example_name}.xlsx")
    if not excel_file.exists():
        print(f"❌ Excel file not found: {excel_file}")
        print("Run with --run option to generate it")
        return False
    
    print(f"\n=== Checking Excel file: {excel_file} ===\n")
    
    try:
        # Load workbook for checks
        workbook = openpyxl.load_workbook(excel_file)
        
        # Run checks
        print("Checking formulas...")
        formula_check = check_formulas(workbook, example_name)
        
        print("Checking string null-termination...")
        string_check = check_string_null_termination(workbook)
        
        print("Checking XML content...")
        xml_check = check_xml_content(example_name)
        
        print("Checking binary compatibility...")
        binary_check = check_binary_compatibility(example_name, REFERENCE_DIR)
        
        print("Comparing with reference file...")
        content_check = compare_with_reference(
            example_name, 
            REFERENCE_DIR, 
            RESULTS_DIR, 
            PROJECT_ROOT, 
            ignore_styles=args.ignore_styles
        )
        
        # Summary
        print("\n=== Check Summary ===")
        print(f"Formula Check: {'✅ PASSED' if formula_check else '❌ FAILED'}")
        print(f"String Null-Termination: {'✅ PASSED' if string_check else '❌ FAILED'}")
        print(f"XML Check: {'✅ PASSED' if xml_check else '❌ FAILED'}")
        print(f"Binary Compatibility: {'✅ PASSED' if binary_check else '❌ FAILED'}")
        print(f"Content Check: {'✅ PASSED' if content_check else '❌ FAILED'}")
        
        if formula_check and string_check and xml_check and binary_check and content_check:
            print("\n✅ All checks passed! The file should pass manual verification.")
            return True
        else:
            print("\n⚠️ Some checks failed. Review issues before manual verification.")
            return False
            
    except Exception as e:
        print(f"❌ Error checking Excel file: {e}")
        traceback.print_exc()
        return False


def check_all_examples(args):
    """Run checks on all examples"""
    # Get all example files
    example_files = list(EXAMPLES_DIR.glob("*.zig"))
    failed_examples = []
    
    for example_file in example_files:
        example_name = example_file.stem
        if example_name == "status":  # Skip status binary
            continue
            
        # Run the example
        if not run_example(example_name, PROJECT_ROOT, quiet=True):
            failed_examples.append((example_name, "Failed to run example"))
            continue
            
        # Check the Excel file
        excel_file = Path(f"{example_name}.xlsx")
        if not excel_file.exists():
            failed_examples.append((example_name, "Excel file not generated"))
            continue
            
        try:
            # Load workbook for checks
            workbook = openpyxl.load_workbook(excel_file)
            
            # Run checks
            formula_check = check_formulas(workbook, example_name)
            string_check = check_string_null_termination(workbook)
            xml_check = check_xml_content(example_name)
            binary_check = check_binary_compatibility(example_name, REFERENCE_DIR)
            content_check = compare_with_reference(
                example_name, 
                REFERENCE_DIR, 
                RESULTS_DIR, 
                PROJECT_ROOT, 
                quiet=True, 
                ignore_styles=args.ignore_styles
            )
            
            if formula_check and string_check and xml_check and binary_check and content_check:
                print(f"✅ {example_name}")
            else:
                failed_examples.append((example_name, "One or more checks failed"))
                # Re-run checks with verbose output to show details
                print(f"\nDetailed output for {example_name}:")
                try:
                    workbook = openpyxl.load_workbook(Path(f"{example_name}.xlsx"))
                    print("\nChecking formulas...")
                    check_formulas(workbook, example_name)
                    print("\nChecking string null-termination...")
                    check_string_null_termination(workbook)
                    print("\nChecking XML content...")
                    check_xml_content(example_name)
                    print("\nChecking binary compatibility...")
                    check_binary_compatibility(example_name, REFERENCE_DIR)
                    print("\nComparing with reference file...")
                    compare_with_reference(
                        example_name, 
                        REFERENCE_DIR, 
                        RESULTS_DIR, 
                        PROJECT_ROOT, 
                        ignore_styles=args.ignore_styles
                    )
                except Exception as e:
                    print(f"Error during detailed check: {e}")
                    traceback.print_exc()
                
        except Exception as e:
            failed_examples.append((example_name, f"Error checking Excel file: {e}"))
    
    # Print detailed failure information if any
    if failed_examples:
        print("\n=== Failed Examples ===")
        for example_name, error in failed_examples:
            print(f"\n❌ {example_name}: {error}")
        
        return False
    else:
        print("\n✅ All examples passed!")
        return True


def main():
    parser = argparse.ArgumentParser(description="Check Excel files for common issues before verification")
    parser.add_argument("example", nargs="?", help="Name of the example to check (without .zig extension)")
    parser.add_argument("--build", action="store_true", help="Build the example before checking")
    parser.add_argument("--run", action="store_true", help="Run the example to generate the Excel file")
    parser.add_argument("--verbose", "-v", action="store_true", help="Show detailed information about checks")
    parser.add_argument("--file-only", action="store_true", help="Skip example file check, just check the Excel file")
    parser.add_argument("--all", action="store_true", help="Check all examples (except status)")
    parser.add_argument("--ignore-styles", action="store_true", help="Ignore style differences in comparison")
    
    args = parser.parse_args()
    
    if args.all:
        if args.example:
            print("Error: Cannot specify both --all and an example name")
            return 1
        
        return 0 if check_all_examples(args) else 1
    
    if not args.example:
        print("Error: Must specify either an example name or --all")
        return 1
    
    return 0 if check_single_example(args.example, args) else 1


if __name__ == "__main__":
    sys.exit(main()) 