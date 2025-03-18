#!/usr/bin/env python3
"""
Unverify examples by removing their verification status.
This is used after API changes to force re-verification of previously verified examples.
"""

import os
import sys
import glob
import argparse
from pathlib import Path

# Set up paths
PROJECT_ROOT = Path(__file__).parent.parent
RESULTS_DIR = PROJECT_ROOT / "testing" / "results"


def list_verified_examples():
    """List all currently verified examples"""
    verified_examples = []
    for example_dir in RESULTS_DIR.glob("*"):
        if not example_dir.is_dir():
            continue
        
        verified_file = example_dir / "verified"
        if verified_file.exists():
            verified_examples.append(example_dir.name)
    
    return sorted(verified_examples)


def unverify_example(example_name):
    """Unverify a specific example by removing its verified status file"""
    example_dir = RESULTS_DIR / example_name
    verified_file = example_dir / "verified"
    
    if not example_dir.exists():
        print(f"🚫 Example '{example_name}' does not exist in results directory")
        return False
    
    if not verified_file.exists():
        print(f"ℹ️ Example '{example_name}' is already unverified")
        return True
    
    try:
        verified_file.unlink()
        print(f"✅ Successfully unverified '{example_name}'")
        return True
    except Exception as e:
        print(f"❌ Error unverifying '{example_name}': {e}")
        return False


def unverify_all_examples():
    """Unverify all examples by removing all verified status files"""
    verified_examples = list_verified_examples()
    
    if not verified_examples:
        print("ℹ️ No verified examples found")
        return True
    
    success = True
    for example in verified_examples:
        if not unverify_example(example):
            success = False
    
    return success


def main():
    parser = argparse.ArgumentParser(description="Unverify examples after API changes")
    parser.add_argument("examples", nargs="*", help="Names of examples to unverify (leave empty to list or use --all)")
    parser.add_argument("--all", action="store_true", help="Unverify all examples")
    parser.add_argument("--list", action="store_true", help="List all verified examples")
    
    args = parser.parse_args()
    
    # Just list verified examples
    if args.list or (not args.examples and not args.all):
        verified = list_verified_examples()
        if verified:
            print(f"Currently verified examples ({len(verified)}):")
            for example in verified:
                print(f"  - {example}")
        else:
            print("No verified examples found")
        return 0
    
    # Unverify all examples
    if args.all:
        if unverify_all_examples():
            return 0
        return 1
    
    # Unverify specific examples
    success = True
    for example in args.examples:
        if not unverify_example(example):
            success = False
    
    return 0 if success else 1


if __name__ == "__main__":
    sys.exit(main()) 