#!/usr/bin/env python3
"""
Example running utilities for the autocheck tool.
"""

import subprocess
from pathlib import Path


def get_relative_path(path, project_root):
    """Convert a path to be relative to the project root if possible"""
    try:
        return path.relative_to(project_root)
    except ValueError:
        # If the path can't be made relative to PROJECT_ROOT, return as is
        return path


def build_example(example_name, project_root):
    """Build the example and return True if successful"""
    print(f"Building example: {example_name}")
    cmd = ["zig", "build", "examples", f"-Dexample={example_name}"]
    result = subprocess.run(cmd, capture_output=True, text=True)
    
    if result.returncode != 0:
        print(f"❌ Build failed for {example_name}")
        print(result.stderr)
        return False
    
    print(f"✅ Build successful for {example_name}")
    return True


def run_example(example_name, project_root, quiet=False):
    """Run the example to generate the Excel file"""
    if not quiet:
        print(f"Running example to generate Excel file: {example_name}.xlsx")
    example_bin = project_root / "zig-out" / "bin" / example_name
    
    if not example_bin.exists():
        print(f"❌ Executable not found at {get_relative_path(example_bin, project_root)}")
        return False
    
    result = subprocess.run([str(example_bin)], capture_output=True, text=True)
    
    if result.returncode != 0:
        print(f"❌ Example execution failed for {example_name}")
        print(result.stderr)
        return False
    
    generated_file = Path(f"{example_name}.xlsx")
    if not generated_file.exists():
        print(f"❌ Excel file not generated: {generated_file}")
        return False
    
    if not quiet:
        print(f"✅ Excel file generated: {generated_file}")
    return True 