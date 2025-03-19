# Welcome, new-hire.

> **Note**: This document is a guide for the AI assistant to understand the project structure and workflow. When starting fresh, this document will help the assistant provide consistent and accurate assistance.

You are writing Zig code. The code produced will provide a high-level, user-friendly ergonomic, and idiomatic API for the production of Excel spreadsheet (.xlsx files).
The high-level API is a wrapper around a lower-level zig binding to a C library, libxlsxwriter. The example programs for this binding will be referred to for conversion into the high-level API.

## Overview
Your task is to create a high-level API that makes Excel file generation simple and intuitive in Zig. You'll do this by:

1. Studying the low-level bindings in `testing/zig-c-binding-examples/`
2. Creating corresponding high-level examples in `examples/`
3. Verifying your work against reference files

You will be running programs from the project root which is the current directory. 

Directories in the project are:

**src/**  : High-level interface wrapper code

**examples/** : Being populated with examples using the new highlevel wrappers being developed. Each example from testing/zig-c-binding-examples/ should be represented by a corresponding example in this directory.

**testing/**  : Files and programs useful for testing and verification

**testing/reference-xls/*.xlsx** : Reference xlsx files that are to be compared to the output generated from the high-level wrappers being developed

**testing/test-output-xls** : Output directory to place verified outputs from examples/*.zig after testing them to make sure they produce correct output

**testing/zig-c-binding-examples**: binding-style zig examples that will be examined for understanding the lower level calls to use when designing the wrapper interface to the bindings. These will not be modified.

**zig-out/bin/** : Output executables from the zig build process (example programs) will be here. When executed, they produce a spreadsheet. This will be compared against a reference spreadsheet in **testing/reference-xls/*.xlsx** .

**utils/status** : A program that will show the current status of progress on creating the examples corresponding to the examples in zig-c-binding-examples.

**utils/autocheck.py** : A Python script that checks Excel files for common issues and compatibility with reference files.

**utils/verify.py** : A script that helps automate the process of taking screenshots of Excel files for manual verification.

**utils/unverify.py** : A script to remove verification status from examples after API changes.

## Workflow

### Development Cycle
Follow this cycle:

1. Check what needs to be done:
```bash
zig build status          # full detailed view
zig build status -- --short  # compact view of implemented/verified examples
```

2. Check status of a specific example:
```bash
zig build status -- hello
```

3. Verify your work:
```bash
zig build run verify -- hello
```

### Automated Excel Checking
Before submitting an example for manual verification, use the `autocheck.py` script to catch common issues:

```bash
python3 utils/autocheck.py tutorial1 --build --run
```

This script performs several checks on the generated Excel file:

1. **Formula checks**: Detects circular references, malformed ranges, and other formula issues
2. **String null-termination**: Finds issues with strings not being properly null-terminated
3. **XML content analysis**: Examines the internal XML structure for encoding problems
4. **Binary compatibility**: Compares file structure with the reference file
5. **Content comparison**: Verifies cell contents match the reference file

The script requires the `openpyxl` Python package:
```bash
pip install openpyxl
```

Common usage patterns:
```bash
# Build, run and check an example
python3 utils/autocheck.py example_name --build --run

# Check an existing Excel file without building
python3 utils/autocheck.py example_name 

# Check any Excel file (doesn't need to be an example)
python3 utils/autocheck.py some_file --file-only

# Check all examples at once
python3 utils/autocheck.py --all

# Ignore style differences in comparison (useful for array formulas and other cases where internal representation may differ)
python3 utils/autocheck.py example_name --ignore-styles
```

Using this tool before manual verification saves time by catching technical issues early, allowing you to fix problems before asking a human to verify the visual appearance.

### Manual Verification Process
The manual verification process helps ensure that the generated Excel files match the reference files:

1. Run the verifier:
```bash
./utils/verify.py hello
```

2. The script will:
   - Build the example
   - Take a screenshot comparing the generated Excel with the reference file
   - Ask you to verify if the output looks correct
   - Save all results in `testing/results/hello/`:
     - `hello.xlsx` (if verified) or `failed-hello.xlsx` (if failed)
     - `comparison_hello.png` (screenshot)
     - `hello_output.txt` (verification result)

3. The status program (`zig build status`) checks for the presence of the output.txt file in the results directory to determine if an example is verified.

4. If you need to unverify examples after API changes, use the unverify.py script:
```bash
# Unverify a specific example
./utils/unverify.py hello

# List all verified examples
./utils/unverify.py --list

# Unverify all examples
./utils/unverify.py --all
```

## Coding Standards

- Use comma (,) after the last parameter in function definitions, struct
  definitions, function calls, etc. so the zig formatter will wrap lines
  and keep the line width less than 80 characters. You can also put a
  newline (\n) character after the '=' in an assignment, and zig will
  move the right side to the next line, and still understand the syntax.
  This is a good way to also shorten line width if needed.

- Refactor common functionality so functions are short while maintaining
  the user-friendliness for the public API calls

## Common Problems & Troubleshooting

When implementing examples, you might encounter issues where your output file doesn't match the reference. Here are some common problems and solutions:

### Zig Language Issues

- **Deprecated std.mem.split**: As of Zig 0.14.0, `std.mem.split` is deprecated. Use `std.mem.splitScalar`, `std.mem.splitAny`, or `std.mem.splitSequence` instead, depending on your use case. Most string splitting with a single character delimiter should use `std.mem.splitScalar`.

### Excel Formula Issues

- **Warning Indicators**: If Excel shows warning indicators in formula cells in your output but not in the reference file, check:
  1. The exact range syntax used in the formula
  2. Whether your range inadvertently creates a circular reference
  3. If you need to dynamically calculate the range based on actual data

- **Formula Not Working**: If formulas appear as text instead of calculations:
  1. Ensure strings are properly null-terminated when sent to the C library
  2. Use `allocator.dupeZ()` to create null-terminated strings
  3. Check the exact format used in the reference example

### String Handling

- Always ensure strings passed to the C library are null-terminated
- Use `allocator.dupeZ()` to create null-terminated strings from slices
- Remember to free these allocations after use with `defer allocator.free()`

### Binary Compatibility

- Excel's file format has many nuances that aren't visible in the displayed content
- Even when cells appear identical, underlying binary encoding matters
- Follow the exact patterns used in the C library for best compatibility

### API Design Tips

- When adding new functions, ensure they work with both row/column indices and cell references
- For methods that take format parameters, make them optional with `?*Format`
- Use builder patterns with method chaining for better ergonomics
- Always free resources with `defer` statements

### Verification Strategy

If verification fails:
1. Compare your implementation with the reference example line by line
2. Run the autocheck.py script to identify technical issues
3. Use verify.py to see exactly what's different visually
4. Sometimes you need to rebuild with the exact same values as the reference
5. For formulas, use the C library's exact syntax rather than creating your own

Remember that the goal is binary compatibility with the reference files, not just visual similarity.
  