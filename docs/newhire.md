# Welcome, new-hire.

> **Note**: This document is a guide for the AI assistant to understand the project structure and workflow. When starting fresh, this document will help the assistant provide consistent and accurate assistance.

You are writing Zig code to provide a high-level, user-friendly ergonomic, and idiomatic API for the production of Excel spreadsheet (.xlsx files).
The high-level API is a wrapper around a lower-level zig binding to a C library, libxlsxwriter. The example programs for this binding will be referred to for conversion into the high-level API.

## Overview

Your task is to create a high-level API that makes Excel file generation simple and intuitive in Zig. You'll do this by:

1. Studying the low-level bindings in `testing/zig-c-binding-examples/`
2. Creating corresponding high-level examples in `examples/`
3. Verifying your work against reference files

You will be running programs from the project root which is the current directory. 

Directories in the project are:

**src/**  : High-level interface wrapper code

**examples/** : Being populated with examples using the new highlevel wrappers being developed (in src/). Each example from `testing/zig-c-binding-examples/` should be represented by a corresponding example in this directory. These are also representative of the examples written in the C language (in `testing/c-examples/`). Both should be referred to when developing wrapper examples based on them.

**testing/**  : Files and programs used for testing and verification

**testing/reference-xls/*.xlsx** : Reference xlsx files that are to be compared to the output generated from the high-level wrappers being developed.

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
./utils/status          # full detailed view

./utils/status --short  # compact view of implemented/verified examples
```

2. Check status of a specific example:
```bash
./utils/status hello
```

### Automated Excel Checking

Before submitting an example for manual verification, use the `autocheck.py` script to catch common issues. This command
should also be run to iterate when developing an example, since
it will show compile errors while developing the code, and
run the autocheck.py if it passes compilation.

```bash
python3 utils/autocheck.py tutorial1 --build --run        # Build, run and check
python3 utils/autocheck.py tutorial1 --ignore-styles      # Ignore style differences
python3 utils/autocheck.py --all                         # Check all examples
python3 utils/autocheck.py --all --build                 # Build all examples
python3 utils/autocheck.py --all --build --force         # Build all examples including broken ones
python3 utils/autocheck.py --list-broken                 # List known broken examples
```

This script performs several checks on the generated Excel file:

1. **Formula checks**: Detects circular references, malformed ranges, and other formula issues
2. **String null-termination**: Finds issues with strings not being properly null-terminated
3. **XML content analysis**: Examines the internal XML structure for encoding problems
4. **Binary compatibility**: Compares file structure with the reference file
5. **Content comparison**: Verifies cell contents match the reference file


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

3. The status program (`./utils/status`) checks for the presence of `testing/results/{example_name}/verified` file in the results directory to determine if an example is verified.

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

- Use comma (,) after the last parameter in function definitions, struct definitions, function calls, etc. so the zig formatter will wrap lines and keep the line width less than 80 characters. You can also put a newline (\n) character after the '=' in an assignment, and zig will move the right side to the next line, and still understand the syntax. This is a good way to also shorten line width if needed.

- Refactor common functionality so functions are short while maintaining the user-friendliness for the public API calls

- When creating new examples in the `examples/` directory, name the output Excel files without the "zig-" prefix (e.g., use "comments1.xlsx" not "zig-comments1.xlsx"). This ensures the output matches the reference files exactly. Many of the .zig examples in `testing/zig-c-binding-examples` use the "zig-examplename.xlsx" format, which is not to be used for generating examples in the high-level API examples.

- When determining how to update an example to fix it, importing "xlsxwriter" is usually the wrong answer. If there are parts of "xlswriter" directly is what the wrapper API is seeking to prevent.

## Common Problems & Troubleshooting

When implementing examples, you might encounter issues where your output file doesn't match the reference. Here are some common problems and solutions:

### Build System

- **DO NOT MODIFY build.zig**: The build system automatically discovers and builds examples from the examples/ directory. There is no need to manually register new examples or modify the build configuration.

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

### Autofilter Implementation

When implementing autofilter functionality, be aware of these key points:

1. **Row Visibility**: Excel doesn't automatically hide filtered rows. You must:
   - Manually hide rows that don't match the filter criteria
   - Apply row hiding BEFORE setting the autofilter and filter conditions
   - Use `worksheet.hideRow()` for individual rows or `worksheet.setDefaultRow()` for bulk hiding

2. **Filter Criteria Application**:
   - Apply the autofilter first (`worksheet.autofilter()`)
   - Then apply filter conditions (`worksheet.filterColumn()`)
   - For multiple conditions, use `filterColumn2()` with appropriate operators (AND/OR)
   - Match the exact filtering logic in your row hiding code

3. **Common Gotchas**:
   - String comparisons must handle null-termination correctly
   - Numeric comparisons should match the exact ranges in filter conditions
   - When using multiple filters, ensure row hiding logic matches ALL filter conditions
   - For blank/non-blank filters, use empty string comparison appropriately

4. **Testing Strategy**:
   - Always compare with reference files both visually and using autocheck
   - Pay special attention to row heights and visibility states
   - Test with various filter combinations to ensure correct row hiding
   - Verify that filter dropdowns show correct state in Excel

Example pattern for implementing filters:
```zig
// 1. Write data
try writeWorksheetData(&worksheet);

// 2. Hide rows that don't match filter criteria
for (data, 0..) |row, i| {
    if (!matchesFilterCriteria(row)) {
        worksheet.hideRow(@intCast(i + 1));
    }
}

// 3. Apply autofilter
try worksheet.autofilter(0, 0, lastRow, lastCol);

// 4. Set filter conditions
try worksheet.filterColumn(0, .{
    .criteria = .equal_to,
    .value_string = "SomeValue",
});
```

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

## Unit Testing

### Unit Test Organization for `zig build test` runs

The project follows a strict test organization pattern:

1. Each API implementation file in `src/` has a corresponding test file prefixed with `test_`. For example:
   - `format.zig` → `test_format.zig`
   - `workbook.zig` → `test_workbook.zig`
   - `worksheet.zig` → `test_worksheet.zig`

2. All test files are located in the `src/` directory alongside their implementation files.

3. Tests are imported and included in the build through `excellent.zig` using a comptime block:
```zig
comptime {
    _ = @import("test_worksheet.zig");
    _ = @import("test_format.zig");
    // ... other test files ...
}
```

### Running Tests
To run the unit tests:

```bash
# Run all tests
zig build test

# Run a specific test file directly *todo*
zig test src/test_worksheet.zig

# Run tests with verbose output
zig build test --verbose
```

### Writing Tests
When writing tests, follow these guidelines:

1. Test files should import their dependencies through `excellent.zig`:
```zig
const std = @import("std");
const excellent = @import("excellent.zig");
const Format = excellent.Format;  // Import what you need
```

2. Each test function should be marked with the `test` keyword and a descriptive name:
```zig
test "format_creation" {
    // Test code here
}
```

3. Always clean up resources in tests using `defer`:
```zig
var workbook = try Workbook.create(std.testing.allocator, "/tmp/test.xlsx");
defer {
    _ = workbook.close() catch {};
    workbook.deinit();
}
```

4. Use the standard testing utilities from `std.testing`:
```zig
try std.testing.expectEqual(expected, actual);
try std.testing.expectError(error.SomeError, function_that_errors());
```

### Test Coverage
The test suite covers:
- API functionality and correctness
- Resource management (memory leaks, proper cleanup)
- Error handling
- Edge cases and invalid inputs

When adding new features, ensure corresponding tests are added to maintain test coverage.
  