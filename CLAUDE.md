# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

Excellent is a high-level Zig wrapper for creating Excel (.xlsx) files. It wraps zig-xlsxwriter (which binds to the C library libxlsxwriter) to provide an ergonomic, idiomatic Zig API.

**Key Principle**: The high-level API should be user-friendly and hide the low-level C bindings. Avoid directly importing "xlsxwriter" in examples - that's what the wrapper exists to prevent.

## Build Commands

### Building
```bash
zig build                      # Build everything (library, examples, utilities)
zig build -Dexample=<name>     # Build and run a specific example
zig build examples             # Build all examples
zig build utils                # Build utility programs
zig build test                 # Run unit tests
zig build coverage             # Run tests with kcov coverage analysis
zig build clean                # Clean build artifacts and .xlsx files
```

### Running Examples
```bash
zig build -Dexample=hello      # Build and run the 'hello' example
./zig-out/bin/<example_name>   # Run a built example directly
```

### Development Tools
```bash
./utils/status                 # Show detailed progress on all examples
./utils/status --short         # Compact view of implemented/verified examples
./utils/status <example_name>  # Check status of specific example
```

## Project Structure

### Core Directories
- **src/**: High-level wrapper API implementation
  - `excellent.zig`: Main entry point, re-exports all public APIs
  - `workbook.zig`: Workbook type and operations
  - `worksheet.zig`: Worksheet type with write operations
  - `format.zig`: Cell formatting (colors, borders, alignment, etc.)
  - `chart.zig`: Chart creation and configuration
  - `date_time.zig`: DateTime handling for Excel
  - `data_validation.zig`: Data validation rules
  - `styled.zig`: Rich text styling
  - `conditional_format.zig`: Conditional formatting
  - Test files: `test_*.zig`

- **examples/**: High-level API examples (the main development focus)
  - These correspond to examples in `testing/zig-c-binding-examples/`
  - Output files should match names in `testing/reference-xls/`

- **testing/**: Testing and verification infrastructure
  - `reference-xls/`: Reference .xlsx files for comparison
  - `test-output-xls/`: Verified output from examples
  - `zig-c-binding-examples/`: Low-level binding examples (read-only, for reference)
  - `results/`: Verification results and screenshots

- **utils/**: Development utilities
  - `status`: Shows example implementation/verification progress
  - `autocheck.py`: Automated Excel file validation
  - `verify.py`: Manual verification workflow with screenshots
  - `unverify.py`: Remove verification status after API changes

## Development Workflow

### Creating/Updating Examples

1. **Check what needs work**:
   ```bash
   ./utils/status --short
   ```

2. **Study the reference implementation**:
   - Look at `testing/zig-c-binding-examples/<name>.zig` for low-level calls
   - Examine `testing/reference-xls/<name>.xlsx` for expected output

3. **Implement the high-level example**:
   - Create/update `examples/<name>.zig`
   - Use only the high-level API from `@import("excellent")`
   - Output filename should match reference (no "zig-" prefix)

4. **Test with autocheck (iterate here)**:
   ```bash
   python3 utils/autocheck.py <name> --build --run
   python3 utils/autocheck.py <name> --ignore-styles  # For array formulas, etc.
   ```

5. **Manual verification (when autocheck passes)**:
   ```bash
   ./utils/verify.py <name>
   ```
   - Takes screenshots comparing output to reference
   - Creates `testing/results/<name>/verified` marker file

6. **Check final status**:
   ```bash
   ./utils/status <name>
   ```

### Testing Commands
```bash
zig build test                          # Run all unit tests
python3 utils/autocheck.py --all        # Check all examples
python3 utils/autocheck.py --list-broken # List known broken examples
```

## Architecture Notes

### Memory Management
- `Workbook` owns all child resources (worksheets, formats, charts, chartsheets, conditional formats)
- Always call `workbook.close()` to write the file
- Call `workbook.deinit()` to clean up memory (closes if still open)
- Typical pattern:
  ```zig
  var workbook = try xlsx.Workbook.create(allocator, "output.xlsx");
  defer workbook.deinit();
  // ... use workbook ...
  try workbook.close();
  ```

### Cell References
- Dual API for cell operations: row/column indices OR cell references
- Index-based: `writeString(0, 0, "Hello", null)` (row 0, col 0)
- Cell reference-based: `writeStringCell("A1", "Hello", null)`
- Use `cell.strToRowCol()` to convert cell references to indices

### Format Objects
- Create formats via `workbook.addFormat()`
- Formats are owned by workbook and cleaned up automatically
- Can be reused across multiple cells/operations

### Date/Time Handling
- Use `DateTime` struct for Excel date/time values
- Must use a format with number formatting to display correctly
- Excel stores dates as numbers (days since 1900-01-01)

## Coding Standards

- **Line width**: Keep lines under 80 characters
  - Add trailing commas in function calls/definitions to enable auto-wrapping
  - Use newlines after `=` in assignments for zig formatter to wrap

- **Error handling**: Use Zig error unions, translate C error codes properly

- **Naming conventions**:
  - Examples output files: Match reference names exactly (no "zig-" prefix)
  - Functions: camelCase
  - Types: PascalCase

- **Testing**: Write tests in `test_*.zig` files, included via `comptime` in `excellent.zig`

## Common Issues

### String Handling
- Strings must be null-terminated for C API
- Worksheet allocator handles duplication internally
- Most string issues show up in autocheck.py

### Formula Issues
- Circular references detected by autocheck
- Use `--ignore-styles` flag when internal representation differs but output is correct
- Array formulas may need special handling

### Verification Workflow
- Always run autocheck before manual verification
- If API changes break examples, use `./utils/unverify.py` to reset verification status
- Verification creates markers in `testing/results/<name>/verified`

## Dependencies

- Zig 0.14.0 (required)
- zig-xlsxwriter: Local dependency (../zig-xlsxwriter)
- Python 3: For utility scripts (autocheck, verify, unverify)
- kcov: For coverage analysis (optional)
