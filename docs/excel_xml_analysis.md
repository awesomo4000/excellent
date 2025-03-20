# Using XML Lint for Excel File Analysis

Excel files (`.xlsx`) are actually ZIP archives containing XML files. When troubleshooting issues in generated Excel files, examining these XML files can provide detailed insights that aren't visible through the Excel UI.

## Basic Extraction and Analysis

1. First, extract the Excel file's contents:
   ```bash
   mkdir excel_contents
   unzip example.xlsx -d excel_contents
   ```

2. The most important files to examine:
   - `xl/workbook.xml` - Workbook structure
   - `xl/worksheets/sheet1.xml` - Content of the first worksheet
   - `xl/styles.xml` - All styling information
   - `xl/sharedStrings.xml` - Text content

3. Use xmllint to format and examine the files:
   ```bash
   xmllint --format excel_contents/xl/styles.xml > styles_formatted.xml
   ```

## Finding Style Issues

When styles aren't appearing correctly:

1. Examine `styles.xml` to identify custom styles:
   ```bash
   xmllint --xpath "//cellXfs/xf" excel_contents/xl/styles.xml | less
   ```

2. Check if a cell references the correct style ID:
   ```bash
   # Find cells with specific style
   xmllint --xpath "//c[@s='5']" excel_contents/xl/worksheets/sheet1.xml
   ```

3. Compare with a reference file:
   ```bash
   diff -u <(xmllint --format reference/xl/styles.xml) <(xmllint --format broken/xl/styles.xml)
   ```

## Diagnosing Formula Problems

For formula issues:

1. Extract and examine formula definitions:
   ```bash
   xmllint --xpath "//f" excel_contents/xl/worksheets/sheet1.xml
   ```

2. Look for truncated strings or missing null terminators:
   ```bash
   # Check for incomplete formulas
   xmllint --xpath "//f[string-length(text()) < 3]" excel_contents/xl/worksheets/sheet1.xml
   ```

## Common Issues and Solutions

- **Missing styles**: Check if style IDs in cells (`s` attribute) match existing styles in `styles.xml`
- **Truncated strings**: Look for strings ending with `...` in `sharedStrings.xml`
- **Invalid formulas**: Search for malformed XML in the formula text
- **Color issues**: Compare the `<color>` elements within styles to see if RGB values are set correctly

## Automation Example

Here's how you might script a check for specific issues:

```bash
#!/bin/bash
# Extract Excel file
unzip -q "$1" -d temp_excel

# Check for malformed formulas
echo "Checking for malformed formulas..."
xmllint --xpath "//f[contains(text(), ':') and not(contains(text(), 'A1:'))]" temp_excel/xl/worksheets/sheet1.xml 2>/dev/null || echo "No malformed formulas found"

# Check for missing style references
echo "Checking for invalid style references..."
styles_count=$(xmllint --xpath "count(//cellXfs/xf)" temp_excel/xl/styles.xml)
xmllint --xpath "//c[@s > $styles_count]" temp_excel/xl/worksheets/sheet1.xml 2>/dev/null || echo "No invalid style references found"

# Clean up
rm -rf temp_excel
```

This approach helps identify issues that aren't apparent from the Excel interface and can be particularly useful for debugging programmatically generated Excel files. 