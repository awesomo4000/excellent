#!/bin/bash

# Check if the correct number of arguments is provided
if [ "$#" -ne 2 ]; then
    echo "Usage: $0 <reference_excel_file> <generated_excel_file>"
    exit 1
fi

# Get absolute paths for both files
REFERENCE_FILE=$(realpath "$1")
GENERATED_FILE=$(realpath "$2")

# Verify that both files exist
if [ ! -f "$REFERENCE_FILE" ]; then
    echo "Error: Reference file not found: $REFERENCE_FILE"
    exit 1
fi

if [ ! -f "$GENERATED_FILE" ]; then
    echo "Error: Generated file not found: $GENERATED_FILE"
    exit 1
fi

# Create a temporary file with proper .xlsx extension
TEMP_BASE=$(mktemp -t excel_XXXXXX)
TEMP_FILE="${TEMP_BASE}.xlsx"
mv "$TEMP_BASE" "$TEMP_FILE"

# Copy the generated file to the temp location
cp "$GENERATED_FILE" "$TEMP_FILE"

# Use osascript to execute AppleScript that positions Excel windows as requested
osascript <<EOF
set displaySize to do shell script "system_profiler SPDisplaysDataType | grep Resolution | head -1"
set screenWidth to word 2 of displaySize
set screenHeight to word 4 of displaySize
set gap to 0
set halfGap to gap / 2
tell application "Microsoft Excel"
    # Activate Excel (bring to front)
    activate
    
    # Get the screen dimensions to calculate window sizes
    
    # Calculate window dimensions - now 1/5 of screen width
    set windowWidth to screenWidth / 2
    set windowHeight to screenHeight * 7 / 8
    
    # Open the reference file and position it at the top left
    open POSIX file "$REFERENCE_FILE"
    set bounds of window 1 to {0, 0, windowWidth-halfGap, windowHeight}
    
    # Open the generated file and position it
    open POSIX file "$TEMP_FILE"
    set bounds of window 1 to {windowWidth + halfGap, 0, windowWidth * 2 + halfGap, windowHeight}
end tell
EOF

# Clean up the temporary file
rm "$TEMP_FILE"
