# Excel Macro Example

This example demonstrates how to create an Excel workbook with VBA macros using the Excellent library.

## Prerequisites

1. Excel installed on your system
2. Python installed on your system (for the vba_extract.py utility)
3. libxlsxwriter's vba_extract.py utility (available in the libxlsxwriter examples directory)

## Creating the VBA Project Binary

Before running the example, you need to create a `vbaProject.bin` file that contains the VBA code. Here's how:

1. Create a new Excel workbook
2. Save it as a macro-enabled workbook (.xlsm extension)
3. Open the VBA editor:
   - Windows: Press Alt+F11
   - Mac: Press Option+F11
4. Insert a new module (Insert > Module)
5. Add the following VBA code:
   ```vba
   Sub say_hello()
       MsgBox "Hello from Zig!"
   End Sub
   ```
6. Save and close the workbook
7. Use the vba_extract.py utility to extract the VBA project:
   ```bash
   python vba_extract.py your_workbook.xlsm
   ```
8. Copy the extracted `vbaProject.bin` file to this directory

## Running the Example

1. Make sure you have the `vbaProject.bin` file in the same directory as `macro.zig`
2. Build and run the example:
   ```bash
   zig build run-macro
   ```
3. Open the generated `macro.xlsm` file
4. Click the "Press Me" button to run the macro

## Notes

- The workbook must be saved with the `.xlsm` extension to support macros
- The macro name in the button options must match the name of your VBA subroutine
- Some systems may show security warnings when opening macro-enabled workbooks
- Make sure to enable macros in Excel to use the button functionality 