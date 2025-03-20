const std = @import("std");
const excellent = @import("excellent");
const Workbook = excellent.Workbook;
const Worksheet = excellent.Worksheet;
const TmpFile = excellent.TmpFile;

// To create a vbaProject.bin file:
// 1. Create a new Excel workbook and save it as .xlsm (macro-enabled workbook)
// 2. Open the VBA editor (Alt+F11 on Windows, Option+F11 on Mac)
// 3. Insert a new module (Insert > Module)
// 4. Add the following code:
//    Sub say_hello()
//        MsgBox "Hello from Zig!"
//    End Sub
// 5. Save and close the workbook
// 6. Use the vba_extract.py utility from libxlsxwriter to extract the vbaProject.bin:
//    python vba_extract.py your_workbook.xlsm
// 7. Copy the extracted vbaProject.bin to this directory

// Embed the VBA project binary
const vba_data = @embedFile("vbaProject.bin");

pub fn main() !void {
    // Create a temporary file for the VBA project using the TmpFile API
    var arena = std.heap.ArenaAllocator.init(std.heap.page_allocator);
    defer arena.deinit();
    const allocator = arena.allocator();

    var tmp_file = try TmpFile.create(allocator, "vba_data");
    defer tmp_file.cleanUp();

    // Write the embedded data to the temporary file
    try tmp_file.write(vba_data);

    // Note the xlsm extension of the filename - required for macro-enabled workbooks
    var workbook = try Workbook.create(allocator, "macro.xlsm");
    defer workbook.deinit();

    var worksheet = try workbook.addWorksheet(null);

    // Set column width
    worksheet.setColumnWidth(0, 0, 30);

    // Add the VBA project to the workbook
    try workbook.addVbaProject(tmp_file.path);

    // Write some text
    try worksheet.writeString(2, 0, "Press the button to say hello.", null);

    // Insert a button that calls the macro
    try worksheet.insertButton(2, 1, .{
        .caption = "Press Me",
        .macro = "say_hello", // This must match the name of your VBA subroutine
        .width = 80,
        .height = 30,
    });

    // Close the workbook
    try workbook.close();
}
