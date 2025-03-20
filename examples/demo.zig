const std = @import("std");
const excellent = @import("excellent");
const Workbook = excellent.Workbook;
const Worksheet = excellent.Worksheet;
const Format = excellent.Format;
const TmpFile = excellent.TmpFile;

// Embed the logo image directly into the executable
const logo_data = @embedFile("logo.png");

pub fn main() !void {
    // Create a temporary file for the logo using the TmpFile API
    var arena = std.heap.ArenaAllocator.init(std.heap.page_allocator);
    defer arena.deinit();
    const allocator = arena.allocator();

    var tmp_file = try TmpFile.create(allocator, "logo_");
    defer tmp_file.cleanUp();

    // Write the embedded data to the temporary file
    try tmp_file.write(logo_data);

    // Create a new workbook and add a worksheet
    var workbook = try Workbook.create(allocator, "demo.xlsx");
    defer {
        _ = workbook.close() catch {};
        workbook.deinit();
    }

    var worksheet = try workbook.addWorksheet(null);

    // Add a format and set its bold property
    var format = try workbook.addFormat();
    _ = format.setBold();

    // Change the column width for clarity
    worksheet.setColumnWidth(0, 0, 20);

    // Write some simple text
    try worksheet.writeString(0, 0, "Hello", null);

    // Text with formatting
    try worksheet.writeString(1, 0, "World", format);

    // Write some numbers
    try worksheet.writeNumber(2, 0, 123, null);
    try worksheet.writeNumber(3, 0, 123.456, null);

    // Insert an image using the temporary file path
    try worksheet.insertImage(1, 2, tmp_file.path);
}
