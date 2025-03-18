const std = @import("std");
const excel = @import("excellent");
const xlsxwriter = @import("xlsxwriter");

pub fn main() !void {
    var gpa = std.heap.GeneralPurposeAllocator(.{}){};
    defer if (gpa.deinit() == .leak) {
        std.debug.panic("leaks detected", .{});
    };
    const allocator = gpa.allocator();

    var workbook = try excel.Workbook.create(allocator, "styled_text.xlsx");
    defer workbook.deinit();

    var worksheet = try workbook.addWorksheet(null);

    // Set column widths for better readability
    worksheet.setColumnWidth(0, 0, 30);
    worksheet.setColumnWidth(1, 1, 20);

    // Create some styled text with different formats
    const bold_text = try excel.StyledText.bold(workbook, "Bold Text");
    const italic_text = try excel.StyledText.italic(workbook, "Italic Text");
    const red_text = try excel.StyledText.colored(workbook, "Red Text", 0xFF0000);
    const bold_italic = try excel.StyledText.boldItalic(workbook, "Bold and Italic");

    // Create a highlighted format
    var highlight_format = try workbook.addFormat();
    _ = highlight_format.setBgColor(0xFFFF00);
    _ = highlight_format.setPattern(1); // LXW_PATTERN_SOLID
    _ = highlight_format.setFontColor(0x000000);
    _ = highlight_format.setBorder(.thin);

    // Create a writer starting at A1
    var writer = worksheet.writer(0, 0, null);

    // Write a header with mixed formatting
    try writer.printStyled("Welcome to {s}!", .{bold_text});
    writer.nextRow();

    // Write text with multiple styled components
    try writer.printStyled("This is {s} and this is {s} and this is {s}", .{ italic_text, red_text, bold_italic });
    writer.nextRow();

    // Write text with cell highlighting - using writeString directly for cell-level formatting
    try worksheet.writeString(writer.current_row, writer.current_col, "Here is a highlighted cell", highlight_format);
    writer.nextRow();

    // Write with default formatting
    try writer.print("Plain text {s}", .{"here"});
    writer.nextRow();

    // Demonstrate using withFormat for a section
    var bold_writer = writer.withFormat(bold_text.style);
    try bold_writer.print("Everything from here is bold: {s}", .{"including this"});
    writer.nextRow();

    // Back to normal formatting
    try writer.print("Back to normal text", .{});

    try workbook.close();
}
