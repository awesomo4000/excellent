const std = @import("std");
const excel = @import("excellent");

pub fn main() !void {
    var gpa = std.heap.GeneralPurposeAllocator(.{}){};
    defer if (gpa.deinit() == .leak) {
        std.debug.panic("leaks detected", .{});
    };
    const allocator = gpa.allocator();

    var workbook = try excel.Workbook.create(
        allocator,
        "format_font.xlsx",
    );
    defer workbook.deinit();

    var worksheet = try workbook.addWorksheet(null);

    // Widen the first column to make the text clearer
    worksheet.setColumnWidth(0, 0, 20);

    // Add formats with method chaining
    var bold = try workbook.addFormat(); // this will deinit the format, so we don't need to do it manually
    _ = bold.setBold();

    var italic = try workbook.addFormat();
    _ = italic.setItalic();

    var bold_italic = try workbook.addFormat();
    _ = bold_italic.setBold().setItalic();

    // Write some formatted strings using cell references
    try worksheet.writeStringCell("A1", "This is bold", bold);
    try worksheet.writeStringCell("A2", "This is italic", italic);
    try worksheet.writeStringCell("A3", "Bold and italic", bold_italic);

    try workbook.close();
}
