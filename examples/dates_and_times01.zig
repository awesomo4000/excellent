const std = @import("std");
const excel = @import("excellent");

pub fn main() !void {
    var gpa = std.heap.GeneralPurposeAllocator(.{}){};
    defer if (gpa.deinit() == .leak) {
        std.debug.panic("leaks detected", .{});
    };
    const allocator = gpa.allocator();

    // A number to display as a date.
    const number = 41333.5;

    // Create a new workbook and add a worksheet.
    var workbook = try excel.Workbook.create(
        allocator,
        "dates_and_times01.xlsx",
    );
    defer workbook.deinit();

    var worksheet = try workbook.addWorksheet(null);

    // Add a format with date formatting.
    var format = try workbook.addFormat();
    _ = try format.setNumFormat("mmm d yyyy hh:mm AM/PM");

    // Widen the first column to make the text clearer.
    worksheet.setColumnWidth(0, 0, 20);

    // Write the number without formatting.
    try worksheet.writeNumber(0, 0, number, null); // 41333.5

    // Write the number with formatting. Note: the worksheet.writeDateTime()
    // function is preferable for writing dates and times. This is for
    // demonstration purposes only.
    try worksheet.writeNumber(1, 0, number, format); // Feb 28 2013 12:00 PM

    try workbook.close();
}
